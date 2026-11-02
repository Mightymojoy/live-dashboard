# -*- coding: utf-8 -*-
"""
飞书知识库 → live_data.json 同步脚本

用法:
    python tools/feishu_sync.py --inspect   探明 wiki 节点类型 / 表字段 / 样例数据，不写任何文件
    python tools/feishu_sync.py --sync      全量拉取 → 校验闸门 → 写 live_data.json 与 sync_meta.json

环境变量（GitHub Actions 里配成 Secrets）:
    FEISHU_APP_ID       自建应用 App ID
    FEISHU_APP_SECRET   自建应用 App Secret
    FEISHU_WIKI_TOKEN   wiki 节点 token（默认已内置当前看板对应的节点）

注意:
    本脚本的字段映射来自 tools/feishu_field_map.json。
    首次使用请先跑 --inspect，按真实列名补全映射后再跑 --sync。
    未跑过 --inspect 的情况下，--sync 会拒绝写入，避免用错列名污染线上数据。
"""
import argparse
import datetime
import json
import os
import sys

try:
    import requests
except ImportError:
    print("[错误] 缺少 requests 依赖，请先执行 pip install requests")
    sys.exit(1)

BASE = "https://open.feishu.cn/open-apis"
DEFAULT_WIKI_TOKEN = "J8sdwGyjYiyeNikTGTuczPJOnkf"
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
MAP_PATH = os.path.join(HERE, "feishu_field_map.json")
DATA_PATH = os.path.join(ROOT, "live_data.json")
META_PATH = os.path.join(ROOT, "sync_meta.json")

# 看板消费的 28 个字段（顺序与字段名必须与现有 live_data.json 一致）
TARGET_FIELDS = [
    "id", "yearMonth", "date", "platformActivity", "liveActivity", "week",
    "session", "timeSlot", "supervisor", "anchor", "startTime", "durationMin",
    "exposureUsers", "uv", "newFans", "buyers", "sales", "refund", "refund2",
    "adCost", "roiPre", "roiPost", "gpm", "uvValue", "cvr", "peakViewers",
    "avgViewers", "avgWatchMin",
]
NUMERIC_FIELDS = {
    "durationMin", "exposureUsers", "uv", "newFans", "buyers", "sales", "refund",
    "refund2", "adCost", "roiPre", "roiPost", "gpm", "uvValue", "cvr",
    "peakViewers", "avgViewers", "avgWatchMin",
}


def log(msg):
    print(msg, flush=True)


def get_env(name, required=True):
    v = os.environ.get(name, "").strip()
    if required and not v:
        log("[错误] 缺少环境变量 %s" % name)
        sys.exit(2)
    return v


def get_token(app_id, app_secret):
    r = requests.post(BASE + "/auth/v3/tenant_access_token/internal",
                      json={"app_id": app_id, "app_secret": app_secret}, timeout=30)
    data = r.json()
    if data.get("code") != 0:
        log("[错误] 获取 tenant_access_token 失败: %s" % data)
        sys.exit(3)
    return data["tenant_access_token"]


def api_get(token, path, params=None):
    r = requests.get(BASE + path, headers={"Authorization": "Bearer " + token},
                     params=params or {}, timeout=60)
    return r.json()


def api_post(token, path, payload, params=None):
    r = requests.post(BASE + path, headers={"Authorization": "Bearer " + token},
                      json=payload, params=params or {}, timeout=60)
    return r.json()


def resolve_node(token, wiki_token):
    data = api_get(token, "/wiki/v2/spaces/get_node", {"token": wiki_token, "obj_type": "wiki"})
    if data.get("code") != 0:
        log("[错误] 解析 wiki 节点失败: %s" % data)
        log("       常见原因：应用未开通 wiki:wiki:readonly，或未被加入该知识库空间")
        sys.exit(4)
    node = data["data"]["node"]
    return node.get("obj_type"), node.get("obj_token"), node.get("title")


def list_bitable_tables(token, app_token):
    tables, page_token = [], None
    while True:
        params = {"page_size": 100}
        if page_token:
            params["page_token"] = page_token
        data = api_get(token, "/bitable/v1/apps/%s/tables" % app_token, params)
        if data.get("code") != 0:
            log("[错误] 读取多维表格数据表失败: %s" % data)
            sys.exit(5)
        items = data["data"].get("items", [])
        tables.extend(items)
        if not data["data"].get("has_more"):
            break
        page_token = data["data"].get("page_token")
    return tables


def list_fields(token, app_token, table_id):
    data = api_get(token, "/bitable/v1/apps/%s/tables/%s/fields" % (app_token, table_id),
                   {"page_size": 200})
    if data.get("code") != 0:
        log("[错误] 读取字段定义失败: %s" % data)
        sys.exit(6)
    return data["data"].get("items", [])


def fetch_records(token, app_token, table_id):
    out, page_token = [], None
    while True:
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        data = api_post(token, "/bitable/v1/apps/%s/tables/%s/records/search" % (app_token, table_id),
                        {"automatic_fields": False}, params)
        if data.get("code") != 0:
            log("[错误] 读取记录失败: %s" % data)
            sys.exit(7)
        out.extend(data["data"].get("items", []))
        if not data["data"].get("has_more"):
            break
        page_token = data["data"].get("page_token")
    return out


def cell_text(v):
    """把飞书单元格值归一成字符串/数字。"""
    if v is None:
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, list):
        parts = []
        for item in v:
            if isinstance(item, dict):
                parts.append(str(item.get("text") or item.get("name") or item.get("file_name") or ""))
            else:
                parts.append(str(item))
        return "".join(parts).strip()
    if isinstance(v, dict):
        for k in ("text", "name", "value", "link"):
            if v.get(k):
                return str(v[k]).strip()
        return ""
    return str(v).strip()


def to_num(v, default=0.0):
    t = str(cell_text(v)).replace(",", "").replace("%", "").replace("¥", "").replace("$", "").strip()
    if t in ("", "-", "--", "/", "N/A", "n/a", "NA", "null", "None", "nan"):
        return default
    try:
        return float(t)
    except ValueError:
        return default


def to_date_pair(v):
    """返回 (MM-DD, YYYY-MM)。支持毫秒时间戳与常见字符串。"""
    raw = v
    if isinstance(v, (int, float)):
        dt = datetime.datetime.fromtimestamp(v / 1000)
        return dt.strftime("%m-%d"), dt.strftime("%Y-%m")
    t = str(cell_text(v))
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S", "%m-%d"):
        try:
            dt = datetime.datetime.strptime(t[:len(fmt) + 2].strip(), fmt)
            if fmt == "%m-%d":
                dt = dt.replace(year=datetime.date.today().year)
            return dt.strftime("%m-%d"), dt.strftime("%Y-%m")
        except ValueError:
            continue
    return "", ""


def load_map():
    if not os.path.exists(MAP_PATH):
        log("[错误] 缺少字段映射文件 tools/feishu_field_map.json")
        sys.exit(8)
    with open(MAP_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    mapping = cfg.get("fields", {})
    missing = [k for k in TARGET_FIELDS if not mapping.get(k)]
    return cfg, mapping, missing


def cmd_inspect(token, obj_type, obj_token, title):
    log("=" * 62)
    log("wiki 节点：%s" % title)
    log("对象类型：%s" % obj_type)
    log("对象 token：%s" % obj_token)
    log("=" * 62)
    if obj_type != "bitable":
        log("[提示] 该节点不是多维表格（bitable），当前脚本只支持多维表格直读。")
        log("       若是电子表格，需要改用 sheets 接口；若是文档，只能导出后解析。")
        return
    tables = list_bitable_tables(token, obj_token)
    for t in tables:
        log("\n数据表：%s（table_id=%s）" % (t.get("name"), t.get("table_id")))
        fields = list_fields(token, obj_token, t["table_id"])
        for f in fields:
            log("    - %-24s type=%s" % (f.get("field_name"), f.get("type")))
        records = fetch_records(token, obj_token, t["table_id"])[:3]
        log("    样例记录（前 3 条）：")
        for r in records:
            log("      " + json.dumps(r.get("fields", {}), ensure_ascii=False)[:300])
    log("\n下一步：把上面的列名填进 tools/feishu_field_map.json 的 fields 映射，再执行 --sync")


def validate(prev_shops, new_shops, new_total, new_cutoff):
    errors = []
    prev_total = sum(len(v) for v in prev_shops.values()) if prev_shops else 0
    if prev_total > 0 and new_total < prev_total * 0.9:
        errors.append("总条数骤降：%d → %d（低于上次的 90%%）" % (prev_total, new_total))
    prev_cut = ""
    for recs in prev_shops.values():
        for r in recs:
            d = str(r.get("yearMonth") or "") + "-" + str(r.get("date") or "")
            if d > prev_cut:
                prev_cut = d
    if prev_cut and new_cutoff and new_cutoff < prev_cut[:7]:
        errors.append("数据截止日期倒退：%s → %s" % (prev_cut[:7], new_cutoff))
    missing_shops = [s for s in prev_shops.keys() if s not in new_shops]
    if missing_shops:
        errors.append("店铺缺失：%s" % "、".join(missing_shops))
    return errors


def cmd_sync(token, obj_type, obj_token):
    if obj_type != "bitable":
        log("[中止] 当前节点类型为 %s，只支持 bitable 直读。" % obj_type)
        sys.exit(9)
    cfg, mapping, missing = load_map()
    if missing:
        log("[中止] 字段映射不完整，缺少以下目标字段的飞书列名：")
        for m in missing:
            log("    - %s" % m)
        log("请先执行 --inspect 查看真实列名，补全 tools/feishu_field_map.json")
        sys.exit(10)

    table_id = cfg["table_id"]
    shop_field = cfg["shopField"]
    log("[1/4] 拉取记录…")
    records = fetch_records(token, obj_token, table_id)
    log("      共 %d 条原始记录" % len(records))

    log("[2/4] 字段映射与清洗…")
    result, skipped = {}, 0
    for item in records:
        f = item.get("fields", {})
        shop = str(cell_text(f.get(shop_field, ""))).strip()
        if not shop:
            skipped += 1
            continue
        row = {}
        for target, src in mapping.items():
            v = f.get(src)
            if target in NUMERIC_FIELDS:
                row[target] = to_num(v)
            elif target in ("date", "yearMonth"):
                continue
            else:
                row[target] = cell_text(v)
        d, ym = to_date_pair(f.get(mapping["date"])) if mapping.get("date") else ("", "")
        row["date"], row["yearMonth"] = d, ym
        if not row.get("id"):
            row["id"] = item.get("record_id", "")
        result.setdefault(shop, []).append(row)
    log("      有效 %d 条，缺少店铺字段被跳过 %d 条" % (sum(len(v) for v in result.values()), skipped))

    log("[3/4] 校验闸门…")
    prev = {}
    if os.path.exists(DATA_PATH):
        with open(DATA_PATH, encoding="utf-8") as fp:
            prev = json.load(fp)
    new_total = sum(len(v) for v in result.values())
    cutoffs = sorted({r.get("yearMonth", "")[:7] for v in result.values() for r in v if r.get("yearMonth")})
    new_cutoff = cutoffs[-1] if cutoffs else ""
    errors = validate(prev, result, new_total, new_cutoff)
    if errors:
        for e in errors:
            log("      [不通过] " + e)
        log("[中止] 校验未通过，不写入任何文件，线上数据保持原样。")
        sys.exit(11)
    log("      通过：%d 条 / 截止 %s / %d 个店铺" % (new_total, new_cutoff, len(result)))

    log("[4/4] 写入文件…")
    tmp = DATA_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fp:
        json.dump(result, fp, ensure_ascii=False, separators=(",", ":"))
    os.replace(tmp, DATA_PATH)
    meta = {
        "status": "ok",
        "source": "飞书知识库",
        "dataCutoff": new_cutoff,
        "recordCount": new_total,
        "syncedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "updatedBy": "GitHub Actions",
        "note": "页面运行时读取本文件，用于显示真实的数据截止时间与同步状态。",
    }
    with open(META_PATH, "w", encoding="utf-8") as fp:
        json.dump(meta, fp, ensure_ascii=False, indent=2)
    log("      完成：live_data.json %d 条，sync_meta.json 已更新" % new_total)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--inspect", action="store_true")
    ap.add_argument("--sync", action="store_true")
    args = ap.parse_args()
    if not args.inspect and not args.sync:
        ap.print_help()
        return

    app_id = get_env("FEISHU_APP_ID")
    app_secret = get_env("FEISHU_APP_SECRET")
    wiki_token = get_env("FEISHU_WIKI_TOKEN", required=False) or DEFAULT_WIKI_TOKEN

    token = get_token(app_id, app_secret)
    obj_type, obj_token, title = resolve_node(token, wiki_token)

    if args.inspect:
        cmd_inspect(token, obj_type, obj_token, title)
    if args.sync:
        cmd_sync(token, obj_type, obj_token)


if __name__ == "__main__":
    main()
