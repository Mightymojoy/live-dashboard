# -*- coding: utf-8 -*-
"""
飞书知识库 → live_data.json 同步脚本

用法:
    python tools/feishu_sync.py --inspect   探明节点类型 / 子表 / 表头 / 样例数据，不写任何文件
    python tools/feishu_sync.py --sync      全量拉取 → 校验闸门 → 写 live_data.json 与 sync_meta.json

环境变量（GitHub Actions 里配成 Secrets）:
    FEISHU_APP_ID       自建应用 App ID
    FEISHU_APP_SECRET   自建应用 App Secret
    FEISHU_WIKI_TOKEN   wiki 节点 token（默认已内置当前看板对应的节点）

依赖: 仅标准库，无需 pip install。

说明:
    字段映射在 tools/feishu_field_map.json。首次使用必须先跑 --inspect，
    按真实表头补全映射后再跑 --sync；映射不完整时脚本拒绝写入。
"""
import argparse
import datetime
import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request

BASE = "https://open.feishu.cn/open-apis"
DEFAULT_WIKI_TOKEN = "J8sdwGyjYiyeNikTGTuczPJOnkf"
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
MAP_PATH = os.path.join(HERE, "feishu_field_map.json")
DATA_PATH = os.path.join(ROOT, "live_data.json")
META_PATH = os.path.join(ROOT, "sync_meta.json")

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
CHUNK_ROWS = 800


def log(msg):
    print(msg, flush=True)


def http_json(method, path, token=None, payload=None, params=None):
    """仅用标准库发请求，避免依赖 requests（本机 venv 可能缺包）。"""
    url = path if path.startswith("http") else BASE + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Content-Type", "application/json; charset=utf-8")
    if token:
        req.add_header("Authorization", "Bearer " + token)
    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        try:
            return json.loads(body)
        except ValueError:
            return {"code": e.code, "msg": body[:200]}
    except urllib.error.URLError as e:
        return {"code": -1, "msg": "网络错误: %s" % e.reason}


def get_env(name, required=True):
    v = os.environ.get(name, "").strip()
    if required and not v:
        log("[错误] 缺少环境变量 %s" % name)
        sys.exit(2)
    return v


def get_token(app_id, app_secret):
    data = http_json("POST", "/auth/v3/tenant_access_token/internal",
                     payload={"app_id": app_id, "app_secret": app_secret})
    if data.get("code") != 0:
        log("[错误] 获取 tenant_access_token 失败: %s" % data)
        sys.exit(3)
    return data["tenant_access_token"]


def api_get(token, path, params=None):
    return http_json("GET", path, token=token, params=params)


def api_post(token, path, payload, params=None):
    return http_json("POST", path, token=token, payload=payload, params=params)


def resolve_node(token, wiki_token):
    data = api_get(token, "/wiki/v2/spaces/get_node", {"token": wiki_token, "obj_type": "wiki"})
    if data.get("code") != 0:
        log("[错误] 解析 wiki 节点失败: %s" % data)
        log("       原因多为：应用未开通 wiki:wiki:readonly，或未被加入该知识库空间")
        sys.exit(4)
    node = data["data"]["node"]
    return node.get("obj_type"), node.get("obj_token"), node.get("title")


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def fail(code, what, data):
    log("[错误] %s 失败: %s" % (what, data))
    if data.get("code") == 99991672:
        log("       缺少权限：请在开发者后台开通 sheets:spreadsheet:readonly 后重试。")
    sys.exit(code)


# ---------------- 电子表格 ----------------

def list_sheets(token, st):
    data = api_get(token, "/sheets/v3/spreadsheets/%s/sheets/query" % st)
    if data.get("code") != 0:
        fail(5, "读取子表列表", data)
    return data["data"].get("sheets", [])


def read_range(token, st, rng):
    path = "/sheets/v2/spreadsheets/%s/values/%s" % (st, rng)
    data = api_get(token, path, {"valueRenderOption": "ToString"})
    if data.get("code") != 0:
        fail(6, "读取单元格范围 %s" % rng, data)
    vr = (data.get("data") or {}).get("valueRange") or {}
    return vr.get("values") or []


def read_sheet(token, st, sheet_id, row_count, col_count):
    rows, start = [], 1
    last_col = col_letter(max(col_count, 1))
    while start <= row_count:
        end = min(start + CHUNK_ROWS - 1, row_count)
        block = read_range(token, st, "%s!A%d:%s%d" % (sheet_id, start, last_col, end))
        if not block:
            break
        rows.extend(block)
        start = end + 1
    return rows


# ---------------- 多维表格 ----------------

def list_bitable_tables(token, app_token):
    tables, page_token = [], None
    while True:
        params = {"page_size": 100}
        if page_token:
            params["page_token"] = page_token
        data = api_get(token, "/bitable/v1/apps/%s/tables" % app_token, params)
        if data.get("code") != 0:
            fail(7, "读取数据表列表", data)
        tables.extend(data["data"].get("items", []))
        if not data["data"].get("has_more"):
            break
        page_token = data["data"].get("page_token")
    return tables


def fetch_bitable_records(token, app_token, table_id):
    out, page_token = [], None
    while True:
        params = {"page_size": 500}
        if page_token:
            params["page_token"] = page_token
        data = api_post(token, "/bitable/v1/apps/%s/tables/%s/records/search" % (app_token, table_id),
                        {"automatic_fields": False}, params)
        if data.get("code") != 0:
            fail(8, "读取记录", data)
        out.extend(data["data"].get("items", []))
        if not data["data"].get("has_more"):
            break
        page_token = data["data"].get("page_token")
    return out


# ---------------- 值归一 ----------------

def cell_text(v):
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
                parts.append(str(item.get("text") or item.get("name") or ""))
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
    if isinstance(v, (int, float)):
        try:
            dt = datetime.datetime.fromtimestamp(v / 1000)
            return dt.strftime("%m-%d"), dt.strftime("%Y-%m")
        except (ValueError, OSError):
            return "", ""
    t = str(cell_text(v)).strip()
    if not t:
        return "", ""
    head = t.split(" ")[0].replace("/", "-")
    parts = head.split("-")
    try:
        if len(parts) == 3:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
        elif len(parts) == 2:
            m, d = int(parts[0]), int(parts[1])
            y = datetime.date.today().year
        else:
            return "", ""
        return "%02d-%02d" % (m, d), "%04d-%02d" % (y, m)
    except ValueError:
        return "", ""


# ---------------- 命令 ----------------

def cmd_inspect(token, obj_type, obj_token, title):
    log("=" * 64)
    log("节点标题：%s" % title)
    log("对象类型：%s" % obj_type)
    log("对象 token：%s" % obj_token)
    log("=" * 64)

    if obj_type == "sheet":
        sheets = list_sheets(token, obj_token)
        for s in sheets:
            gp = s.get("grid_properties") or {}
            log("\n子表：%s" % s.get("title"))
            log("  sheet_id=%s  行=%s  列=%s" % (s.get("sheet_id"), gp.get("row_count"), gp.get("column_count")))
            cols = int(gp.get("column_count") or 1)
            rows = read_range(token, obj_token, "%s!A1:%s6" % (s.get("sheet_id"), col_letter(cols)))
            for i, row in enumerate(rows):
                cells = ["" if c is None else str(c) for c in row]
                log("  第%d行: %s" % (i + 1, " | ".join(cells)[:420]))
        log("\n下一步：把上面的表头填进 tools/feishu_field_map.json，再执行 --sync")
        return

    if obj_type == "bitable":
        for t in list_bitable_tables(token, obj_token):
            log("\n数据表：%s（table_id=%s）" % (t.get("name"), t.get("table_id")))
            recs = fetch_bitable_records(token, obj_token, t["table_id"])[:3]
            for r in recs:
                log("  样例: " + json.dumps(r.get("fields", {}), ensure_ascii=False)[:400])
        log("\n下一步：把上面的列名填进 tools/feishu_field_map.json，再执行 --sync")
        return

    log("[提示] 节点类型 %s 暂不支持自动同步，只能导出后人工处理。" % obj_type)


def load_map():
    if not os.path.exists(MAP_PATH):
        log("[错误] 缺少字段映射文件 tools/feishu_field_map.json")
        sys.exit(9)
    with open(MAP_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    mapping = cfg.get("fields", {})
    missing = [k for k in TARGET_FIELDS if not str(mapping.get(k, "")).strip()]
    return cfg, mapping, missing


def validate(prev, result, new_total, new_cutoff):
    errors = []
    prev_total = sum(len(v) for v in prev.values()) if prev else 0
    if prev_total > 0 and new_total < prev_total * 0.9:
        errors.append("总条数骤降：%d → %d（低于上次的 90%%）" % (prev_total, new_total))
    prev_cut = ""
    for recs in prev.values():
        for r in recs:
            c = str(r.get("yearMonth") or "") + "-" + str(r.get("date") or "")
            if c > prev_cut:
                prev_cut = c
    if prev_cut and new_cutoff and new_cutoff < prev_cut[:7]:
        errors.append("数据截止月份倒退：%s → %s" % (prev_cut[:7], new_cutoff))
    missing_shops = [s for s in prev.keys() if s not in result]
    if missing_shops:
        errors.append("店铺缺失：%s" % "、".join(missing_shops))
    return errors


def build_from_rows(rows, cfg, mapping):
    header_row = int(cfg.get("headerRow", 1))
    if len(rows) < header_row:
        log("[中止] 子表行数不足，读不到表头")
        sys.exit(10)
    header = [str(c).strip() if c is not None else "" for c in rows[header_row - 1]]
    idx = {}
    for target, src in mapping.items():
        if not str(src).strip():
            continue
        if src not in header:
            log("[中止] 表头里找不到列「%s」（映射自 %s）" % (src, target))
            log("       实际表头：%s" % " | ".join([h for h in header if h]))
            sys.exit(11)
        idx[target] = header.index(src)

    shop_field = cfg.get("shopField", "店铺")
    if shop_field not in header:
        log("[中止] 表头里找不到店铺列「%s」" % shop_field)
        sys.exit(12)
    shop_idx = header.index(shop_field)

    result, skipped = {}, 0
    for row in rows[header_row:]:
        def get(target):
            i = idx.get(target)
            return row[i] if i is not None and i < len(row) else None

        shop = str(cell_text(row[shop_idx] if shop_idx < len(row) else "")).strip()
        if not shop:
            skipped += 1
            continue
        rec = {}
        for target in TARGET_FIELDS:
            if target in ("date", "yearMonth"):
                continue
            if target in idx:
                v = get(target)
                rec[target] = to_num(v) if target in NUMERIC_FIELDS else cell_text(v)
            else:
                rec[target] = 0.0 if target in NUMERIC_FIELDS else ""
        d, ym = to_date_pair(get("date")) if "date" in idx else ("", "")
        rec["date"], rec["yearMonth"] = d, ym
        result.setdefault(shop, []).append(rec)
    return result, skipped


def cmd_sync(token, obj_type, obj_token):
    cfg, mapping, missing = load_map()
    if missing:
        log("[中止] 字段映射不完整，以下目标字段还没有对应飞书列名：")
        for m in missing:
            log("    - %s" % m)
        log("请先执行 --inspect 查看真实表头，补全 tools/feishu_field_map.json")
        sys.exit(13)

    log("[1/3] 拉取数据…")
    if obj_type == "sheet":
        sheets = list_sheets(token, obj_token)
        want_id = cfg.get("sheetId") or ""
        want_title = cfg.get("sheetTitle") or ""
        target = None
        for s in sheets:
            if want_id and s.get("sheet_id") == want_id:
                target = s
                break
            if want_title and s.get("title") == want_title:
                target = s
                break
        if not target:
            log("[中止] 找不到指定子表（sheetId=%s / sheetTitle=%s）" % (want_id, want_title))
            log("       可用子表：%s" % "、".join([s.get("title", "") for s in sheets]))
            sys.exit(14)
        gp = target.get("grid_properties") or {}
        rows = read_sheet(token, obj_token, target["sheet_id"],
                          int(gp.get("row_count") or 0), int(gp.get("column_count") or 1))
        log("      子表「%s」共 %d 行" % (target.get("title"), len(rows)))
        result, skipped = build_from_rows(rows, cfg, mapping)
    elif obj_type == "bitable":
        log("[中止] 当前节点是多维表格，但映射配置按电子表格模式编写，请核对 feishu_field_map.json")
        sys.exit(15)
    else:
        log("[中止] 不支持的节点类型：%s" % obj_type)
        sys.exit(16)

    new_total = sum(len(v) for v in result.values())
    log("      有效 %d 条，缺少店铺字段跳过 %d 条" % (new_total, skipped))

    log("[2/3] 校验闸门…")
    prev = {}
    if os.path.exists(DATA_PATH):
        try:
            with open(DATA_PATH, encoding="utf-8") as fp:
                prev = json.load(fp)
        except ValueError:
            prev = {}
    cutoffs = sorted({r.get("yearMonth", "")[:7] for v in result.values() for r in v if r.get("yearMonth")})
    new_cutoff = cutoffs[-1] if cutoffs else ""
    errors = validate(prev, result, new_total, new_cutoff)
    if errors:
        for e in errors:
            log("      [不通过] " + e)
        log("[中止] 校验未通过，不写任何文件，线上数据保持原样。")
        sys.exit(17)
    log("      通过：%d 条 / 截止 %s / %d 个店铺" % (new_total, new_cutoff, len(result)))

    log("[3/3] 写入文件…")
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
