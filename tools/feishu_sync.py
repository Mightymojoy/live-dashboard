# -*- coding: utf-8 -*-
"""
飞书知识库 → live_data.json 同步脚本

用法:
    python tools/feishu_sync.py --inspect   探明节点类型 / 子表 / 表头 / 样例数据，不写任何文件
    python tools/feishu_sync.py --sync      全量拉取 → 校验闸门 → 写 live_data.json 与 sync_meta.json
    python tools/feishu_sync.py --sync --dry-run
                                            只拉取并比对，不覆盖任何文件

环境变量（GitHub Actions 里配成 Secrets）:
    FEISHU_APP_ID       自建应用 App ID
    FEISHU_APP_SECRET   自建应用 App Secret
    FEISHU_WIKI_TOKEN   wiki 节点 token（默认已内置当前看板对应的节点）

依赖: 仅标准库，无需 pip install。

口径说明（与本地 convert_excel_to_json.py 保持一致，2026-11-02 实测确认）:
    1. 用 valueRenderOption=UnformattedValue 取值。ToString 会返回公式文本
       （如 TEXT(J2,"mm-dd")），FormattedValue 会把百分比变成 "1.10%" 而现有
       数据里 cvr 是小数 0.011，故只有 UnformattedValue 口径一致。
    2. 店铺名来自子表标题，源表本身没有「店铺」列。
    3. 「日期」列为空的行跳过（对应表格末尾的预留空行）。
    4. 空值按字段分别落为 null 或 0，与现有 live_data.json 完全一致。
"""
import argparse
import datetime
import json
import os
import re
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

# 输出字段顺序（与现有 live_data.json 完全一致，前端已依赖此结构）
TARGET_FIELDS = [
    "id", "yearMonth", "date", "platformActivity", "liveActivity", "week",
    "session", "timeSlot", "supervisor", "anchor", "startTime", "durationMin",
    "exposureUsers", "uv", "newFans", "buyers", "sales", "refund", "refund2",
    "adCost", "roiPre", "roiPost", "gpm", "uvValue", "cvr", "peakViewers",
    "avgViewers", "avgWatchMin",
]

# 内置映射：飞书表头（已去换行/空格）→ 输出字段。2026-11-02 由 --inspect 实测确认。
DEFAULT_FIELD_MAP = {
    "yearMonth": "年份月份",
    "date": "日期",
    "platformActivity": "平台活动",
    "liveActivity": "直播间活动",
    "week": "周",
    "session": "场次",
    "timeSlot": "早中晚",
    "supervisor": "场控",
    "anchor": "主播",
    "startTime": "开播时间",
    "durationMin": "直播时长分钟",
    "exposureUsers": "曝光人数",
    "uv": "UV观看人数",
    "newFans": "新增粉丝数",
    "buyers": "成交人数",
    "sales": "销售额",
    "refund": "退款金额（绩效统计/实时更新）",
    "refund2": "退款金额（下播数据/不更新）",
    "adCost": "广告消耗",
    "roiPre": "退前ROI",
    "roiPost": "退后ROI",
    "gpm": "GPM",
    "uvValue": "UV价值",
    "cvr": "观看成交率(人数)",
    "peakViewers": "最高在线人数",
    "avgViewers": "平均在线人数",
    "avgWatchMin": "人均观看时长(分钟)",
}

# 子表顺序（也是 id 前缀与输出顺序的依据）
DEFAULT_SHEETS = ["摩登新贵女", "轻熟质享客", "云端商务家", "视频号"]

NUMERIC_FIELDS = {
    "session", "durationMin", "exposureUsers", "uv", "newFans", "buyers", "sales",
    "refund", "refund2", "adCost", "roiPre", "roiPost", "gpm", "uvValue", "cvr",
    "peakViewers", "avgViewers", "avgWatchMin",
}
# 空值时的落盘值：0 表示「业务上按零计」，None 表示「无此数据」
ZERO_WHEN_EMPTY = {"session", "durationMin", "sales", "refund", "refund2", "adCost"}
DEFAULT_WHEN_EMPTY = {"session": 1.0}

EMPTY_TOKENS = {"", "-", "--", "/", ".", "N/A", "n/a", "NA", "null", "None", "nan",
                "NaN", "#DIV/0!", "#VALUE!", "#N/A", "#REF!", "#NAME?"}

EXCEL_EPOCH = datetime.datetime(1899, 12, 30)
CHUNK_ROWS = 800
WS_RE = re.compile(r"[\s\u3000]+")
BEIJING_TZ = datetime.timezone(datetime.timedelta(hours=8))


def now_beijing():
    """统一用北京时间。

    GitHub Actions 的 ubuntu runner 是 UTC，直接用 datetime.now() 会让
    sync_meta.json 的 syncedAt 比实际早 8 小时（实测首跑写入 06:12，
    实际为北京 14:12），看板上会显示错误时间。
    """
    return datetime.datetime.now(BEIJING_TZ)


def log(msg):
    print(msg, flush=True)


def http_json(method, path, token=None, payload=None, params=None):
    """仅用标准库发请求，避免依赖 requests（Actions 侧也无需 pip install）。"""
    url = path if path.startswith("http") else BASE + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Content-Type", "application/json; charset=utf-8")
    if token:
        req.add_header("Authorization", "Bearer " + token)
    try:
        with urllib.request.urlopen(req, timeout=120) as resp:
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
    c = data.get("code")
    if c == 99991672:
        log("       缺少权限：请在开发者后台开通 sheets:spreadsheet:readonly 并发布版本。")
        app_id = os.environ.get("FEISHU_APP_ID", "").strip()
        if app_id:
            log("       授权链接：https://open.feishu.cn/app/%s/auth"
                "?q=sheets:spreadsheet:readonly&op_from=openapi&token_type=tenant" % app_id)
    elif c in (1254302, 1310249):
        log("       该电子表格未授权给本应用：请在表格「···」→ 文档协作者里添加应用。")
    sys.exit(code)


def norm_header(s):
    """表头去换行/空格：'直播时长\\n分钟' → '直播时长分钟'。"""
    return WS_RE.sub("", str(s if s is not None else ""))


# ---------------- 电子表格读取 ----------------

def list_sheets(token, st):
    data = api_get(token, "/sheets/v3/spreadsheets/%s/sheets/query" % st)
    if data.get("code") != 0:
        fail(5, "读取子表列表", data)
    return data["data"].get("sheets", [])


def read_range(token, st, rng):
    """UnformattedValue：取原始值而非公式文本，口径与现有 JSON 一致。"""
    path = "/sheets/v2/spreadsheets/%s/values/%s" % (st, rng)
    data = api_get(token, path, {"valueRenderOption": "UnformattedValue"})
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


# ---------------- 值归一 ----------------

def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        t = v.strip()
        return "" if t in EMPTY_TOKENS else t
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


def as_text(v):
    """文本字段统一转字符串。

    飞书把「平台活动=618」「主播=67」这类纯数字单元格返回为数值型，
    若直接透传会得到 618 而非 "618"，与既有 JSON 的字符串口径不一致。
    """
    t = cell_text(v)
    if t is None or t == "":
        return None
    if isinstance(t, bool):
        return str(t)
    if isinstance(t, float):
        return str(int(t)) if t.is_integer() else str(t)
    if isinstance(t, int):
        return str(t)
    return str(t)


def to_num(v, default=None):
    if v is None:
        return default
    if isinstance(v, bool):
        return float(v)
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).replace("\u3000", "").replace(",", "").replace("%", "")
    for sym in ("\uffe5", "\xa5", "$", " "):
        t = t.replace(sym, "")
    t = t.strip()
    if t in EMPTY_TOKENS:
        return default
    try:
        return float(t)
    except ValueError:
        return default


def serial_to_dt(serial):
    """Excel 序列号 → datetime。45110.4705555556 → 2023-07-03 11:17:36

    注意：必须先把秒数四舍五入再构造 timedelta。直接用 timedelta(days=float)
    会因二进制浮点误差把 11:17:36 算成 11:17:35（实测 3937 条记录全部差 1 秒）。
    """
    total_seconds = round(float(serial) * 86400)
    return EXCEL_EPOCH + datetime.timedelta(seconds=total_seconds)


def date_pair(v):
    """返回 (date 'MM-DD', yearMonth 'YYYY-MM')。支持序列号与文本。"""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        dt = serial_to_dt(v)
        return dt.strftime("%m-%d"), dt.strftime("%Y-%m")
    t = str(cell_text(v) or "").strip()
    if not t:
        return "", ""
    head = t.split(" ")[0].replace("/", "-").replace("年", "-").replace("月", "-").replace("日", "")
    parts = [p for p in head.split("-") if p]
    try:
        if len(parts) == 3:
            y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
            return "%02d-%02d" % (m, d), "%04d-%02d" % (y, m)
        if len(parts) == 2:
            m, d = int(parts[0]), int(parts[1])
            return "%02d-%02d" % (m, d), ""
    except ValueError:
        pass
    return "", ""


def start_time_str(v, fallback_date, year_month):
    """开播时间：数值走序列号还原时分秒，文本原样保留。

    源表里有少量历史脏值（如 '2025/10/9/15:00:00'），既有 JSON 就是原样存下来的。
    这里不再做格式改写，避免同步后与本地转换结果产生无谓差异。
    """
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if v and v > 20000:
            return serial_to_dt(v).strftime("%Y-%m-%d %H:%M:%S")
    t = str(cell_text(v) or "").strip()
    if t and not t.startswith("#"):
        return t
    ym = year_month or "2026-01"
    d = fallback_date or "01-01"
    return "%s-%s 00:00:00" % (ym[:4] if ym else "2026", d)


# ---------------- 子表 → 记录 ----------------

def load_config():
    cfg = {}
    if os.path.exists(MAP_PATH):
        try:
            with open(MAP_PATH, encoding="utf-8") as f:
                cfg = json.load(f) or {}
        except ValueError:
            log("[警告] feishu_field_map.json 不是合法 JSON，改用内置映射")
            cfg = {}
    field_map = dict(DEFAULT_FIELD_MAP)
    for k, v in (cfg.get("fields") or {}).items():
        if str(v).strip():
            field_map[k] = norm_header(v)
    sheets = cfg.get("sheets") or DEFAULT_SHEETS
    return field_map, [str(s) for s in sheets]


def build_records(rows, field_map, shop, header_row=1):
    if len(rows) < header_row:
        log("[中止] 子表「%s」读不到表头" % shop)
        sys.exit(10)
    header = [norm_header(c) for c in rows[header_row - 1]]

    idx = {}
    for target, src in field_map.items():
        if src not in header:
            log("[中止] 子表「%s」表头里找不到列「%s」（映射至 %s）" % (shop, src, target))
            log("       实际表头：%s" % " | ".join([h for h in header if h]))
            sys.exit(11)
        idx[target] = header.index(src)

    def cell(row, target):
        i = idx.get(target)
        if i is None or i >= len(row):
            return None
        return row[i]

    records, skipped = [], 0
    for seq, row in enumerate(rows[header_row:], start=1):
        raw_date = cell(row, "date")
        # 「日期」为空 → 表格末尾预留空行，直接跳过（与本地转换脚本一致）
        if raw_date is None or str(cell_text(raw_date)).strip() == "":
            skipped += 1
            continue
        d, ym_from_date = date_pair(raw_date)
        if not d:
            skipped += 1
            continue
        ym_raw = cell(row, "yearMonth")
        _, ym_from_a = date_pair(ym_raw) if ym_raw is not None else ("", "")
        ym = ym_from_a or ym_from_date

        rec = {
            "id": "%s-%04d" % (shop[:2], seq),
            "yearMonth": ym,
            "date": d,
            "startTime": start_time_str(cell(row, "startTime"), d, ym),
        }
        for target in TARGET_FIELDS:
            if target in rec:
                continue
            v = cell(row, target)
            if target in NUMERIC_FIELDS:
                if v is None or str(cell_text(v)).strip() == "":
                    rec[target] = DEFAULT_WHEN_EMPTY.get(target, 0.0 if target in ZERO_WHEN_EMPTY else None)
                else:
                    rec[target] = to_num(v, DEFAULT_WHEN_EMPTY.get(target, 0.0 if target in ZERO_WHEN_EMPTY else None))
            else:
                rec[target] = as_text(v)
        records.append(rec)
    return records, skipped


# ---------------- 命令 ----------------

def cmd_inspect(token, obj_type, obj_token, title):
    log("=" * 64)
    log("节点标题：%s" % title)
    log("对象类型：%s" % obj_type)
    log("对象 token：%s" % obj_token)
    log("=" * 64)

    if obj_type != "sheet":
        log("[提示] 当前节点类型为 %s，本脚本只支持电子表格（sheet）。" % obj_type)
        return

    field_map, want_sheets = load_config()
    sheets = list_sheets(token, obj_token)
    names = [s.get("title") for s in sheets]
    log("子表总数：%d → %s" % (len(sheets), "、".join(names)))
    missing = [w for w in want_sheets if w not in names]
    if missing:
        log("[注意] 配置里的子表在飞书侧不存在：%s" % "、".join(missing))

    for s in sheets:
        gp = s.get("grid_properties") or {}
        cols = int(gp.get("column_count") or 1)
        rows = int(gp.get("row_count") or 0)
        header = (read_range(token, obj_token, "%s!A1:%s1" % (s["sheet_id"], col_letter(cols))) or [[]])[0]
        h = [norm_header(c) for c in header]
        log("\n子表：%s  sheet_id=%s  行=%d 列=%d" % (s.get("title"), s.get("sheet_id"), rows, cols))
        absent = [t for t, src in field_map.items() if src not in h]
        log("  必需列缺失：%s" % ("无" if not absent else "、".join(absent)))
        sample = read_range(token, obj_token, "%s!A1:%s3" % (s["sheet_id"], col_letter(cols)))
        if len(sample) >= 2:
            pairs = []
            for t in ("date", "yearMonth", "anchor", "sales", "cvr", "startTime"):
                if t in field_map and field_map[t] in h:
                    i = h.index(field_map[t])
                    val = sample[1][i] if i < len(sample[1]) else None
                    pairs.append("%s=%r" % (t, val))
            log("  样例(第2行)：%s" % "  ".join(pairs))
    log("\n下一步：python tools/feishu_sync.py --sync --dry-run 先比对条数与截止日期。")


def summarize(result):
    total = sum(len(v) for v in result.values())
    cutoff = ""
    for recs in result.values():
        for r in recs:
            c = "%s-%s" % (r.get("yearMonth") or "", r.get("date") or "")
            if c > cutoff:
                cutoff = c
    day = cutoff.split("-")[-1] if "-" in cutoff else ""
    cutoff = "%s-%s" % (cutoff[:7], day) if day else cutoff[:7]
    return total, cutoff


def validate(prev, result):
    errors = []
    if not prev:
        return errors
    prev_total = sum(len(v) for v in prev.values())
    new_total = sum(len(v) for v in result.values())
    if prev_total and new_total < prev_total * 0.9:
        errors.append("总条数骤降：%d → %d（低于上次的 90%%）" % (prev_total, new_total))
    _, prev_cut = summarize(prev)
    _, new_cut = summarize(result)
    if prev_cut and new_cut and new_cut[:7] < prev_cut[:7]:
        errors.append("数据截止月份倒退：%s → %s" % (prev_cut, new_cut))
    missing = [s for s in prev.keys() if s not in result]
    if missing:
        errors.append("店铺缺失：%s" % "、".join(missing))
    return errors


def cmd_sync(token, obj_type, obj_token, dry_run=False):
    if obj_type != "sheet":
        log("[中止] 当前节点类型 %s，本脚本只支持电子表格。" % obj_type)
        sys.exit(15)

    field_map, want_sheets = load_config()
    sheets = {s.get("title"): s for s in list_sheets(token, obj_token)}

    log("[1/3] 拉取数据…")
    result, detail = {}, []
    for shop in want_sheets:
        s = sheets.get(shop)
        if not s:
            log("      [跳过] 飞书侧没有子表「%s」" % shop)
            continue
        gp = s.get("grid_properties") or {}
        rows = read_sheet(token, obj_token, s["sheet_id"],
                          int(gp.get("row_count") or 0), int(gp.get("column_count") or 1))
        recs, skipped = build_records(rows, field_map, shop)
        result[shop] = recs
        detail.append("      %s: %d 条（跳过空行 %d）" % (shop, len(recs), skipped))
    for line in detail:
        log(line)

    new_total, new_cut = summarize(result)
    log("      合计 %d 条 / 截止 %s / %d 个店铺" % (new_total, new_cut, len(result)))

    log("[2/3] 校验闸门…")
    prev = {}
    if os.path.exists(DATA_PATH):
        try:
            with open(DATA_PATH, encoding="utf-8") as fp:
                prev = json.load(fp)
        except ValueError:
            prev = {}
    errors = validate(prev, result)
    if errors:
        for e in errors:
            log("      [不通过] " + e)
        log("[中止] 校验未通过，不写任何文件，线上数据保持原样。")
        sys.exit(17)
    prev_total = sum(len(v) for v in prev.values()) if prev else 0
    log("      通过（上次 %d 条 → 本次 %d 条）" % (prev_total, new_total))

    if dry_run:
        log("[dry-run] 不写文件。若要落盘，去掉 --dry-run 重跑。")
        return

    log("[3/3] 写入文件…")
    tmp = DATA_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fp:
        json.dump(result, fp, ensure_ascii=False, indent=2)
    os.replace(tmp, DATA_PATH)

    meta = {
        "status": "ok",
        "source": "飞书知识库",
        "dataCutoff": new_cut,
        "recordCount": new_total,
        "syncedAt": now_beijing().strftime("%Y-%m-%d %H:%M"),
        "updatedBy": os.environ.get("SYNC_ACTOR", "本地脚本"),
        "note": "页面运行时读取本文件，用于显示真实的数据截止时间与同步状态。",
    }
    with open(META_PATH, "w", encoding="utf-8") as fp:
        json.dump(meta, fp, ensure_ascii=False, indent=2)
    log("      完成：live_data.json %d 条，sync_meta.json 已更新" % new_total)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--inspect", action="store_true")
    ap.add_argument("--sync", action="store_true")
    ap.add_argument("--dry-run", action="store_true", help="只比对，不覆盖文件")
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
        cmd_sync(token, obj_type, obj_token, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
