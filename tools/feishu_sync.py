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
import time
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
ERROR_PATH = os.path.join(ROOT, "sync_error.json")
PROBE_PATH = os.path.join(ROOT, "sync_probe.json")

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
CHUNK_ROWS = 400

# 飞书「数据未就绪」类错误：单次请求计算量过大或前置操作未完成。
# 90235 = Data not ready（电子表格读值实测命中）
# 1254607 = Data not ready（多维表格文档中同义错误码）
# 1255040 = Request timed out（服务端计算超时）
# -1：本地连接超时/网络错误（读值时飞书可能长时间不回包）
TRANSIENT_CODES = {90235, 1254607, 1255040, -1}

# 本次运行里有多少数据块是靠 FormattedValue 才取到的（口径与原来不同，需比对）
FALLBACK = {"count": 0, "blocks": []}
# 读值单次超时。遇 90235 时飞书要算到超时才回包（实测吃满 120 秒），
# 压到 45 秒能把"失败链"从十分钟级降到分钟级；正常读 400 行 x 27 列仅 1-3 秒。
READ_TIMEOUT = 45
PROBE_TIMEOUT = 25


class NotReady(Exception):
    """拆到最小粒度仍报 Data not ready：数据区域在飞书侧未就绪，需定点诊断。"""

    def __init__(self, rng, data):
        self.rng = rng
        self.data = data if isinstance(data, dict) else {"raw": str(data)}
        Exception.__init__(self, "读取 %s 时飞书返回 code=%s" % (rng, self.data.get("code")))
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


def http_json(method, path, token=None, payload=None, params=None, timeout=120):
    """仅用标准库发请求，避免依赖 requests（Actions 侧也无需 pip install）。

    timeout 可按调用点收紧：读值时若遇到 90235，飞书会先算到超时才回包
    （实测吃满 120 秒），把超时压低能显著缩短失败链耗时。
    """
    url = path if path.startswith("http") else BASE + path
    if params:
        url += "?" + urllib.parse.urlencode(params)
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Content-Type", "application/json; charset=utf-8")
    if token:
        req.add_header("Authorization", "Bearer " + token)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "ignore")
        try:
            return json.loads(body)
        except ValueError:
            return {"code": e.code, "msg": body[:200]}
    except urllib.error.URLError as e:
        return {"code": -1, "msg": "网络错误: %s" % e.reason}
    except (TimeoutError, OSError) as e:
        return {"code": -1, "msg": "连接超时: %s" % e}


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


def api_get(token, path, params=None, timeout=120):
    return http_json("GET", path, token=token, params=params, timeout=timeout)


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


def write_error(what, data, diag=None):
    """把失败原因落盘，让页面直接可见，不必翻 Actions 日志。

    两条通道：
    1) sync_error.json —— 完整错误（含飞书 code/msg + 定点诊断明细），供排查；
    2) sync_meta.json 的 status/error —— 页面状态条直接渲染。
    工作流用 if: always() 提交，失败也会被推到 dist。
    """
    try:
        stamp = now_beijing().strftime("%Y-%m-%d %H:%M")
        code = data.get("code") if isinstance(data, dict) else None
        payload = {
            "status": "fail",
            "failedAt": stamp,
            "stage": what,
            "feishuCode": code,
            "feishu": data,
            "updatedBy": os.environ.get("SYNC_ACTOR", "本地脚本"),
        }
        if diag:
            payload["diagnosis"] = list(diag)
        with open(ERROR_PATH, "w", encoding="utf-8") as fp:
            json.dump(payload, fp, ensure_ascii=False, indent=2)

        meta = {}
        if os.path.exists(META_PATH):
            try:
                with open(META_PATH, encoding="utf-8") as fp:
                    meta = json.load(fp)
            except ValueError:
                meta = {}
        meta["status"] = "fail"
        meta["error"] = "%s 失败（飞书 code=%s）" % (what, code)
        if diag:
            meta["errorHint"] = sum(1 for x in diag if " → OK" in x)
            meta["errorDetail"] = "已定位明细见 sync_error.json 的 diagnosis"
        meta["failedAt"] = stamp
        meta["updatedBy"] = payload["updatedBy"]
        with open(META_PATH, "w", encoding="utf-8") as fp:
            json.dump(meta, fp, ensure_ascii=False, indent=2)
    except Exception as exc:  # 落盘失败不能掩盖原始错误
        log("[警告] 失败原因落盘时出错: %s" % exc)


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
    # Actions 日志下载需要认证，而 check-run annotation 匿名可读 ——
    # 把错误详情打成 annotation，下次失败可远程读到，不必让人翻页面截图。
    try:
        brief = json.dumps(data, ensure_ascii=False)
    except Exception:
        brief = str(data)
    print("::error title=飞书接口失败::%s | 返回 %s" % (what, brief[:800]), flush=True)
    write_error(what, data)
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


def read_range(token, st, rng, trace="", allow_retry=True):
    """UnformattedValue：取原始值而非公式文本，口径与现有 JSON 一致。

    失败后按错误类型分别处置：
    - 90235 / 1254607 / 1255040（Data not ready、内部计算超时）：属于
      「单次请求太重或前置操作未完成」，官方文档明确建议降低请求量后重试。
      这里先整块等 3 秒重试一次，仍失败则把等待时间往上翻倍，累计 4 次。
      （2026-11-09 起因源表从 67 列涨到 91 列，A1:CM800 一次要 7.8 万个格子，
       连续 5 次同步全部卡在这里。）
    - 其他错误：直接判定为真错误退出，不掩盖权限/结构问题。
    """
    path = "/sheets/v2/spreadsheets/%s/values/%s" % (st, rng)
    params = {"valueRenderOption": "UnformattedValue"}
    waits = (3, 6, 12, 20) if allow_retry else ()
    data = api_get(token, path, params)
    for i, wait in enumerate(waits):
        if data.get("code") == 0 or data.get("code") not in TRANSIENT_CODES:
            break
        log("      [重试] %s 返回 code=%s（数据未就绪），%d 秒后第 %d 次重试"
            % (rng, data.get("code"), wait, i + 1))
        time.sleep(wait)
        data = api_get(token, path, params)
    if data.get("code") != 0:
        what = "读取单元格范围 %s" % rng
        if trace:
            what += "（%s）" % trace
        fail(6, what, data)
    vr = (data.get("data") or {}).get("valueRange") or {}
    return vr.get("values") or []


def read_block(token, st, sheet_id, c1, c2, r1, r2, trace, depth=0):
    """读一个矩形块：UnformattedValue 最多重试 3 次，仍不行才对该块降级。

    2026-11-09 三次实测把结论钉死了：
    1) 同一请求两次结果互换（第 100 行一次 OK 一次 90235）⇒ 失败是散发的、
       瞬时的，不是体量问题 ⇒ 拆块无效，已停用；
    2) 一轮 40 个块里只有 14 个需要降级 ⇒ UnformattedValue 大多数时候是好的
       ⇒ 不该一失败就整轮换模式；
    3) 整轮换 FormattedValue 后，销售额被显示格式取整（5706.29→5706.0，
       1156 组里 293 组有差）⇒ FormattedValue 不能作为常规口径。

    因此：主路径坚持 UnformattedValue + 重试；降级只作用于个别块，且会被
    记录下来，由上层决定是否允许落盘（有降级就只写探测报告，不覆盖数据）。
    """
    rng = "%s!%s%d:%s%d" % (sheet_id, col_letter(c1), r1, col_letter(c2), r2)
    path = "/sheets/v2/spreadsheets/%s/values/%s" % (st, rng)
    params = {"valueRenderOption": "UnformattedValue"}

    data = {}
    for attempt, wait in enumerate((0, 3, 8)):
        if wait:
            time.sleep(wait)
        data = api_get(token, path, params, timeout=READ_TIMEOUT)
        if data.get("code") == 0:
            return (data.get("data") or {}).get("valueRange", {}).get("values") or []
        if data.get("code") not in TRANSIENT_CODES:
            what = "读取单元格范围 %s" % rng
            if trace:
                what += "（%s）" % trace
            fail(6, what, data)
        log("      [重试 %d/3] %s 返回 code=%s" % (attempt + 1, rng, data.get("code")))

    # 该块三次都失败 → 只对这一块降级，并记下来（数值可能被显示格式取整）
    alt = api_get(token, path, {"valueRenderOption": "FormattedValue"},
                  timeout=READ_TIMEOUT)
    if alt.get("code") == 0:
        vals = (alt.get("data") or {}).get("valueRange", {}).get("values") or []
        FALLBACK["count"] += 1
        if len(FALLBACK["blocks"]) < 40:
            FALLBACK["blocks"].append(rng)
        log("      [降级] %s 该块改用 FormattedValue（%d 行），数值可能被显示格式取整"
            % (rng, len(vals)))
        return vals

    raise NotReady(rng, data)


def read_sheet(token, st, sheet_id, row_count, col_count, title="", need_headers=None):
    """只读「表头里真正需要的那几列」，而不是整张表的全部列。

    源表右侧可能被追加了大量辅助列（2026-11-09 实测从 67 列变 91 列，
    多出的 24 列我们一个都不用），全量读会把单次请求撑爆。因此：
    1) 先读第 1 行表头（只 1 行，代价极小），按表头名定位所需列的列号；
    2) 数据区只取「最小列号..最大列号」这一条竖带，行按 CHUNK_ROWS 分块；
    3) 返回时把收窄后的表头拼回第一行，build_records 仍按表头名定位列。
    """
    last_col = max(int(col_count or 1), 1)
    trace = "子表「%s」sheet_id=%s grid=%s行×%s列" % (title, sheet_id, row_count, last_col)

    head_block = read_block(token, st, sheet_id, 1, last_col, 1, 1, trace + " 表头")
    head_cells = list(head_block[0]) if head_block else []
    header = [norm_header(c) for c in head_cells]

    lo, hi = 1, last_col
    if need_headers:
        wanted = [norm_header(h) for h in need_headers]
        pos = [header.index(h) + 1 for h in wanted if h in header]
        missing = [h for h in wanted if h not in header]
        if missing:
            log("      [提示] 子表「%s」表头缺 %d 个映射列：%s"
                % (title, len(missing), " / ".join(missing)))
        if pos:
            lo, hi = min(pos), min(max(pos), last_col)
        # 源表被加了列时把列名报出来（annotation 匿名可读），
        # 便于直接判断「是新增辅助列」还是「关键列被改名/删除」。
        extra = [h for h in header if h and h not in wanted]
        if extra:
            log("      [提示] 子表「%s」另有 %d 个未使用列：%s"
                % (title, len(extra), " / ".join(extra[:15]) + (" ..." if len(extra) > 15 else "")))
            print("::notice title=源表存在未使用列::子表「%s」共 %d 列，其中 %d 列未被映射使用：%s"
                  % (title, last_col, len(extra), " / ".join(extra[:30])), flush=True)
    if hi - lo + 1 < last_col:
        log("      %s：grid %d 列 → 实取 %s..%s 共 %d 列"
            % (title, last_col, col_letter(lo), col_letter(hi), hi - lo + 1))

    rows, start = [], 2
    while start <= row_count:
        end = min(start + CHUNK_ROWS - 1, row_count)
        block = read_block(token, st, sheet_id, lo, hi, start, end, trace)
        if not block:
            break
        rows.extend(block)
        start = end + 1

    head_slice = head_cells[lo - 1:hi] if head_cells else []
    return [head_slice] + rows


def diagnose(token, st, sheet_id, title, last_col, row_count, sheets, want_sheets,
             tried_rng="", budget_sec=150):
    """定点探查 90235：是整表未就绪、某几行坏了，还是取值模式的问题。

    重要约束（2026-11-09 实测）：飞书返回 90235 之前要先把本次计算跑完，
    单次耗时约 30 秒才超时。所以探针必须「少而准」，且带总时长预算，
    否则诊断本身会把作业顶过 timeout。探针按信息量从高到低排列，
    预算用尽即停，结论仍然成立（前面几条已能定性）。
    """
    lines = []
    nc = max(int(last_col or 1), 1)
    last = col_letter(nc)
    started = time.time()

    def probe(label, rng, render="UnformattedValue"):
        if time.time() - started > budget_sec:
            return "timeout"
        params = {"valueRenderOption": render} if render else None
        try:
            d = api_get(token, "/sheets/v2/spreadsheets/%s/values/%s" % (st, rng),
                        params, timeout=PROBE_TIMEOUT)
        except Exception as exc:
            lines.append("%s → 请求异常 %s" % (label, exc))
            return "error"
        code = d.get("code") or 0
        if code == 0:
            vals = (d.get("data") or {}).get("valueRange", {}).get("values") or []
            lines.append("%s → OK（%d 行）" % (label, len(vals)))
            return 0
        lines.append("%s → code=%s %s" % (label, code, d.get("msg", "")))
        return code

    # 按信息量排序：先做「同形状 × 不同取值模式」的横向对照，这是能一次性
    # 区分「模式问题」与「随机问题」的关键；再探备用通道；最后看是否全库一致。
    # ① 基准：同宽度只读 1 行
    probe("1 行 A2:%s2" % last, "%s!A2:%s2" % (sheet_id, last))
    # ② 同形状 5 行，四种取值模式横评（tried_rng 用实际失败块，形状更贴近现场）
    five = "%s!A2:%s6" % (sheet_id, last)
    probe("5 行 UnformattedValue", five, "UnformattedValue")
    probe("5 行 FormattedValue", five, "FormattedValue")
    probe("5 行 ToString", five, "ToString")
    probe("5 行 不传 render", five, None)
    # ③ 同一请求原样重复一次：两次结果不同 = 随机，可直接判死「体量说」
    probe("5 行 UnformattedValue 重复", five, "UnformattedValue")
    # ④ 备用通道：把整表导出成 xlsx（绕开读值接口）。仅探「能否建任务」，
    #    不下载、不落盘；返回 99991672 说明缺导出权限，需在开发者后台开。
    if time.time() - started <= budget_sec:
        try:
            d = http_json("POST", "/drive/v1/export_tasks", token=token,
                          payload={"file_extension": "xlsx", "token": st, "type": "sheet"},
                          timeout=PROBE_TIMEOUT)
            lines.append("导出接口 xlsx → code=%s %s" % (d.get("code") or 0, d.get("msg", "")))
        except Exception as exc:
            lines.append("导出接口 xlsx → 异常 %s" % exc)
    # ⑤ 其他子表 5 行：判断是全库一致还是只有本表
    for other in want_sheets:
        s = sheets.get(other)
        if not s or s.get("sheet_id") == sheet_id:
            continue
        gp = s.get("grid_properties") or {}
        wide = col_letter(max(int(gp.get("column_count") or 1), 1))
        probe("另一子表「%s」5 行" % other, "%s!A2:%s6" % (s["sheet_id"], wide))
    # ⑥ 失败块本体：确认当前这次到底是不是也失败
    if tried_rng:
        probe("失败块 UnformattedValue", tried_rng, "UnformattedValue")
    if time.time() - started > budget_sec:
        lines.append("（诊断已用 %ds，达到预算上限，后续探针跳过）" % int(time.time() - started))
    return lines


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


def build_full_date(rec):
    """把 (yearMonth, date) 还原成 YYYY-MM-DD。

    yearMonth 形如 2026-11，date 形如 11-03，年份取 yearMonth 前 4 位、月日取 date。
    注意不能直接字符串拼接后截取：'2026-11'+'-'+'12-30' = '2026-11-12-30'，
    截 [:7] 会得到 '2026-11'，月份信息被 date 的月份覆盖（曾把 12-30 显示成 11-30）。
    """
    ym = str(rec.get("yearMonth") or "")
    dt = str(rec.get("date") or "")
    if not dt:
        return ""
    if len(dt) >= 5 and dt[2] == "-":
        year = ym[:4] if len(ym) >= 4 and ym[:4].isdigit() else ""
        return ("%s-%s" % (year, dt)) if year else dt
    if len(dt) == 10:
        return dt
    return ""


def summarize(result, today=None):
    """返回 (总条数, 数据截止日, 未来日期记录数)。

    截止日只统计「不晚于今天」的记录：源表里常有人预填/复制出错成未来日期
    （实测 2026-12-30 那种脏行），若计入会让看板显示"数据截止 12-30"，
    与实际已发生的直播场次不符。这类记录本身保留，只是不参与截止日计算。
    """
    if today is None:
        today = now_beijing().date()
    total = sum(len(v) for v in result.values())
    cutoff = ""
    future = 0
    for recs in result.values():
        for r in recs:
            full = build_full_date(r)
            if not full:
                continue
            try:
                d = datetime.date.fromisoformat(full)
            except ValueError:
                continue
            if d > today:
                future += 1
                continue
            if full > cutoff:
                cutoff = full
    return total, cutoff, future


def validate(prev, result):
    errors = []
    if not prev:
        return errors
    prev_total = sum(len(v) for v in prev.values())
    new_total = sum(len(v) for v in result.values())
    if prev_total and new_total < prev_total * 0.9:
        errors.append("总条数骤降：%d → %d（低于上次的 90%%）" % (prev_total, new_total))
    _, prev_cut, _ = summarize(prev)
    _, new_cut, _ = summarize(result)
    if prev_cut and new_cut and new_cut[:7] < prev_cut[:7]:
        errors.append("数据截止月份倒退：%s → %s" % (prev_cut, new_cut))
    missing = [s for s in prev.keys() if s not in result]
    if missing:
        errors.append("店铺缺失：%s" % "、".join(missing))
    return errors


AGG_FIELDS = ("sales", "uv", "buyers", "exposureUsers")


def aggregate(records):
    """按 (店铺, 日期) 汇总关键数值，作为两条取数路径的比对指纹。

    用汇总而不是逐行比对，是因为行序/空行会让 id 漂移；而按「店铺+日期」
    汇总后，任何取数口径变化（百分比被渲染、金额被四舍五入）都会在合计上
    体现出来，且与行序无关。
    """
    out = {}
    for shop, recs in records.items():
        for r in recs:
            k = (shop, str(r.get("date") or ""))
            a = out.setdefault(k, {"n": 0})
            a["n"] += 1
            for f in AGG_FIELDS:
                v = r.get(f)
                a[f] = a.get(f, 0.0) + (float(v) if isinstance(v, (int, float)) else 0.0)
    return out


def compare_records(prev, cur, tol_rel=1e-9):
    """比对两条取数路径的结果，返回差异摘要（用于决定能否切换口径）。"""
    a, b = aggregate(prev), aggregate(cur)
    keys = sorted(set(a) & set(b))
    diffs = []
    for k in keys:
        if a[k]["n"] != b[k]["n"]:
            diffs.append({"shop": k[0], "date": k[1], "field": "count",
                          "prev": a[k]["n"], "cur": b[k]["n"]})
        for f in AGG_FIELDS:
            x, y = a[k].get(f, 0.0), b[k].get(f, 0.0)
            if abs(x - y) > max(1e-9, abs(x) * tol_rel):
                diffs.append({"shop": k[0], "date": k[1], "field": f,
                              "prev": round(x, 6), "cur": round(y, 6),
                              "delta": round(y - x, 6)})
    return {
        "comparedGroups": len(keys),
        "onlyInPrev": len(set(a) - set(b)),
        "onlyInCur": len(set(b) - set(a)),
        "diffCount": len(diffs),
        "sample": diffs[:30],
    }


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
        rc, cc = int(gp.get("row_count") or 0), int(gp.get("column_count") or 1)
        try:
            rows = read_sheet(token, obj_token, s["sheet_id"], rc, cc,
                              title=shop, need_headers=list(field_map.values()))
        except NotReady as exc:
            log("      [诊断] 子表「%s」拆到最小粒度仍报 Data not ready，开始定点探查…" % shop)
            lines = diagnose(token, obj_token, s["sheet_id"], shop, cc, rc,
                             sheets, want_sheets, exc.rng)
            for ln in lines:
                log("             " + ln)
            print("::error title=90235 定位::%s" % "%0A".join(lines)[:2200], flush=True)
            write_error("读取单元格范围 %s（子表「%s」）" % (exc.rng, shop), exc.data, lines)
            sys.exit(6)
        recs, skipped = build_records(rows, field_map, shop)
        result[shop] = recs
        detail.append("      %s: %d 条（跳过空行 %d）" % (shop, len(recs), skipped))
    for line in detail:
        log(line)

    new_total, new_cut, new_future = summarize(result)
    log("      合计 %d 条 / 截止 %s / %d 个店铺" % (new_total, new_cut, len(result)))
    if new_future:
        log("      [提示] 有 %d 条记录日期晚于今天，已排除在截止日之外（可能是源表预填或录入错误）"
            % new_future)

    log("[2/3] 校验闸门…")
    prev = {}
    if os.path.exists(DATA_PATH):
        try:
            with open(DATA_PATH, encoding="utf-8") as fp:
                prev = json.load(fp)
        except ValueError:
            prev = {}

    # 本次有数据块是靠 FormattedValue 才取到的 —— 口径与历史数据不同，
    # 先做一致性比对、只写探测报告，绝不覆盖 live_data.json。
    # 只有比对通过（差异 0 或可解释）才值得切换取数口径。
    if FALLBACK["count"]:
        cmp = compare_records(prev, result)
        probe = {
            "mode": "formatted-fallback",
            "at": now_beijing().strftime("%Y-%m-%d %H:%M"),
            "fallbackBlocks": FALLBACK["count"],
            "fallbackSamples": FALLBACK["blocks"][:12],
            "recordCount": new_total,
            "perStore": {k: len(v) for k, v in result.items()},
            "vsPrev": cmp,
            "perStorePrev": {k: len(v) for k, v in (prev or {}).items()},
            "note": "本次未覆盖 live_data.json。vsPrev.diffCount=0 才说明 "
                    "FormattedValue 与 UnformattedValue 口径一致，可以切换。",
        }
        with open(PROBE_PATH, "w", encoding="utf-8") as fp:
            json.dump(probe, fp, ensure_ascii=False, indent=2)
        log("      [探测] %d 个块降级为 FormattedValue" % FALLBACK["count"])
        log("      [探测] 与上次数据比对：共同分组 %d，差异 %d 项（仅上次有 %d / 仅本次有 %d）"
            % (cmp["comparedGroups"], cmp["diffCount"], cmp["onlyInPrev"], cmp["onlyInCur"]))
        for s in cmp["sample"][:5]:
            log("             %s %s %s: %s → %s" % (s["shop"], s["date"], s["field"],
                                                    s.get("prev"), s.get("cur")))
        print("::notice title=FormattedValue 一致性比对::共同分组 %d，差异 %d 项，"
              "仅上次有 %d 组、仅本次有 %d 组；明细见 sync_probe.json"
              % (cmp["comparedGroups"], cmp["diffCount"], cmp["onlyInPrev"], cmp["onlyInCur"]),
              flush=True)
        log("[探测结束] 未写入 live_data.json。比对结果确认后再切换口径。")
        return

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
        "futureRecords": new_future,
        "note": "页面运行时读取本文件，用于显示真实的数据截止时间与同步状态。",
    }
    if new_future:
        meta["warning"] = "源表有 %d 条记录日期晚于今天，未计入截止日，请检查是否为预填或录入错误" % new_future
    with open(META_PATH, "w", encoding="utf-8") as fp:
        json.dump(meta, fp, ensure_ascii=False, indent=2)
    # 成功即清掉上一次的失败记录与探测报告，避免页面/仓库残留旧结论
    for stale in (ERROR_PATH, PROBE_PATH):
        if os.path.exists(stale):
            try:
                os.remove(stale)
            except OSError:
                pass
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
