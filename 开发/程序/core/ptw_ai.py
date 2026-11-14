# -*- coding: utf-8 -*-
"""内置 AI（B2）：让平台自己写口径归类 / 表2 结论 / OAR 变动原因。

设计原则（沿用 MDMQ 的 AI 通道）：
  1. **平台不联网也能用**：默认关闭，用户填 Key 或选免费额度后才启用。
  2. **只产文字与分类，不产数字**：金额、比例、合计一律由平台公式计算，
     AI 输出的文字里若出现数字，视为违规由调用方丢弃（见 `has_number`）。
  3. **结果落在与原「外部 AI 旁路」完全相同的文件**（`*_result.json`），
     所以内置 AI 与外置 AI 可以互相替换、也可以都不跑（保留人工填写）。
  4. 失败要可诊断：错误分类（网络/鉴权/限流/超长/格式），不是一句「失败」。

三条通道与既有任务卡一一对应：
  exp-oar-bucket   口径归类   → oar_map_ai_result.json
  exp-table2-concl 表2 结论   → table2_concl_result.json
  exp-oar-reason   OAR 变动原因 → oar_concl_result.json
"""
import io
import json
import os
import re
import time
import urllib.error
import urllib.request

from . import paths

PROVIDERS = {
    "zhipu": {
        "label": "智谱 GLM-4.7-Flash（免费额度，需注册 Key）",
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "model": "glm-4.7-flash",
        "extra": {"thinking": {"type": "disabled"}},
        "fix": True,          # base_url / model 由程序锁定
    },
    "custom": {
        "label": "自定义 OpenAI 兼容（本地 Ollama / 公司网关 / 其他厂商）",
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o-mini",
        "extra": {},
        "fix": False,
    },
}

SETTINGS = os.path.join(paths.DATA_DIR, "settings.json")
TASK_DIR = os.path.join(paths.DATA_DIR, "bkd_work", "ai_tasks")
MAX_CHARS = 12000          # 单次请求的正文上限（约 6k tokens）
TIMEOUT = 60


class AIError(Exception):
    def __init__(self, kind, msg, detail=""):
        super().__init__(msg)
        self.kind = kind
        self.msg = msg
        self.detail = detail or ""


# ---------------------------------------------------------------- 配置
def load_settings():
    try:
        with io.open(SETTINGS, encoding="utf-8") as f:
            return json.load(f) or {}
    except Exception:
        return {}


def save_settings(patch):
    s = load_settings()
    s.update({k: v for k, v in (patch or {}).items() if k.startswith("ai_")})
    os.makedirs(os.path.dirname(SETTINGS), exist_ok=True)
    with io.open(SETTINGS, "w", encoding="utf-8") as f:
        json.dump(s, f, ensure_ascii=False, indent=2)
    return status()


def cfg():
    s = load_settings()
    prov = (s.get("ai_provider") or "zhipu").strip()
    p = PROVIDERS.get(prov) or PROVIDERS["zhipu"]
    return {
        "enabled": bool(s.get("ai_enabled")),
        "provider": prov if prov in PROVIDERS else "zhipu",
        "label": p["label"],
        "key": (s.get("ai_key") or "").strip(),
        "base_url": (p["base_url"] if p["fix"]
                     else (s.get("ai_base_url") or p["base_url"]).strip().rstrip("/")),
        "model": (p["model"] if p["fix"] else (s.get("ai_model") or p["model"]).strip()),
        "extra": p.get("extra") or {},
        "timeout": int(s.get("ai_timeout") or TIMEOUT),
    }


def status():
    c = cfg()
    return {"providers": [{"key": k, "label": v["label"], "fix": v["fix"],
                           "base_url": v["base_url"], "model": v["model"]}
                          for k, v in PROVIDERS.items()],
            "enabled": c["enabled"], "provider": c["provider"],
            "model": c["model"], "base_url": c["base_url"],
            "has_key": bool(c["key"]), "tasks": list_tasks()}


# ---------------------------------------------------------------- 网络
def _post(url, payload, key, timeout, extra=None):
    body = dict(payload)
    if extra:
        body.update(extra)
    req = urllib.request.Request(
        url, data=json.dumps(body, ensure_ascii=False).encode("utf-8"), method="POST",
        headers={"Content-Type": "application/json", "Authorization": "Bearer " + key,
                 "User-Agent": "PTW/%s" % _ver()})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8", "replace"))
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            detail = e.read().decode("utf-8", "replace")[:400]
        except Exception:
            pass
        if e.code in (401, 403):
            raise AIError("auth", "鉴权失败（Key 无效或无权限）", detail)
        if e.code == 429:
            raise AIError("rate", "触发限流或额度用尽", detail)
        raise AIError("http", "HTTP %d" % e.code, detail)
    except urllib.error.URLError as e:
        raise AIError("net", "连不上服务（%s）" % (e.reason,), "")
    except Exception as e:
        raise AIError("net", "请求失败：%s" % e, "")


def _ver():
    try:
        from app_version import APP_VERSION
        return APP_VERSION
    except Exception:
        return "dev"


def chat(system, user, max_tokens=1200):
    """单轮对话，返回纯文本。"""
    c = cfg()
    if not c["enabled"]:
        raise AIError("off", "内置 AI 未启用：请到「AI 助手」页选择服务商并填 Key")
    if not c["key"]:
        raise AIError("nokey", "还没有填 API Key")
    if len(user) > MAX_CHARS:
        user = user[:MAX_CHARS] + "\n…（内容过长已截断）"
    payload = {"model": c["model"], "temperature": 0.2, "max_tokens": max_tokens,
               "messages": [{"role": "system", "content": system},
                            {"role": "user", "content": user}]}
    j = _post(c["base_url"] + "/chat/completions", payload, c["key"], c["timeout"], c["extra"])
    try:
        txt = j["choices"][0]["message"]["content"]
    except Exception:
        raise AIError("format", "返回格式不是 OpenAI 兼容结构", json.dumps(j)[:300])
    return (txt or "").strip()


def test():
    t0 = time.time()
    try:
        out = chat("你只回答两个字：可用", "连通性测试", max_tokens=16)
        return {"ok": True, "ms": int((time.time() - t0) * 1000), "reply": out[:40]}
    except AIError as e:
        return {"ok": False, "kind": e.kind, "error": e.msg, "detail": e.detail,
                "ms": int((time.time() - t0) * 1000)}


# ---------------------------------------------------------------- 任务
TASKS = {
    "exp-oar-bucket": {"label": "口径归类（哪些科目算薪酬/折旧/广告费）",
                       "pending": "oar_bucket_pending.json",
                       "result": "oar_map_ai_result.json"},
    "exp-table2-concl": {"label": "BKD 表2 变动结论",
                         "pending": "table2_pending.json",
                         "result": "table2_concl_result.json"},
    "exp-oar-reason": {"label": "OAR 变动分析（分析段文字）",
                       "pending": "oar_reason_pending.json",
                       "result": "oar_concl_result.json"},
}


def _pending_path(task):
    return os.path.join(TASK_DIR, TASKS[task]["pending"])


def _result_path(task):
    """结果侧文件与「任务卡」约定同一位置，保证外部 AI 与内置 AI 可互换。"""
    p = os.path.join(paths.DATA_DIR, "bkd_work", TASKS[task]["result"])
    return p


def list_tasks():
    """三条通道的待办条数与是否可跑（给界面显示）。"""
    out = []
    for k, v in TASKS.items():
        p = _pending_path(k)
        n = None
        if os.path.exists(p):
            try:
                j = json.load(io.open(p, encoding="utf-8"))
                n = len(j.get("items") or j.get("candidates") or [])
            except Exception:
                n = -1
        out.append({"id": k, "label": v["label"], "pending": n,
                    "file": v["pending"], "done": os.path.exists(_result_path(k))})
    return out


_RULE = """硬性规则（违反会被程序丢弃）：
1. 只输出 JSON，不要 Markdown 代码块、不要解释、不要前后寒暄。
2. 不要编造任何数字。金额、比例、行数由程序提供，你只做归类与文字表述。
3. 文字要像审计师写给复核人看的：说清是什么业务、为什么变动、是否需要追加程序。
   不知道就明说「需人工判断」，不要编业务原因。
4. 用简体中文。"""


def _parse_json(txt):
    t = (txt or "").strip()
    t = re.sub(r"^```(?:json)?\s*|\s*```$", "", t, flags=re.S)
    i, k = t.find("{"), t.rfind("}")
    if i < 0 or k <= i:
        raise AIError("format", "AI 没有返回 JSON", t[:300])
    try:
        return json.loads(t[i:k + 1])
    except Exception as e:
        raise AIError("format", "JSON 解析失败：%s" % e, t[:300])


_BUCKETS = ("工资薪酬", "固定资产折旧费", "使用权资产折旧费", "无形资产摊销费",
            "长期待摊费用摊销费", "租赁费", "预提费用（含预提和冲销）", "其他")


def run_task(task, limit=60, progress=None):
    """执行一条通道：读 pending → 调 AI → 写 result。返回摘要。"""
    if task not in TASKS:
        raise AIError("arg", "未知任务：%s" % task)
    p = _pending_path(task)
    if not os.path.exists(p):
        raise AIError("nofile", "待办清单不存在（先在一键生成里产出）：%s" % p)
    doc = json.load(io.open(p, encoding="utf-8"))
    items = doc.get("items") or doc.get("candidates") or []
    if not items:
        return {"ok": True, "task": task, "total": 0, "written": 0,
                "note": "清单为空（没有待判断的科目）"}
    items = items[:int(limit or 60)]

    if task == "exp-oar-bucket":
        sys_p = ("你是审计费用口径归类助手。把每个费用科目名归到下列口径之一，并给一句理由。\n"
                 "可选口径：" + "、".join(_BUCKETS) + "\n"
                 "归类要点：工资/奖金/社保/公积金/福利/劳务/董事津贴 → 工资薪酬；"
                 "股权激励、股份支付 → 工资薪酬；折旧摊销按资产类型分开；"
                 "明细科目名（差旅费_xx、办公费_xx）一律归「其他」，不要把明细名当口径。\n"
                 '输出：{"items":[{"company":"","name":"","bucket":"","reason":"","confidence":0.8}]}\n'
                 + _RULE)
        usr = json.dumps({"items": [{"company": i.get("company") or i.get("entity") or "",
                                     "name": i.get("name") or i.get("subject") or "",
                                     "amount": i.get("amount")} for i in items]},
                         ensure_ascii=False)
        j = _parse_json(chat(sys_p, usr, 2000))
        got = [x for x in (j.get("items") or []) if (x.get("bucket") in _BUCKETS)]
        out = {"items": [{"company": x.get("company", ""), "name": x.get("name", ""),
                          "bucket": x.get("bucket"), "reason": (x.get("reason") or "")[:200],
                          "confidence": x.get("confidence"),
                          "source": "内置AI"} for x in got],
               "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"), "engine": "builtin"}
    elif task == "exp-table2-concl":
        sys_p = ("你是审计费用分析助手。为每个费用口径写一段变动结论（80–160 字）："
                 "说明业务性质、变动主因、是否一次性、是否需要追加程序。\n"
                 '输出：{"page|口径":"文字", ...} 或 {"items":[{"key":"页|口径","text":"..."}]}\n'
                 + _RULE)
        usr = json.dumps({"items": [{k: i.get(k) for k in
                                     ("page", "bucket", "cur", "prior", "delta", "pct", "subjects")}
                                    for i in items]}, ensure_ascii=False)
        j = _parse_json(chat(sys_p, usr, 2000))
        if isinstance(j.get("items"), list):
            out = {str(x.get("key") or x.get("bucket") or ""): (x.get("text") or "").strip()
                   for x in j["items"] if (x.get("text") or "").strip()}
        else:
            out = {str(k): str(v).strip() for k, v in j.items() if isinstance(v, str)}
        out = {"conclusions": out, "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
               "engine": "builtin"}
    else:                                   # exp-oar-reason
        sys_p = ("你是审计分析性复核助手。为每个 OAR 口径写一段分析（100–200 字），"
                 "结合给出的序时账证据（对手方/业务性质/金额占比）说明变动原因。\n"
                 '输出：{"items":[{"key":"页|口径","text":"..."}]}\n' + _RULE)
        usr = json.dumps({"items": [{k: i.get(k) for k in
                                     ("page", "bucket", "cur", "prior", "delta", "pct",
                                      "evidence", "subjects")} for i in items]},
                         ensure_ascii=False)
        j = _parse_json(chat(sys_p, usr, 2500))
        out = {"conclusions": {str(x.get("key")): (x.get("text") or "").strip()
                               for x in (j.get("items") or []) if (x.get("text") or "").strip()},
               "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"), "engine": "builtin"}

    rp = _result_path(task)
    os.makedirs(os.path.dirname(rp), exist_ok=True)
    with io.open(rp, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    written = len(out.get("items") or out.get("conclusions") or {})
    return {"ok": True, "task": task, "total": len(items), "written": written,
            "result_file": rp}


def has_number(text):
    """给调用方做最后一道闸：AI 文字里出现「数字+元/万/%」就认为是编造，丢弃。"""
    return bool(re.search(r"\d[\d,\.]*\s*(元|万元|%|亿)", str(text or "")))
