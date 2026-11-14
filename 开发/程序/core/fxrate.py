# -*- coding: utf-8 -*-
"""人民币汇率中间价：从人民银行公告页抓取，并按币种取用。

为什么要自动抓：C.02a 外币测算的汇率列原来是留空标黄让人手填。手工填的汇率
来源无法追溯，而且一张底稿上往往有 5~8 个币种，抄错一个整行差异就出来了。
人民银行每个工作日都发《人民币汇率中间价公告》，正文就是一句话，机器可读。

**抓不到不影响生成**：离线、公告页改版、网络不通时返回空，调用方退回"留空标黄"，
绝不编一个汇率出来——汇率写错比留空危险，留空看得出要补，写错了看不出。
"""

import io
import json
import os
import re
import time
import urllib.parse
import urllib.request
from datetime import date as _date_cls, timedelta as _timedelta

#: HTTPS 上下文要用 ssl。**不能让它把整个模块拖死**：
#: 分发包曾经把 libssl-3.dll / libcrypto-3.dll 当"用不到的文件"删掉，
#: 结果 `import ssl` 直接抛 ImportError，本模块 import 失败 →
#: C.02a 外币测算整张表都没生成（2026-11-07 客户实测）。
#: 这里降级为"没有 ssl 就说明环境缺件"，抓取时按失败处理并说清原因，
#: 主表照常生成、汇率留空标黄，由人补。
try:
    import ssl
except Exception:                                     # pragma: no cover
    ssl = None

#: 人民银行公告里用的币种名 ← 账套里常见的写法
PBC_ALIAS = {
    "人民币": "人民币", "CNY": "人民币", "RMB": "人民币",
    "美元": "美元", "美金": "美元", "USD": "美元",
    "港元": "港元", "港币": "港元", "香港元": "港元", "HKD": "港元",
    "欧元": "欧元", "EUR": "欧元",
    "日元": "日元", "日圆": "日元", "日本日圆": "日元", "JPY": "日元",
    "英镑": "英镑", "GBP": "英镑",
    "新加坡元": "新加坡元", "新加坡币": "新加坡元", "SGD": "新加坡元",
    "澳大利亚元": "澳大利亚元", "澳币": "澳大利亚元", "AUD": "澳大利亚元",
    "新西兰元": "新西兰元", "NZD": "新西兰元",
    "瑞士法郎": "瑞士法郎", "CHF": "瑞士法郎",
    "加拿大元": "加拿大元", "加元": "加拿大元", "CAD": "加拿大元",
    "墨西哥比索": "墨西哥比索", "墨西哥币": "墨西哥比索", "MXN": "墨西哥比索",
    "巴西雷亚尔": "巴西雷亚尔", "BRL": "巴西雷亚尔",
    "挪威克朗": "挪威克朗", "NOK": "挪威克朗",
    "瑞典克朗": "瑞典克朗", "SEK": "瑞典克朗",
    "丹麦克朗": "丹麦克朗", "DKK": "丹麦克朗",
    "韩元": "韩元", "KRW": "韩元",
    "泰铢": "泰铢", "THB": "泰铢",
    "土耳其里拉": "土耳其里拉", "TRY": "土耳其里拉",
    "俄罗斯卢布": "俄罗斯卢布", "RUB": "俄罗斯卢布",
    "南非兰特": "南非兰特", "ZAR": "南非兰特",
    "澳门元": "澳门元", "MOP": "澳门元",
}

#: 人民银行首页的「人民币汇率中间价」栏目（每天更新，指向当日公告）
INDEX_URL = "https://www.pbc.gov.cn/zhengcehuobisi/125207/125217/125925/index.html"

_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")


def _plain(html):
    """HTML → 纯文本（保留中文标点，去掉标签与多余空白）。"""
    t = re.sub(r"(?is)<(script|style).*?</\1>", " ", html)
    t = re.sub(r"<[^>]+>", " ", t)
    t = t.replace("&nbsp;", " ").replace("&amp;", "&")
    return re.sub(r"[\s\u3000]+", " ", t)


def parse_rates(text):
    """从公告正文解析 {币种: 1 单位该币种 = X 人民币}。

    公告里有两种句式，都要认：
      1美元对人民币6.7351元            ← 正向（含 100日元对人民币…）
      人民币1元对1.2005澳门元          ← 反向，要取倒数
    """
    out = {}
    body = _plain(text)
    seg = re.search(r"人民币汇率中间价为(.{20,2000})", body)
    if not seg:
        return out
    s = seg.group(1)

    # 正向：<N><币种>对人民币<汇率>元   （N 可能是 1 或 100）
    for m in re.finditer(r"(\d+)\s*([\u4e00-\u9fa5]{2,8})对人民币([\d.]+)元", s):
        base = float(m.group(1))
        cur = m.group(2)
        try:
            v = float(m.group(3)) / (base or 1)
        except (TypeError, ValueError, ZeroDivisionError):
            continue
        if v > 0:
            out[cur] = round(v, 6)

    # 反向：人民币1元对<汇率><币种>
    for m in re.finditer(r"人民币\s*1\s*元对\s*([\d.]+)\s*([\u4e00-\u9fa5]{2,8})", s):
        try:
            v = float(m.group(1))
        except (TypeError, ValueError):
            continue
        cur = m.group(2)
        if v > 0 and cur not in out:
            out[cur] = round(1.0 / v, 6)
    return out


def _get(url, timeout=25):
    req = urllib.request.Request(url, headers={"User-Agent": _UA,
                                               "Accept-Language": "zh-CN,zh;q=0.9"})
    if ssl is None:
        # 缺 libssl/libcrypto：https 根本建不起来，说清原因比抛个谜之 ImportError 强
        raise RuntimeError("运行环境缺少 OpenSSL（libssl-3.dll / libcrypto-3.dll），"
                           "无法访问 https 汇率公告页")
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
        return r.read()


#: 栏目页的静态分页模板：17105 是该栏目的 moduleid，N 从 1 开始，按日期倒序。
#: 第 1 页与 index.html 等价；N 超过 totalpage 返回 404。
LIST_URL = "https://www.pbc.gov.cn/zhengcehuobisi/125207/125217/125925/17105-%d.html"


def _list_items(html):
    """从栏目页（任意一页）解析出 [(日期, 标题, 绝对URL), ...]。

    必须**限定在 <font class="newslist_style"> 这一段里**解析：
    每页头部还有 <meta name="createDate" content="2026-11-10 ...">，
    整页找日期会多出一条"当天"，看起来像置顶公告，实际是页面生成时间。
    """
    out = []
    for seg in html.split('<font class="newslist_style"')[1:]:
        seg = seg[:1500]
        m = re.search(r'href="(/zhengcehuobisi/125207/125217/125925/[^"]+/index\.html)"', seg)
        t = re.search(r'title="([^"]*)"', seg)
        d = re.search(r'<span class="hui12">(\d{4}-\d{1,2}-\d{1,2})</span>', seg)
        if not (m and t and d):
            continue
        if "中间价" not in (t.group(1) or ""):
            continue
        day = d.group(1)
        y, mo, dd = day.split("-")
        out.append(("%s-%02d-%02d" % (y, int(mo), int(dd)),
                    t.group(1), "https://www.pbc.gov.cn" + m.group(1)))
    return out


def _url_cache_path(cache_dir):
    return os.path.join(cache_dir, "fx", "pbc_urls.json") if cache_dir else ""


def _url_cache_get(path, ymd):
    if not path or not os.path.isfile(path):
        return {}
    try:
        with io.open(path, encoding="utf-8") as fh:
            d = json.load(fh) or {}
        it = d.get(ymd) or {}
        if it.get("url"):
            return {"url": it["url"], "title": it.get("title") or "", "date": ymd,
                    "_cached": True}
    except Exception:
        pass
    return {}


def _url_cache_put(path, ymd, url, title):
    if not path:
        return
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        d = {}
        if os.path.isfile(path):
            try:
                with io.open(path, encoding="utf-8") as fh:
                    d = json.load(fh) or {}
            except Exception:
                d = {}
        d[ymd] = {"url": url, "title": title,
                  "foundAt": time.strftime("%Y-%m-%d %H:%M:%S")}
        with io.open(path, "w", encoding="utf-8") as fh:
            json.dump(d, fh, ensure_ascii=False, indent=1)
    except Exception:
        pass


def find_announcement_on(date, cache_dir="", timeout=20, max_pages=8):
    """定位**指定日期**的人行中间价公告页，返回 {url,title,date}；没有就返回 {}。

    只认这一天，**绝不退到相邻交易日**：截一张别的日期的公告图，等于给底稿
    放一张日期不符的证据。

    做法：栏目页是静态分页（每页 20 条、按日期倒序），先线性翻前 max_pages 页
    （目标日期通常在 1~3 页内），还没找到且总页数更多时，再对剩下的页做二分。
    找到的 URL 会写进 <cache_dir>/fx/pbc_urls.json，同一天下次零请求。
    """
    d = str(date or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})$", d)
    if not m:
        return {}
    y, mo, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
    want = "%04d-%02d-%02d" % (y, mo, dd)
    ymd = "%04d%02d%02d" % (y, mo, dd)

    cache = _url_cache_path(cache_dir)
    hit = _url_cache_get(cache, ymd)
    if hit:
        return hit

    try:
        html = _get(INDEX_URL, timeout).decode("utf-8", errors="replace")
    except Exception:
        return {}
    tm = re.search(r'totalpage="(\d+)"', html)
    total = int(tm.group(1)) if tm else 1

    def check(items):
        for day, title, url in items:
            if day == want:
                return {"url": url, "title": title, "date": want}
        return None

    def scan(n):
        """取第 n 页并解析；取不到返回 []。"""
        if n <= 1:
            return _list_items(html)
        try:
            h = _get(LIST_URL % n, timeout).decode("utf-8", errors="replace")
        except Exception:
            return []
        return _list_items(h)

    # ① 线性：目标日期通常离今天不超过一两个月，1~3 页就命中
    for n in range(1, min(total, max_pages) + 1):
        items = scan(n)
        if not items:
            return {}
        got = check(items)
        if got:
            _url_cache_put(cache, ymd, got["url"], got["title"])
            return got
        newest, oldest = items[0][0], items[-1][0]
        if oldest < want < newest:
            # 落在这页的日期区间内却没有这一条 → 该日未发布（周末/节假日）
            return {}
        if want > newest:
            # 已经翻过头（目标比这一页还新）→ 该日没有公告
            return {}

    # ② 二分：日期严格倒序，页号越大越旧
    lo, hi = max_pages + 1, total
    while lo <= hi:
        mid = (lo + hi) // 2
        items = scan(mid)
        if not items:
            return {}
        got = check(items)
        if got:
            _url_cache_put(cache, ymd, got["url"], got["title"])
            return got
        newest, oldest = items[0][0], items[-1][0]
        if oldest < want < newest:
            return {}
        if want < oldest:
            lo = mid + 1
        else:
            hi = mid - 1
    return {}


def find_announcement(date, timeout=25, max_pages=1):
    """在栏目页里找该日期的公告链接。date 形如 2026-09-30。

    默认只查第 1 页——`fetch()` 走这里，要保证取汇率这条链路**不变慢**；
    截图需要回翻历史页时用 `find_announcement_on(max_pages=8)`。
    """
    return (find_announcement_on(date, "", timeout, max_pages) or {}).get("url", "")



def fetch(date="", url="", cache_dir="", timeout=25, back_days=7):
    """抓某日的中间价，返回 dict（含 _date/_url/_raw）。抓不到返回 {}。

    缓存放在 cache_dir/fx/<date>.json，同一天只联网一次——汇率一天只发一次，
    而且生成底稿时会被反复调用。

    两条数据通道，顺序固定：

      1. **人行公告页**（首选）。公告正文是权威原文，还能截图留证。
      2. **交易中心历史接口**（兜底）。人行栏目页只列最近约 20 个工作日，
         资产负债表日稍微早一点（比如月末结账拖了一个月才做底稿）就翻不到
         那条公告了，此时 `find_announcement` 返回空、汇率整列为空。
         旧版之所以"成功"，往往是因为当时那份日期的缓存文件已经存在，
         命中缓存就不联网了——换台机器或换个数据目录就立刻失效。
    """
    cache = ""
    if cache_dir and date:
        try:
            cache = os.path.join(cache_dir, "fx", "%s.json" % str(date).replace("-", ""))
            if os.path.isfile(cache):
                with io.open(cache, encoding="utf-8") as fh:
                    d = json.load(fh)
                if d.get("rates"):
                    d["_cached"] = True
                    return d
        except Exception:
            cache = ""

    target = url or (find_announcement(date, timeout) if date else "")
    if target:
        try:
            html = _get(target, timeout).decode("utf-8", errors="replace")
        except Exception:
            html = ""
        rates = parse_rates(html) if html else {}
        if rates:
            body = _plain(html)
            m = re.search(r"(\d{4}年\d{1,2}月\d{1,2}日)", body)
            d = {"rates": rates, "url": target,
                 "date": m.group(1) if m else "",
                 "raw": re.search(r"人民币汇率中间价为(.{0,2000})", body).group(0) if
                        re.search(r"人民币汇率中间价为(.{0,2000})", body) else "",
                 "source": "PBC",
                 "requestedDate": str(date),
                 "fetchedAt": time.strftime("%Y-%m-%d %H:%M:%S")}
            _save_cache(cache, d)
            return d

    # 人行公告页这条走不通（旧公告已被翻出栏目页 / 页面改版 / 网络原因）→ 交易中心兜底
    if date:
        d = fetch_cfets(date, timeout, back_days=back_days)
        if d.get("rates"):
            _save_cache(cache, d)
            return d
    return {}


def _save_cache(cache, d):
    if not cache:
        return
    try:
        os.makedirs(os.path.dirname(cache), exist_ok=True)
        with io.open(cache, "w", encoding="utf-8") as fh:
            json.dump({k: v for k, v in d.items() if k != "_cached"},
                      fh, ensure_ascii=False, indent=1)
    except Exception:
        pass


#: 交易中心历史中间价接口（一次可取当日全部 25 个币种）
CFETS_URL = "https://www.chinamoney.com.cn/ags/ms/cm-u-bk-ccpr/CcprHisNew"
CFETS_PAGE = "https://www.chinamoney.com.cn/chinese/bkccpr/"

#: 交易中心的币种代码 → 人行公告里的中文名（与人行口径保持一致，才能对上 PBC_ALIAS）
CFETS_NAMES = {
    "USD": "美元", "EUR": "欧元", "JPY": "日元", "HKD": "港元", "GBP": "英镑",
    "AUD": "澳大利亚元", "NZD": "新西兰元", "SGD": "新加坡元", "CHF": "瑞士法郎",
    "CAD": "加拿大元", "MOP": "澳门元", "MYR": "林吉特", "RUB": "俄罗斯卢布",
    "ZAR": "南非兰特", "KRW": "韩元", "AED": "阿联酋迪拉姆", "SAR": "沙特里亚尔",
    "HUF": "匈牙利福林", "PLN": "波兰兹罗提", "DKK": "丹麦克朗", "SEK": "瑞典克朗",
    "NOK": "挪威克朗", "TRY": "土耳其里拉", "MXN": "墨西哥比索", "THB": "泰铢",
    "BRL": "巴西雷亚尔",
}


def _cfets_query(date, timeout=30):
    """查某一天的中间价原始响应。返回 (head, records)。"""
    body = urllib.parse.urlencode({
        "startDate": date, "endDate": date, "currency": "",
        "pageNum": "1", "pageSize": "15"}).encode()
    req = urllib.request.Request(
        CFETS_URL, data=body,
        headers={"User-Agent": _UA,
                 "Content-Type": "application/x-www-form-urlencoded",
                 "Referer": CFETS_PAGE})
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
        d = json.loads(r.read().decode("utf-8", errors="replace"))
    data = (d or {}).get("data") or {}
    return (data.get("head") or []), ((d or {}).get("records") or [])


def fetch_cfets(date, timeout=20, back_days=3):
    """交易中心历史中间价，返回结构与 fetch() 相同；取不到返回 {}。

    列名形如 `USD/CNY`（1 外币兑人民币）、`100JPY/CNY`（百单位）、
    `CNY/MOP`（1 人民币兑外币，要取倒数）——三种都要认，否则日元会差 100 倍、
    澳门元/韩元这类反向报价会整行算错。

    该日无数据（周末、休市）时按自然日往前退，最多 back_days 天。实际使用的日期
    记在 `date`、原始请求日期记在 `requestedDate`、后退天数记在 `fallbackDays`，
    底稿的 Notes 会把这件事写出来，不允许静默换日期。
    """
    if not date:
        return {}
    try:
        y, m, d = (str(date).split("-") + ["", "", ""])[:3]
        cur = _date_cls(int(y), int(m), int(d))
    except Exception:
        return {}

    for i in range(back_days + 1):
        day = cur - _timedelta(days=i)
        iso = day.strftime("%Y-%m-%d")
        try:
            head, recs = _cfets_query(iso, timeout)
        except Exception:
            recs = []
            head = []
        if not recs:
            continue
        vals = recs[0].get("values") or []
        rates = {}
        for token, val in zip(head, vals):
            try:
                v = float(val)
            except (TypeError, ValueError):
                continue
            if "/" not in token:
                continue
            a, b = token.split("/", 1)
            mult = 1.0
            mm = re.match(r"(\d+)([A-Z]{2,4})$", a)
            if mm:
                mult, a = float(mm.group(1)), mm.group(2)
            name = CFETS_NAMES.get(a) or CFETS_NAMES.get(b)
            if not name:
                continue
            if b == "CNY" and v > 0:                 # 1 外币 = v 人民币
                rates[name] = round(v / (mult or 1), 6)
            elif a == "CNY" and v > 0:               # 1 人民币 = v 外币
                rates[name] = round(1.0 / v, 6)
        if not rates:
            continue
        out = {"rates": rates, "url": CFETS_PAGE,
               "date": "%d年%d月%d日" % (day.year, day.month, day.day),
               "requestedDate": str(date),
               "source": "CFETS",
               "raw": "中国外汇交易中心受权公布人民币汇率中间价（%s）" % iso,
               "fetchedAt": time.strftime("%Y-%m-%d %H:%M:%S")}
        if i:
            out["fallbackDays"] = i
        return out
    return {}


#: 本机浏览器的常见安装位置（无头截图用，不依赖 playwright 包）
_BROWSERS = (
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
)


def _browser():
    """找一个可用于无头截图的浏览器；都没有返回空。"""
    for p in _BROWSERS:
        if os.path.isfile(p):
            return p
    return ""


def shot_path(date, cache_dir):
    """汇率公告的截图路径。没有就返回空。

    容错写法：`pbc_20260930.png` 是约定名，但人手工放的时候常写成
    `pbc_2026-09-30.png`、扩展名大写；这几种都认，避免"图明明放了却读不到"。
    """
    if not (date and cache_dir):
        return ""
    ymd = str(date).replace("-", "")
    d = os.path.join(cache_dir, "fx")
    for name in ("pbc_%s.png" % ymd, "pbc_%s.png" % date,
                 "pbc_%s.PNG" % ymd, "pbc_%s.PNG" % date):
        p = os.path.join(d, name)
        if os.path.isfile(p):
            return p
    return ""


def shot_cmd(date, url="", cache_dir=""):
    """给出可直接复制执行的无头截图命令（写进任务卡，替 AI 省掉"怎么截"）。"""
    ymd = str(date or "").replace("-", "")
    if not (ymd and cache_dir):
        return ""
    out = os.path.join(cache_dir, "fx", "pbc_%s.png" % ymd)
    exe = _browser() or r"msedge.exe"
    return ('"%s" --headless=new --disable-gpu --hide-scrollbars '
            '--window-size=1100,900 --screenshot="%s" "%s"'
            % (exe, os.path.normpath(out), url or ""))


def shot(date, url="", cache_dir="", timeout=90):
    """确保有该日的公告截图，返回路径；做不到返回空。

    优先用已有的 png（人手工截的、或上次截的）。没有才尝试调 playwright-cli——
    内置运行时没有浏览器驱动，只有开发机/装了 playwright 的机器能自动截。
    两种都不行就返回空，底稿里只留来源链接，不报错。
    """
    import shutil
    import subprocess
    if not (date and cache_dir):
        return ""
    # 必须用绝对路径：Edge/Chrome 的 --screenshot 按浏览器自己的工作目录解析相对路径，
    # 传相对路径会"返回 0 但文件不在预期位置"，看起来像截图失败。
    out = os.path.abspath(os.path.join(cache_dir, "fx",
                                       "pbc_%s.png" % str(date).replace("-", "")))
    if os.path.isfile(out):
        return out
    if not url:
        return ""
    try:
        os.makedirs(os.path.dirname(out), exist_ok=True)
    except Exception:
        return ""
    # ① playwright-cli：开发机常装，能跑就跑
    if shutil.which("playwright-cli"):
        try:
            sess = "ptwfx%d" % (int(time.time()) % 100000)
            base = ["playwright-cli", "--browser=msedge", "-s=%s" % sess]
            subprocess.run(base + ["open", url], timeout=timeout,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            time.sleep(4)
            subprocess.run(base + ["resize", "1100", "900"], timeout=30,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            subprocess.run(base + ["screenshot", "--filename", out], timeout=timeout,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if os.path.isfile(out) and os.path.getsize(out) > 12288:
                return out
        except Exception:
            pass
    # ② 本机 Edge/Chrome 无头截图：不依赖任何 Python 包，客户机装了浏览器就能用
    exe = _browser()
    if not exe:
        return ""
    for flag in ("--headless=new", "--headless"):
        try:
            subprocess.run([exe, flag, "--disable-gpu", "--hide-scrollbars",
                            "--window-size=1100,900", "--screenshot=" + out, url],
                           timeout=timeout,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            if os.path.isfile(out) and os.path.getsize(out) > 12288:
                return out
        except Exception:
            continue
    return ""


def rate_for(currency, table):
    """按账套里的币种写法取汇率。人民币恒为 1；取不到返回 None（不要编）。"""
    if not currency:
        return None
    name = PBC_ALIAS.get(str(currency).strip().upper(), None) or \
        PBC_ALIAS.get(str(currency).strip(), None) or str(currency).strip()
    if name == "人民币":
        return 1.0
    v = (table or {}).get("rates", {}).get(name)
    if isinstance(v, (int, float)) and v > 0:
        return float(v)
    return None
