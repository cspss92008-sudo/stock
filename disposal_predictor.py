# -*- coding: utf-8 -*-
"""
隔日高機率進處置股票 自動篩選程式  (上市 TWSE + 上櫃 TPEx)
=====================================================================
規則依據：證交所「公布或通知注意交易資訊暨處置作業要點」第6條 (115.08.03 版)
  處置條件：
    A. 連續 3 個營業日 以「第一款」(最近六日累積漲跌幅異常) 公布注意
    B. 連續 5 個營業日            以第一款~第八款 公布注意
    C. 最近 10 個營業日內有 6 日   以第一款~第八款 公布注意
    D. 最近 30 個營業日內有 12 日  以第一款~第八款 公布注意
  → 本程式找出「今天再差 1 天就達標」的股票，即明日只要再被公布注意就進處置。
  (第九~十三款：量、週轉率、借券、當沖…不計入處置基數，程式會排除)

門檻價估算：
  第一款：明日收盤 > 6日前基準價 × 1.32  或  > max(基準×1.25, 基準+50元)
          (未計入「與大盤/類股平均差幅≧20%」條件，屬保守估計)
  第六款：股價淨值比 ≧ 6  → 門檻價 ≈ 每股淨值 × 6 (仍需當日週轉率≧5% 等條件)

使用：
  pip install requests pandas openpyxl
  python disposal_predictor.py            # 以最近一個營業日為基準
  python disposal_predictor.py 20260903   # 指定基準日
輸出：console 表格 + disposal_candidates_YYYYMMDD.xlsx
"""
import re
import sys
import json
import time
import datetime as dt
from collections import defaultdict
from pathlib import Path

import requests
import pandas as pd

# ---------------------------------------------------------------- 設定
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept": "application/json"}
CACHE_DIR = Path("cache"); CACHE_DIR.mkdir(exist_ok=True)
SLEEP = 0.8            # 對交易所禮貌一點，避免被擋
LOOKBACK = 30          # 需回溯的營業日數 (30日內12次)

# 交易所端點 (若交易所改版，只需改這裡；用瀏覽器 F12 → Network 看實際 URL)
TWSE_NOTICE  = "https://www.twse.com.tw/rwd/zh/announcement/notice?date={d8}&response=json"
TWSE_PUNISH  = "https://www.twse.com.tw/rwd/zh/announcement/punish?startDate={d8}&endDate={d8}&response=json"
TWSE_CAL     = "https://www.twse.com.tw/rwd/zh/afterTrading/MI_INDEX?date={d8}&type=IND&response=json"
TWSE_STOCKDAY= "https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date={d8}&stockNo={code}&response=json"
TWSE_BWIBBU  = "https://openapi.twse.com.tw/v1/exchangeReport/BWIBBU_ALL"     # 本益比/淨值比 (全部上市)
TPEX_NOTICE  = "https://www.tpex.org.tw/www/zh-tw/bulletin/attention?date={dslash}&response=json"
TPEX_PUNISH  = "https://www.tpex.org.tw/www/zh-tw/bulletin/disposal?date={dslash}&response=json"
TPEX_STOCKDAY= "https://www.tpex.org.tw/www/zh-tw/afterTrading/tradingStock?code={code}&date={dslash}&response=json"
TPEX_PBR     = "https://www.tpex.org.tw/openapi/v1/tpex_mainboard_peratio_analysis"   # 上櫃本益比/淨值比

# ---------------------------------------------------------------- 工具
def get_json(url, cache_key=None):
    """GET + 本機快取 (歷史日期資料不會變，快取後重跑很快)"""
    if cache_key:
        f = CACHE_DIR / (re.sub(r"[^\w.-]", "_", cache_key) + ".json")
        if f.exists():
            return json.loads(f.read_text("utf-8"))
    for _ in range(3):
        try:
            r = requests.get(url, headers=HEADERS, timeout=20)
            r.raise_for_status()
            js = r.json()
            if cache_key:
                f.write_text(json.dumps(js, ensure_ascii=False), "utf-8")
            time.sleep(SLEEP)
            return js
        except Exception as e:
            print(f"  [retry] {url} -> {e}")
            time.sleep(2)
    return None

def d8(d):  return d.strftime("%Y%m%d")
def dsl(d): return d.strftime("%Y/%m/%d")

def is_trading_day(d):
    js = get_json(TWSE_CAL.format(d8=d8(d)), f"cal_{d8(d)}")
    return bool(js) and js.get("stat") == "OK"

def trading_days(end, n):
    """回傳 end 往前 n 個營業日 (舊→新)"""
    days, d = [], end
    while len(days) < n:
        if d.weekday() < 5 and is_trading_day(d):
            days.append(d)
        d -= dt.timedelta(days=1)
    return days[::-1]

def latest_trading_day(base=None):
    d = base or dt.date.today()
    while not (d.weekday() < 5 and is_trading_day(d)):
        d -= dt.timedelta(days=1)
    return d

# ---------------------------------------------------------------- 注意款別解析
CN = {"一":1,"二":2,"三":3,"四":4,"五":5,"六":6,"七":7,"八":8,"九":9,"十":10,"十一":11,"十二":12,"十三":13}

def classify_segment(s):
    """把一段注意交易資訊文字對應到第幾款 (關鍵字優先序很重要)"""
    if "存託憑證" in s or "溢價" in s or "折價" in s:          return 8
    if "借券" in s:                                            return 12
    if "當日沖銷" in s or "當沖" in s:                          return 13
    if "券資比" in s:                                          return 7
    if "本益比" in s or "淨值比" in s:                          return 6
    if "證券商" in s and "累積" in s:                          return 5
    if "價差" in s and "元" in s:                              return 11
    if "起迄兩個營業日" in s and "百分比" in s:                 return 2
    if "累積" in s and ("漲" in s or "跌" in s):
        if "成交量" in s and "放大" in s:                      return 3
        if "週轉率" in s:                                      return 4
        return 1
    if "累積週轉率" in s or ("週轉率" in s and "超過" in s):    return 10
    if "成交量" in s and "放大" in s:                          return 9
    return None

def parse_clauses(text):
    """回傳 set，例如 {1,4}。先找『第X款』，找不到再用關鍵字。"""
    text = str(text or "")
    found = {CN[m] for m in re.findall(r"第([一二三四五六七八九十]+)款", text) if m in CN}
    if found:
        return found
    # 依「一、二、三、」或換行切段
    segs = re.split(r"(?:^|\s)[一二三四五六七八九十]+、|\n|；", text)
    for seg in segs:
        c = classify_segment(seg)
        if c: found.add(c)
    return found

# ---------------------------------------------------------------- 注意股 / 處置股 抓取
def rows_to_records(js):
    """TWSE/TPEx 的 json 通常是 {fields:[...], data:[[...]]} 或 {tables:[{fields,data}]}"""
    if not js: return []
    tables = js.get("tables") or [js]
    out = []
    for t in tables:
        fields, data = t.get("fields") or [], t.get("data") or []
        for row in data:
            out.append(dict(zip(fields, row)))
    return out

def find_col(rec, *keys):
    for k in rec:
        if any(x in k for x in keys):
            return rec[k]
    return ""

def fetch_notice(d):
    """回傳 list of (code, name, market, clauses:set, raw_text)"""
    res = []
    for market, url, key in (("上市", TWSE_NOTICE.format(d8=d8(d)), f"twse_notice_{d8(d)}"),
                             ("上櫃", TPEX_NOTICE.format(dslash=dsl(d)), f"tpex_notice_{d8(d)}")):
        for rec in rows_to_records(get_json(url, key)):
            code = str(find_col(rec, "證券代號", "代號", "Code")).strip()
            name = str(find_col(rec, "證券名稱", "名稱", "Name")).strip()
            text = " ".join(str(v) for v in rec.values())
            if not re.fullmatch(r"\d{4}[A-Z0-9]?", code):   # 排除權證/ETF等非4碼
                continue
            res.append((code, name, market, parse_clauses(text), text))
    return res

def fetch_punished_now(d):
    """今天仍在處置中的代號 (處置期間注意日數不列入基數，且已在處置就不用預測)"""
    codes = set()
    for url, key in ((TWSE_PUNISH.format(d8=d8(d)), f"twse_punish_{d8(d)}"),
                     (TPEX_PUNISH.format(dslash=dsl(d)), f"tpex_punish_{d8(d)}")):
        for rec in rows_to_records(get_json(url, key)):
            code = str(find_col(rec, "證券代號", "代號", "Code")).strip()
            period = str(find_col(rec, "處置期間", "期間"))
            # 期間格式多為 "115/09/01～115/09/05"，若含今天以後日期則仍在處置
            m = re.findall(r"(\d{3})/(\d{2})/(\d{2})", period)
            if len(m) >= 2:
                y, mo, da = map(int, m[1]); end = dt.date(y + 1911, mo, da)
                if end >= d: codes.add(code)
            elif code:
                codes.add(code)
    return codes

# ---------------------------------------------------------------- 價格 / 淨值
def close_history(code, market, d):
    """回傳 [(date, close)] 最近兩個月，舊→新"""
    closes = []
    for m in (0, 1):
        first = (d.replace(day=1) - dt.timedelta(days=1)).replace(day=1) if m else d.replace(day=1)
        if market == "上市":
            js = get_json(TWSE_STOCKDAY.format(d8=d8(first), code=code), f"twse_day_{code}_{d8(first)[:6]}")
        else:
            js = get_json(TPEX_STOCKDAY.format(dslash=dsl(first), code=code), f"tpex_day_{code}_{d8(first)[:6]}")
        for rec in rows_to_records(js):
            ds = str(find_col(rec, "日期", "Date"))
            cl = str(find_col(rec, "收盤", "Close")).replace(",", "")
            mm = re.match(r"(\d{2,3})/(\d{1,2})/(\d{1,2})", ds)
            try:
                y, mo, da = map(int, mm.groups()); y = y + 1911 if y < 1911 else y
                closes.append((dt.date(y, mo, da), float(cl)))
            except Exception:
                pass
    closes = sorted(set(closes))
    return [c for c in closes if c[0] <= d]

def load_bvps():
    """每股淨值 ≈ 收盤價 / 股價淨值比  (上市 openapi 有收盤價欄；上櫃另抓)"""
    bv = {}
    for rec in (get_json(TWSE_BWIBBU, f"bwibbu_{d8(dt.date.today())}") or []):
        try:
            bv[rec["Code"]] = float(rec["ClosingPrice"]) / float(rec["PBratio"])
        except Exception:
            pass
    for rec in (get_json(TPEX_PBR, f"tpex_pbr_{d8(dt.date.today())}") or []):
        try:
            code = rec.get("SecuritiesCompanyCode") or rec.get("Code")
            pbr = float(rec.get("PriceBookRatio") or rec.get("PBratio"))
            price = float(str(rec.get("ClosingPrice") or rec.get("Close") or "nan").replace(",", ""))
            bv[code] = price / pbr
        except Exception:
            pass
    return bv

# ---------------------------------------------------------------- 主邏輯
def main(base_date=None):
    today = latest_trading_day(base_date)
    days = trading_days(today, LOOKBACK)
    print(f"基準日 {today}，回溯 {len(days)} 個營業日 ({days[0]} ~ {days[-1]})")

    # hist[code][date] = set(clauses)
    hist = defaultdict(dict); info = {}
    for d in days:
        for code, name, mkt, cls, raw in fetch_notice(d):
            hist[code][d] = cls
            info[code] = (name, mkt)
    punished = fetch_punished_now(today)
    print(f"注意股累計 {len(hist)} 檔；目前處置中 {len(punished)} 檔 (排除)")

    idx = {d: i for i, d in enumerate(days)}
    rows = []
    for code, dmap in hist.items():
        if code in punished: continue
        name, mkt = info[code]
        base_days = [d for d in days if d in dmap and dmap[d] & set(range(1, 9))]   # 一~八款
        c1_days   = [d for d in days if d in dmap and 1 in dmap[d]]

        def consec(dlist):   # 到今天為止連續幾天
            n = 0
            for d in reversed(days):
                if d in dlist: n += 1
                else: break
            return n

        cons_all, cons_c1 = consec(base_days), consec(c1_days)
        cnt10 = len([d for d in base_days if idx[d] >= len(days) - 9])   # 今天+前9日 → 明日視窗
        cnt30 = len([d for d in base_days if idx[d] >= len(days) - 29])

        triggers = []
        if cons_c1 == 2:                triggers.append("A 連3日第一款 (今已連2)")
        if cons_all == 4:               triggers.append("B 連5日一~八款 (今已連4)")
        if cnt10 == 5:                  triggers.append("C 10日內6次 (已5次)")
        if cnt30 == 11:                 triggers.append("D 30日內12次 (已11次)")
        if not triggers:
            continue

        today_cls = sorted(dmap.get(today, set()))
        last10 = days[-10:]
        rows.append(dict(代號=code, 名稱=name, 市場=mkt, 今日款別=",".join(map(str, today_cls)),
                         _grid=[{"d": d.strftime("%m/%d"), "hit": d in base_days, "c1": d in c1_days} for d in last10],
                         連續一至八款=cons_all, 連續第一款=cons_c1, 近10日次數=cnt10, 近30日次數=cnt30,
                         觸發條件="；".join(triggers), 需第一款=any(t.startswith("A") for t in triggers)))

    if not rows:
        print("今日無「差一天進處置」候選。"); return

    # 門檻價
    bvps = load_bvps()
    for r in rows:
        r["第一款門檻價"] = r["第六款門檻價(PBR6)"] = r["今日收盤"] = None
        try:
            ch = close_history(r["代號"], r["市場"], today)
            if len(ch) >= 6:
                r["今日收盤"] = ch[-1][1]
                base = ch[-5][1]            # 明日的6日視窗 = 明日+今天+前4日 → 基準為 5 個營業日前收盤
                th = min(base * 1.32, max(base * 1.25, base + 50))
                r["第一款門檻價"] = round(th, 2)
        except Exception as e:
            print(f"  價格抓取失敗 {r['代號']}: {e}")
        if r["代號"] in bvps:
            r["第六款門檻價(PBR6)"] = round(bvps[r["代號"]] * 6, 2)

    df = pd.DataFrame(rows)
    grids = {r["代號"]: r.pop("_grid") for r in rows}
    df = df.drop(columns=["_grid"])
    df["風險等級"] = df["觸發條件"].apply(lambda s: "高" if ("A" in s or "B" in s) else "中")
    df["備註"] = df.apply(lambda r:
        f"明日收盤漲過 {r['第一款門檻價']} 即達第一款" if r["需第一款"] and r["第一款門檻價"]
        else "明日再以一~八款任一款公布注意即進處置", axis=1)
    df = df.drop(columns=["需第一款"]).sort_values(["風險等級", "連續一至八款"], ascending=[True, False])

    pd.set_option("display.width", 200); pd.set_option("display.max_columns", None)
    print(df.to_string(index=False))
    out = f"disposal_candidates_{d8(today)}.xlsx"
    df.to_excel(out, index=False)
    # 給網頁用的 JSON (docs/data.json → GitHub Pages)
    Path("docs").mkdir(exist_ok=True)
    payload = {"base_date": today.isoformat(),
               "generated": dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
               "items": [dict(r, grid=grids[r["代號"]]) for r in df.where(pd.notnull(df), None).to_dict("records")]}
    Path("docs/data.json").write_text(json.dumps(payload, ensure_ascii=False, indent=1), "utf-8")
    print(f"\n已輸出 {out} 與 docs/data.json")

if __name__ == "__main__":
    bd = dt.datetime.strptime(sys.argv[1], "%Y%m%d").date() if len(sys.argv) > 1 else None
    main(bd)
