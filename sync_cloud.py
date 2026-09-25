"""
自動從 Google 雲端試算表同步最新數據並重新產生 data.js 與 data.json
雲端網址: https://docs.google.com/spreadsheets/d/1wut51zEYI7Ij0aBbCx-b7fpEA_n0lhdMht7PtEjnGHY/edit?gid=933525900#gid=933525900
"""

import urllib.request
import urllib.parse
import json
import re
import os

SPREADSHEET_ID = "1wut51zEYI7Ij0aBbCx-b7fpEA_n0lhdMht7PtEjnGHY"

def parse_date(v):
    if v is None:
        return None
    s = str(v).split('.')[0].strip()
    if len(s) == 6 and s.isdigit():
        return f"20{s[0:2]}-{s[2:4]}-{s[4:6]}"
    elif len(s) == 8 and s.isdigit():
        return f"{s[0:4]}-{s[4:6]}-{s[6:8]}"
    return None

def fetch_gviz_sheet(sheet_name):
    url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/gviz/tq?tqx=out:json&sheet=" + urllib.parse.quote(sheet_name)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    raw = urllib.request.urlopen(req).read().decode("utf-8")
    match = re.search(r"google\.visualization\.Query\.setResponse\((.*)\);", raw, re.DOTALL)
    if not match:
        raise ValueError(f"無法解析工作表: {sheet_name}")
    data = json.loads(match.group(1))
    return data["table"]["rows"]

def sync():
    print("正在從 Google 雲端試算表下載最新資料...")

    # 1. 就醫、照顧花費
    rows1 = fetch_gviz_sheet("就醫、照顧花費記帳")
    s1_records = []
    idx1 = 1
    for r in rows1:
        c = r.get("c", [])
        if len(c) < 5 or not c[1]:
            continue
        d = parse_date(c[1].get("v"))
        if not d:
            continue
        raw_item = str(c[2].get("v", "") if len(c) > 2 and c[2] else "").strip()
        amt = float(c[4].get("v", 0) or 0) if len(c) > 4 and c[4] else 0.0
        note = str(c[5].get("v", "") or "") if len(c) > 5 and c[5] else ""
        mom = c[6].get("v") if len(c) > 6 and c[6] else None
        
        notes = [note.strip()] if note.strip() else []
        if mom is not None:
            try:
                amt += float(mom)
                notes.append("老媽")
            except ValueError:
                notes.append(f"老媽: {mom}")

        if any(k in raw_item for k in ["看診", "門診", "回診", "急診"]):
            cat = "看診"
            if raw_item != "看診":
                notes.insert(0, f"原項目: {raw_item}")
        elif any(k in raw_item for k in ["長照"]):
            cat = "長照費"
            if raw_item != "長照費":
                notes.insert(0, f"原項目: {raw_item}")
        elif any(k in raw_item for k in ["住院", "急診預繳"]):
            cat = "住院費"
            if raw_item != "住院費":
                notes.insert(0, f"原項目: {raw_item}")
        elif any(k in raw_item for k in ["看護"]):
            cat = "看護費"
            if raw_item != "看護費":
                notes.insert(0, f"原項目: {raw_item}")
        else:
            cat = "其他"
            if raw_item and raw_item != "其他":
                notes.insert(0, f"原項目: {raw_item}")

        s1_records.append({
            "id": f"med_{idx1}",
            "date": d,
            "item": cat,
            "amount": int(amt) if amt.is_integer() else amt,
            "note": "; ".join(notes)
        })
        idx1 += 1

    # 2. 家用花費、繳款紀錄
    rows2 = fetch_gviz_sheet("家用花費、繳款紀錄")
    s2_records = []
    idx2 = 1
    for r in rows2:
        c = r.get("c", [])
        if len(c) < 6 or not c[1]:
            continue
        d = parse_date(c[1].get("v"))
        if not d:
            continue
        raw_item = str(c[2].get("v", "") if len(c) > 2 and c[2] else "").strip()
        amt = float(c[5].get("v", 0) or 0) if len(c) > 5 and c[5] else 0.0
        note = str(c[6].get("v", "") or "") if len(c) > 6 and c[6] else ""
        notes = [note.strip()] if note.strip() else []

        if "台電" in raw_item:
            cat = "台電"
        elif "中華電信" in raw_item:
            cat = "中華電信"
        elif "瓦斯" in raw_item:
            cat = "瓦斯"
        elif "北水" in raw_item or "水" in raw_item:
            cat = "北水"
        elif "轉帳" in raw_item:
            cat = "房貸轉帳"
        elif "墓園管理費" in raw_item:
            cat = "墓園管理費"
            if "土城" in raw_item:
                notes.insert(0, "土城")
            elif "八里" in raw_item:
                notes.insert(0, "八里")
        else:
            cat = "其他"
            if raw_item and raw_item != "其他":
                notes.insert(0, f"原項目: {raw_item}")

        s2_records.append({
            "id": f"house_{idx2}",
            "date": d,
            "item": cat,
            "amount": int(amt) if amt.is_integer() else amt,
            "note": "; ".join(notes)
        })
        idx2 += 1

    # 3. 老哥轉帳記錄
    rows4 = fetch_gviz_sheet("老哥轉帳記錄")
    s4_records = []
    idx4 = 1
    for r in rows4:
        c = r.get("c", [])
        if len(c) < 7 or not c[1]:
            continue
        d = parse_date(c[1].get("v"))
        if not d:
            continue
        twd = c[6].get("v") if len(c) > 6 and c[6] else None
        note = str(c[7].get("v", "") or "") if len(c) > 7 and c[7] else ""
        if twd is not None and str(twd).strip():
            try:
                amt = float(twd)
                s4_records.append({
                    "id": f"transfer_{idx4}",
                    "date": d,
                    "amount": int(amt) if amt.is_integer() else amt,
                    "note": note.strip()
                })
                idx4 += 1
            except ValueError:
                pass

    data = {
        "medicalCare": s1_records,
        "household": s2_records,
        "transfers": s4_records
    }

    with open("data.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    with open("data.js", "w", encoding="utf-8") as f:
        f.write("// 預設初始歷史資料集 (由 Google 雲端試算表自動同步)\n")
        f.write("window.INITIAL_DATA = ")
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write(";\n")

    print(f"同步完成！")
    print(f"  就醫照顧: {len(s1_records)} 筆, 總額 NT$ {sum(x['amount'] for x in s1_records):,}")
    print(f"  家用繳款: {len(s2_records)} 筆, 總額 NT$ {sum(x['amount'] for x in s2_records):,}")
    print(f"  老哥轉帳: {len(s4_records)} 筆, 總額 NT$ {sum(x['amount'] for x in s4_records):,}")

if __name__ == "__main__":
    sync()
