"""
報帳收據 AI 自動辨識（GitHub Actions 排程執行）

流程：
  1. 從 Firestore 找出還沒辨識過的收據（receipts，status = submitted）
  2. 把 receiptFiles/{id}_{i} 的分段組回原始檔（JPG / PDF）
  3. 交給 Gemini 讀出日期、廠商、金額、幣別、類別、稅號、發票號碼
  4. 只填「空白」或「系統預設」的欄位——人工填過的欄位絕不覆蓋
  5. 在收據上記錄 aiFilled（哪些欄位是 AI 填的）、ocrStatus、ocr（原始辨識結果）

需要的 GitHub Secrets：
  FIREBASE_SERVICE_ACCOUNT  Firebase 服務帳號金鑰（整份 JSON 內容）
  GEMINI_API_KEY            Gemini API 金鑰

手動測試：python scripts/receipt_ocr.py --dry-run   （只印結果，不寫回）
"""
import argparse
import base64
import datetime as dt
import json
import os
import re
import sys
import time

import requests
import firebase_admin
from firebase_admin import credentials, firestore

GEMINI_MODELS = [
    m for m in [
        os.environ.get("GEMINI_MODEL"),
        "gemini-3.5-flash",
        "gemini-3.1-flash-lite",
        "gemini-2.5-flash-lite",
    ] if m
]
GEMINI_BASE = "https://generativelanguage.googleapis.com/v1beta/models"
TIMEOUT = (10, 90)
MAX_PER_RUN = 40          # 每次最多處理幾張，避免一次用光免費額度
MAX_TRIES = 3             # 同一張失敗幾次後放棄，改標「無法辨識」
CATEGORIES = ["transport", "meals", "office", "postage", "entertain", "misc"]
CURRENCIES = ["TRY", "USD", "EUR", "TWD"]

PROMPT = """\
You read expense receipts for an office in Istanbul, Türkiye. Most documents are
Turkish (fiş, e-Arşiv fatura, e-Fatura, taxi/BiTaksi receipts, bank slips), some
are English or Chinese. Extract the fields below from the attached document.

Rules:
- Only report what is printed on the document. Never guess. If a field is not
  visible or not legible, return null for it.
- amount = the final total actually paid (TOPLAM / GENEL TOPLAM / ÖDENECEK TUTAR),
  tax included, as a number with a dot decimal separator (Turkish "1.234,50"
  means 1234.50).
- currency: TRY for TL / ₺ / TRY; otherwise USD, EUR or TWD; null if unclear.
- date: the document/issue date as YYYY-MM-DD (Turkish dates are DD.MM.YYYY).
- vendor: the seller's trade name, short (e.g. "Migros", "BiTaksi", "PTT"),
  not the full legal company title.
- category: one of transport, meals, office, postage, entertain, misc
  (transport = taxi, fuel, parking, tickets; meals = restaurants, cafés, food;
  office = stationery, supplies, printing; postage = cargo, courier, PTT,
  telecom; entertain = gifts, hosting guests; misc = anything else).
- tax_id: seller VKN/TCKN if printed; invoice_no: fiş/fatura number if printed.
- is_receipt: false if the image is not a receipt or invoice at all.
- confidence: 0 to 1, how sure you are about amount and date together.

Return JSON only:
{"is_receipt": bool, "date": str|null, "vendor": str|null, "amount": number|null,
 "currency": str|null, "category": str|null, "tax_id": str|null,
 "invoice_no": str|null, "confidence": number, "note": str|null}
"""


def log(msg):
    print(msg, flush=True)


def init_db():
    raw = os.environ.get("FIREBASE_SERVICE_ACCOUNT")
    if not raw:
        # 還沒設定好 Secrets 時安靜結束，不要讓排程每 15 分鐘寄一封失敗通知
        log("尚未設定 FIREBASE_SERVICE_ACCOUNT，略過本次執行")
        sys.exit(0)
    cred = credentials.Certificate(json.loads(raw))
    firebase_admin.initialize_app(cred)
    return firestore.client()


def load_file(db, rid, chunks):
    parts = []
    for i in range(chunks or 1):
        snap = db.collection("receiptFiles").document(f"{rid}_{i}").get()
        if not snap.exists:
            raise ValueError(f"缺少檔案分段 {rid}_{i}")
        parts.append(snap.to_dict().get("data", ""))
    data_url = "".join(parts)
    m = re.match(r"^data:([^;]+);base64,(.*)$", data_url, re.S)
    if not m:
        raise ValueError("檔案格式不是 data URL")
    return m.group(1), m.group(2)          # mime, base64 字串


def parse_json(text):
    text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text.strip())
    return json.loads(text)


def call_gemini(mime, b64, api_key):
    body = {
        "contents": [{"parts": [
            {"inline_data": {"mime_type": mime, "data": b64}},
            {"text": PROMPT},
        ]}],
        "generationConfig": {"responseMimeType": "application/json", "temperature": 0},
    }
    last = None
    for model in GEMINI_MODELS:
        for attempt in range(1, 4):
            try:
                r = requests.post(f"{GEMINI_BASE}/{model}:generateContent",
                                  params={"key": api_key}, json=body, timeout=TIMEOUT)
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                last = f"{model}: {type(e).__name__}"      # 不印網址，避免洩漏金鑰
                time.sleep(2 ** attempt)
                continue
            if r.status_code in (404, 429):
                last = f"{model}: HTTP {r.status_code}"
                break                                       # 換下一個模型
            if r.status_code >= 500 or r.status_code == 408:
                last = f"{model}: HTTP {r.status_code}"
                time.sleep(2 ** attempt)
                continue
            if r.status_code != 200:
                last = f"{model}: HTTP {r.status_code} {r.text[:200]}"
                break
            try:
                text = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                return parse_json(text), model
            except (KeyError, IndexError, json.JSONDecodeError) as e:
                last = f"{model}: 回應無法解析（{type(e).__name__}）"
                continue
    raise RuntimeError(last or "Gemini 沒有回應")


def clean(res):
    """把模型回傳整理成可以直接寫進 Firestore 的值，不合格的一律當 null。"""
    out = {}
    d = res.get("date")
    if isinstance(d, str) and re.match(r"^\d{4}-\d{2}-\d{2}$", d):
        try:
            day = dt.date.fromisoformat(d)
            if dt.date(2020, 1, 1) <= day <= dt.date.today() + dt.timedelta(days=1):
                out["date"] = d
        except ValueError:
            pass
    v = res.get("vendor")
    if isinstance(v, str) and v.strip():
        out["vendor"] = v.strip()[:60]
    a = res.get("amount")
    if isinstance(a, (int, float)) and 0 < a < 10_000_000:
        out["amount"] = round(float(a), 2)
    c = res.get("currency")
    if c in CURRENCIES:
        out["currency"] = c
    cat = res.get("category")
    if cat in CATEGORIES:
        out["category"] = cat
    for k in ("tax_id", "invoice_no"):
        if isinstance(res.get(k), str) and res[k].strip():
            out[k] = res[k].strip()[:40]
    return out


def plan_update(r, got):
    """決定哪些欄位可以由 AI 填：只填空白欄位或系統自動給的預設值。"""
    patch, filled = {}, []
    # 日期：同仁上傳時系統自動填「上傳當天」（dateSource = upload），可以換成收據上的日期
    if "date" in got and (not r.get("date") or r.get("dateSource", "upload") == "upload"):
        if got["date"] != r.get("date"):
            patch["date"] = got["date"]
        patch["dateSource"] = "ai"
        filled.append("date")
    if "vendor" in got and not r.get("vendor"):
        patch["vendor"] = got["vendor"]; filled.append("vendor")
    amount_empty = r.get("amount") in (None, "")
    if "amount" in got and amount_empty:
        patch["amount"] = got["amount"]; filled.append("amount")
        # 金額是 AI 填的，幣別預設值 TRY 也一起換成收據上的幣別
        if "currency" in got:
            patch["currency"] = got["currency"]; filled.append("currency")
    if "category" in got and not r.get("category"):
        patch["category"] = got["category"]; filled.append("category")
    for k in ("tax_id", "invoice_no"):
        if k in got and not r.get(k):
            patch[k] = got[k]
    return patch, filled


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="只印出辨識結果，不寫回 Firestore")
    ap.add_argument("--redo", help="指定收據 id 重新辨識")
    args = ap.parse_args()

    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        log("尚未設定 GEMINI_API_KEY，略過本次執行")
        sys.exit(0)
    db = init_db()

    if args.redo:
        snap = db.collection("receipts").document(args.redo).get()
        todo = [snap] if snap.exists else []
    else:
        todo = [s for s in db.collection("receipts").where("status", "==", "submitted").stream()
                if (s.to_dict().get("ocrStatus") in (None, "pending"))
                and (s.to_dict().get("ocrTries") or 0) < MAX_TRIES]
    todo = todo[:MAX_PER_RUN]
    log(f"待辨識：{len(todo)} 張")

    done = failed = 0
    for snap in todo:
        r, rid = snap.to_dict(), snap.id
        ref = db.collection("receipts").document(rid)
        try:
            mime, b64 = load_file(db, rid, r.get("chunks"))
            res, model = call_gemini(mime, b64, api_key)
            got = clean(res)
            conf = res.get("confidence") if isinstance(res.get("confidence"), (int, float)) else None
            if res.get("is_receipt") is False:
                patch, filled, status = {}, [], "unreadable"
            else:
                patch, filled = plan_update(r, got)
                status = "done" if ("amount" in got or "vendor" in got) else "unreadable"
            ocr = {"model": model, "at": firestore.SERVER_TIMESTAMP, "confidence": conf,
                   "raw": {k: res.get(k) for k in ("date", "vendor", "amount", "currency",
                                                    "category", "tax_id", "invoice_no", "note")}}
            update = {**patch, "ocrStatus": status, "ocr": ocr,
                      "aiFilled": sorted(set((r.get("aiFilled") or []) + filled))}
            log(f"  {rid}: {status} · 填入 {filled or '無'} · 信心 {conf} · {model}")
            if not args.dry_run:
                ref.update(update)
            done += 1
        except Exception as e:  # 單張失敗不影響其他張
            tries = (r.get("ocrTries") or 0) + 1
            log(f"  {rid}: 失敗（第 {tries} 次）{type(e).__name__}: {str(e)[:200]}")
            if not args.dry_run:
                ref.update({"ocrTries": tries,
                            "ocrStatus": "failed" if tries >= MAX_TRIES else "pending",
                            "ocrError": str(e)[:300]})
            failed += 1
    log(f"完成：成功 {done} 張，失敗 {failed} 張")


if __name__ == "__main__":
    main()
