# 부킹 1건에 걸린 운임 세팅 중 적용할 세팅을 고른다. (조회 전용, 전산 적용은 하지 않음)
# 사용법: python lookup.py <엑셀경로> <부킹번호>
# 판정 규칙
#   1) 세팅 실화주명 == 부킹 실화주명 인 세팅이 정확히 1개 -> 그 세팅 적용
#   2) 위가 없고, 부킹에 걸린 세팅이 '실화주명 빈칸 1개뿐' -> 그 세팅 적용
#   3) 그 외(세팅 없음 / 다른 화주명 세팅이 섞임 / 빈칸 2개 이상 / 실화주 세팅 2개 이상) -> 적용 보류
#   부킹요청자나 다른 화주 세팅은 절대 대신 쓰지 않는다.
import sys
import json
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")


def sheet_rows(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    out = []
    for r in rows[1:]:
        if all(v is None for v in r):
            continue
        out.append({header[i]: r[i] for i in range(len(header))})
    return out


def clean(v):
    return str(v).strip() if v is not None else ""


def fail(msg):
    print(json.dumps({"error": msg}, ensure_ascii=False))
    sys.exit(1)


def main():
    if len(sys.argv) != 3:
        fail("사용법: python lookup.py <엑셀경로> <부킹번호>")

    path, bk_no = sys.argv[1], sys.argv[2].strip().upper()
    wb = load_workbook(path, data_only=True)
    for name in ("부킹목록", "등록운임"):
        if name not in wb.sheetnames:
            fail(f"'{name}' 시트가 없습니다.")

    found = [b for b in sheet_rows(wb["부킹목록"]) if clean(b.get("부킹번호")).upper() == bk_no]
    if not found:
        fail(f"부킹번호 {bk_no}을(를) 부킹목록에서 찾지 못했습니다.")
    if len(found) > 1:
        fail(f"부킹번호 {bk_no}이(가) 부킹목록에 {len(found)}번 있습니다. 원본 확인 필요.")

    b = found[0]
    shipper = clean(b.get("실화주명"))
    qty = b.get("수량")
    booking_info = {k: (clean(v) if v is not None else "[확인 필요]") for k, v in b.items()}

    # 이 부킹에 걸린 세팅 모으기
    settings = {}
    for r in sheet_rows(wb["등록운임"]):
        if clean(r.get("부킹번호")).upper() != bk_no:
            continue
        sno = clean(r.get("세팅번호")) or "[세팅번호 없음]"
        s = settings.setdefault(sno, {"세팅번호": sno, "실화주명": clean(r.get("실화주명")), "운임": []})
        unit = clean(r.get("단위")).upper()
        price = r.get("단가")
        if price is None or price == "":
            price, amount, q = "[확인 필요]", "[확인 필요]", "[확인 필요]"
        elif unit == "BL":
            q, amount = 1, price
        elif unit == "CNTR" and isinstance(qty, (int, float)):
            q, amount = qty, price * qty
        else:
            q, amount = "[확인 필요]", "[확인 필요]"
        s["운임"].append({
            "운임항목": clean(r.get("운임항목")),
            "통화": clean(r.get("통화")) or "[확인 필요]",
            "단가": price,
            "단위": unit or "[확인 필요]",
            "수량": q,
            "금액": amount,
            "등록자": clean(r.get("등록자")) or "[확인 필요]",
        })

    all_settings = list(settings.values())
    summary = [{"세팅번호": s["세팅번호"], "실화주명": s["실화주명"] or "(빈칸)"} for s in all_settings]
    match = [s for s in all_settings if shipper and s["실화주명"] == shipper]
    blank = [s for s in all_settings if not s["실화주명"]]

    chosen, reason = None, ""
    if not shipper:
        reason = "부킹에 실화주명이 없음"
    elif len(match) == 1:
        chosen, reason = match[0], "실화주명이 같은 세팅"
    elif len(match) > 1:
        reason = f"실화주명이 같은 세팅이 {len(match)}건"
    elif not all_settings:
        reason = "부킹에 걸린 운임 세팅이 없음"
    elif len(all_settings) == 1 and len(blank) == 1:
        chosen, reason = blank[0], "실화주명 빈칸 세팅 1건뿐"
    elif len(blank) >= 2:
        reason = f"실화주 세팅 없음, 실화주명 빈칸 세팅이 {len(blank)}건"
    else:
        reason = "실화주 세팅 없음, 다른 화주명 세팅이 걸려 있음"

    result = {
        "부킹": booking_info,
        "실화주명": shipper or "[확인 필요]",
        "판정": "적용 가능" if chosen else "적용 보류",
        "사유": reason,
        "걸린 세팅 목록": summary,
        "적용 세팅": chosen,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
