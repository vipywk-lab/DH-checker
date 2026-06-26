# ==========================================
# bx_checker.py
# 버전: v3.0 (2026-06-25)
# 문의: 승무계획팀
#
# [지원 항공사]
#   에어부산 / 대한항공 / 진에어 / 파라타항공  → 자동조회
#   티웨이항공                                 → 반자동 (아래 참고)
#
# [티웨이항공 안내]
#   티웨이 홈페이지는 Akamai Bot Manager 보안이 적용되어
#   자동조회가 불가합니다.
#   대신 다음과 같이 처리됩니다:
#     1. 티웨이 예약이 있으면 조회 완료 후 건별로 Chrome 탭이 자동으로 열립니다.
#     2. 팝업창이 뜨며 예약번호·성·이름 복사 버튼이 제공됩니다.
#     3. 복사 버튼 클릭 → Chrome 탭에서 해당 칸에 Ctrl+V 하면 됩니다.
#     4. [다음 →] 버튼으로 건별로 순서대로 처리합니다.
# ==========================================
__version__ = "3.1"
VERSION_URL  = "https://raw.githubusercontent.com/vipywk-lab/DH-checker/main/bx_checker.py"

import asyncio
import random
import webbrowser
import re
import logging
import os
import subprocess
from datetime import datetime, timedelta
import openpyxl
from tkinter import filedialog, messagebox
import tkinter as tk
from playwright.async_api import async_playwright, TimeoutError as PWTimeout
try:
    from playwright_stealth import Stealth
    STEALTH_AVAILABLE = True
except ImportError:
    STEALTH_AVAILABLE = False


# ==========================================
# Playwright Chromium 최초 1회 자동 설치
# ==========================================
chromium_path = os.path.expanduser("~\\AppData\\Local\\ms-playwright")
if not os.path.exists(chromium_path):
    print("기반 시스템(브라우저)을 설치 중입니다. 최초 1회만 진행되며 시간이 조금 걸릴 수 있습니다...")
    try:
        subprocess.run(["playwright", "install", "chromium"], check=True)
        print("설치 완료!\n")
    except Exception as e:
        print(f"설치 중 오류가 발생했습니다: {e}")

# ==========================================
# 로깅 설정 (오류 발생 시 텍스트 파일로 저장)
# ==========================================
logging.basicConfig(
    filename=f"에러로그_{datetime.now().strftime('%Y%m%d')}.txt",
    level=logging.ERROR,
    format='%(asctime)s [%(levelname)s] %(message)s',
    encoding='utf-8'
)

root = tk.Tk()
root.withdraw()
messagebox.showinfo("안내", "xlsm 파일을 선택해주세요")
EXCEL_PATH = filedialog.askopenfilename(
    title="xlsm 파일 선택",
    filetypes=[("Excel files", "*.xlsm *.xlsx")]
)
if not EXCEL_PATH:
    raise SystemExit("파일을 선택하지 않았습니다.")

SHEET_NAME = "검증대상"
RESULT_COL = 8
DETAIL_COL = 9
BX_URL     = "https://www.airbusan.com/web/individual/reserve/index"
KE_URL     = "https://www.koreanair.com/reservation/search"
LJ_URL     = "https://www.jinair.com/booking/index"
WE_URL     = "https://www.parataair.com/ko/login/viewLogin.do?tab=2#"
TW_URL     = "https://www.twayair.com/app/reservation/searchMemberBooking#none"
HEADLESS   = False
DELAY_MIN  = 1.0
DELAY_MAX  = 2.0

DOMESTIC_AIRPORTS = {"PUS","CJU","TAE","CJJ","HIN","RSU","KPO","MWX","GMP","ICN"}


def check_for_update():
    """GitHub raw URL에서 최신 버전 확인 후 콘솔 알림"""
    import urllib.request
    try:
        with urllib.request.urlopen(VERSION_URL, timeout=5) as resp:
            for line in resp.read().decode("utf-8").splitlines():
                if line.startswith("__version__"):
                    latest = line.split("=")[1].strip().strip('"').strip("'")
                    def _ver(v):
                        try:
                            return tuple(int(x) for x in v.split("."))
                        except Exception:
                            return (0,)
                    if _ver(latest) > _ver(__version__):
                        print(f"\n{'!'*50}")
                        print(f"  ⚠️  업데이트 필요: 현재 v{__version__} → 최신 v{latest}")
                        print(f"  아래 링크에서 최신 버전을 받아주세요:")
                        print(f"  https://github.com/vipywk-lab/DH-checker")
                        print(f"{'!'*50}\n")
                        input("업데이트 후 다시 실행해주세요. 엔터 누르면 종료...")
                        raise SystemExit("구버전 실행 차단")
                    else:
                        print(f"✅ 최신 버전입니다 (v{__version__})")
                    return
    except Exception:
        print("⚠️  버전 확인 실패 (네트워크 연결 없음 — 무시하고 계속 진행)\n")


def get_check_mode():
    """실행 시 조회 범위 선택 팝업 — 3가지 옵션"""
    import calendar
    from tkinter import simpledialog
    today    = datetime.now()
    # 이번달 말일
    this_last = calendar.monthrange(today.year, today.month)[1]
    days_to_eom = this_last - today.day
    # 다음달 말일
    next_month = today.month % 12 + 1
    next_year  = today.year + (1 if today.month == 12 else 0)
    next_last  = calendar.monthrange(next_year, next_month)[1]
    days_to_nom = days_to_eom + next_last  # 이번달 잔여 + 다음달 전체

    choice = simpledialog.askstring(
        "조회 범위 선택",
        f"조회 범위를 입력하세요.\n\n"
        f"  1  →  오늘부터 5일 (일상 조회)\n"
        f"  2  →  이번달 말일까지 ({today.month}월 {this_last}일)\n"
        f"  3  →  다음달 말일까지 ({next_year}년 {next_month}월 {next_last}일)\n\n"
        f"※ 2·3번은 딜레이가 자동으로 늘어납니다.",
        parent=root
    )

    if choice == "2":
        return days_to_eom, 3.0, 6.0
    elif choice == "3":
        return days_to_nom, 3.0, 6.0
    else:  # 1 또는 기타
        return 5, 1.0, 2.0


def split_korean_name(name):
    name = name.strip()
    # 동명이인 구분자 제거: 이경수A → 이경수, 박지연B → 박지연
    name = re.sub(r'[A-Za-z]+$', '', name).strip()
    return name[0], name[1:]


def parse_dep_date(dep_time_str):
    s = str(dep_time_str).strip()
    try:
        return datetime.strptime(s[:7], "%d%b%y")
    except:
        pass
    try:
        return datetime.strptime(s[:9], "%d-%b-%y")
    except:
        pass
    try:
        n = float(s)
        if 40000 < n < 60000:
            return datetime(1899, 12, 30) + timedelta(days=int(n))
    except:
        pass
    return None


def is_within_check_range(dep_time_str, end_date):
    dep_date = parse_dep_date(str(dep_time_str))
    if not dep_date:
        return True
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    return today <= dep_date <= end_date


def is_international(dep, arr):
    return dep not in DOMESTIC_AIRPORTS or arr not in DOMESTIC_AIRPORTS


def load_targets(path, sheet, end_date):
    wb = openpyxl.load_workbook(path, keep_vba=True)
    if sheet not in wb.sheetnames:
        messagebox.showerror(
            "시트 없음",
            f"선택한 파일에 [{sheet}] 시트가 없습니다.\n\n"
            f"① DH_자동화.xlsm에서 VBA 매크로를 먼저 실행해 [{sheet}] 시트를 생성하세요.\n"
            f"② 올바른 파일을 선택했는지 확인하세요.\n\n"
            f"선택한 파일: {os.path.basename(path)}\n"
            f"현재 시트 목록: {', '.join(wb.sheetnames)}"
        )
        raise SystemExit(f"[{sheet}] 시트 없음 → 프로그램 종료")
    ws = wb[sheet]
    targets = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        kor_name, airline, pnr, dep, arr, dep_time, eng_name = (list(row) + [None]*7)[:7]
        if not all([kor_name, airline, pnr]):
            continue
        if airline not in ("에어부산", "대한항공", "진에어", "파라타항공", "티웨이항공"):
            continue
        if not re.match(r'^[A-Z0-9]{6}$', str(pnr).strip().upper()):
            continue
        if not is_within_check_range(str(dep_time or ""), end_date):
            continue
        last, first = split_korean_name(str(kor_name))
        targets.append({
            "kor_name": str(kor_name),
            "airline" : str(airline),
            "last"    : last,
            "first"   : first,
            "eng_name": str(eng_name).strip().replace("-", " ") if eng_name else "",
            "pnr"     : str(pnr).strip().upper(),
            "dep"     : str(dep or ""),
            "arr"     : str(arr or ""),
            "dep_time": str(dep_time or ""),
            "result"  : None,
            "detail"  : None,
        })
    return targets


def save_results(path, sheet, targets):
    """PNR 기반으로 정확하게 매칭해서 저장 + 확인필요 요약 시트 생성"""
    wb = openpyxl.load_workbook(path, keep_vba=True)
    ws = wb[sheet]
    ws.cell(1, RESULT_COL).value = "검증결과"
    ws.cell(1, DETAIL_COL).value = "조회내용"

    # PNR + 항공사 + 이름 조합으로 정확히 매칭
    result_map = {}
    for t in targets:
        key = (t["pnr"], t["airline"], t["kor_name"])
        result_map[key] = (t["result"], t["detail"])

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        airline  = row[1].value
        pnr      = str(row[2].value).strip().upper() if row[2].value else ""
        kor_name = str(row[0].value).strip() if row[0].value else ""
        if airline not in ("에어부산", "대한항공", "진에어", "파라타항공", "티웨이항공"):
            continue
        key = (pnr, airline, kor_name)
        if key in result_map:
            result, detail = result_map[key]
            ws.cell(row_idx, RESULT_COL).value = result
            ws.cell(row_idx, DETAIL_COL).value = detail

    # ── 확인필요 요약 시트 생성 ──
    summary_name = "확인필요_요약"
    if summary_name in wb.sheetnames:
        del wb[summary_name]

    ws_summary = wb.create_sheet(summary_name)
    ws_summary.append(["탑승객", "항공사", "PNR", "검증결과", "조회내용"])
    ws_summary.row_dimensions[1].height = 20

    error_count = 0
    for t in targets:
        res = str(t["result"])
        if any(kw in res for kw in ["불일치", "예약없음", "PNR오류", "오류", "타임아웃", "수동확인필요"]):
            ws_summary.append([
                t["kor_name"],
                t["airline"],
                t["pnr"],
                t["result"],
                t["detail"]
            ])
            error_count += 1

    if error_count == 0:
        ws_summary.append(["✅ 모든 예약이 정상적으로 확인되었습니다!"])
    else:
        # 헤더 볼드
        from openpyxl.styles import Font
        for cell in ws_summary[1]:
            cell.font = Font(bold=True)

    ws_summary.column_dimensions["A"].width = 12
    ws_summary.column_dimensions["B"].width = 12
    ws_summary.column_dimensions["C"].width = 10
    ws_summary.column_dimensions["D"].width = 15
    ws_summary.column_dimensions["E"].width = 40

    # ── 저장 (엑셀 열려있으면 팝업 안내) ──
    while True:
        try:
            wb.save(path)
            print(f"\n저장 완료: {path}")
            print(f"→ [확인필요_요약] 시트에서 {error_count}건 확인하세요!" if error_count > 0 else "→ 모든 예약 정상!")
            break
        except (PermissionError, Exception) as e:
            messagebox.showerror(
                "저장 오류",
                f"엑셀 파일 저장에 실패했습니다!\n\n"
                f"파일이 열려있다면 닫고 '확인'을 눌러주세요.\n\n"
                f"{path}\n\n오류: {str(e)[:80]}"
            )


async def check_bx(page, target):
    pnr      = target["pnr"]
    eng_name = target.get("eng_name", "")
    dep      = target["dep"]
    arr      = target["arr"]

    intl = is_international(dep, arr)

    if intl and eng_name:
        parts = eng_name.split("/")
        last  = parts[0].strip() if len(parts) >= 1 else target["last"]
        first = parts[1].strip() if len(parts) >= 2 else target["first"]
    else:
        last  = target["last"]
        first = target["first"]
    #print(f"  → [BX 디버그] PNR:{pnr} | 국제선:{intl} | 성:{last} | 이름:{first}")
    try:
        await page.goto(BX_URL, wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(1500)

        # 클라우드플레어 감지 → 사람이 직접 캡챠 풀도록 안내
        CF_KEYWORDS = ["보안 확인 수행 중", "사람인지 확인하십시오", "Checking your browser",
                       "DDoS protection", "보안 서비스", "악의적인 봇", "Cloudflare",
                       "Just a moment", "Enable JavaScript", "cf-browser-verification",
                       "Verify you are human", "Security check"]
        body_check = await page.inner_text("body")
        if any(kw in body_check for kw in CF_KEYWORDS):
            print(f"\n{'='*50}")
            print(f"  ⚠️  [에어부산] 클라우드플레어 보안 확인이 필요합니다!")
            print(f"  → 열린 브라우저에서 '사람인지 확인하십시오' 체크박스를 클릭해주세요.")
            print(f"  → 완료 후 여기서 엔터를 눌러주세요.")
            print(f"{'='*50}")
            await asyncio.get_event_loop().run_in_executor(None, input, "  [확인 후 엔터] ")
            # 통과됐는지 재확인
            body_check2 = await page.inner_text("body")
            if any(kw in body_check2 for kw in CF_KEYWORDS):
                return "⏱️ 타임아웃", "클라우드플레어 차단 미해제 → 재실행 필요"

        await page.click("text=예약번호로 조회", timeout=5000)
        await page.wait_for_timeout(800)

        await page.locator("input[placeholder*='예약번호']").first.fill(pnr)
        await page.wait_for_timeout(300)
        await page.locator("input[placeholder='성']").first.fill(last)
        await page.wait_for_timeout(300)
        await page.locator("input[placeholder='이름']").first.fill(first)
        await page.wait_for_timeout(300)

        await page.evaluate("document.querySelector('.buttonOfflineCheckin').click()")
        await page.wait_for_timeout(2000)

        body_text = await page.inner_text("body")
        if "해당 예약번호가 확인되지 않습니다" in body_text:
            try:
                await page.click("button:has-text('확인')", timeout=2000)
            except:
                pass
            return "❌ PNR오류", "해당 예약번호 확인 불가"

        try:
            await page.wait_for_selector("text=항공권 구매완료", timeout=10000)
        except:
            pass
        await page.wait_for_timeout(1000)

        html_content = await page.inner_text("body")

        # 결과 파싱 전 클라우드플레어 재체크
        if any(kw in html_content for kw in CF_KEYWORDS):
            return "⏱️ 타임아웃", "클라우드플레어 차단 → 재실행 필요"
        if any(kw in html_content for kw in ["조회 결과가 없", "예약 내역이 없", "일치하는 예약"]):
            return "❌ 예약없음", "조회결과 없음"

        flt_match   = re.search(r'BX\s*\d{3,4}', html_content)
        date_match  = re.search(r'(\d{4}-\d{2}-\d{2})', html_content)

        # 구간: 한글 도시명 (국내 + 에어부산 국제선 취항지)
        BX_CITIES = (
            '부산|서울|김포|제주|대구|광주|청주|인천'
            '|후쿠오카|오사카|삿포로|도쿄|나리타'
            '|다낭|나트랑|냐짱|보홀|세부|칼리보|비엔티안|치앙마이'
            '|타이베이|가오슝|홍콩|마카오|칭다오|옌지|장자제|시안|상하이'
        )
        route_match = re.search(
            rf'({BX_CITIES}).{{1,15}}({BX_CITIES})',
            html_content
        )

        if route_match:
            route_found = route_match.group()
        else:
            # fallback: 공항코드로 매칭 (한글 도시명 미표시 대비)
            codes = re.findall(
                r'(?<![A-Z0-9])(PUS|GMP|ICN|CJU|TAE|CJJ|HIN|RSU|KPO|MWX'
                r'|FUK|KIX|CTS|NRT|HND|NGO'
                r'|DAD|CXR|TAG|CEB|KLO|VTE|CNX'
                r'|TPE|KHH|HKG|MFM|TAO|YNJ|DYG|XIY|PVG)(?![A-Z0-9])',
                html_content
            )
            route_found = f"{codes[0]}→{codes[1]}" if len(codes) >= 2 else "구간미확인"

        flt_found   = flt_match.group().replace(" ", "") if flt_match else "편명미확인"
        date_found  = date_match.group() if date_match else "날짜미확인"

        detail = f"{flt_found} | {date_found} | {route_found}"

        dep_date = parse_dep_date(target["dep_time"])
        mismatch = []
        if dep_date and date_found != "날짜미확인":
            try:
                site_date = datetime.strptime(date_found, "%Y-%m-%d")
                if dep_date.date() != site_date.date():
                    mismatch.append(
                        f"날짜불일치(PDC:{dep_date.strftime('%m/%d')} vs 사이트:{site_date.strftime('%m/%d')})"
                    )
            except:
                pass

        if mismatch:
            return "⚠️ 불일치", detail + " | " + " / ".join(mismatch)

        if flt_found == "편명미확인" and date_found == "날짜미확인":
            return "💥 오류", "파싱 실패 (CF 차단 또는 페이지 미로딩)"

        return "✅ 확인완료", detail

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception as e:
        logging.error(f"에어부산 조회 실패 | PNR: {pnr} | 탑승객: {last}{first}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"


async def check_ke(page, target):
    pnr      = target["pnr"]
    eng_name = target.get("eng_name", "")
    dep      = target["dep"]
    arr      = target["arr"]

    # 국내선=한글, 국제선=영문 (BX/LJ와 동일)
    intl = is_international(dep, arr)
    if intl and eng_name:
        parts = eng_name.split("/")
        last  = parts[0].strip() if len(parts) >= 1 else target["last"]
        first = parts[1].strip() if len(parts) >= 2 else target["first"]
    else:
        last  = target["last"]
        first = target["first"]

    dep_date = parse_dep_date(target["dep_time"])
    if not dep_date:
        return "💥 오류", "출발일 파싱 실패"

    try:
        await page.goto(KE_URL, wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(2000)

        try:
            await page.click("button:has-text('동의합니다')", timeout=3000)
            await page.wait_for_timeout(500)
        except:
            pass

        await page.locator("input[maxlength='13']").first.fill(pnr)
        await page.wait_for_timeout(300)

        await page.click("button[data-dialog-id='#dialog-datepicker1']", timeout=5000)
        await page.wait_for_timeout(1500)

        dep_day   = str(dep_date.day)
        dep_month = dep_date.month
        dep_year  = dep_date.year

        await page.evaluate(f"""
            (function() {{
                var tds = document.querySelectorAll('td.datepicker__td.-available[role="button"]');
                for (var td of tds) {{
                    var span = td.querySelector('span[aria-hidden="true"]');
                    if (!span) continue;
                    if (span.textContent.trim() !== '{dep_day}') continue;
                    var container = td.closest('[id*="month{dep_year}{dep_month:02d}"]');
                    if (container) {{
                        td.click();
                        return;
                    }}
                }}
                for (var td of tds) {{
                    var span = td.querySelector('span[aria-hidden="true"]');
                    if (span && span.textContent.trim() === '{dep_day}') {{
                        td.click();
                        return;
                    }}
                }}
            }})();
        """)
        await page.wait_for_timeout(800)

        await page.locator("input[autocomplete='family-name']").first.fill(last)
        await page.wait_for_timeout(300)
        await page.locator("input[autocomplete='given-name']").first.fill(first)
        await page.wait_for_timeout(300)

        await page.click("button:has-text('조회')", timeout=5000)

        try:
            await page.wait_for_selector(".journey-info__date", timeout=20000)
        except:
            pass
        await page.wait_for_timeout(2000)

        html_content = await page.inner_text("body")

        if any(kw in html_content for kw in ["조회 결과가 없", "예약을 찾을 수 없", "확인되지 않", "일치하는 예약"]):
            return "❌ PNR오류", "예약 확인 불가"

        flt_match = re.search(r'KE\s*\d{3,4}', html_content)
        flt_found = flt_match.group().replace(" ", "") if flt_match else "편명미확인"

        date_match = re.search(r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', html_content)
        if date_match:
            y = date_match.group(1)
            m = date_match.group(2).zfill(2)
            d = date_match.group(3).zfill(2)
            date_found = f"{y}-{m}-{d}"
        else:
            date_found = "날짜미확인"

        airports = re.findall(
            r'\b(PUS|GMP|ICN|CJU|TAE|CJJ|HIN|RSU|KPO|MWX'
            r'|NRT|HND|KIX|NGO|FUK|CTS|OKA'
            r'|BKK|CNX|HKT|SGN|HAN|DAD|CXR|MNL|CEB|CRK'
            r'|TPE|HKG|MFM|SIN|KUL|PNH|REP|VTE|RGN|DPS|PQC'
            r'|PEK|PVG|CAN|TAO|SZX|WEH|YNJ|HRB|SHE)\b',
            html_content
        )
        if len(airports) >= 2:
            route_found = f"{airports[0]}→{airports[1]}"
        else:
            route_found = "구간미확인"

        detail = f"{flt_found} | {date_found} | {route_found}"

        mismatch = []
        if dep_date and date_found != "날짜미확인":
            try:
                site_date = datetime.strptime(date_found, "%Y-%m-%d")
                if dep_date.date() != site_date.date():
                    mismatch.append(
                        f"날짜불일치(PDC:{dep_date.strftime('%m/%d')} vs 사이트:{site_date.strftime('%m/%d')})"
                    )
            except:
                pass

        if mismatch:
            return "⚠️ 불일치", detail + " | " + " / ".join(mismatch)

        return "✅ 확인완료", detail

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception as e:
        logging.error(f"대한항공 조회 실패 | PNR: {pnr} | 탑승객: {last}{first}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"


async def check_lj(page, target):
    pnr      = target["pnr"]
    last     = target["last"]
    first    = target["first"]
    eng_name = target.get("eng_name", "")
    dep      = target["dep"]
    arr      = target["arr"]

    dep_date = parse_dep_date(target["dep_time"])
    if not dep_date:
        return "💥 오류", "출발일 파싱 실패"

    intl = is_international(dep, arr)

    if intl and eng_name:
        parts       = eng_name.split("/")
        input_last  = parts[0].strip() if len(parts) >= 1 else last
        input_first = parts[1].strip() if len(parts) >= 2 else first
    else:
        input_last  = last
        input_first = first

    try:
        await page.goto(LJ_URL, wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(2000)

        await page.click("text=예약조회", timeout=5000)
        await page.wait_for_timeout(800)

        await page.locator("input[placeholder*='6자리']").first.fill(pnr)
        await page.wait_for_timeout(300)

        await page.locator("#lastName_resv").fill(input_last)
        await page.wait_for_timeout(300)

        await page.locator("#firstName_resv").fill(input_first)
        await page.wait_for_timeout(300)

        await page.click("#departureDate_resv", timeout=5000)
        await page.wait_for_timeout(2000)

        dep_str = dep_date.strftime("%Y.%m.%d")
        iframe_locator = page.frame_locator("iframe[src*='basicCalendarLayer']")

        try:
            date_input = iframe_locator.locator(f"input[name='hiddenDate'][value^='{dep_str}']")
            await date_input.evaluate("el => el.parentElement.click()")
        except:
            pass

        await page.wait_for_timeout(800)
        await page.click("button[role='login-button']", timeout=5000)

        try:
            await page.wait_for_selector(f"text={pnr}", timeout=15000)
        except:
            pass
        await page.wait_for_timeout(2000)

        html_content = await page.inner_text("body")

        if pnr not in html_content:
            return "💥 오류", "결과 페이지 로드 실패"

        if any(kw in html_content for kw in ["조회 결과가 없", "예약 내역이 없", "확인되지 않"]):
            return "❌ PNR오류", "예약 확인 불가"

        flt_match = re.search(r'LJ\d{3,4}', html_content)
        flt_found = flt_match.group() if flt_match else "편명미확인"

        date_match = re.search(r'(\d{4})\.(\d{2})\.(\d{2})\(', html_content)
        if date_match:
            y = date_match.group(1)
            m = date_match.group(2).zfill(2)
            d = date_match.group(3).zfill(2)
            date_found = f"{y}-{m}-{d}"
        else:
            date_found = "날짜미확인"

        airports = re.findall(
            r'\b(PUS|GMP|ICN|CJU|TAE|CJJ|HIN|RSU|KPO|MWX'
            r'|CNX|BKK|HKT|NRT|HND|KIX|NGO|CTS|FUK|OKA'
            r'|DAD|SGN|HAN|CXR|PQC|CEB|KLO|TAG|MNL'
            r'|TPE|HKG|MFM|SIN|DPS|GUM)\b',
            html_content
        )
        if len(airports) >= 2:
            route_found = f"{airports[0]}→{airports[1]}"
        else:
            route_found = "구간미확인"

        detail = f"{flt_found} | {date_found} | {route_found}"

        mismatch = []
        if dep_date and date_found != "날짜미확인":
            try:
                site_date = datetime.strptime(date_found, "%Y-%m-%d")
                if dep_date.date() != site_date.date():
                    mismatch.append(
                        f"날짜불일치(PDC:{dep_date.strftime('%m/%d')} vs 사이트:{site_date.strftime('%m/%d')})"
                    )
            except:
                pass

        if mismatch:
            return "⚠️ 불일치", detail + " | " + " / ".join(mismatch)

        return "✅ 확인완료", detail

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception as e:
        logging.error(f"진에어 조회 실패 | PNR: {pnr} | 탑승객: {input_last}{input_first}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"


async def check_we(page, target, we_email):
    pnr      = target["pnr"]
    kor_name = target["kor_name"]

    if not we_email:
        return "⚠️ 수동확인필요", "파라타항공-이메일 미입력"

    dep_date = parse_dep_date(target["dep_time"])
    # dep_date 파싱 실패해도 조회는 진행 (날짜 불일치 검사만 스킵)

    # alert 팝업 대비 (틀린 PNR 등) - 메시지 캡처 후 자동 닫기
    dialog_msgs = []
    async def _on_dialog(dialog):
        dialog_msgs.append(dialog.message)
        await dialog.accept()
    page.on("dialog", _on_dialog)

    try:
        await page.goto(WE_URL, wait_until="domcontentloaded", timeout=20000)
        await page.wait_for_timeout(1500)

        # 홈페이지 비회원 탭 클릭
        await page.click("a[href='#nonmember']", timeout=5000)
        await page.wait_for_timeout(800)

        # 이메일 + 예약번호 입력
        await page.fill("#userEmail", we_email)
        await page.wait_for_timeout(300)
        await page.fill("#reservationNum", pnr)
        await page.wait_for_timeout(300)

        # 예약조회 버튼 클릭
        await page.click("#nonMemberResvSearchBtn", timeout=5000)

        # 결과 페이지 대기
        try:
            await page.wait_for_url("**/viewReservationDetail.do**", timeout=15000)
        except:
            pass
        await page.wait_for_timeout(2000)

        current_url = page.url
        html_content = await page.inner_text("body")

        # 실패 판정
        if "viewReservationDetail.do" not in current_url:
            if dialog_msgs:
                return "❌ PNR오류", f"알림: {dialog_msgs[0][:40]}"
            if any(kw in html_content for kw in ["일치하는 예약", "확인되지 않", "조회 결과가 없", "예약 내역이 없"]):
                return "❌ PNR오류", "예약 확인 불가"
            return "💥 오류", "결과 페이지 이동 실패"

        # 편명 파싱 (WE208 형태)
        flt_match = re.search(r'WE\s*\d{3,4}', html_content)
        flt_found = flt_match.group().replace(" ", "") if flt_match else "편명미확인"

        # 날짜 파싱: "출발일시" 라벨 뒤의 날짜만 (예약일 오인식 방지)
        date_match = re.search(r'출발일시[\s\S]{0,20}?(\d{4})\.(\d{2})\.(\d{2})', html_content)
        if not date_match:
            # fallback: 요일괄호가 바로 붙은 날짜 (출발일시 형식: 2026.06.15(월))
            date_match = re.search(r'(\d{4})\.(\d{2})\.(\d{2})\(', html_content)
        if date_match:
            date_found = f"{date_match.group(1)}-{date_match.group(2)}-{date_match.group(3)}"
        else:
            date_found = "날짜미확인"

        # 구간 파싱 (한글이 코드에 바로 붙어도 매칭: 푸꾸옥PQC서울 등)
        airports = re.findall(
            r'(?<![A-Z0-9])(PUS|GMP|ICN|CJU|TAE|CJJ|HIN|RSU|KPO|MWX|PQC|RGN|DPS|MFM|HKG)(?![A-Z0-9])',
            html_content
        )
        route_found = f"{airports[0]}→{airports[1]}" if len(airports) >= 2 else "구간미확인"

        detail = f"{flt_found} | {date_found} | {route_found}"

        # 날짜 불일치 검사
        mismatch = []
        if dep_date and date_found != "날짜미확인":
            try:
                site_date = datetime.strptime(date_found, "%Y-%m-%d")
                if dep_date.date() != site_date.date():
                    mismatch.append(
                        f"날짜불일치(PDC:{dep_date.strftime('%m/%d')} vs 사이트:{site_date.strftime('%m/%d')})"
                    )
            except:
                pass

        if mismatch:
            return "⚠️ 불일치", detail + " | " + " / ".join(mismatch)

        return "✅ 확인완료", detail

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception:
        logging.error(f"파라타항공 조회 실패 | PNR: {pnr} | 탑승객: {kor_name}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"
    finally:
        page.remove_listener("dialog", _on_dialog)


async def check_tw(page, target):
    """TW는 Akamai Premium으로 자동조회 불가 → 건별 URL 파라미터로 Chrome 탭 오픈"""
    import urllib.parse
    pnr      = target["pnr"]
    kor_name = target["kor_name"]
    eng_name = target.get("eng_name", "")
    dep      = target["dep"]
    arr      = target["arr"]

    # 국내선=한글, 국제선=영문
    intl = is_international(dep, arr)
    if intl and eng_name:
        parts    = eng_name.split("/")
        name_last  = parts[0].strip() if len(parts) >= 1 else target["last"]
        name_first = parts[1].strip() if len(parts) >= 2 else target["first"]
        name_display = f"{name_last}/{name_first} (영문)"
    else:
        name_last  = target["last"]
        name_first = target["first"]
        name_display = f"{kor_name} (한글)"

    return "⚠️ 수동확인필요", f"티웨이-수동조회: {name_display} | PNR: {pnr}"


async def run_check(page, target, we_email=""):
    """단일 조회 실행 + 재시도 로직"""
    airline  = target["airline"]
    eng_name = target.get("eng_name", "")

    if airline == "에어부산":
        result, detail = await check_bx(page, target)
    elif airline == "대한항공":
        result, detail = await check_ke(page, target)
    elif airline == "진에어":
        result, detail = await check_lj(page, target)
    elif airline == "파라타항공":
        result, detail = await check_we(page, target, we_email)
    elif airline == "티웨이항공":
        result, detail = await check_tw(page, target)
    else:
        return "⬜ 미지원", "지원 항공사 아님"

    # ── 국내선 한글 조회 실패 시 영문으로 재시도 (외국인 승무원 대비) ──
    # 파라타 제외 / 영문명 있을 때 / PNR오류·예약없음일 때만
    intl = is_international(target["dep"], target["arr"])
    if (
        airline in ("에어부산", "대한항공", "진에어", "티웨이항공")
        and not intl
        and eng_name
        and any(kw in result for kw in ["PNR오류", "예약없음"])
    ):
        # 영문 이름으로 강제 전환한 임시 target 복사
        parts = eng_name.split("/")
        tmp = dict(target)
        tmp["last"]  = parts[0].strip() if len(parts) >= 1 else target["last"]
        tmp["first"] = parts[1].strip() if len(parts) >= 2 else target["first"]
        # 국제선 분기를 타지 않도록 dep/arr을 해외로 속이지 않고
        # check_* 함수 내 intl=False → 한글 사용이 되므로
        # last/first만 덮어쓴 tmp를 넘기면 영문으로 입력됨
        await asyncio.sleep(1)
        if airline == "에어부산":
            r2, d2 = await check_bx(page, tmp)
        elif airline == "대한항공":
            r2, d2 = await check_ke(page, tmp)
        elif airline == "진에어":
            r2, d2 = await check_lj(page, tmp)
        elif airline == "티웨이항공":
            r2, d2 = await check_tw(page, tmp)
        if "확인완료" in r2 or "불일치" in r2:
            result = r2
            detail = "[영문재시도] " + d2

    # ── 타임아웃/오류 시 1회 재시도 ──
    elif "타임아웃" in result or ("오류" in result and "파싱" not in detail):
        await asyncio.sleep(2)
        if airline == "에어부산":
            result, detail = await check_bx(page, target)
        elif airline == "대한항공":
            result, detail = await check_ke(page, target)
        elif airline == "진에어":
            result, detail = await check_lj(page, target)
        elif airline == "파라타항공":
            result, detail = await check_we(page, target, we_email)
        elif airline == "티웨이항공":
            result, detail = await check_tw(page, target)
        if "타임아웃" not in result:
            detail = "[재시도 성공] " + detail

    return result, detail



def _show_tw_popup(tw_manual, root):
    """티웨이 수동확인 건별 팝업 — PNR/이름 복사 버튼"""
    import tkinter as tk

    total = len(tw_manual)

    for idx, t in enumerate(tw_manual):
        # 이름 결정 (국내=한글, 국제=영문)
        intl = is_international(t["dep"], t["arr"])
        eng  = t.get("eng_name", "")
        if intl and eng:
            parts      = eng.split("/")
            name_last  = parts[0].strip()
            name_first = parts[1].strip() if len(parts) >= 2 else ""
            name_label = f"{name_last} / {name_first}  (영문)"
        else:
            name_last  = t["last"]
            name_first = t["first"]
            name_label = f"{name_last} / {name_first}  (한글)"

        pnr = t["pnr"]

        # 팝업 뜨기 직전에 Chrome 탭 열기
        import urllib.parse
        params = urllib.parse.urlencode({
            "pnrNumber": pnr,
            "lastName": name_last,
            "firstName": name_first,
        })
        webbrowser.open(f"https://www.twayair.com/app/reservation/searchMemberBooking?{params}")

        # ── 팝업 윈도우 ──
        win = tk.Toplevel(root)
        win.title(f"티웨이 수동확인  [{idx+1}/{total}]")
        win.resizable(False, False)
        win.attributes("-topmost", True)

        pad = dict(padx=16, pady=6)

        tk.Label(win, text=f"[{idx+1}/{total}]  {t['kor_name']}  |  {t['dep']}→{t['arr']}  {t['dep_time']}",
                 font=("맑은 고딕", 10, "bold")).grid(row=0, column=0, columnspan=3, pady=(14, 4), padx=16)

        # 구분선
        tk.Frame(win, height=1, bg="#cccccc").grid(row=1, column=0, columnspan=3, sticky="ew", padx=16)

        # 예약번호 행
        tk.Label(win, text="예약번호", width=8, anchor="e").grid(row=2, column=0, **pad)
        tk.Label(win, text=pnr, font=("Consolas", 13, "bold"), fg="#c0392b", width=10, anchor="w").grid(row=2, column=1, **pad)

        def _copy_pnr(p=pnr):
            root.clipboard_clear(); root.clipboard_append(p); root.update()
            btn_pnr.config(text="✓ 복사됨", fg="green")
            win.after(1200, lambda: btn_pnr.config(text="복사", fg="black"))
        btn_pnr = tk.Button(win, text="복사", width=6, command=_copy_pnr)
        btn_pnr.grid(row=2, column=2, **pad)

        # 이름 행
        tk.Label(win, text="이  름", width=8, anchor="e").grid(row=3, column=0, **pad)
        tk.Label(win, text=name_label, font=("맑은 고딕", 11), width=22, anchor="w").grid(row=3, column=1, **pad)

        def _copy_name(last=name_last, first=name_first):
            root.clipboard_clear(); root.clipboard_append(last); root.update()
            btn_last.config(text="✓ 성 복사됨", fg="green")
            win.after(1200, lambda: btn_last.config(text="성 복사", fg="black"))
        btn_last = tk.Button(win, text="성 복사", width=6, command=_copy_name)
        btn_last.grid(row=3, column=2, **pad)

        # 이름(first) 행
        tk.Label(win, text="", width=8).grid(row=4, column=0)
        tk.Label(win, text="↑ 성 복사 후 → 이름 복사", font=("맑은 고딕", 9), fg="#888888", anchor="w").grid(row=4, column=1, sticky="w", padx=16)

        def _copy_first(first=name_first):
            root.clipboard_clear(); root.clipboard_append(first); root.update()
            btn_first.config(text="✓ 이름 복사됨", fg="green")
            win.after(1200, lambda: btn_first.config(text="이름 복사", fg="black"))
        btn_first = tk.Button(win, text="이름 복사", width=8, command=_copy_first)
        btn_first.grid(row=4, column=2, padx=16, pady=2)

        # 구분선
        tk.Frame(win, height=1, bg="#cccccc").grid(row=5, column=0, columnspan=3, sticky="ew", padx=16, pady=(6,0))

        # 다음/건너뛰기 버튼
        done = tk.BooleanVar(value=False)
        btn_frame = tk.Frame(win)
        btn_frame.grid(row=6, column=0, columnspan=3, pady=12)

        next_text = "다음 →" if idx < total - 1 else "완료 ✓"
        tk.Button(btn_frame, text=next_text, width=10,
                  command=lambda: done.set(True)).pack(side="left", padx=8)
        tk.Button(btn_frame, text="건너뛰기", width=8, fg="#888888",
                  command=lambda: done.set(True)).pack(side="left", padx=8)

        win.protocol("WM_DELETE_WINDOW", lambda: done.set(True))

        # 화면 중앙 배치
        win.update_idletasks()
        w, h = win.winfo_width(), win.winfo_height()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        win.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        win.wait_variable(done)
        win.destroy()


async def main():
    print(f"{'='*54}")
    print("  ✈️  타사 예약 자동 검증 시스템  v3.1")
    print("      2026-06-27  |  문의: 승무계획팀")
    print(f"{'='*54}")
    print("  [자동조회]  에어부산 / 대한항공 / 진에어 / 파라타항공")
    print("  [반자동]    티웨이항공  → Chrome 탭 자동 오픈 + 복사 팝업")
    print(f"{'─'*54}")
    print("  v3.1  조회 범위 3가지 선택 (5일/이번달말/다음달말)")
    print("  v3.0  티웨이 Chrome 탭 자동 오픈 + 건별 복사 팝업")
    print("  v2.1  에어부산 CF 우회 / 파라타 자동조회 / 월말조회")
    print("        대한항공 국제선 개선 / 동명이인·하이픈 처리")
    print(f"{'='*54}\n")

    check_for_update()

    # 조회 범위 선택 팝업
    check_days, delay_min, delay_max = get_check_mode()
    today    = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    end_date = today + timedelta(days=check_days)
    mode_label = f"{end_date.month}월 {end_date.day}일까지" if check_days > 5 else f"오늘~5일 이내"

    targets = load_targets(EXCEL_PATH, SHEET_NAME, end_date)
    total   = len(targets)

    if total == 0:
        print(f"검증 대상이 없습니다. ({mode_label})")
        input("엔터 누르면 종료...")
        return

    bx_cnt = sum(1 for t in targets if t["airline"] == "에어부산")
    ke_cnt = sum(1 for t in targets if t["airline"] == "대한항공")
    lj_cnt = sum(1 for t in targets if t["airline"] == "진에어")
    we_cnt = sum(1 for t in targets if t["airline"] == "파라타항공")
    tw_cnt = sum(1 for t in targets if t["airline"] == "티웨이항공")

    print(f"검증 대상: {total}건 ({mode_label})")
    print(f"  에어부산: {bx_cnt}건 | 대한항공: {ke_cnt}건 | 진에어: {lj_cnt}건 | 파라타항공: {we_cnt}건 | 티웨이: {tw_cnt}건")
    print(f"  딜레이: {delay_min}~{delay_max}초")
    print(f"{'='*50}\n")

    # 파라타항공 건수 있으면 이메일 입력 팝업
    we_email = ""
    if we_cnt > 0:
        from tkinter import simpledialog
        we_email = simpledialog.askstring(
            "파라타항공 조회 이메일",
            f"파라타항공 예약 {we_cnt}건이 있습니다.\n\n예약 시 사용한 이메일 주소를 입력하세요.\n(홈페이지 비회원 조회용)",
            parent=root
        )
        if not we_email:
            print("⚠️  이메일 미입력 → 파라타항공 건은 수동확인으로 처리됩니다.")
        else:
            print(f"파라타항공 조회 이메일: {we_email}\n")

    # 시스템 Chrome 경로 자동 탐색
    CHROME_PATHS = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expanduser(r"~\AppData\Local\Google\Chrome\Application\chrome.exe"),
    ]
    chrome_exe = next((p for p in CHROME_PATHS if os.path.exists(p)), None)

    if chrome_exe:
        print(f"시스템 Chrome 사용: {chrome_exe}")
    else:
        print("⚠️  Chrome 미발견 → Playwright Chromium으로 실행 (에어부산 캡챠 발생 가능)")

    LAUNCH_ARGS = [
        "--disable-blink-features=AutomationControlled",
        "--disable-infobars",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-extensions",
    ]
    CONTEXT_OPTS = dict(
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/124.0.0.0 Safari/537.36"
        ),
        locale="ko-KR",
        timezone_id="Asia/Seoul",
        viewport={"width": 1280, "height": 800},
        java_script_enabled=True,
    )

    async with async_playwright() as p:
        launch_kwargs = dict(headless=HEADLESS, args=LAUNCH_ARGS)
        if chrome_exe:
            launch_kwargs["executable_path"] = chrome_exe
        browser = await p.chromium.launch(**launch_kwargs)
        context = await browser.new_context(**CONTEXT_OPTS)

        # ── playwright-stealth 적용 ──
        if STEALTH_AVAILABLE:
            stealth = Stealth(
                navigator_languages_override=("ko-KR", "ko"),
                navigator_platform_override="Win32",
                navigator_webdriver=True,
                chrome_runtime=True,
            )
            print("playwright-stealth 적용 완료")
        else:
            stealth = None
            await context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['ko-KR','ko','en-US','en']});
                window.chrome = {runtime: {}};
            """)
            print("⚠️  playwright-stealth 미설치 → 기본 우회 모드")

        page = await context.new_page()
        if stealth:
            await stealth.apply_stealth_async(page)

        pnr_cache = {}

        for i, target in enumerate(targets, 1):
            airline = target["airline"]
            pnr     = target["pnr"]
            print(f"[{i:02d}/{total}] {target['kor_name']:5} | {airline} | {pnr} | ", end="", flush=True)

            if pnr in pnr_cache:
                result, detail = pnr_cache[pnr]
                target["result"] = result
                target["detail"] = detail
                print(f"[캐시] {result}  {detail}")
            else:
                result, detail = await run_check(page, target, we_email)
                target["result"] = result
                target["detail"] = detail
                if "확인완료" in result:
                    pnr_cache[pnr] = (result, detail)
                print(f"{result}  {detail}")
                if i < total:
                    if target["airline"] == "에어부산":
                        await asyncio.sleep(random.uniform(5.0, 10.0))
                    else:
                        await asyncio.sleep(random.uniform(delay_min, delay_max))

        await browser.close()

    save_results(EXCEL_PATH, SHEET_NAME, targets)

    confirmed = sum(1 for t in targets if t["result"] and "확인완료"    in str(t["result"]))
    mismatch  = sum(1 for t in targets if t["result"] and "불일치"       in str(t["result"]))
    no_rsv    = sum(1 for t in targets if t["result"] and "예약없음"     in str(t["result"]))
    pnr_err   = sum(1 for t in targets if t["result"] and "PNR오류"     in str(t["result"]))
    manual    = sum(1 for t in targets if t["result"] and "수동확인필요" in str(t["result"]))
    error     = sum(1 for t in targets if t["result"] and ("오류" in str(t["result"]) or "타임아웃" in str(t["result"])))

    print(f"\n{'='*50}")
    print(f"✅ 확인완료      : {confirmed}건")
    print(f"⚠️  불일치       : {mismatch}건  ← 즉시 확인!")
    print(f"❌ 예약없음      : {no_rsv}건   ← 즉시 확인!")
    print(f"❌ PNR오류       : {pnr_err}건  ← 즉시 확인!")
    print(f"⚠️  수동확인필요  : {manual}건  ← 티웨이항공 직접 조회 필요!")

    # TW 수동확인 팝업 (건별 복사 버튼)
    tw_manual = [t for t in targets if "티웨이" in str(t["airline"]) and "수동확인필요" in str(t["result"])]
    if tw_manual:
        # 같은 PNR은 팝업 1번만 (동행 탑승객은 탑승객 추가 버튼으로 처리)
        seen_pnr = set()
        tw_dedup = []
        for t in tw_manual:
            if t["pnr"] not in seen_pnr:
                seen_pnr.add(t["pnr"])
                tw_dedup.append(t)
        print(f"\n{'─'*52}")
        print(f"  📋 티웨이항공 수동확인 팝업 실행 중 ({len(tw_dedup)}건 / 전체 {len(tw_manual)}명)")
        print(f"{'─'*52}")
        _show_tw_popup(tw_dedup, root)

    print(f"💥 오류/재시도   : {error}건")
    input("\n엔터 누르면 종료...")


if __name__ == "__main__":
    asyncio.run(main())