import asyncio
import random
import re
import logging
import os
import sys
import json
import subprocess
import threading
import time
import uuid
import webbrowser
from datetime import datetime, timedelta
import openpyxl
from tkinter import filedialog, messagebox, ttk
import tkinter as tk
# Patchright(자동화 흔적을 브라우저 내부에서 제거한 Playwright 호환판)가 설치돼 있으면 우선 사용
# — 에어부산 보안확인이 체크해도 무한반복되는 문제 대응 (v3.15.0). 없으면 기존 Playwright
try:
    from patchright.async_api import async_playwright, TimeoutError as PWTimeout
    PATCHRIGHT = True
except ImportError:
    from playwright.async_api import async_playwright, TimeoutError as PWTimeout
    PATCHRIGHT = False
try:
    from playwright_stealth import Stealth
    STEALTH_AVAILABLE = not PATCHRIGHT  # Patchright 사용 시 stealth 병행 금지(충돌)
except ImportError:
    STEALTH_AVAILABLE = False

__version__ = "3.18.0"
VERSION_URL = "https://raw.githubusercontent.com/vipywk-lab/DH-checker/main/bx_checker.py"
NAS_PATH    = r"\\10.223.120.38\종합통제\24. 승무계획팀\29.자동화\DH 조회 자동화"
GITHUB_URL  = "https://github.com/vipywk-lab/DH-checker"

# 실행 시 콘솔에 표시되는 이번 버전 변경사항 (유저용 — 기술 용어 지양, 짧게)
LATEST_CHANGELOG = (
    "  - [사용법 변경] 에어부산은 평소 쓰는 크롬에서 확장 프로그램이 자동 조회합니다.\n"
    "    (최초 1회 'DH조회_확장' 설치 필요 — 폴더 안 설치방법.txt 참고)\n"
    "    조회 시작 시 에어부산 창이 열리면 '사람인지 확인'만 통과 → 팝업 [확인].\n"
    "    확장이 없는 PC는 자동으로 수동확인 팝업으로 진행됩니다."
)

# 클라우드플레어 감지 키워드 (전역 — 모든 항공사 조회 함수에서 공유)
CF_KEYWORDS = ["보안 확인 수행 중", "사람인지 확인하십시오", "Checking your browser",
               "DDoS protection", "보안 서비스", "악의적인 봇", "Cloudflare"]


def _is_reliable_result(flt_found, route_found):
    """
    편명·구간을 둘 다 못 읽었으면 실제 예약 정보를 읽은 게 아니라
    (알려지지 않은 오류 문구의 실패 페이지 등) 실패일 가능성이 높음.
    실패 키워드 목록에 없다고 무조건 "확인완료"로 넘기지 않기 위한 안전장치.
    """
    return not (flt_found == "편명미확인" and route_found == "구간미확인")


# ==========================================
# 체인지로그
# v3.18.0 (2026-09-26) — 에어부산: 평소 크롬 + 확장 프로그램 자동조회
#   - v3.17.0 수동확인은 사용자 부담이 커서, 평소 크롬(디버그 포트 없음 → 보안확인 통과됨)
#     안에서 확장 프로그램이 입력·조회하는 방식으로 자동화
#   - 구조: 이 프로그램이 127.0.0.1:38517에 작은 창구를 열고, 확장이 1초마다 다음 건을
#     받아가서 PNR/성/이름 입력 → 조회 → 결과 화면 글자를 돌려줌. 판정은 기존 로직 그대로
#     (결과 판정 부분을 _parse_bx_text로 분리해 크롬조종·확장 방식이 공유)
#   - BX_MODE 스위치: "ext"(기본) / "manual" / "auto"(크롬 조종 — 현재 차단됨, 복귀용)
#   - 확장이 없거나 응답이 없으면 에어부산은 자동으로 수동확인 팝업으로 대체
#   - 조회 도중 에어부산 탭이 닫히면 20초 후 자동으로 다시 열어줌
#   - 확장은 이 프로그램이 실행 중일 때만 동작 (평소 에어부산 이용엔 영향 없음)
#   - 차단 방지 딜레이(5~10초)는 확장 방식에도 그대로 적용
# v3.17.0 (2026-09-26) — 에어부산 수동확인 방식으로 전환
#   - 원인 확정 (사용자 PC 테스트): 같은 빈 프로필이라도
#       A) 디버그 포트 없이 띄운 크롬 → 보안확인 통과
#       B) 디버그 포트(--remote-debugging-port)를 연 크롬 → 무한반복
#       C) A에서 통과한 프로필에 포트만 열어 다시 띄움 → 다시 막힘
#     → 프로그램이 조종할 수 있는 크롬은 통과 쿠키가 있어도 에어부산을 못 지나감.
#     v3.15.0 Patchright(포트 대신 파이프 사용)도 이미 실패했으므로 자동조회 불가로 판단
#   - 에어부산을 티웨이와 같은 수동확인(기본 브라우저 + 복사 버튼 팝업)으로 전환.
#     성/이름은 자동조회 때와 같은 규칙(국제선+영문명→영문, 그 외→한글)으로 표시
#   - 티웨이 팝업 코드를 공용 함수(_manual_check)로 정리해 두 항공사가 같이 사용
#   - 자동조회 코드(check_bx, 크롬 먼저 띄우기)는 삭제하지 않고 BX_AUTO 스위치로 보존
#     → 클라우드플레어 정책이 바뀌면 True로 바꿔 복귀 가능
#   - 수동확인 건은 조회 간 대기(딜레이) 생략
# v3.16.6 (2026-09-26) — 조회용 크롬에서만 보안확인 무한반복 수정
#   - 증상: 첫 크롬 창에서 '사람인지 확인' 체크 → 뱅뱅 돌다가 다시 확인 화면 반복.
#     같은 PC·같은 IP의 평소 크롬에서는 체크하면 바로 통과됨
#   - 원인(추정): 조회용 전용 프로필을 계속 재사용하면서, 테스트 중 실패한 보안확인
#     기록(클라우드플레어 쿠키)이 그 프로필에 쌓임 → 그 프로필 자체가 의심받음
#   - 수정: 실행할 때마다 완전히 새 프로필(run_날짜시각 폴더)로 크롬을 띄움.
#     이전 실행에서 남은 조회용 크롬은 먼저 닫고, 예전 프로필 폴더는 자동 삭제.
#     (v3.16.4의 '남은 크롬 재사용'은 오염된 상태를 이어받을 수 있어 폐기)
# v3.16.5 (2026-09-26) — 배포 전 전체 점검에서 나온 안정성 보완 3건
#   - 보안확인 대기 중 자동 새로고침 간격 15초 → 30초 (사람이 체크박스를 누르는
#     도중에 새로고침돼서 확인이 끊기는 것 방지)
#   - 크롬 연결 확인 요청이 회사 프록시 설정을 타지 않도록 함 (다른 PC에서
#     프록시 때문에 연결 실패 → 자동화 브라우저로 대체되는 상황 예방)
#   - 사람이 보안확인을 끝낸 탭을 닫아버려도, 브라우저가 살아있으면 새 탭으로
#     교체해서 계속 진행 (파라타·티웨이는 이 탭을 직접 쓰기 때문에 전체가
#     중단되던 문제 예방)
# v3.16.4 (2026-09-26) — "이어서 조회" 시 보안확인 통과 불가 수정
#   - 증상: 처음 실행 땐 정상인데, 중간에 끊겨서 "이어서 조회"로 재실행하면
#     에어부산 보안확인 체크박스를 눌러도 넘어가지 않음
#   - 원인: 이전 실행이 중간에 끊기면(콘솔 창 닫힘 등) 조회용 크롬이 종료되지 않고
#     남아있음. 재실행 시 포트 기록 파일을 무조건 지운 뒤 크롬을 다시 실행했는데,
#     같은 프로필의 크롬이 이미 떠 있으면 새 크롬은 기존 크롬에 창만 추가하고
#     끝나서 새 포트 기록이 생기지 않음 → 20초 대기 후 연결 실패 → 자동화
#     브라우저(기존 방식)로 대체 실행 → 이 브라우저에선 체크해도 통과 불가
#   - 수정: 재실행 시 남아있는 조회용 크롬이 응답하면 그 크롬을 그대로 재사용.
#     종료할 때도 크롬 본체에 직접 종료 명령을 보내서 재사용한 크롬까지 확실히 닫음
# v3.16.3 (2026-09-26) — 브라우저 통째로 닫힘 대응 + 오류 감지 정확도 개선
#   - 증상: 에어부산 조회 중 보안확인에 걸린 뒤, 브라우저 자체가 완전히 닫혀서
#     그 이후 모든 건이 연쇄적으로 실패함 (TargetClosedError)
#   - 추정 원인: 조회용 탭 하나만 열려있는 상태에서 그 탭이 어떤 이유로든
#     닫히면(사람이 실수로 닫거나, 사이트 쪽에서 닫히거나) 크롬은 '마지막 창이
#     닫혔다'고 판단해 프로그램 전체를 종료해버림. 사람이 최초 보안확인을
#     통과시킨 탭 외에는 상시로 열려있는 탭이 없었던 게 근본 원인
#   - 수정 1: 아무것도 하지 않는 빈 탭을 하나 띄워두고 절대 닫지 않음 →
#     다른 탭이 전부 닫혀도 크롬 자체는 계속 살아있음
#   - 수정 2: 각 항공사 조회 함수의 오류 처리가 "브라우저 닫힘" 오류까지
#     뭉뚱그려 "💥 오류"로 삼켜버려서, 정작 필요한 시점에 메인 루프가
#     "브라우저 닫힘"을 못 알아채고 있었음 → 이제 정확히 구분해서 즉시 감지
# v3.16.2 (2026-09-26) — 조회 도중 재등장하는 에어부산 보안확인 새로고침 자동화
#   - 최초 진입 시 보안확인(사람이 직접 크롬을 통과)은 v3.16.1로 잘 해결됨
#   - 그런데 조회 도중 2건째 이후 새 탭에서 보안확인이 다시 뜨는 경우가 있었고,
#     이땐 자동감지 루프가 그냥 가만히 기다리기만 해서 화면이 멈춘 채로 최대
#     3분 타임아웃까지 흘러가버림 — 정작 필요했던 "새로고침" 동작이 빠져있었음
#   - 수정: 대기 중 15초마다 자동으로 새로고침(F5와 동일 효과)을 시도하도록 추가.
#     화면을 앞으로 가져오는 동작(bring_to_front)은 유지 — 사용자가 직접
#     체크박스를 누르거나 새로고침해도 무방함 (둘 다 동시에 동작 가능)
# v3.16.1 (2026-09-26) — 에어부산 보안확인 화면 멈춤 현상 개선
#   - 증상: 크롬 창이 열리고 보안확인(사람인지 확인) 화면에서 멍하게 멈춰있다가,
#     새로고침하면 그제서야 확인 화면이 다시 뜨면서 정상 통과됨
#   - 추정 원인: v3.16.0에서 자동화 연결용 디버그 포트를 9222(자동화 도구들이
#     쓰는 가장 널리 알려진 기본값)로 고정해뒀음. 보안검사 스크립트가 페이지
#     안에서 이 포트가 열려있는지 확인해 자동화로 의심했을 가능성이 있음
#     (100% 확정된 원인은 아니며, 실사용 결과로 검증 필요)
#   - 수정: 포트를 9222로 고정하지 않고 크롬이 매번 랜덤으로 고르게 변경
#     (크롬 표준 기능인 --remote-debugging-port=0 사용)
# v3.16.0 (2026-09-25) — 에어부산 보안확인: "사람이 먼저 통과 → 그 창에 연결" 방식
#   - v3.14~3.15(stealth 조정, Patchright)로도 체크 후 무한반복 해결 안 됨
#   - 에어부산 건이 있으면 크롬을 자동화 없이 '일반 실행'으로 먼저 띄우고,
#     사용자가 보안확인을 직접 통과한 뒤 [확인] → 프로그램이 그 크롬에 연결해 조회
#     (보안확인 시점엔 자동화 연결이 아예 없어서 일반 사용자와 동일)
#   - 전용 프로필(DH_checker_profile) 사용, 조회 끝나면 그 크롬은 자동 종료
#   - 연결 실패/크롬 미설치/에어부산 0건이면 기존 방식 그대로
# v3.15.0 (2026-09-25) — 에어부산 보안확인 무한반복 대응: Patchright 도입
#   - v3.14.1 이후에도 체크박스를 눌러도 넘어가지 않음 → stealth 문제가 아니라
#     Playwright 자체의 자동화 흔적(브라우저 조종 채널 신호, 자동화 플래그)을
#     Cloudflare가 감지하는 것으로 판단
#   - Patchright: Playwright와 사용법이 100% 같고, 이 흔적들을 브라우저 내부에서
#     제거한 호환판. 설치돼 있으면 자동으로 사용, 없으면 기존 Playwright로 동작
#   - Patchright 사용 시 stealth·커스텀 인자·고정 창크기는 끔 (충돌 방지 권장사항)
#   - 사람이 체크박스를 누르는 구조는 그대로 — 캡챠 자동풀이는 넣지 않음
# v3.14.1 (2026-09-25) — 에어부산 보안확인 무한반복 수정 (v3.14.0 회귀)
#   - v3.14.0에서 stealth를 모든 탭에 적용하도록 바꾼 뒤, 에어부산 확인창에서
#     체크해도 넘어가지 않고 계속 확인만 반복되는 현상 발생
#   - 원인: Cloudflare가 stealth의 가짜 브라우저 정보(플랫폼·플러그인 위장 등)를
#     실제 Chrome 정보와 안 맞는 이상 신호로 판단. v3.13 이전엔 조회용 새 탭에
#     stealth가 안 걸려 있어서 우연히 체크가 통과되고 있었던 것
#   - 수정: stealth 적용 범위를 원래대로(첫 탭만) 되돌림. 전용 프로필·UA 고정 제거·
#     통과 자동감지는 유지
# v3.14.0 (2026-09-25) — 에어부산 보안확인(캡챠) 빈도 감소
#   - [주원인 수정] 위장설정(stealth)이 첫 탭에만 적용되고, 조회마다 새로 여는
#     탭에는 적용되지 않던 버그 → 브라우저 전체(context)에 적용하도록 변경
#   - User-Agent를 Chrome/124로 고정하던 것 제거 — 실제 설치된 Chrome 버전과
#     안 맞아서 오히려 봇 의심 신호가 되고 있었음
#   - 전용 브라우저 프로필(~\AppData\Local\DH_checker_profile) 사용 → 한 번
#     통과한 보안확인 기록이 다음 실행에도 유지됨. 평소 Chrome과는 분리됨.
#     프로필 사용 불가 시(동시 실행 등) 기존 일회성 방식으로 자동 전환
#   - 보안확인 화면이 뜨면 통과 여부를 1초마다 자동 감지(최대 3분) →
#     콘솔로 돌아가 엔터 누르는 단계 제거
#   - 캡챠 자동 클릭/외부 캡챠풀이 서비스는 의도적으로 넣지 않음(약관·안정성)
# v3.13.1 (2026-09-21) — 윈도우 계정 변경 후 브라우저 설치 오류 수정
#   - 증상: 윈도우 계정을 바꾼 뒤 실행하면 "기반 시스템(브라우저)을 설치 중입니다"
#     단계에서 "[WinError 2] 지정된 파일을 찾을 수 없습니다" 오류 발생
#   - 원인: playwright 설치 명령을 "PATH에 등록된 playwright 실행파일 이름"으로
#     찾아 실행했는데, 이 실행파일은 계정별 폴더에 설치되고 PATH 등록도
#     계정마다 따로라서, 계정을 바꾸면 못 찾는 경우가 있었음
#   - 수정: 지금 실행 중인 파이썬 자체를 통해 playwright 모듈을 직접 호출하도록
#     변경 — 어떤 계정에서 실행하든 항상 같은 방식으로 동작함
# v3.13.0 (2026-09-12) — 조회 속도 개선 (IP밴 위험 없는 범위)
#   - 캐시 제거(v3.11.0) 후 전수조회로 늘어난 시간 부담 완화 목적
#   - 항공사 사이트 사이의 딜레이(차단 방지용)는 절대 건드리지 않음 — 오직
#     우리 코드 내부의 불필요한 고정대기(입력 필드 사이 300ms 등)만 정리
#     · Playwright의 fill()/click()은 요소가 준비될 때까지 자동으로 기다리므로
#       입력 직후의 고정 대기는 대부분 중복이었음 → 제거
#   - 건당 내부 고정대기 총합: 에어부산 6.2초→3.8초, 대한항공 8.7초→5.7초,
#     진에어 8.5초→4.9초, 파라타 4.9초→2.7초, 제주항공 5.2초→4.1초
#   - 제주항공은 과거 타이밍 관련 버그가 많았던 곳이라 명백히 중복인 곳만
#     최소한으로 정리하고 나머지는 그대로 유지(보수적 접근)
# v3.12.0 (2026-09-12) — 진행상황 표시창 추가 (Gemini 코드리뷰 반영)
#   - 조회 시작하면 화면 오른쪽 위에 작은 창이 떠서 진행률 바(N/전체건),
#     방금 조회한 사람과 결과, 확인완료/불일치/오류·기타 실시간 집계를 보여줌
#   - "화면이 멈춘 건지 도는 건지 모르겠다"는 불안감 해소 목적
#   - 이 창을 닫아도(X 버튼) 조회 자체엔 영향 없음, 콘솔이 계속 진행함
# v3.11.2 (2026-09-12) — 데이터 보호 강화 (Gemini 코드리뷰 반영)
#   - 엑셀 저장 실패 팝업을 재시도/취소 방식으로 개선. 이전엔 파일이 계속
#     잠겨있으면 탈출구 없이 무한 반복됐음 → 이제 [취소] 선택 시
#     "검증결과_백업_YYYYMMDD_HHMMSS.xlsx"로 새 파일 저장 후 안전하게 종료
#   - 조회 결과를 매 건마다 가벼운 JSONL 백업 파일("결과백업_*.jsonl")에도
#     한 줄씩 남김. 엑셀 파일 자체가 손상되는 최악의 경우에도 이 파일로 복구 가능
# v3.11.1 (2026-09-02) — [중요] 두 개의 안정성 수정 병합
#   - (다른 세션에서 작업) 이어서 조회 기능의 실제 안전성 확보:
#     기존엔 전체 조회가 끝난 뒤 한 번에 저장해서, 창닫기/작업관리자 종료/
#     정전 등으로 예고 없이 죽으면 그때까지 결과가 통째로 사라지고
#     "이어서 조회"도 무용지물이었음. 이제 한 건 처리할 때마다 조용히
#     엑셀에 즉시 반영 → 어떤 방식으로 죽어도 재실행 시 정확히 다음 건부터 이어짐
#   - (이 세션에서 작업) v3.11.0의 PNR 캐시 제거를 유지한 채 위 수정과 병합
# v3.11.0 (2026-09-02) — [중요] PNR 캐시 기능 완전 제거
#   - v3.10.1에서 동승자 이름 확인 로직을 넣었음에도 오탐 사례가 추가로
#     보고되어, 캐시 자체를 없애고 같은 예약번호라도 사람마다 항상 개별
#     조회하도록 변경. 이 도구는 속도보다 정확도가 우선이라는 판단.
#   - 조회 시간은 늘지만 v3.10.0의 "이어서 조회" 기능으로 중단 시
#     재조회 부담은 크지 않음
# v3.10.3 (2026-09-02) — 대한항공 쿠키 확인창 클릭 안 되던 문제 수정
#   - 실제 버튼 구조가 예상과 달라(class "-cta", 텍스트는 내부 span에 위치)
#     클릭이 안 되던 문제 확인, data-click-name="Accept all" 속성 기반으로 수정
# v3.10.2 (2026-09-02) — 대한항공 쿠키 확인창 자동 처리
#   - 대한항공 쿠키 확인창 클릭이 안 되던 문제 수정 (버튼 구조 재확인)
# v3.10.1 (2026-08-26) — [중요] 캐시 오판정 수정 (사용자 리포트)
#   - 같은 예약번호(PNR)면 이전 결과를 그대로 재사용하면서 그 사람이 실제로
#     해당 예약 명단에 있는지는 확인하지 않아, 엉뚱한 사람도 '확인완료'로
#     넘어가던 문제 수정
#   - 이제 캐시 사용 전 결과 화면에 해당 탑승객(한글명/영문명)이 있는지 확인하고,
#     확인되지 않으면 캐시를 쓰지 않고 개별 조회함 (화면 표시: [캐시·동승확인])
# v3.10.0 (2026-08-26) — 이어서 조회 기능 추가
#   - 실행 시 이미 '확인완료'된 건이 있으면 건너뛸지 묻는 팝업 표시
#     (한 달치 400건 규모는 1시간 가까이 걸려, 중단 시 처음부터 다시 하던 문제 해소)
#   - 오류·불일치·PNR오류·미조회 건은 건너뛰지 않고 항상 다시 조회
#   - 건너뛴 건의 기존 결과는 그대로 유지되고 요약 시트에도 정상 반영됨
# v3.9.4 (2026-08-26) — 파일 선택 관련 안내 개선
#   - 조회 시작 시 실제로 읽는 엑셀 파일의 전체 경로를 콘솔에 표시
#     (화면에서 편집한 파일과 다른 파일을 선택해 "대상 0건"이 뜨던 혼란 방지)
#   - 대상 0건일 때 원인 체크리스트(파일 확인/저장 여부/조회 범위) 안내 추가
# v3.9.3 (2026-07-24) — 안정성 보강 (전체 재검토)
#   - 이름 칸이 비었거나 알파벳만 있는 경우 IndexError로 그 건이 오류나던 문제 방어
#   - 조회 결과가 엑셀 행과 매칭 안 되면(이름/PNR 불일치) 조용히 빈칸으로 남던 것을
#     콘솔에 매칭 실패 건을 안내하도록 개선
#   - GitHub 파일에서 버전 정보를 못 찾을 때 안내 출력 (버전체크 무력화 방지)
# v3.9.2 (2026-07-24) — 전체 재검토로 발견한 버그 수정
#   - 조회 도중 예외가 나면 프로그램이 통째로 죽고 그때까지의 결과가
#     하나도 저장되지 않던 문제 수정 (이제 항상 엑셀 저장 + 미조회 건 표시)
#   - 제주항공 재조회(영문/타임아웃) 분기가 빠져 있어 프로그램이 멈추던 문제 수정
#   - 대상 0건일 때 오류로 종료되던 문제 수정
#   - PNR오류가 오류 건수에 중복 집계되던 문제 수정
# v3.9.1 (2026-07-24) — [매우 중요] PNR 오판정 버그 수정
#   - 에어부산/대한항공: 조회마다 같은 브라우저 탭을 재사용하고 있었는데,
#     잘못된 PNR 조회가 실패해도 화면이 안 바뀌면 직전 사람의 결과 화면이
#     그대로 남아있어 그걸 "확인완료"로 잘못 읽는 문제가 있었음
#     → 조회마다 새 탭 사용 + 조회한 PNR이 실제 결과 화면에 있는지 확인하는
#       이중 안전장치 추가 (파라타항공에도 PNR 검증 추가)
# v3.9.0 (2026-07-07) — [중요] 오판정 방지 안전장치 추가
#   - 에어부산/대한항공/진에어/파라타항공/제주항공 5개 항공사 전부 해당:
#     실패 문구를 못 찾았다고 해서 무조건 "확인완료"로 반환하지 않도록 수정.
#     편명·구간을 둘 다 못 읽은 경우(=제대로 된 예약 페이지를 읽은 게 아닐 가능성)
#     "PNR오류"로 표시해 즉시 확인하도록 변경
# v3.8.x (2026-07-04) — 제주항공 조회 기능 추가
#   - PNR/성명 자동입력, 조회결과 자동판정(날짜·구간까지 정확히 확인)
#   - 달력 날짜 선택만 사람이 클릭(안내 팝업 표시), 나머지는 자동 진행
#   - 사이트 광고 팝업 자동 닫기, 구버전 안내에 NAS 경로/GitHub 링크 반영
# v3.7.0 (2026-07-04) — [중요] 진에어 날짜 오판정 수정
#   - 예약일을 출발일로 착각해서 멀쩡한 예약도 "날짜불일치"로 잘못 뜨던 문제 수정
# v3.6.0 (2026-07-03) — 제주항공 조회 최초 추가
# v3.5.0~3.5.2 (2026-07-03)
#   - GitHub 버전체크 연동(구버전이면 실행 시 자동 안내 후 차단)
#   - 실행창 제목표시줄에 버전 자동 표시
# v3.4.0~3.4.1 (2026-07-03) — 티웨이항공 조회 추가
#   - 자동조회 불가 사이트라 팝업 안내 + 항목별 [복사] 버튼으로 지원
# v3.3.0~3.3.2 (2026-07-01~03)
#   - 진에어 조회 오류 수정(정상 예약도 PNR오류로 뜨던 문제)
#   - 조회 범위 선택 1/2/3 버튼 복원, 오류 로그 저장 위치 고정
# v3.2.1 이하 — 초기 안정화 (생략)
# ==========================================
# Playwright Chromium 최초 1회 자동 설치
# ==========================================
def check_for_update():
    """GitHub raw URL에서 최신 버전 확인 — 구버전이면 실행 차단"""
    import urllib.request
    try:
        with urllib.request.urlopen(VERSION_URL, timeout=5) as resp:
            found_version = False
            for line in resp.read().decode("utf-8").splitlines():
                if line.startswith("__version__"):
                    found_version = True
                    latest = line.split("=")[1].strip().strip('"').strip("'")

                    def _ver(v):
                        try:
                            return tuple(int(x) for x in v.split("."))
                        except Exception:
                            return (0,)

                    if _ver(latest) > _ver(__version__):
                        print(f"\n{'!'*50}")
                        print(f"  ⚠️  업데이트 필요: 현재 v{__version__} → 최신 v{latest}")
                        print(f"  최신 파일로 교체가 필요합니다. 아래 둘 중 하나로 받아주세요.")
                        print(f"  1) NAS: {NAS_PATH}")
                        print(f"  2) GitHub: {GITHUB_URL}")
                        print(f"{'!'*50}\n")
                        input("업데이트 후 다시 실행해주세요. 엔터 누르면 종료...")
                        raise SystemExit("구버전 실행 차단")
                    else:
                        print(f"✅ 최신 버전입니다 (v{__version__})")
                    return
            # 파일은 받았으나 __version__ 줄을 못 찾은 경우 (포맷 변경 등)
            if not found_version:
                print("⚠️  버전 정보를 확인하지 못했습니다 (버전 체크 건너뜀)\n")
    except SystemExit:
        raise
    except Exception:
        print("⚠️  버전 확인 실패 (네트워크 연결 없음 — 무시하고 계속 진행)\n")


check_for_update()

chromium_path = os.path.expanduser("~\\AppData\\Local\\ms-playwright")
if not os.path.exists(chromium_path):
    print("기반 시스템(브라우저)을 설치 중입니다. 최초 1회만 진행되며 시간이 조금 걸릴 수 있습니다...")
    try:
        # sys.executable(현재 실행 중인 파이썬 경로)로 모듈 호출 → PATH에 의존하지 않음
        # (윈도우 계정을 바꾸면 playwright.exe가 새 계정 PATH에 없어 [WinError 2]가 나던 문제 방지)
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
        print("설치 완료!\n")
    except Exception as e:
        print(f"설치 중 오류가 발생했습니다: {e}")

# ==========================================
# 로깅 설정 (오류 발생 시 텍스트 파일로 저장)
# ==========================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
LOG_PATH = os.path.join(SCRIPT_DIR, f"에러로그_{datetime.now().strftime('%Y%m%d')}.txt")
# 조회 결과 원본 백업 (엑셀 손상 시 복구용) — 실행할 때마다 새 파일
JSONL_PATH = os.path.join(SCRIPT_DIR, f"결과백업_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jsonl")
logging.basicConfig(
    filename=LOG_PATH,
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
# 에어부산 조회 방식 (v3.18.0)
#   "ext"    : 평소 크롬 + 확장 프로그램으로 자동 조회 (기본값, 확장 없으면 수동확인으로 자동 대체)
#   "manual" : 수동확인 팝업 (v3.17.0 방식)
#   "auto"   : 프로그램이 크롬 조종 (클라우드플레어에 막혀 현재 사용 불가 — 정책 바뀌면 복귀용)
BX_MODE    = "ext"
KE_URL     = "https://www.koreanair.com/reservation/search"
LJ_URL     = "https://www.jinair.com/booking/index"
WE_URL     = "https://www.parataair.com/ko/login/viewLogin.do?tab=2#"
TW_URL     = "https://www.twayair.com/app/reservation/searchMemberBooking"
JJ_URL     = "https://www.jejuair.net/ko/ibe/mypage/viewOnOffReservationList.do"

# 국내/국제선 공항코드 (전역 — 여러 항공사 조회 함수에서 공유)
AIRPORT_CODES = (
    r'PUS|GMP|ICN|CJU|TAE|CJJ|HIN|RSU|KPO|MWX'
    r'|CNX|BKK|HKT|NRT|HND|KIX|NGO|CTS|FUK|OKA'
    r'|DAD|SGN|HAN|CXR|PQC|CEB|KLO|TAG|MNL'
    r'|TPE|HKG|MFM|SIN|DPS|GUM'
)
HEADLESS   = False
DELAY_MIN  = 1.0
DELAY_MAX  = 2.0

DOMESTIC_AIRPORTS = {"PUS","CJU","TAE","CJJ","HIN","RSU","KPO","MWX","GMP","ICN"}


def create_progress_window(total):
    """
    조회 진행 상황을 보여주는 작은 창.
    콘솔 화면이 멈춘 건지 도는 건지 불안할 때 참고용 — 필수는 아니라
    창을 닫아도(X 버튼) 조회 자체는 계속 진행됨.
    """
    win = tk.Toplevel(root)
    win.title("타사 예약 검증 — 진행 상황")
    win.geometry("420x170")
    win.resizable(False, False)
    # 콘솔/브라우저 창을 가리지 않도록 오른쪽 위 구석에 배치
    win.geometry("+%d+%d" % (win.winfo_screenwidth() - 440, 40))

    tk.Label(win, text="조회 진행 중입니다...", font=("맑은 고딕", 11, "bold")).pack(pady=(14, 4))

    count_label = tk.Label(win, text=f"0 / {total}건", font=("맑은 고딕", 10))
    count_label.pack()

    bar = ttk.Progressbar(win, length=380, maximum=total, value=0)
    bar.pack(pady=8)

    current_label = tk.Label(win, text="대기 중...", font=("맑은 고딕", 9), fg="#444",
                              wraplength=380, justify="left")
    current_label.pack()

    summary_label = tk.Label(win, text="확인완료 0 | 불일치 0 | 오류·기타 0",
                              font=("맑은 고딕", 9), fg="#666")
    summary_label.pack(pady=(8, 0))

    win.update_idletasks()
    return {
        "win": win, "bar": bar, "count": count_label,
        "current": current_label, "summary": summary_label,
        "n_ok": 0, "n_mismatch": 0, "n_other": 0,
    }


def update_progress_window(state, i, total, target, result):
    """매 건 처리 직후 진행창 내용 갱신 (실패해도 조회 흐름엔 영향 없음)"""
    if not state:
        return
    try:
        if "확인완료" in result:
            state["n_ok"] += 1
        elif "불일치" in result:
            state["n_mismatch"] += 1
        else:
            state["n_other"] += 1

        state["bar"]["value"] = i
        state["count"].config(text=f"{i} / {total}건")
        state["current"].config(text=f"방금 조회: {target['kor_name']} ({target['airline']}) → {result}")
        state["summary"].config(
            text=f"확인완료 {state['n_ok']} | 불일치 {state['n_mismatch']} | 오류·기타 {state['n_other']}"
        )
        state["win"].update_idletasks()
        state["win"].update()
    except Exception:
        pass  # 창을 닫았거나 갱신 실패해도 조회는 계속 진행


def close_progress_window(state):
    if not state:
        return
    try:
        state["win"].destroy()
    except Exception:
        pass


def ask_resume(done_count, total_count):
    """
    이미 '확인완료'된 건이 있으면 건너뛸지 묻는 팝업.
    반환값: True(건너뛰기) / False(전체 다시 조회)
    """
    result = [None]
    popup = tk.Toplevel(root)
    popup.title("이어서 조회")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()

    tk.Label(
        popup,
        text=(
            f"이미 조회가 끝난 건이 {done_count}건 있습니다.\n"
            f"(전체 {total_count}건 중)\n\n"
            "건너뛰고 나머지만 조회하면 시간을 크게 줄일 수 있습니다.\n"
            "※ 오류·불일치·미조회 건은 건너뛰지 않고 다시 조회합니다.\n\n"
            "어떻게 할까요?"
        ),
        justify="left", padx=20, pady=15
    ).pack()

    btn_frame = tk.Frame(popup)
    btn_frame.pack(pady=(0, 15))

    tk.Button(
        btn_frame, text=f"  이어서 조회 ({total_count - done_count}건)  ", width=22,
        command=lambda: [result.__setitem__(0, True), popup.destroy()]
    ).pack(side="left", padx=6)
    tk.Button(
        btn_frame, text=f"  처음부터 다시 ({total_count}건)  ", width=22,
        command=lambda: [result.__setitem__(0, False), popup.destroy()]
    ).pack(side="left", padx=6)

    popup.wait_window()
    # 팝업을 그냥 닫으면 안전하게 '이어서 조회'로 처리
    return True if result[0] is None else result[0]


def get_check_mode():
    """실행 시 조회 범위 선택 팝업 — 1:5일 / 2:이번달말 / 3:다음달"""
    import calendar
    today = datetime.now()

    # 이번달 말
    this_last = calendar.monthrange(today.year, today.month)[1]
    days_to_eom = this_last - today.day

    # 다음달
    if today.month == 12:
        next_year, next_month = today.year + 1, 1
    else:
        next_year, next_month = today.year, today.month + 1
    next_last = calendar.monthrange(next_year, next_month)[1]

    result = [None]
    popup = tk.Toplevel(root)
    popup.title("조회 범위 선택")
    popup.resizable(False, False)
    popup.grab_set()

    tk.Label(
        popup,
        text=(
            "조회 범위를 선택하세요.\n\n"
            f"  1.  오늘부터 5일\n"
            f"  2.  이번달 말까지 ({today.month}월 {this_last}일, 약 {days_to_eom}일)\n"
            f"  3.  다음달 ({next_year}년 {next_month}월 1일 ~ {next_last}일)\n\n"
            "※ 2·3번은 딜레이가 자동으로 늘어납니다."
        ),
        justify="left",
        padx=20, pady=15
    ).pack()

    btn_frame = tk.Frame(popup)
    btn_frame.pack(pady=(0, 15))

    for n in (1, 2, 3):
        tk.Button(
            btn_frame, text=f"  {n}번  ", width=8,
            command=lambda v=n: [result.__setitem__(0, v), popup.destroy()]
        ).pack(side="left", padx=8)

    popup.wait_window()

    if result[0] == 2:
        end = datetime(today.year, today.month, this_last)
        return "this", today, end, 3.0, 6.0
    elif result[0] == 3:
        start = datetime(next_year, next_month, 1)
        end   = datetime(next_year, next_month, next_last)
        return "next", start, end, 3.0, 6.0
    else:  # 1번 또는 팝업 강제 종료
        end = today + timedelta(days=5)
        return "5d", today, end, 1.0, 2.0


def split_korean_name(name):
    name = name.strip()
    # 동명이인 구분자 제거: 이경수A → 이경수, 박지연B → 박지연
    stripped = re.sub(r'[A-Za-z]+$', '', name).strip()
    # 접미사 제거 후 비었으면(알파벳만 있거나 빈 값) 원본 사용, 그래도 비면 빈 문자열 반환
    base = stripped if stripped else name
    if not base:
        return "", ""
    return base[0], base[1:]


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


def is_within_check_range(dep_time_str, start_date, end_date):
    dep_date = parse_dep_date(str(dep_time_str))
    if not dep_date:
        return True
    return start_date.replace(hour=0, minute=0, second=0, microsecond=0) <= dep_date <= end_date


def is_international(dep, arr):
    return dep not in DOMESTIC_AIRPORTS or arr not in DOMESTIC_AIRPORTS


def load_targets(path, sheet, start_date, end_date):
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
        vals = (list(row) + [None]*9)[:9]
        kor_name, airline, pnr, dep, arr, dep_time, eng_name = vals[:7]
        prev_result, prev_detail = vals[7], vals[8]
        if not all([kor_name, airline, pnr]):
            continue
        if airline not in ("에어부산", "대한항공", "진에어", "제주항공", "파라타항공", "티웨이항공"):
            continue
        if not re.match(r'^[A-Z0-9]{6}$', str(pnr).strip().upper()):
            continue
        if not is_within_check_range(str(dep_time or ""), start_date, end_date):
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
            "prev_result": str(prev_result) if prev_result else "",
            "prev_detail": str(prev_detail) if prev_detail else "",
        })
    return targets


def _append_jsonl_backup(target):
    """
    조회 결과 한 건을 JSONL(한 줄에 하나씩) 형태로 즉시 남겨둠.
    엑셀이 동시접근/비정상종료 등으로 손상되더라도, 이 파일로 결과를 복구할 수 있음.
    실패해도 조회 흐름을 막지 않도록 조용히 넘어감.
    """
    try:
        with open(JSONL_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps({
                "시각"   : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "탑승객" : target.get("kor_name", ""),
                "항공사" : target.get("airline", ""),
                "PNR"   : target.get("pnr", ""),
                "결과"   : target.get("result", ""),
                "내용"   : target.get("detail", ""),
            }, ensure_ascii=False) + "\n")
    except Exception:
        pass


def save_results(path, sheet, targets, silent=False):
    """PNR 기반으로 정확하게 매칭해서 저장 + 확인필요 요약 시트 생성.

    silent=True: 조회 도중 매 건마다 호출되는 "중간 저장" 모드.
      - 결과 컬럼만 빠르게 기록하고, 무거운 요약시트 재생성은 생략.
      - 파일이 열려있어 저장 실패해도 팝업 없이 조용히 넘어감(조회를 막지 않기 위함).
        다음 건 처리 후 다시 저장을 시도하므로, 결국 파일이 닫히는 순간 반영됨.
    silent=False(기본): 조회 종료 시 호출되는 "최종 저장". 요약시트 생성 +
      파일이 열려있으면 팝업으로 안내하고 닫을 때까지 기다림(기존 동작 유지).
    """
    wb = openpyxl.load_workbook(path, keep_vba=True)
    ws = wb[sheet]
    ws.cell(1, RESULT_COL).value = "검증결과"
    ws.cell(1, DETAIL_COL).value = "조회내용"

    # PNR + 항공사 + 이름 조합으로 정확히 매칭
    result_map = {}
    for t in targets:
        key = (t["pnr"], t["airline"], t["kor_name"])
        result_map[key] = (t["result"], t["detail"])

    matched_keys = set()
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        airline  = row[1].value
        pnr      = str(row[2].value).strip().upper() if row[2].value else ""
        kor_name = str(row[0].value).strip() if row[0].value else ""
        if airline not in ("에어부산", "대한항공", "진에어", "제주항공", "파라타항공", "티웨이항공"):
            continue
        key = (pnr, airline, kor_name)
        if key in result_map:
            result, detail = result_map[key]
            ws.cell(row_idx, RESULT_COL).value = result
            ws.cell(row_idx, DETAIL_COL).value = detail
            matched_keys.add(key)

    if silent:
        # 중간 저장: 실패해도 조회 흐름을 막지 않도록 조용히 넘어감
        try:
            wb.save(path)
        except Exception:
            pass
        return

    # 조회는 했으나 엑셀 행과 매칭 안 된 건 경고 (이름/PNR이 조회 중 미묘하게 달라진 경우)
    unmatched = [t for t in targets if (t["pnr"], t["airline"], t["kor_name"]) not in matched_keys]
    if unmatched:
        print(f"\n⚠️  결과 {len(unmatched)}건이 엑셀 행과 매칭되지 않아 기록 못 함:")
        for t in unmatched:
            print(f"     - {t['kor_name']} | {t['airline']} | {t['pnr']}")
        print("   (엑셀 원본의 이름/PNR과 조회 데이터가 다른지 확인 필요)")

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
        if any(kw in res for kw in ["불일치", "예약없음", "PNR오류", "오류", "타임아웃", "수동확인필요", "미조회"]):
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

    # ── 저장 (엑셀 열려있으면 재시도/취소 안내, 취소 시 백업 파일로 저장) ──
    while True:
        try:
            wb.save(path)
            print(f"\n저장 완료: {path}")
            print(f"→ [확인필요_요약] 시트에서 {error_count}건 확인하세요!" if error_count > 0 else "→ 모든 예약 정상!")
            break
        except Exception:
            retry = messagebox.askretrycancel(
                "저장 오류",
                f"엑셀 파일이 열려있거나 저장할 수 없습니다!\n\n{path}\n\n"
                f"[다시 시도] 파일을 닫고 다시 눌러주세요.\n"
                f"[취소] 원본 대신 새 파일로 백업 저장합니다 (결과는 유실되지 않습니다)."
            )
            if not retry:
                folder = os.path.dirname(path)
                stamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
                backup_path = os.path.join(folder, f"검증결과_백업_{stamp}.xlsx")
                try:
                    wb.save(backup_path)
                    print(f"\n⚠️  원본 저장 실패 → 백업 파일로 저장했습니다:")
                    print(f"   {backup_path}")
                    messagebox.showinfo(
                        "백업 저장 완료",
                        f"원본 대신 아래 파일로 저장했습니다:\n\n{backup_path}\n\n"
                        f"내용을 확인하고 원본 파일에 옮겨주세요."
                    )
                except Exception:
                    print(f"\n💥 백업 저장도 실패했습니다. 로그 파일을 확인해주세요: {LOG_PATH}")
                    logging.error("최종 저장 및 백업 저장 모두 실패", exc_info=True)
                break


def _cdp_ready(port):
    import urllib.request
    try:
        # 프록시 무시 — 회사 PC에 프록시가 설정돼 있으면 내 PC(127.0.0.1)로 가는
        # 요청까지 프록시로 보내져서 연결 확인이 실패할 수 있음
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        opener.open(f"http://127.0.0.1:{port}/json/version", timeout=1)
        return True
    except Exception:
        return False


async def _read_devtools_port(profile_dir, timeout_sec=20):
    """
    --remote-debugging-port=0으로 실행하면 크롬이 실제로 사용할 포트를 직접 골라서
    <프로필폴더>/DevToolsActivePort 파일의 첫 줄에 적어둠 (크롬 자체 표준 동작).
    이 파일이 생길 때까지 기다렸다가 포트 번호를 읽어옴.
    """
    port_file = os.path.join(profile_dir, "DevToolsActivePort")
    for _ in range(int(timeout_sec / 0.5)):
        try:
            with open(port_file, "r", encoding="utf-8") as f:
                line = f.readline().strip()
                if line.isdigit():
                    return int(line)
        except FileNotFoundError:
            pass
        await asyncio.sleep(0.5)
    return None


async def _cleanup_old_profiles(p, profile_dir):
    """
    이전 실행에서 남은 조회용 크롬을 닫고 예전 프로필 폴더를 정리 (실패해도 무시).
    - 프로필 폴더 바로 아래(v3.16.0~3.16.5 방식)와 run_* 하위 폴더(v3.16.6~) 모두 확인
    """
    import shutil
    candidates = [profile_dir] + [
        os.path.join(profile_dir, d) for d in os.listdir(profile_dir)
        if d.startswith("run_") and os.path.isdir(os.path.join(profile_dir, d))
    ]
    closed_any = False
    for d in candidates:
        try:
            with open(os.path.join(d, "DevToolsActivePort"), "r", encoding="utf-8") as f:
                line = f.readline().strip()
            if line.isdigit() and _cdp_ready(int(line)):
                br = await p.chromium.connect_over_cdp(f"http://127.0.0.1:{line}")
                try:
                    cdp = await br.new_browser_cdp_session()
                    await cdp.send("Browser.close")
                    closed_any = True
                except Exception:
                    pass
                try:
                    await br.close()
                except Exception:
                    pass
        except Exception:
            pass
    if closed_any:
        print("이전에 남아있던 조회용 크롬을 닫았습니다.")
        await asyncio.sleep(2)  # 크롬이 파일 잠금을 풀 시간
    for d in candidates[1:]:
        shutil.rmtree(d, ignore_errors=True)


async def attach_real_chrome(p, chrome_exe, profile_dir):
    """
    (v3.16.0) 크롬을 '일반 실행'으로 먼저 띄우고 → 사용자가 에어부산 보안확인을 직접 통과 →
    그 다음에 프로그램이 그 창에 붙어서 조회. 보안확인 시점엔 자동화 연결이 전혀 없으므로
    일반 사용자와 완전히 동일한 상태에서 통과됨.
    (v3.16.1) 디버그 포트를 9222로 고정하지 않고 크롬이 직접 랜덤 포트를 고르게 함
    — 9222는 자동화 도구들이 쓰는 가장 잘 알려진 기본값이라, 보안검사 스크립트가
    이 포트가 열려있는지 페이지 안에서 탐지해 자동화로 판단할 가능성이 있음.
    실패 시 (None, None, None) 반환 → 호출부에서 기존 방식으로 대체.
    """
    os.makedirs(profile_dir, exist_ok=True)

    # (v3.16.6) 매 실행마다 "완전히 새 프로필"로 크롬을 띄움.
    # 같은 프로필을 계속 쓰면, 여러 번 실패한 보안확인 기록(쿠키)이 쌓여서 클라우드플레어가
    # 그 프로필 자체를 의심 → 평소 크롬에선 통과되는데 조회용 크롬에서만 체크가 무한반복됨.
    # 이전 실행에서 남은 조회용 크롬은 먼저 닫고, 예전 프로필은 지움.
    await _cleanup_old_profiles(p, profile_dir)
    run_dir = os.path.join(profile_dir, f"run_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
    os.makedirs(run_dir, exist_ok=True)

    try:
        proc = subprocess.Popen([
            chrome_exe,
            "--remote-debugging-port=0",   # 0 = 크롬이 알아서 안 쓰는 포트를 랜덤으로 선택
            f"--user-data-dir={run_dir}",
            "--no-first-run",
            "--no-default-browser-check",
            "--window-size=1280,800",
            BX_URL,
        ])
    except Exception as e:
        logging.warning(f"크롬 일반 실행 실패: {e}")
        return None, None, None

    cdp_port = await _read_devtools_port(run_dir)
    if cdp_port is None or not _cdp_ready(cdp_port):
        logging.warning("크롬 연결 포트 확인 실패")
        print("⚠️  조회용 크롬 연결 실패 → 기존 방식으로 실행합니다.")
        try:
            proc.terminate()
        except Exception:
            pass
        return None, None, None

    print("\n" + "=" * 50)
    print("  [에어부산] 열린 크롬 창에서 보안확인을 직접 통과해주세요.")
    print("  → 예약조회 화면이 보이면 팝업의 [확인]을 누르세요.")
    print("=" * 50)
    messagebox.showinfo(
        "에어부산 보안확인",
        "방금 열린 크롬 창에서\n\n"
        "1. '사람인지 확인'이 뜨면 체크해서 통과\n"
        "2. 에어부산 예약조회 화면이 정상적으로 보이면\n"
        "   이 창의 [확인]을 눌러주세요.\n\n"
        "※ 크롬 창은 닫지 마세요. 조회가 끝나면 자동으로 닫힙니다."
    )
    try:
        browser = await p.chromium.connect_over_cdp(f"http://127.0.0.1:{cdp_port}")
        context = browser.contexts[0] if browser.contexts else await browser.new_context()
        print("✅ 보안확인 통과된 크롬 창에 연결 완료\n")
        return context, browser, proc
    except Exception as e:
        logging.warning(f"크롬 연결(CDP) 실패: {e}")
        print("⚠️  크롬 연결 실패 → 기존 방식으로 실행합니다.")
        try:
            proc.terminate()
        except Exception:
            pass
        return None, None, None


def _reraise_if_closed(exc):
    """
    (v3.16.3) 브라우저/컨텍스트 자체가 닫혀서 난 오류는 이번 한 건만의 문제가
    아니라 이후 모든 조회가 똑같이 실패할 상황임. 그런데 각 check_* 함수의
    'except Exception'이 이런 오류까지 "💥 오류"로 뭉뚱그려버리면, 메인 루프가
    "브라우저 닫힘"을 감지 못 하고 남은 건들을 하나씩 헛되이 재시도하게 됨.
    → 메시지에 closed가 들어있으면 삼키지 않고 그대로 다시 던져서
      메인 루프가 즉시 감지하고 남은 건을 "미조회"로 정리하도록 함.
    """
    if "closed" in str(exc).lower():
        raise exc


def _parse_bx_text(html_content, target, pnr):
    """에어부산 결과 화면 텍스트 → (결과, 상세). 자동조회·확장조회 공용 (v3.18.0 분리)"""
    # 결과 파싱 전 클라우드플레어 재체크
    if any(kw in html_content for kw in CF_KEYWORDS):
        return "⏱️ 타임아웃", "클라우드플레어 차단 → 재실행 필요"
    if any(kw in html_content for kw in ["조회 결과가 없", "예약 내역이 없", "일치하는 예약"]):
        return "❌ 예약없음", "조회결과 없음"

    # 조회한 PNR이 실제로 결과 화면에 있는지 확인 (잔류 화면 오판정 방지 2중 안전장치)
    if pnr.upper() not in html_content.upper():
        return "❌ PNR오류", "조회한 PNR이 결과 화면에서 확인되지 않음"

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

    if not _is_reliable_result(flt_found, route_found):
        return "❌ PNR오류", f"예약 확인 불가 (편명/구간 모두 미확인) | {detail}"

    return "✅ 확인완료", detail


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

    # 조회마다 새 탭 사용 — 이전 사람의 결과 화면이 남아있는 상태에서
    # 이번 조회가 실패하면 그 잔류 화면을 잘못 읽어버리는 문제 방지
    bx_page = await page.context.new_page()
    try:
        await bx_page.goto(BX_URL, wait_until="domcontentloaded", timeout=20000)
        await bx_page.wait_for_timeout(1000)

        # 클라우드플레어 감지 → 사람이 직접 캡챠 풀도록 안내
        body_check = await bx_page.inner_text("body")
        if any(kw in body_check for kw in CF_KEYWORDS):
            print(f"\n{'='*50}")
            print(f"  ⚠️  [에어부산] 클라우드플레어 보안 확인이 필요합니다!")
            print(f"  → 열린 브라우저에서 '사람인지 확인하십시오' 체크박스를 클릭해주세요.")
            print(f"  → 화면이 멈춘 것처럼 안 넘어가면 새로고침(F5)해주세요 — 자동으로도")
            print(f"    주기적으로 새로고침을 시도합니다.")
            print(f"  → 통과되면 자동으로 이어서 진행됩니다 (최대 3분 대기).")
            print(f"{'='*50}")
            try:
                await bx_page.bring_to_front()
            except Exception:
                pass
            # 통과 여부 자동 감지 (v3.14.0: 콘솔로 돌아와 엔터 누를 필요 없음)
            # (v3.16.2) 확인 화면이 멈춘 채로 안 넘어가는 경우가 있어 — 사람이 직접
            # 새로고침하면 넘어가는 게 확인된 증상 → 30초마다 자동으로 새로고침도 시도
            # (15초는 사람이 체크하는 도중에 새로고침돼 끊길 수 있어 v3.16.5에서 30초로)
            # (사용자가 그 사이에 직접 체크박스를 누르고 있어도 문제 없음)
            passed = False
            for i in range(180):
                await bx_page.wait_for_timeout(1000)
                if i > 0 and i % 30 == 0:
                    try:
                        await bx_page.reload(wait_until="domcontentloaded", timeout=10000)
                    except Exception:
                        pass
                try:
                    body_check2 = await bx_page.inner_text("body")
                except Exception:
                    continue  # 새로고침/통과 직후 페이지 전환 중
                if not any(kw in body_check2 for kw in CF_KEYWORDS):
                    passed = True
                    break
            if not passed:
                return "⏱️ 타임아웃", "클라우드플레어 차단 미해제 → 재실행 필요"
            print("  ✅ 보안확인 통과 → 조회 계속")
            await bx_page.wait_for_load_state("domcontentloaded")
            await bx_page.wait_for_timeout(1000)

        await bx_page.click("text=예약번호로 조회", timeout=5000)
        await bx_page.wait_for_timeout(500)

        # Playwright의 fill()은 요소가 조작 가능해질 때까지 자동으로 기다리므로
        # 입력 사이 고정 대기는 불필요 (v3.13.0에서 제거 — 조회 속도 개선)
        await bx_page.locator("input[placeholder*='예약번호']").first.fill(pnr)
        await bx_page.locator("input[placeholder='성']").first.fill(last)
        await bx_page.locator("input[placeholder='이름']").first.fill(first)

        await bx_page.evaluate("document.querySelector('.buttonOfflineCheckin').click()")
        await bx_page.wait_for_timeout(1500)

        body_text = await bx_page.inner_text("body")
        if "해당 예약번호가 확인되지 않습니다" in body_text:
            try:
                await bx_page.click("button:has-text('확인')", timeout=2000)
            except:
                pass
            return "❌ PNR오류", "해당 예약번호 확인 불가"

        try:
            await bx_page.wait_for_selector("text=항공권 구매완료", timeout=10000)
        except:
            pass
        await bx_page.wait_for_timeout(800)

        html_content = await bx_page.inner_text("body")

        return _parse_bx_text(html_content, target, pnr)

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception as e:
        _reraise_if_closed(e)
        logging.error(f"에어부산 조회 실패 | PNR: {pnr} | 탑승객: {last}{first}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"
    finally:
        try:
            await bx_page.close()
        except Exception:
            pass


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

    ke_page = await page.context.new_page()
    try:
        await ke_page.goto(KE_URL, wait_until="domcontentloaded", timeout=20000)
        await ke_page.wait_for_timeout(1200)

        try:
            await ke_page.click("button:has-text('동의합니다')", timeout=3000)
            await ke_page.wait_for_timeout(500)
        except:
            pass

        try:
            await ke_page.click('button[data-click-name="Accept all"]', timeout=3000)
            await ke_page.wait_for_timeout(500)
        except:
            try:
                await ke_page.click("button:has-text('모든 쿠키 허용')", timeout=2000)
                await ke_page.wait_for_timeout(500)
            except:
                pass

        # fill()은 자동으로 요소 준비상태를 기다리므로 입력 직후 고정대기 제거
        await ke_page.locator("input[maxlength='13']").first.fill(pnr)

        await ke_page.click("button[data-dialog-id='#dialog-datepicker1']", timeout=5000)
        await ke_page.wait_for_timeout(1000)

        dep_day   = str(dep_date.day)
        dep_month = dep_date.month
        dep_year  = dep_date.year

        await ke_page.evaluate(f"""
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
        await ke_page.wait_for_timeout(500)

        await ke_page.locator("input[autocomplete='family-name']").first.fill(last)
        await ke_page.locator("input[autocomplete='given-name']").first.fill(first)

        await ke_page.click("button:has-text('조회')", timeout=5000)

        try:
            await ke_page.wait_for_selector(".journey-info__date", timeout=20000)
        except:
            pass
        await ke_page.wait_for_timeout(1500)

        html_content = await ke_page.inner_text("body")

        if any(kw in html_content for kw in ["조회 결과가 없", "예약을 찾을 수 없", "확인되지 않", "일치하는 예약"]):
            return "❌ PNR오류", "예약 확인 불가"

        # 조회한 PNR이 실제로 결과 화면에 있는지 확인 (잔류 화면 오판정 방지 2중 안전장치)
        if pnr.upper() not in html_content.upper():
            return "❌ PNR오류", "조회한 PNR이 결과 화면에서 확인되지 않음"

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

        if not _is_reliable_result(flt_found, route_found):
            return "❌ PNR오류", f"예약 확인 불가 (편명/구간 모두 미확인) | {detail}"

        return "✅ 확인완료", detail

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception as e:
        _reraise_if_closed(e)
        logging.error(f"대한항공 조회 실패 | PNR: {pnr} | 탑승객: {last}{first}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"
    finally:
        try:
            await ke_page.close()
        except Exception:
            pass


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

    # 진에어는 SPA 구조 — 이전 조회 결과가 페이지에 잔류하므로 매 조회마다 새 탭 사용
    lj_page = await page.context.new_page()
    try:
        await lj_page.goto(LJ_URL, wait_until="domcontentloaded", timeout=20000)
        await lj_page.wait_for_timeout(1200)

        await lj_page.click("text=예약조회", timeout=5000)
        await lj_page.wait_for_timeout(500)

        # fill()은 요소 준비상태를 자동 대기하므로 입력 사이 고정대기 제거
        await lj_page.locator("input[placeholder*='6자리']").first.fill(pnr)
        await lj_page.locator("#lastName_resv").fill(input_last)
        await lj_page.locator("#firstName_resv").fill(input_first)

        await lj_page.click("#departureDate_resv", timeout=5000)
        await lj_page.wait_for_timeout(1200)

        dep_str = dep_date.strftime("%Y.%m.%d")
        iframe_locator = lj_page.frame_locator("iframe[src*='basicCalendarLayer']")

        try:
            date_input = iframe_locator.locator(f"input[name='hiddenDate'][value^='{dep_str}']")
            await date_input.evaluate("el => el.parentElement.click()")
        except:
            pass

        await lj_page.wait_for_timeout(500)
        await lj_page.click("button[role='login-button']", timeout=5000)

        # 진에어 내부 예약번호 ≠ 조회 PNR — "여정 예약정보" 텍스트로 성공 판정
        try:
            await lj_page.wait_for_selector("text=여정 예약정보", timeout=15000)
        except:
            pass
        await lj_page.wait_for_timeout(1500)

        html_content = await lj_page.inner_text("body")

        # CF 재체크
        if any(kw in html_content for kw in CF_KEYWORDS):
            return "⏱️ 타임아웃", "클라우드플레어 차단 → 재실행 필요"

        # 성공 판정: "여정 예약정보" 존재 여부
        if "여정 예약정보" not in html_content:
            if any(kw in html_content for kw in ["조회 결과가 없", "예약 내역이 없", "확인되지 않"]):
                return "❌ PNR오류", "예약 확인 불가"
            return "❌ PNR오류", "예약 확인 불가 (PNR 미조회)"

        flt_match = re.search(r'LJ\d{3,4}', html_content)
        flt_found = flt_match.group() if flt_match else "편명미확인"

        # 페이지에 "예약일 YYYY.MM.DD(요일)"가 실제 출발일보다 먼저 나와서
        # 첫 매치만 쓰면 예약일을 출발일로 착각함 — 예약일 다음 매치를 사용
        date_matches = re.findall(r'(\d{4})\.(\d{2})\.(\d{2})\(', html_content)
        if len(date_matches) >= 2:
            y, m, d = date_matches[1]
            date_found = f"{y}-{m.zfill(2)}-{d.zfill(2)}"
        elif len(date_matches) == 1:
            y, m, d = date_matches[0]
            date_found = f"{y}-{m.zfill(2)}-{d.zfill(2)}"
        else:
            date_found = "날짜미확인"

        airports = re.findall(
            r'(?<![A-Z0-9])(PUS|GMP|ICN|CJU|TAE|CJJ|HIN|RSU|KPO|MWX'
            r'|CNX|BKK|HKT|NRT|HND|KIX|NGO|CTS|FUK|OKA'
            r'|DAD|SGN|HAN|CXR|PQC|CEB|KLO|TAG|MNL'
            r'|TPE|HKG|MFM|SIN|DPS|GUM)(?![A-Z0-9])',
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

        if not _is_reliable_result(flt_found, route_found):
            return "❌ PNR오류", f"예약 확인 불가 (편명/구간 모두 미확인) | {detail}"

        return "✅ 확인완료", detail

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception as e:
        _reraise_if_closed(e)
        logging.error(f"진에어 조회 실패 | PNR: {pnr} | 탑승객: {input_last}{input_first}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"
    finally:
        try:
            await lj_page.close()
        except Exception:
            pass


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
        await page.wait_for_timeout(1000)

        # 홈페이지 비회원 탭 클릭
        await page.click("a[href='#nonmember']", timeout=5000)
        await page.wait_for_timeout(500)

        # 이메일 + 예약번호 입력 (fill()이 요소 준비상태 자동대기하므로 사이 고정대기 제거)
        await page.fill("#userEmail", we_email)
        await page.fill("#reservationNum", pnr)

        # 예약조회 버튼 클릭
        await page.click("#nonMemberResvSearchBtn", timeout=5000)

        # 결과 페이지 대기
        try:
            await page.wait_for_url("**/viewReservationDetail.do**", timeout=15000)
        except:
            pass
        await page.wait_for_timeout(1200)

        current_url = page.url
        html_content = await page.inner_text("body")

        # 실패 판정
        if "viewReservationDetail.do" not in current_url:
            if dialog_msgs:
                return "❌ PNR오류", f"알림: {dialog_msgs[0][:40]}"
            if any(kw in html_content for kw in ["일치하는 예약", "확인되지 않", "조회 결과가 없", "예약 내역이 없"]):
                return "❌ PNR오류", "예약 확인 불가"
            return "💥 오류", "결과 페이지 이동 실패"

        # 조회한 PNR이 실제로 결과 화면에 있는지 확인 (잔류 화면 오판정 방지 2중 안전장치)
        if pnr.upper() not in html_content.upper():
            return "❌ PNR오류", "조회한 PNR이 결과 화면에서 확인되지 않음"

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

        if not _is_reliable_result(flt_found, route_found):
            return "❌ PNR오류", f"예약 확인 불가 (편명/구간 모두 미확인) | {detail}"

        return "✅ 확인완료", detail

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception as e:
        _reraise_if_closed(e)
        logging.error(f"파라타항공 조회 실패 | PNR: {pnr} | 탑승객: {kor_name}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"
    finally:
        try:
            page.remove_listener("dialog", _on_dialog)
        except Exception:
            pass


async def _manual_check(page, target, airline_label, url, fields, guide_extra=""):
    """
    수동 확인 공용 팝업 — 평소 쓰는 기본 브라우저로 조회 페이지를 열고,
    항목별 [복사] 버튼으로 사람이 직접 입력·확인한 뒤 결과를 선택.
    (티웨이: Akamai 차단 / 에어부산: v3.17.0부터 클라우드플레어 차단으로 사용)
    """
    # 자동화 창은 수동확인에 쓰지 않음 — 혼동 방지용 안내 문구 표시
    try:
        await page.goto(
            "data:text/html,"
            "<html><body style='font-family:sans-serif;padding:60px;"
            "font-size:22px;color:#333;text-align:center;'>"
            "이 창은 자동화 전용입니다.<br><br>"
            f"{airline_label} 조회는<br>"
            "<b>새로 열린 별도의 브라우저 창</b>에서 진행해주세요."
            "</body></html>",
            timeout=5000
        )
    except Exception:
        pass  # 안내 문구 표시 실패해도 조회 자체엔 영향 없음

    # 조회 페이지를 시스템 기본 브라우저의 새 창으로 오픈
    webbrowser.open(url, new=2)

    result_box = [None]
    popup = tk.Toplevel()
    popup.title(f"{airline_label} 수동 확인")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()

    tk.Label(
        popup,
        text=(
            f"{airline_label}은(는) 보안 정책상 자동 조회가 불가합니다.\n"
            "방금 새로 열린 별도의 브라우저 창에서\n"
            "(자동화 창 아님 — about:blank 창은 무시하세요)\n"
            "아래 항목을 [복사] 버튼으로 복사해 붙여넣어\n"
            "직접 조회한 뒤 결과를 선택해주세요."
            + guide_extra
        ),
        justify="left", padx=20
    ).pack(pady=(15, 8))

    field_frame = tk.Frame(popup)
    field_frame.pack(padx=20, pady=5)

    def _copy(value):
        popup.clipboard_clear()
        popup.clipboard_append(value)

    def _add_field(row, label, value):
        tk.Label(field_frame, text=label, width=8, anchor="w").grid(row=row, column=0, sticky="w", pady=3)
        e = tk.Entry(field_frame, width=26)
        e.insert(0, value)
        e.config(state="readonly")
        e.grid(row=row, column=1, padx=6)
        tk.Button(field_frame, text="복사", width=6,
                  command=lambda v=value: _copy(v)).grid(row=row, column=2)

    for i, (label, value) in enumerate(fields):
        _add_field(i, label, value)

    btn_frame = tk.Frame(popup)
    btn_frame.pack(pady=(10, 15))

    def _choose(v):
        result_box[0] = v
        popup.destroy()

    tk.Button(btn_frame, text="✅ 확인완료", width=12,
              command=lambda: _choose("ok")).pack(side="left", padx=6)
    tk.Button(btn_frame, text="❌ PNR오류", width=12,
              command=lambda: _choose("pnr_error")).pack(side="left", padx=6)
    tk.Button(btn_frame, text="⏭ 보류(건너뛰기)", width=14,
              command=lambda: _choose("skip")).pack(side="left", padx=6)

    popup.wait_window()

    if result_box[0] == "ok":
        return "✅ 확인완료", "[수동확인] 예약 확인됨"
    elif result_box[0] == "pnr_error":
        return "❌ PNR오류", "[수동확인] 예약 확인 불가"
    else:
        return "⚠️ 수동확인필요", "[수동확인] 보류됨 — 재확인 필요"


async def check_tw(page, target):
    """
    티웨이항공 — Akamai 봇 차단으로 Playwright 자동 조회/자동입력 불가.
    Chrome 탭을 열고, 팝업에서 항목별 [복사] 버튼으로 사람이 직접 붙여넣도록 함.
    (자동 붙여넣기는 Akamai에 감지되어 사용 불가)
    """
    pnr      = target["pnr"]
    kor_name = target["kor_name"]
    eng_name = target.get("eng_name", "")
    dep_time = target.get("dep_time", "")
    dep      = target.get("dep", "")
    arr      = target.get("arr", "")

    fields = [("PNR", pnr), ("한글성명", kor_name)]
    if eng_name:
        parts = eng_name.split("/")
        eng_last  = parts[0].strip() if len(parts) >= 1 else eng_name
        eng_first = parts[1].strip() if len(parts) >= 2 else ""
        fields.append(("영문성", eng_last))
        if eng_first:
            fields.append(("영문이름", eng_first))
    fields.append(("구간", f"{dep} → {arr}"))
    fields.append(("출발일", dep_time))

    return await _manual_check(page, target, "티웨이항공", TW_URL, fields)


# ==========================================
# (v3.18.0) 에어부산 — 크롬 확장 프로그램 연동
# 디버그 포트가 열린 크롬은 클라우드플레어를 못 지나가므로, 사용자가 평소 쓰는 크롬에
# 설치된 확장("DH 조회 도우미")이 입력·조회를 하고, 결과 화면 글자를 여기로 보내줌.
# 판정은 기존 파이썬 로직(_parse_bx_text) 그대로 사용.
# 통신: 이 프로그램이 내 PC 안에서만 열리는 작은 창구(127.0.0.1:EXT_PORT)를 열고,
#       확장이 1초마다 "다음 건 있어?"를 물어보고 결과를 돌려주는 구조
# ==========================================
EXT_PORT = 38517


class _ExtBridge:
    def __init__(self):
        self.lock      = threading.Lock()
        self.current   = None   # 확장에 넘길 조회 1건
        self.handed_at = None   # 확장이 가져간 시각
        self.results   = {}     # id → 결과 화면 텍스트
        self.last_poll = 0.0    # 확장이 마지막으로 물어본 시각
        self.server    = None
        self.available = False  # 확장이 실제로 동작 중인지

    def start(self):
        import http.server
        bridge = self

        class Handler(http.server.BaseHTTPRequestHandler):
            def log_message(self, *a):
                pass

            def _send(self, obj):
                data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "application/json; charset=utf-8")
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Access-Control-Allow-Headers", "Content-Type")
                self.send_header("Content-Length", str(len(data)))
                self.end_headers()
                self.wfile.write(data)

            def do_OPTIONS(self):
                self._send({})

            def do_GET(self):
                if not self.path.startswith("/next"):
                    return self._send({"wait": True})
                with bridge.lock:
                    bridge.last_poll = time.time()
                    cur = bridge.current
                    # 아직 안 가져갔거나, 가져간 지 40초가 지나도 결과가 없으면(페이지 꼬임) 다시 넘김
                    if cur and (bridge.handed_at is None or time.time() - bridge.handed_at > 40):
                        bridge.handed_at = time.time()
                        return self._send({"target": cur})
                return self._send({"wait": True})

            def do_POST(self):
                try:
                    n = int(self.headers.get("Content-Length", 0))
                    body = json.loads(self.rfile.read(n).decode("utf-8"))
                    with bridge.lock:
                        bridge.last_poll = time.time()
                        if bridge.current and body.get("id") == bridge.current["id"]:
                            bridge.results[body["id"]] = body.get("text", "")
                except Exception:
                    pass
                self._send({"ok": True})

        try:
            self.server = http.server.ThreadingHTTPServer(("127.0.0.1", EXT_PORT), Handler)
        except OSError as e:
            logging.warning(f"확장 연동 창구 열기 실패 (포트 {EXT_PORT} 사용 중): {e}")
            return False
        threading.Thread(target=self.server.serve_forever, daemon=True).start()
        return True

    def stop(self):
        if self.server:
            try:
                self.server.shutdown()
                self.server.server_close()
            except Exception:
                pass
            self.server = None

    def seen_recently(self, sec):
        return time.time() - self.last_poll < sec


EXT = _ExtBridge()


async def setup_bx_extension():
    """에어부산 건이 있을 때 조회 시작 전에 1회: 창구 열기 → 평소 크롬에 에어부산 열기 → 확장 동작 확인"""
    if not EXT.start():
        print("⚠️  확장 연동 창구를 열 수 없음 (프로그램이 이미 켜져 있는지 확인) → 에어부산은 수동확인으로 진행")
        return
    webbrowser.open(BX_URL, new=2)
    print("\n" + "=" * 50)
    print("  [에어부산] 평소 쓰는 크롬에 에어부산 창이 열렸습니다.")
    print("  → '사람인지 확인'이 뜨면 체크해서 통과한 뒤 팝업의 [확인]을 누르세요.")
    print("  → 조회가 끝날 때까지 그 에어부산 탭은 닫지 마세요.")
    print("=" * 50)
    messagebox.showinfo(
        "에어부산 준비",
        "평소 쓰는 크롬에 에어부산 창이 열렸습니다.\n\n"
        "1. '사람인지 확인'이 뜨면 체크해서 통과\n"
        "2. 예약조회 화면이 보이면 이 창의 [확인]\n\n"
        "※ 조회가 끝날 때까지 그 에어부산 탭은 닫지 마세요.\n"
        "   (에어부산 입력·조회는 확장 프로그램이 자동으로 합니다)"
    )
    for _ in range(20):  # 최대 10초 동안 확장 응답 확인
        if EXT.seen_recently(5):
            EXT.available = True
            print("✅ 에어부산 확장 연결 확인 — 에어부산은 평소 크롬에서 자동 조회합니다\n")
            return
        await asyncio.sleep(0.5)
    print("⚠️  확장 프로그램 응답 없음 → 에어부산은 수동확인으로 진행합니다.")
    print("   (확장 미설치, 또는 에어부산 탭이 예약조회 화면이 아닌 경우)\n")


async def check_bx_ext(page, target):
    """에어부산 확장 조회 1건. 확장이 없으면 수동확인으로 대체"""
    if not EXT.available:
        return await check_bx_manual(page, target)

    pnr      = target["pnr"]
    eng_name = target.get("eng_name", "")
    if is_international(target["dep"], target["arr"]) and eng_name:
        parts = eng_name.split("/")
        last  = parts[0].strip() if len(parts) >= 1 else target["last"]
        first = parts[1].strip() if len(parts) >= 2 else target["first"]
    else:
        last  = target["last"]
        first = target["first"]

    tid = uuid.uuid4().hex
    with EXT.lock:
        EXT.current   = {"id": tid, "pnr": pnr, "last": last, "first": first}
        EXT.handed_at = None
    try:
        reopened = False
        start = time.time()
        while time.time() - start < 120:
            with EXT.lock:
                text = EXT.results.pop(tid, None)
                handed = EXT.handed_at is not None
            if text is not None:
                if text == "__FIELD_NOT_FOUND__":
                    return "💥 오류", "확장: 입력칸을 찾지 못함 (에어부산 화면 변경 가능성)"
                if "해당 예약번호가 확인되지 않습니다" in text:
                    return "❌ PNR오류", "해당 예약번호 확인 불가"
                return _parse_bx_text(text, target, pnr)
            # 20초 동안 확장이 안 가져가고 연락도 없으면 → 에어부산 탭이 닫힌 것으로 보고 다시 열기
            if not handed and not reopened and time.time() - start > 20 and not EXT.seen_recently(15):
                print("\n  ⚠️  에어부산 탭 응답 없음 → 다시 엽니다 (보안확인이 뜨면 통과해주세요)")
                webbrowser.open(BX_URL, new=2)
                reopened = True
            await asyncio.sleep(0.5)
        return "⏱️ 타임아웃", "확장 응답 없음 — 에어부산 탭이 열려있는지 확인"
    finally:
        with EXT.lock:
            EXT.current = None
            EXT.results.pop(tid, None)


async def check_bx_manual(page, target):
    """
    (v3.17.0) 에어부산 수동 확인.
    테스트 결과 디버그 포트가 열린 크롬(= 프로그램이 조종 가능한 모든 크롬)은
    클라우드플레어 보안확인을 통과하지 못함 — 통과 쿠키가 있어도 다시 막힘.
    평소 쓰는 크롬에서는 정상 통과되므로 티웨이와 같은 수동 확인 방식으로 전환.
    입력값(성/이름)은 자동조회 때와 동일한 규칙: 국제선+영문명 있으면 영문, 아니면 한글
    """
    pnr      = target["pnr"]
    eng_name = target.get("eng_name", "")
    dep      = target.get("dep", "")
    arr      = target.get("arr", "")

    if is_international(dep, arr) and eng_name:
        parts = eng_name.split("/")
        last  = parts[0].strip() if len(parts) >= 1 else target["last"]
        first = parts[1].strip() if len(parts) >= 2 else target["first"]
    else:
        last  = target["last"]
        first = target["first"]

    fields = [
        ("PNR", pnr),
        ("성", last),
        ("이름", first),
        ("구간", f"{dep} → {arr}"),
        ("출발일", target.get("dep_time", "")),
    ]
    guide = (
        "\n\n※ '사람인지 확인'이 뜨면 체크해서 통과한 뒤 조회하세요.\n"
        "   [예약번호로 조회] 탭에서 입력하면 됩니다."
    )
    return await _manual_check(page, target, "에어부산", BX_URL, fields, guide)


async def _dismiss_ad_popup(p):
    """
    그루비 광고 팝업 닫기 — 확인된 실물 셀렉터 하나만 정확히 겨냥.
    실물 img 태그: <img src="...groobee.io/image/close/..." alt="닫기" class="img_999999">
    실패해도 조회에 영향 없게 조용히 넘어감.
    """
    close_selector = "img[src*='groobee.io/image/close']"

    async def _try_close(scope):
        try:
            loc = scope.locator(close_selector).first
            if not await loc.is_visible(timeout=500):
                return False
            # 1) 일반 클릭 시도
            try:
                await loc.click(timeout=800)
                return True
            except Exception:
                pass
            # 2) 부모 요소 클릭 (img가 클릭 이벤트를 부모에 위임하는 경우)
            try:
                await loc.evaluate("el => (el.closest('a,button,div[onclick]') || el.parentElement).click()")
                return True
            except Exception:
                pass
            # 3) JS 직접 클릭 (오버레이/z-index 무시)
            try:
                await loc.evaluate("el => el.click()")
                return True
            except Exception:
                pass
        except Exception:
            pass
        return False

    # 최대 3회 시도 (팝업이 여러 개 겹쳐 있는 경우 대비)
    for _ in range(3):
        closed_any = False
        # 메인 프레임 + 모든 iframe 순회 (그루비는 iframe에 들어있기도 함)
        for scope in [p] + list(p.frames):
            if await _try_close(scope):
                closed_any = True
                await p.wait_for_timeout(400)
                break
        if not closed_any:
            break


def _prompt_jj_calendar(pnr, kor_name, target_date_display):
    """
    제주항공 달력 자동 클릭이 불안정해서, 입력값은 자동으로 채워두고
    달력 날짜 선택만 사람이 직접 하도록 안내하는 팝업.
    나머지(조회, 결과판정)는 계속 자동 진행됨.
    """
    result_box = [None]
    popup = tk.Toplevel()
    popup.title("제주항공 - 날짜 선택 필요")
    popup.resizable(False, False)
    popup.attributes("-topmost", True)
    popup.grab_set()

    tk.Label(
        popup,
        text=(
            "※ 제주항공은 보안 정책상 날짜 자동 선택이 안 됩니다.\n"
            "   (오류 아님 — 아래 순서대로만 해주시면 됩니다)\n\n"
            f"  탑승객 : {kor_name}\n"
            f"  PNR    : {pnr}\n"
            f"  출발일 : {target_date_display}\n\n"
            "① 지금 열려있는 자동화 Chrome 창을 클릭해서 앞으로 가져오세요\n"
            "   (이미 달력이 떠 있는 상태입니다)\n"
            "② 달력에서 위 '출발일' 날짜만 클릭하세요\n"
            "③ 아래 [완료] 버튼을 눌러주세요\n"
            "   → '선택'/'조회' 버튼 클릭 및 나머지는 자동으로 진행됩니다\n\n"
            "(혹시 실수로 선택/조회까지 눌러도 문제없이 진행됩니다)\n"
            "날짜가 헷갈리거나 실수했으면 [건너뛰기]를 눌러주세요."
        ),
        justify="left", padx=20, pady=15
    ).pack()

    btn_frame = tk.Frame(popup)
    btn_frame.pack(pady=(0, 15))

    def _choose(v):
        result_box[0] = v
        popup.destroy()

    tk.Button(btn_frame, text="✅ 완료 — 계속 진행", width=16,
              command=lambda: _choose("done")).pack(side="left", padx=6)
    tk.Button(btn_frame, text="⏭ 건너뛰기", width=12,
              command=lambda: _choose("skip")).pack(side="left", padx=6)

    popup.wait_window()
    return result_box[0]


async def check_jj(page, target):
    """
    제주항공 — 조회 페이지(viewOnOffReservationList.do)에서 입력 후
    결과 페이지(viewReservationDetail.do)로 이동.
    ※ 진입 시 마케팅 팝업이 뜰 수 있어 자동 닫기 시도 후 진행.
    ※ 달력 날짜 선택은 "일(day) 숫자" 텍스트 매칭 방식 — 검증 필요.
    """
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

    jj_page = await page.context.new_page()
    try:
        await jj_page.goto(JJ_URL, wait_until="domcontentloaded", timeout=20000)
        await jj_page.wait_for_timeout(1000)

        # 진입 시 뜨는 마케팅 팝업(그루비 등)이 폼을 가려 클릭이 막히는 문제 방지
        await _dismiss_ad_popup(jj_page)

        await jj_page.locator("#recordLocatorLabel").fill(pnr)
        await jj_page.locator("#psInputLastName_1").fill(input_last)

        await jj_page.locator("#psInputFirstName_1").fill(input_first)
        await jj_page.locator("#psInputFirstName_1").press("Tab")
        await jj_page.wait_for_timeout(300)

        # 탑승일자 달력 선택 — 자동 클릭이 계속 불안정해서 사람이 직접 선택하도록 전환
        # (입력값은 이미 자동으로 채워짐, 달력 클릭만 사람이 하고 나머지는 자동 진행)
        await jj_page.click("#boardingDateBtn", timeout=5000)
        await jj_page.wait_for_selector("#datepicker01", timeout=5000)
        await jj_page.wait_for_timeout(500)
        await _dismiss_ad_popup(jj_page)

        target_date_display = dep_date.strftime("%Y-%m-%d (%a)")
        choice = _prompt_jj_calendar(pnr, target["kor_name"], target_date_display)

        if choice != "done":
            return "⚠️ 수동확인필요", "[제주달력] 사용자가 건너뜀"

        # 광고가 다시 떴을 수 있어 한 번 더 정리
        await _dismiss_ad_popup(jj_page)

        # 사람이 사이트 "선택" 버튼을 안 눌렀을 경우 대비해서 자동화가 눌러줌 (있으면)
        try:
            choose_btn = jj_page.locator("#chooseDepDateBtn")
            if await choose_btn.is_visible(timeout=500):
                await choose_btn.click(timeout=2000)
                await jj_page.wait_for_timeout(500)
        except Exception:
            pass

        # 상태 확인 — 두 가지 경우 모두 정상 진행:
        # (1) 검색 폼에서 날짜가 채워짐  (2) 이미 조회까지 눌러서 결과 페이지로 넘어감
        date_selected = False
        already_searched = False
        try:
            selectdate_val = await jj_page.evaluate(
                "() => { const el = document.querySelector('#selectDate'); return el ? el.value : null; }"
            )
            if selectdate_val:
                date_selected = True
            elif selectdate_val is None:
                # #selectDate 자체가 없다 = 검색 폼을 벗어난 것 (이미 결과 페이지로 이동)
                already_searched = True
        except Exception:
            pass

        if not date_selected and not already_searched:
            return "❌ PNR오류", "달력 날짜 선택 실패"

        await jj_page.wait_for_timeout(300)

        # 조회 버튼 — 사람이 이미 눌렀으면 중복 클릭하지 않음
        if not already_searched:
            # id="searchResvBtn" (입력 검증 통과 전까지 disabled)
            try:
                await jj_page.wait_for_selector("#searchResvBtn:not([disabled])", timeout=5000)
            except Exception:
                pass
            await jj_page.click("#searchResvBtn", timeout=5000)

        # 조회 성공 시 viewReservationDetail.do로 실제 페이지 이동(navigate)이 발생함
        # 텍스트만 기다리면 이동 중인 중간 상태를 잘못 캡처할 수 있어 URL 이동을 우선 대기
        # (URL이 안 바뀌는 SPA 방식일 수도 있어 짧게 시도 후 텍스트 대기로 넘어감)
        try:
            await jj_page.wait_for_url(re.compile(r"viewReservationDetail"), timeout=8000)
        except Exception:
            pass

        # URL 이동 후에도 렌더링 시간 필요 — "탑승객 정보" 텍스트로 최종 확인
        try:
            await jj_page.wait_for_selector("text=탑승객 정보", timeout=10000)
        except Exception:
            pass
        await jj_page.wait_for_timeout(1500)

        html_content = await jj_page.inner_text("body")

        if any(kw in html_content for kw in CF_KEYWORDS):
            return "⏱️ 타임아웃", "보안 확인 필요 → 재실행 필요"

        # 성공 판정: "탑승객 정보" 존재 여부
        if "탑승객 정보" not in html_content:
            return "❌ PNR오류", "예약 확인 불가"

        flt_match = re.search(r'7C\d{3,4}', html_content)
        flt_found = flt_match.group() if flt_match else "편명미확인"

        # 페이지에 "예약일 YYYY.MM.DD(요일)"가 실제 출발일보다 먼저 나와서
        # 첫 매치만 쓰면 예약일을 출발일로 착각함 — 예약일 다음 매치를 사용
        date_matches = re.findall(r'(\d{4})\.(\d{2})\.(\d{2})\(', html_content)
        if len(date_matches) >= 2:
            y, m, d = date_matches[1]
            date_found = f"{y}-{m.zfill(2)}-{d.zfill(2)}"
        elif len(date_matches) == 1:
            y, m, d = date_matches[0]
            date_found = f"{y}-{m.zfill(2)}-{d.zfill(2)}"
        else:
            date_found = "날짜미확인"

        # 이 사이트는 공항코드(PQC/ICN) 대신 한글 도시명("푸꾸옥","서울")으로 표시함
        try:
            city_titles = await jj_page.locator(".boarding__info-title").all_inner_texts()
            city_titles = [c.strip() for c in city_titles if c.strip()]
        except Exception:
            city_titles = []

        if len(city_titles) >= 2:
            route_found = f"{city_titles[0]}→{city_titles[1]}"
        else:
            # 혹시 코드로 표시되는 경우 대비한 폴백
            airports = re.findall(
                rf'(?<![A-Z0-9])({AIRPORT_CODES})(?![A-Z0-9])',
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
            except Exception:
                pass

        if mismatch:
            return "⚠️ 불일치", detail + " | " + " / ".join(mismatch)

        if not _is_reliable_result(flt_found, route_found):
            return "❌ PNR오류", f"예약 확인 불가 (편명/구간 모두 미확인) | {detail}"

        return "✅ 확인완료", detail

    except PWTimeout:
        return "⏱️ 타임아웃", "재시도 필요"
    except Exception as e:
        _reraise_if_closed(e)
        logging.error(f"제주항공 조회 실패 | PNR: {pnr} | 탑승객: {input_last}{input_first}", exc_info=True)
        return "💥 오류", "시스템 로그 확인 필요"
    finally:
        try:
            await jj_page.close()
        except Exception:
            pass


async def _bx_check(page, target):
    """에어부산 자동 조회 — 방식(BX_MODE)에 따라 확장 또는 크롬 조종"""
    if BX_MODE == "ext":
        return await check_bx_ext(page, target)
    return await check_bx(page, target)


async def run_check(page, target, we_email=""):
    """단일 조회 실행 + 재시도 로직"""
    airline  = target["airline"]
    eng_name = target.get("eng_name", "")

    if airline == "에어부산" and (BX_MODE == "manual" or (BX_MODE == "ext" and not EXT.available)):
        # 수동확인 방식 — 자동 재시도/영문재시도 로직 대상 아님
        return await check_bx_manual(page, target)
    elif airline == "에어부산":
        result, detail = await _bx_check(page, target)
    elif airline == "대한항공":
        result, detail = await check_ke(page, target)
    elif airline == "진에어":
        result, detail = await check_lj(page, target)
    elif airline == "제주항공":
        result, detail = await check_jj(page, target)
    elif airline == "파라타항공":
        result, detail = await check_we(page, target, we_email)
    elif airline == "티웨이항공":
        # 수동확인 방식 — 자동 재시도/영문재시도 로직 대상 아님
        return await check_tw(page, target)
    else:
        return "⬜ 미지원", "지원 항공사 아님"

    # ── 국내선 한글 조회 실패 시 영문으로 재시도 (외국인 승무원 대비) ──
    # 파라타 제외 / 영문명 있을 때 / PNR오류·예약없음일 때만
    intl = is_international(target["dep"], target["arr"])
    if (
        airline in ("에어부산", "대한항공", "진에어", "제주항공")
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
        r2, d2 = result, detail   # 분기 누락 시 NameError 방지
        if airline == "에어부산":
            r2, d2 = await _bx_check(page, tmp)
        elif airline == "대한항공":
            r2, d2 = await check_ke(page, tmp)
        elif airline == "진에어":
            r2, d2 = await check_lj(page, tmp)
        elif airline == "제주항공":
            r2, d2 = await check_jj(page, tmp)
        if "확인완료" in r2 or "불일치" in r2:
            result = r2
            detail = "[영문재시도] " + d2

    # ── 타임아웃/오류 시 1회 재시도 ──
    elif "타임아웃" in result or ("오류" in result and "PNR" not in result and "파싱" not in detail):
        await asyncio.sleep(2)
        if airline == "에어부산":
            result, detail = await _bx_check(page, target)
        elif airline == "대한항공":
            result, detail = await check_ke(page, target)
        elif airline == "진에어":
            result, detail = await check_lj(page, target)
        elif airline == "제주항공":
            result, detail = await check_jj(page, target)
        elif airline == "파라타항공":
            result, detail = await check_we(page, target, we_email)
        if "확인완료" in result or "불일치" in result:
            detail = "[재시도 성공] " + detail

    return result, detail


async def main():
    os.system(f"title 타사 예약 자동 검증 시스템 v{__version__}")
    print(f"{'='*50}")
    print(f"✈️  타사 예약 자동 검증 시스템 v{__version__}")
    print("문의: 승무계획팀")
    print(f"{'='*50}")
    print(f"오류 로그 저장 위치: {LOG_PATH}")
    print(f"\n[이번 버전 변경사항]\n{LATEST_CHANGELOG}\n")
    print(f"{'='*50}\n")

    # 조회 범위 선택 팝업
    mode, start_date, end_date, delay_min, delay_max = get_check_mode()
    if mode == "5d":
        mode_label = "오늘~5일 이내"
    elif mode == "this":
        mode_label = f"이번달 말까지 ({end_date.month}월 {end_date.day}일)"
    else:
        mode_label = f"다음달 ({start_date.month}월 {start_date.day}일 ~ {end_date.month}월 {end_date.day}일)"

    targets = load_targets(EXCEL_PATH, SHEET_NAME, start_date, end_date)
    total   = len(targets)

    print(f"📂 선택한 파일: {EXCEL_PATH}")
    print(f"   시트: {SHEET_NAME}\n")

    if total == 0:
        print(f"검증 대상이 없습니다. ({mode_label})")
        print("\n⚠️  데이터가 분명히 있는데 0건이면 아래를 확인하세요:")
        print("   1) 위에 표시된 '선택한 파일'이 실제로 데이터를 넣은 그 파일이 맞나요?")
        print("      (같은 이름의 다른 파일/예전 파일을 골랐을 수 있습니다)")
        print("   2) 엑셀에서 데이터 입력 후 저장(Ctrl+S)했나요?")
        print(f"   3) 조회 범위({mode_label})에 실제 출발일이 포함되나요?")
        input("\n엔터 누르면 종료...")
        return

    # ── 이미 조회 완료된 건이 있으면 이어서 할지 확인 ──
    done_targets = [t for t in targets if "확인완료" in t["prev_result"]]
    if done_targets:
        if ask_resume(len(done_targets), total):
            for t in done_targets:
                t["result"] = t["prev_result"]
                t["detail"] = t["prev_detail"]
                t["skipped"] = True
            print(f"⏭  이미 완료된 {len(done_targets)}건은 건너뜁니다. "
                  f"(남은 조회: {total - len(done_targets)}건)\n")

    pending = [t for t in targets if not t.get("skipped")]

    bx_cnt = sum(1 for t in pending if t["airline"] == "에어부산")
    ke_cnt = sum(1 for t in pending if t["airline"] == "대한항공")
    lj_cnt = sum(1 for t in pending if t["airline"] == "진에어")
    jj_cnt = sum(1 for t in pending if t["airline"] == "제주항공")
    we_cnt = sum(1 for t in pending if t["airline"] == "파라타항공")
    tw_cnt = sum(1 for t in pending if t["airline"] == "티웨이항공")

    print(f"검증 대상: {len(pending)}건 ({mode_label})")
    print(f"  에어부산: {bx_cnt}건 | 대한항공: {ke_cnt}건 | 진에어: {lj_cnt}건 | 제주항공: {jj_cnt}건 | 파라타항공: {we_cnt}건 | 티웨이: {tw_cnt}건")
    print(f"  딜레이: {delay_min}~{delay_max}초")
    print(f"{'='*50}\n")

    # 건너뛴 것만 있고 조회할 게 없으면 바로 저장하고 종료
    if not pending:
        print("모두 이미 조회 완료된 상태입니다. 결과만 다시 저장합니다.\n")
        save_results(EXCEL_PATH, SHEET_NAME, targets)
        input("\n엔터 누르면 종료...")
        return

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

    # 조회 전용 브라우저 프로필 (v3.14.0)
    # 보안확인 통과 쿠키(cf_clearance)가 다음 실행에도 남아 → 캡챠 재등장 빈도 감소
    # 사용자의 평소 Chrome 프로필과 분리 (Chrome 켜져 있어도 충돌 없음)
    bot_profile = os.path.expanduser(r"~\AppData\Local\DH_checker_profile")

    if chrome_exe:
        print(f"시스템 Chrome 사용: {chrome_exe}")
    else:
        print("⚠️  Chrome 미발견 → Playwright Chromium으로 실행 (에어부산 캡챠 발생 가능)")

    # Chrome 쿠키 파일 임시 복사 (원본 잠금 회피)
    import shutil, tempfile
    async with async_playwright() as p:
        browser     = None
        chrome_proc = None
        context     = None
        # (v3.16.0) 에어부산 건이 있으면: 크롬 일반 실행 → 사용자가 보안확인 통과 → 그 창에 연결
        if BX_MODE == "auto" and chrome_exe and bx_cnt > 0:
            context, browser, chrome_proc = await attach_real_chrome(p, chrome_exe, bot_profile)
        cdp_mode = context is not None

        if not cdp_mode:
            launch_kwargs = dict(
                headless=HEADLESS,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--disable-infobars",
                    "--no-sandbox",
                    "--disable-dev-shm-usage",
                    "--disable-extensions",
                    "--window-size=1280,800",
                ]
            )
            if chrome_exe:
                launch_kwargs["executable_path"] = chrome_exe
            if PATCHRIGHT:
                # Patchright는 자체적으로 자동화 흔적을 처리하므로 커스텀 인자를 최소화해야 함
                launch_kwargs["args"] = ["--window-size=1280,800"]

            # user_agent 고정값 제거 (v3.14.0): 예전엔 Chrome/124로 고정해뒀는데
            # 실제 설치된 Chrome 버전과 달라서 오히려 봇 의심 신호가 됐음 → 실제 값 그대로 사용
            ctx_kwargs = dict(
                locale="ko-KR",
                timezone_id="Asia/Seoul",
                viewport={"width": 1280, "height": 800},
                java_script_enabled=True,
            )
            if PATCHRIGHT:
                ctx_kwargs["viewport"] = None  # 실제 창 크기 그대로 (Patchright 권장)
            try:
                os.makedirs(bot_profile, exist_ok=True)
                context = await p.chromium.launch_persistent_context(
                    bot_profile, **launch_kwargs, **ctx_kwargs
                )
                print("전용 브라우저 프로필 사용 (보안확인 통과 기록 유지)")
            except Exception as e:
                # 프로필 잠김(프로그램 2개 동시 실행 등) → 기존 방식으로 실행
                logging.warning(f"전용 프로필 실행 실패 → 일회성 모드: {e}")
                print("⚠️  전용 프로필 사용 불가 → 일회성 모드로 실행 (프로그램이 이미 켜져 있는지 확인)")
                browser = await p.chromium.launch(**launch_kwargs)
                context = await browser.new_context(**ctx_kwargs)

        # playwright-stealth 적용 (클라우드플레어 핑거프린트 우회)
        if cdp_mode:
            stealth = None  # 실제 크롬에 연결한 상태 → 위장 불필요
        elif PATCHRIGHT:
            stealth = None
            print("Patchright 모드 (자동화 흔적 제거 브라우저)")
        elif STEALTH_AVAILABLE:
            stealth = Stealth(
                navigator_languages_override=("ko-KR", "ko"),
                navigator_platform_override="Win32",
                navigator_webdriver=True,
                chrome_runtime=True,
            )
            print("playwright-stealth 적용 완료 (클라우드플레어 우회 시도)")
        else:
            stealth = None
            await context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                Object.defineProperty(navigator, 'plugins', {get: () => [1,2,3,4,5]});
                Object.defineProperty(navigator, 'languages', {get: () => ['ko-KR','ko','en-US','en']});
                window.chrome = {runtime: {}};
            """)
            print("⚠️  playwright-stealth 미설치 → 기본 우회 모드")

        page = context.pages[0] if context.pages else await context.new_page()

        # stealth는 첫 탭(page)에만 적용 — v3.14.0에서 context 전체로 넓혔다가
        # 에어부산 보안확인이 체크 후에도 무한반복되는 문제 발생 → v3.14.1에서 원복.
        # (Cloudflare가 stealth의 가짜 브라우저 정보를 오히려 이상 신호로 판단.
        #  조회용 새 탭은 위장 없는 순수 Chrome 상태가 사람 체크 통과에 유리)
        if stealth:
            await stealth.apply_stealth_async(page)

        # (v3.16.3) 안전핀 탭 — 크롬은 열려있던 탭이 0개가 되면(사람이 실수로
        # 첫 탭을 닫거나, 보안검사 페이지가 탭을 강제로 닫는 경우 등) 프로그램
        # 전체가 종료돼버림. 절대 건드리지 않는 빈 탭을 하나 띄워서 항상 최소
        # 1개는 열려있게 만들어 이 문제를 원천 차단함.
        try:
            _anchor_page = await context.new_page()
            await _anchor_page.goto("about:blank")
        except Exception:
            pass

        if BX_MODE == "ext" and bx_cnt > 0:
            await setup_bx_extension()

        pending_total = len(pending)
        progress_state = create_progress_window(pending_total)

        for i, target in enumerate(pending, 1):
            airline = target["airline"]
            pnr     = target["pnr"]
            print(f"[{i:02d}/{pending_total}] {target['kor_name']:5} | {airline} | {pnr} | ", end="", flush=True)

            # v3.11.0: PNR 캐시 제거 — 같은 예약번호라도 사람마다 매번 개별 조회함.
            # (동승자 이름 검증으로도 오탐 사례가 나와서, 속도보다 정확도를 우선함.
            #  대신 이어서 조회 기능으로 중단 시 재조회 부담을 줄임)
            # (v3.16.5) 메인 탭이 닫혀 있으면(사람이 보안확인 끝난 탭을 닫은 경우 등)
            # 브라우저는 살아있으니 새 탭으로 교체 — 파라타·티웨이는 메인 탭을 직접 쓰기 때문
            try:
                if page.is_closed():
                    page = await context.new_page()
            except Exception:
                pass  # 브라우저 자체가 죽은 경우 → 아래 조회에서 closed로 감지되어 안전 종료
            try:
                result, detail = await run_check(page, target, we_email)
            except Exception as exc:
                logging.error(f"조회 중 예외 | {airline} | PNR: {pnr}", exc_info=True)
                result, detail = "💥 오류", "조회 중 오류 발생 (로그 확인)"
                if "closed" in str(exc).lower():
                    target["result"] = result
                    target["detail"] = detail
                    print(f"{result}  {detail}")
                    update_progress_window(progress_state, i, pending_total, target, result)
                    print("\n⚠️  브라우저가 닫혀 남은 건은 진행할 수 없습니다.")
                    print("→ 여기까지의 결과는 엑셀에 저장됩니다.\n")
                    for rest in pending[i:]:
                        rest["result"] = "⬜ 미조회"
                        rest["detail"] = "브라우저 종료로 미처리"
                    break
            target["result"] = result
            target["detail"] = detail
            print(f"{result}  {detail}")
            update_progress_window(progress_state, i, pending_total, target, result)

            # 매 건마다 조용히 중간 저장. 프로그램이 창닫기/강제종료/정전 등으로
            # 예고 없이 죽어도, 재실행 시 여기까지의 결과는 "이어서 조회"로 살아남음.
            save_results(EXCEL_PATH, SHEET_NAME, targets, silent=True)
            # 엑셀 자체가 손상되는 최악의 경우를 대비한 가벼운 원본 백업
            _append_jsonl_backup(target)

            if i < pending_total:
                # 수동확인 건(티웨이, 수동모드 에어부산)은 사람이 직접 조회하므로 대기 불필요
                manual = target["airline"] == "티웨이항공" or (
                    target["airline"] == "에어부산" and
                    (BX_MODE == "manual" or (BX_MODE == "ext" and not EXT.available)))
                if manual:
                    pass
                # 에어부산 자동조회(확장 포함) 시엔 차단 방지 딜레이 더 늘림
                elif target["airline"] == "에어부산":
                    await asyncio.sleep(random.uniform(5.0, 10.0))
                else:
                    await asyncio.sleep(random.uniform(delay_min, delay_max))

        close_progress_window(progress_state)
        EXT.stop()

        if cdp_mode:
            # 조회용 크롬 종료 — 이전 실행에서 살아남은 크롬을 재사용한 경우엔
            # chrome_proc이 그 크롬 본체가 아니므로, 연결 끊기 전에 크롬에 직접 종료 명령을 보냄
            closed_ok = False
            try:
                cdp = await browser.new_browser_cdp_session()
                await cdp.send("Browser.close")
                closed_ok = True
            except Exception:
                pass
            try:
                await browser.close()  # 연결 해제
            except Exception:
                pass
            if not closed_ok and chrome_proc:
                try:
                    chrome_proc.terminate()
                except Exception:
                    pass
        else:
            try:
                await context.close()
                if browser:
                    await browser.close()
            except Exception:
                pass  # 이미 닫혀있으면 무시

    save_results(EXCEL_PATH, SHEET_NAME, targets)

    confirmed = sum(1 for t in targets if t["result"] and "확인완료"    in str(t["result"]))
    mismatch  = sum(1 for t in targets if t["result"] and "불일치"       in str(t["result"]))
    no_rsv    = sum(1 for t in targets if t["result"] and "예약없음"     in str(t["result"]))
    pnr_err   = sum(1 for t in targets if t["result"] and "PNR오류"     in str(t["result"]))
    manual    = sum(1 for t in targets if t["result"] and "수동확인필요" in str(t["result"]))
    # "오류"는 "PNR오류"의 부분문자열이라 그대로 세면 중복 집계됨 → PNR오류는 제외
    error     = sum(1 for t in targets if t["result"] and
                    (("오류" in str(t["result"]) and "PNR오류" not in str(t["result"]))
                     or "타임아웃" in str(t["result"])))
    skipped   = sum(1 for t in targets if t["result"] and "미조회" in str(t["result"]))

    print(f"\n{'='*50}")
    print(f"✅ 확인완료      : {confirmed}건")
    print(f"⚠️  불일치       : {mismatch}건  ← 즉시 확인!")
    print(f"❌ 예약없음      : {no_rsv}건   ← 즉시 확인!")
    print(f"❌ PNR오류       : {pnr_err}건  ← 즉시 확인!")
    print(f"⚠️  수동확인필요  : {manual}건  ← 파라타항공 직접 조회 필요!")
    print(f"💥 오류/재시도   : {error}건")
    if skipped:
        print(f"⬜ 미조회        : {skipped}건  ← 중단되어 조회 못함, 재실행 필요!")
    input("\n엔터 누르면 종료...")


if __name__ == "__main__":
    asyncio.run(main())