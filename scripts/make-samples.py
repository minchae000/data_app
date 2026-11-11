"""가상 샘플을 만듭니다 (2차 — 「자주 검사 CHECK SHEET」 양식).

실행: python3 scripts/make-samples.py      (reportlab · openpyxl · Pillow, 스캔 흉내는 poppler 의 pdftoppm 이 있으면 씀)
글꼴: 인쇄된 양식 글자는 나눔고딕(NanumGothic-Regular.ttf, OFL)을 씁니다. 없으면 FONT 환경변수로 TTF 경로를 주세요.

수강생이 보내 준 검사표(10장 ①)의 「칸 배치」만 본떠 새로 그렸습니다. 모델 · 품번 · 품명 · 관리 번호 · 치수 · 도면은 모두 지어낸 것이고,
실제 검사표 그림 · 값은 리포에 넣지 않습니다.

만드는 것
- samples/공유폴더_샘플/*.pdf   검사표 6장. 파일 이름은 규칙 없이(10장 ② 의 예처럼) — 「모델-호기」 「img00001」 「모델 상부 #호기」 「모델 하부센터 (날짜)」 「날짜~날짜 작업」
    · 글자 층 있음 4장 — 태블릿 · 엑셀로 채워 PDF 로 저장한 경우(값까지 글자). 도구가 칸 자리로 바로 읽음
    · 그림만 2장 — 손으로 적어 스캔한 경우(글자 층 없음). 사람이 옮기거나 AI 반자동
    · 여러 쪽 1장(2026-11-10 답 ④) — 「모델-호기」 이름의 합본: 1쪽 자주 검사(글자 층) + 2쪽 시트 부착 체크시트(스캔, 90° 누운 채)
- samples/출하일정_샘플.xlsx     출하일 · 모델 · 품번 · 호기 · 수량 (호기 범위 「315015~315016」 한 줄 포함)
- samples/출하일정_초대형형식_샘플.xlsx  초대형 출하일정 엑셀의 칸 모양(2026-11-10 답 ③ 사진의 열 이름 · 두 줄 머리 · 「월/일」 출하예정)만 본뜬 가상 표
- js/sample-data.js             파일 목록 + 정답(시험 · 「가상 AI 답」 용) + 가릴 영역
- js/sample-pdfs.js             PDF 의 base64 사본(file:// 로 열어도 샘플이 열리도록, 필요할 때만 불러옴)
"""
import base64
import datetime as dt
import io
import json
import math
import os
import random
import shutil
import subprocess
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from PIL import Image, ImageFilter
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = Path(__file__).resolve().parent.parent
OUT_PDF = ROOT / 'samples' / '공유폴더_샘플'
FONT = os.environ.get('FONT') or str(Path.home() / 'Library/Fonts/NanumGothic-Regular.ttf')
pdfmetrics.registerFont(TTFont('KR', FONT))
W, H = landscape(A4)

# ── 지어낸 모델 · 품번 · 항목 ──────────────────────────
MODELS = {
    'A': {'model': 'MODEL-A100 / MODEL-A100E', 'part': 'QA-70011', 'name': '하부 프레임 C', 'mgmt': 'QF-CS-011', 'process': 'MODEL-A100 / MODEL-A100E 하부 프레임 사상',
          'items': [
              {'no': '1', 'label': ['가이드 홀', '방향(전·후)', '확인할 것'], 'tool': '육안', 'subs': []},
              {'no': '2', 'label': ['420±1'], 'tool': '줄자', 'subs': ['LH', 'RH']},
              {'no': '3', 'label': ['85±2'], 'tool': '줄자', 'subs': ['LH', 'RH']},
              {'no': '4', 'label': ['1860±5'], 'tool': '줄자', 'subs': ['LH', 'RH']},
              {'no': '5', 'label': ['2740±6'], 'tool': '줄자', 'subs': ['FR', 'RE']},
              {'no': '6', 'label': ['외 관'], 'tool': '육안', 'subs': []}]},
    'B': {'model': 'MODEL-B200', 'part': 'QA-80022', 'name': '상부 브래킷 B', 'mgmt': 'QF-CS-022', 'process': 'MODEL-B200 상부 브래킷 용접',
          'items': [
              {'no': '1', 'label': ['슬롯 방향', '확인할 것'], 'tool': '육안', 'subs': []},
              {'no': '2', 'label': ['1860±5'], 'tool': '줄자', 'subs': ['LH', 'RH']},
              {'no': '3', 'label': ['48 +1.0/-0.5'], 'tool': '버니어', 'subs': ['LH', 'RH']},
              {'no': '4', 'label': ['평탄도', '1.5 MAX'], 'tool': '게이지', 'subs': []},
              {'no': '5', 'label': ['2740±6'], 'tool': '줄자', 'subs': ['FR', 'RE']},
              {'no': '6', 'label': ['외 관'], 'tool': '육안', 'subs': []}]},
}


def label_text(it):
    return ' '.join(it['label'])


def cells(m):
    out = []
    for it in MODELS[m]['items']:
        for s in (it['subs'] or ['']):
            out.append((it['no'] + (' ' + s if s else ''), it))
    return out


rnd = random.Random(1013)


def good_values(m):
    """공차 안쪽의 그럴듯한 값"""
    v = {}
    for k, it in cells(m):
        lab = label_text(it)
        if it['tool'] == '육안':
            v[k] = 'OK'
        elif '±' in lab:
            n, t = lab.split('±')
            n, t = float(n), float(t)
            v[k] = ('%.1f' % (n + rnd.uniform(-0.6, 0.6) * t)).rstrip('0').rstrip('.')
        elif '+1.0/-0.5' in lab:
            v[k] = '%.1f' % (48 + rnd.uniform(-0.3, 0.7))
        elif 'MAX' in lab:
            v[k] = '%.1f' % rnd.uniform(0.3, 1.2)
    return v


# 검사표 6장 — 이름 · 종류 · 모델 · 작성일 · 호기별 값 · 문제점
SHEETS = []


def add(name, kind, m, date, hogis, problems=None):
    SHEETS.append({'name': name, 'kind': kind, 'm': m, 'date': date, 'hogis': hogis, 'problems': problems or {}})


g = good_values
v1, v2, v3 = g('A'), g('A'), g('A')
v3['4 RH'] = 'V'                                   # 값 없이 체크(V)만 → 설정 「체크만」 이 합격이면 합격
add('MODEL-A100-314501.pdf', 'text', 'A', '2026-10-19', [('314501', v1), ('314502', v2), ('314503', v3)])
b1, b2 = g('B'), g('B')
b1['3 LH'] = '49.3'                                # 48 +1.0/-0.5 → 상한 49 초과 = 불합격
add('img00001.pdf', 'text', 'B', '2026-10-20', [('315010', b1), ('315011', b2)], {'3': '재가공 후 재측정'})
c1, c2 = g('B'), g('B')
c1['5 FR'] = '2746'                                # 상한과 같은 값 → 합격(경계 포함)
c2['2 RH'] = ''                                    # 빈 칸 → 보류(미판정)
add('MODEL-B200 하부센터 (2026.10.12).pdf', 'text', 'B', '2026-10-12', [('315012', c1), ('315013', c2)])
d1 = g('A')
d1['2 LH'] = '421'                                 # 상한과 같은 값 → 합격
add('MODEL-A100 상부 #314504.pdf', 'scan', 'A', '2026-10-21', [('314504', d1)])
e1 = g('A')
e1['6'] = 'NG'                                     # 외관 NG → 불합격
add('1013~1015 작업.pdf', 'scan', 'A', '2026-10-15', [('314505', e1)], {'6': '찍힘 — 보수 후 재검'})
f1 = g('A')
add('img00002.pdf', 'text', 'A', '2026-10-08', [('314499', f1)])   # 출하일정에 없는 호기
m1 = g('A')
add('MODEL-A100-314506.pdf', 'multi', 'A', '2026-11-02', [('314506', m1)])   # 합본 — 1쪽 자주 검사 + 2쪽 시트(아래 SEAT)

# 시트 부착 체크시트(2쪽) — 모델링 뷰의 풍선마다 체크. 풍선 24개 중 1개는 체크 없음 → 불합격
SEAT = {'process': '시트 부착', 'date': '2026-11-02', 'mgmt': 'QF-ST-006', 'rev': '00', 'model': 'MODEL-A100', 'name': '사이드 프레임',
        'part': 'QA-70011', 'hogi': '314506', 'missing': ('VIEW B', 2), 'balloons': 24}

# ── 손글씨 획 (1단계 샘플과 같은 방식) ──────────────────
def arc(cx, cy, rx, ry, a0, a1, n=14):
    n = max(5, int(n * 0.6))
    return [(cx + rx * math.cos(math.radians(a0 + (a1 - a0) * i / n)), cy + ry * math.sin(math.radians(a0 + (a1 - a0) * i / n))) for i in range(n + 1)]


G = {
    '0': (0.62, [arc(0.3, 0.5, 0.27, 0.5, 95, 455, 22)]),
    '1': (0.42, [[(0.08, 0.78), (0.27, 1.0), (0.27, 0.0)]]),
    '2': (0.62, [arc(0.3, 0.72, 0.26, 0.27, 160, -40, 12) + [(0.04, 0.0), (0.6, 0.0)]]),
    '3': (0.62, [arc(0.28, 0.76, 0.24, 0.24, 150, -90, 12) + arc(0.28, 0.27, 0.29, 0.27, 90, -160, 12)]),
    '4': (0.64, [[(0.46, 0.0), (0.46, 1.0), (0.02, 0.32), (0.62, 0.32)]]),
    '5': (0.62, [[(0.56, 1.0), (0.12, 1.0), (0.07, 0.56)] + arc(0.3, 0.31, 0.29, 0.31, 125, -155, 14)]),
    '6': (0.62, [arc(0.55, 0.42, 0.47, 0.56, 80, 180, 8) + arc(0.32, 0.29, 0.25, 0.29, 180, 540, 18)]),
    '7': (0.6, [[(0.03, 1.0), (0.58, 1.0), (0.22, 0.0)]]),
    '8': (0.62, [arc(0.3, 0.76, 0.21, 0.24, -90, 270, 16), arc(0.3, 0.27, 0.27, 0.27, 90, 450, 18)]),
    '9': (0.62, [arc(0.3, 0.7, 0.26, 0.29, 0, 370, 18) + [(0.56, 0.62), (0.44, 0.0)]]),
    '.': (0.25, [[(0.08, 0.02), (0.13, 0.05), (0.1, 0.0)]]),
    '-': (0.5, [[(0.06, 0.48), (0.44, 0.5)]]),
    'O': (0.74, [arc(0.36, 0.5, 0.34, 0.5, 100, 470, 22)]),
    'K': (0.62, [[(0.06, 1.0), (0.06, 0.0)], [(0.58, 1.0), (0.08, 0.42), (0.6, 0.0)]]),
    'N': (0.68, [[(0.05, 0.0), (0.05, 1.0), (0.62, 0.0), (0.62, 1.0)]]),
    'G': (0.72, [arc(0.38, 0.5, 0.34, 0.5, 40, 340, 18) + [(0.68, 0.42), (0.42, 0.42)]]),
    '✓': (0.7, [[(0.04, 0.5), (0.26, 0.08), (0.7, 1.0)]]),
    ' ': (0.35, []),
}
INK = (0.12, 0.16, 0.42)


class Pen:
    def __init__(self, c, seed):
        self.c = c
        self.r = random.Random(seed)
        self.slant = self.r.uniform(0.12, 0.26)

    def width(self, text, size):
        return sum((G.get(ch, G['0'])[0] + 0.06) * size for ch in text)

    def write(self, x, y, text, size=8.5):
        c, r = self.c, self.r
        c.saveState()
        c.setStrokeColorRGB(*INK)
        c.setLineCap(1)
        c.setLineJoin(1)
        base = y + r.uniform(-0.5, 0.5)
        for ch in text:
            w, strokes = G.get(ch, G['0'])
            hgt = size * r.uniform(0.9, 1.1)
            c.setLineWidth(r.uniform(0.7, 1.0))
            ph, amp = r.uniform(0, 6.3), hgt * 0.025
            for st in strokes:
                pts = []
                for i, (px, py) in enumerate(st):
                    X = x + (px + py * self.slant) * hgt + amp * math.sin(ph + i * 0.9)
                    Y = base + py * hgt + amp * math.cos(ph + i * 0.7)
                    pts.append((X, Y))
                pa = c.beginPath()
                pa.moveTo(*pts[0])
                for q in pts[1:]:
                    pa.lineTo(*q)
                c.drawPath(pa, stroke=1, fill=0)
            x += w * hgt * r.uniform(0.95, 1.1) + hgt * 0.06
        c.restoreState()

    def scribble(self, x, y, w, h):
        c, r = self.c, self.r
        c.saveState()
        c.setStrokeColorRGB(*INK)
        c.setLineWidth(0.9)
        pa = c.beginPath()
        pa.moveTo(x, y + h * 0.4)
        for i in range(1, 11):
            pa.lineTo(x + w * i / 10 + r.uniform(-2, 2), y + h * (0.5 + 0.45 * math.sin(i * 1.7 + r.uniform(0, 1))))
        c.drawPath(pa, stroke=1, fill=0)
        c.restoreState()


# ── 양식 자리 (가로 A4, 위에서부터 잰 값 → reportlab 은 아래에서부터라 Y() 로 바꿈) ──
def Y(top):
    return H - top


L0, R0 = 24, W - 24
TB = {'x': [430, 452, 532, 572, 594, 630, 666, 702, 738, 774, R0]}   # No · 검사 항목 · 측정기 · 구분 · 호기 ×5 · 문제점
TB_TOP = 214          # 표 머리 위 끝
HEAD_H = 34
ROW_H = {'1': 46, '2': 36, '3': 36, '4': 36, '5': 36, '6': 22}
SIGN_BOXES = []


def text(c, x, top, s, size=8, center=False, right=False):
    c.setFont('KR', size)
    if center:
        c.drawCentredString(x, Y(top), s)
    elif right:
        c.drawRightString(x, Y(top), s)
    else:
        c.drawString(x, Y(top), s)


def rect(c, x0, top, x1, bottom, lw=0.7):
    c.setLineWidth(lw)
    c.rect(x0, Y(bottom), x1 - x0, bottom - top)


def hline(c, x0, x1, top, lw=0.5):
    c.setLineWidth(lw)
    c.line(x0, Y(top), x1, Y(top))


def vline(c, x, t0, t1, lw=0.5):
    c.setLineWidth(lw)
    c.line(x, Y(t0), x, Y(t1))


def fill_value(c, pen, kind, x, top, s, size=8):
    """칸 가운데(x)에 값 — 글자 층(text)이면 글꼴로, 스캔(scan)이면 손글씨 획으로"""
    if not s:
        return
    if kind == 'text':
        text(c, x, top, s, size, center=True)
    else:
        pen.write(x - pen.width(s, size * 1.05) / 2, Y(top) - 1, s, size * 1.05)


def draw_sheet(c, sh, seed):
    m = MODELS[sh['m']]
    kind = sh['kind']
    pen = Pen(c, seed)
    c.setStrokeColorRGB(0.15, 0.15, 0.15)
    c.setFillColorRGB(0.1, 0.1, 0.1)
    rect(c, L0, 20, R0, H - 22, 1.0)
    # 제목 · 작성일 · 제작 공정명
    c.setFont('KR', 22)
    c.drawString(40, Y(62), '자주 검사 CHECK SHEET')
    vline(c, 310, 20, 92, 0.8)
    text(c, 320, 48, '작성일 :', 10)
    text(c, 320, 72, '제작 공정명 :', 10)
    if kind == 'text':
        d = dt.date.fromisoformat(sh['date'])
        text(c, 382, 48, '%d년 %d월 %d일' % (d.year, d.month, d.day), 10)
        text(c, 392, 72, m['process'], 9)
    else:
        pen.write(382, Y(48), sh['date'].replace('-', '.'), 10)
        text(c, 392, 72, m['process'], 9)   # 공정명은 인쇄된 양식 글자
    # 결재 칸(오른쪽 위) — 이름 · 서명은 가릴 영역
    bx = R0 - 150
    rect(c, bx, 26, R0 - 6, 86)
    hline(c, bx, R0 - 6, 40)
    for i, t in enumerate(['입 안', '심 사', '결 정']):
        text(c, bx + 25 + i * 48, 36, t, 7.5, center=True)
        if i:
            vline(c, bx + i * 48, 26, 86)
        if kind == 'scan':
            pen.scribble(bx + 8 + i * 48, Y(76), 30, 18)
    SIGN_BOXES.append(('결재 칸', bx - 3, 23, R0 - 3, 89))
    hline(c, L0, R0, 92, 1.0)
    # 머리 표 — 모델명 · 품번 / 품명 · 관리 번호
    rows = [(92, 118, ('모델명', m['model'], '품번', m['part'])), (118, 144, ('품명', m['name'], '관리 번호', m['mgmt']))]
    for t0, t1, (a, av, b, bv) in rows:
        rect(c, L0, t0, 420, t1)
        for x in (84, 250, 300):
            vline(c, x, t0, t1)
        mid = (t0 + t1) / 2 + 3
        text(c, 54, mid, a, 8.5, center=True)
        text(c, 275, mid, b, 8.5, center=True)
        text(c, 167, mid, av, 8, center=True)
        text(c, 360, mid, bv, 8, center=True)
    # 작업자 준수 사항(오른쪽 위) — 지어낸 안내
    text(c, 434, 104, '※ 작업자 준수 사항', 8)
    for i, t in enumerate(['1. 선입 선출 순서로 호기를 관리한다', '2. 제품에 적힌 호기를 호기수 칸에 적는다', '3. 사상 뒤 외관을 검사한다', '4. 고친 치수는 문제점 칸에 적는다']):
        text(c, 440, 117 + i * 11, t, 7)
    # 도면(왼쪽) — 지어낸 단순한 틀
    c.setLineWidth(0.8)
    rect(c, 70, 190, 330, 430)
    c.circle(200, Y(310), 70)
    c.circle(200, Y(310), 30)
    for (x0, t0, x1, t1) in [(70, 190, 110, 230), (290, 190, 330, 230), (70, 390, 110, 430), (290, 390, 330, 430)]:
        rect(c, x0, t0, x1, t1, 0.5)
    text(c, 200, 182, 'RH', 8, center=True)
    text(c, 200, 446, 'LH', 8, center=True)
    text(c, 56, 312, '전', 8, center=True)
    text(c, 344, 312, '후', 8, center=True)
    hline(c, 70, 330, 462, 0.4)
    text(c, 200, 472, '(5) 2740', 8, center=True)
    vline(c, 352, 190, 430, 0.4)
    text(c, 368, 300, '(2)', 7, center=True)
    text(c, 368, 312, '420', 7, center=True)
    # 개정 표(왼쪽 아래)
    rect(c, 200, 494, 420, 540)
    for x in (230, 300):
        vline(c, x, 494, 540)
    for tp in (508, 524):
        hline(c, 200, 420, tp)
    text(c, 215, 504, 'Rev.', 7, center=True)
    text(c, 265, 504, '개 정 일', 7, center=True)
    text(c, 360, 504, '개 정 내 용', 7, center=True)
    text(c, 215, 519, '0', 7, center=True)
    text(c, 265, 519, '2024. 03. 02', 7, center=True)
    text(c, 360, 519, '최초 제정', 7, center=True)
    rect(c, L0, 494, 190, 540)
    for i, t in enumerate(['입 안', '심 사', '결 정']):
        text(c, L0 + 27 + i * 55, 504, t, 7.5, center=True)
        if i:
            vline(c, L0 + i * 55, 494, 540)
        if kind == 'scan':
            pen.scribble(L0 + 10 + i * 55, Y(534), 34, 16)
    hline(c, L0, 190, 508)
    SIGN_BOXES.append(('아래 결재 칸', L0 - 2, 491, 193, 543))
    # 항목 표(오른쪽)
    x = TB['x']
    top = TB_TOP
    rect(c, x[0], top, R0, top + HEAD_H)
    for xi in x[1:-1]:
        if xi not in (594,):
            vline(c, xi, top, top + HEAD_H)
    c.setLineWidth(0.5)
    c.line(x[2], Y(top), x[4], Y(top + HEAD_H))       # 대각선 — 위 오른쪽 「호기수」 · 아래 왼쪽 「측정기」
    hm = top + HEAD_H / 2 + 3
    text(c, (x[0] + x[1]) / 2, hm, 'No.', 8, center=True)
    text(c, (x[1] + x[2]) / 2, hm, '검사 항목', 8, center=True)
    text(c, x[4] - 2, top + 12, '호기수', 7.5, right=True)
    text(c, x[2] + 4, top + HEAD_H - 6, '측정기', 7.5)
    text(c, (x[-2] + x[-1]) / 2, top + 14, '문제점', 7.5, center=True)
    text(c, (x[-2] + x[-1]) / 2, top + 25, '(조치사항)', 7, center=True)
    for i, (hno, _) in enumerate(sh['hogis']):
        fill_value(c, pen, kind, (x[4 + i] + x[5 + i]) / 2, hm, hno, 8)
    t = top + HEAD_H
    vals = {hno: v for hno, v in sh['hogis']}
    for it in m['items']:
        hgt = ROW_H[it['no']]
        rect(c, x[0], t, R0, t + hgt)
        for xi in x[1:-1]:
            vline(c, xi, t, t + hgt)
        mid = t + hgt / 2 + 3
        text(c, (x[0] + x[1]) / 2, mid, it['no'], 8, center=True)
        nl = len(it['label'])
        for j, ln in enumerate(it['label']):
            text(c, (x[1] + x[2]) / 2, mid + (j - (nl - 1) / 2) * 10, ln, 7.5, center=True)
        text(c, (x[2] + x[3]) / 2, mid, it['tool'], 7.5, center=True)
        subs = it['subs'] or ['']
        sh_h = hgt / len(subs)
        for k, sb in enumerate(subs):
            st = t + k * sh_h
            if sb:
                if k:
                    hline(c, x[3], x[-2], st, 0.4)
                text(c, (x[3] + x[4]) / 2, st + sh_h / 2 + 3, sb, 7, center=True)
            ck = it['no'] + (' ' + sb if sb else '')
            for i, (hno, _) in enumerate(sh['hogis']):
                fill_value(c, pen, kind, (x[4 + i] + x[5 + i]) / 2, st + sh_h / 2 + 3, vals[hno].get(ck, ''), 8)
        pr = sh['problems'].get(it['no'])
        if pr:
            if kind == 'text':
                for j, part in enumerate([pr[:6], pr[6:12], pr[12:]]):
                    if part:
                        text(c, (x[-2] + x[-1]) / 2, mid - 8 + j * 9, part.strip(), 6.5, center=True)
            else:
                pen.scribble(x[-2] + 6, Y(mid + 4), 40, 8)
        t += hgt
    # 검사자 줄 — 서명은 읽지 않고, AI 에 보낼 그림에서는 가림
    rect(c, x[0], t, R0, t + 22)
    vline(c, x[3], t, t + 22)
    text(c, (x[0] + x[3]) / 2, t + 14, '검사자', 8, center=True)
    for i in range(len(sh['hogis'])):
        vline(c, x[5 + i], t, t + 22)
        pen.scribble(x[4 + i] + 6, Y(t + 18), 24, 12)
    SIGN_BOXES.append(('검사자 서명 줄', x[3] - 2, t - 2, x[-2] + 2, t + 24))
    c.setFont('KR', 6.5)
    c.setFillColorRGB(0.4, 0.4, 0.4)
    c.drawString(L0 + 4, 26, '가상 샘플 — 모델 · 품번 · 치수 · 값은 모두 지어낸 것입니다.')


def draw_seat(c, seed):
    """시트 부착 체크시트(가로) — 머리 줄(관리 번호 · REV · MODEL · PART NAME · P/No. · 호기수 · 검사일 · 결재) + 모델링 뷰 4칸의 풍선"""
    S = SEAT
    pen = Pen(c, seed)
    c.setStrokeColorRGB(0.15, 0.15, 0.15)
    c.setFillColorRGB(0.1, 0.1, 0.1)
    rect(c, L0, 20, R0, H - 22, 1.0)
    c.setFont('KR', 18)
    c.drawString(40, Y(52), '시트 부착 체크시트 (SEAT)')
    heads = [('관리 번호', 90), ('REV.', 40), ('MODEL', 90), ('PART NAME', 100), ('P/No.', 80), ('호기수', 70), ('검사일', 70)]
    x = L0 + 10
    for (hd, wd), v in zip(heads, [S['mgmt'], S['rev'], S['model'], S['name'], S['part'], None, None]):
        rect(c, x, 64, x + wd, 84)
        rect(c, x, 84, x + wd, 108)
        text(c, x + wd / 2, 78, hd, 7.5, center=True)
        if v:
            text(c, x + wd / 2, 100, v, 7.5, center=True)
        x += wd
    pen.write(L0 + 10 + 400 + 8, Y(102), S['hogi'], 9)          # 호기수 칸
    pen.write(L0 + 10 + 470 + 12, Y(102), '11.02', 9)           # 검사일 칸
    bx = x + 10                                       # 결재 칸(입안 · 심사 · 결정) — 시트 양식은 자리가 다름
    rect(c, bx, 56, R0 - 10, 108)
    hline(c, bx, R0 - 10, 70)
    for i, t in enumerate(['입 안', '심 사', '결 정']):
        text(c, bx + 26 + i * 52, 66, t, 7.5, center=True)
        if i:
            vline(c, bx + i * 52, 56, 108)
        pen.scribble(bx + 8 + i * 52, Y(100), 32, 18)
    SIGN_BOXES.append(('시트 결재 칸', bx - 3, 53, R0 - 7, 111))
    r = random.Random(seed)
    views = [('VIEW A', 40, 130, 390, 330), ('VIEW B', 410, 130, 800, 330), ('VIEW C', 40, 350, 390, 550), ('VIEW D', 410, 350, 800, 550)]
    n = 0
    for vi, (vn, x0, t0, x1, t1) in enumerate(views):
        rect(c, x0, t0, x1, t1, 0.6)
        text(c, (x0 + x1) / 2, t1 - 6, vn, 8, center=True)
        c.setLineWidth(0.5)
        rect(c, x0 + 30, t0 + 40, x1 - 30, t1 - 40, 0.5)        # 지어낸 단순한 틀
        for k in range(6):
            n += 1
            bxp = x0 + 50 + k * (x1 - x0 - 100) / 5
            btp = t0 + 30 + (k % 2) * 110
            c.setStrokeColorRGB(0.75, 0.1, 0.1)
            c.setDash(2, 2)
            c.circle(bxp, Y(btp), 9)
            c.setDash()
            c.setStrokeColorRGB(0.15, 0.15, 0.15)
            if not (vn == S['missing'][0] and k == S['missing'][1]):
                pen.write(bxp - 4, Y(btp) - 3, '✓', 9)
    assert n == S['balloons']
    c.setFont('KR', 6.5)
    c.setFillColorRGB(0.4, 0.4, 0.4)
    c.drawString(L0 + 4, 26, '가상 샘플 — 모델 · 품번 · 그림 · 값은 모두 지어낸 것입니다.')


def raster(pdf_bytes, seed, turn=0):
    """벡터 PDF → 흐리고 살짝 기운 스캔 그림(JPEG). turn=90 이면 시계 반대로 눕힘(돌아간 스캔 흉내)"""
    r = random.Random(seed)
    tmp = Path(tempfile.mkdtemp())
    try:
        (tmp / 'a.pdf').write_bytes(pdf_bytes)
        subprocess.run(['pdftoppm', '-r', '120', '-gray', '-png', '-singlefile', str(tmp / 'a.pdf'), str(tmp / 'a')], check=True)
        im = Image.open(tmp / 'a.png').convert('L')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    im = im.rotate(r.uniform(-0.7, 0.7), resample=Image.BICUBIC, fillcolor=255, expand=False)
    if turn:
        im = im.rotate(turn, expand=True)
    im = im.filter(ImageFilter.GaussianBlur(0.5))
    px = im.load()
    for _ in range(2500):
        x, y = r.randrange(im.width), r.randrange(im.height)
        px[x, y] = max(0, px[x, y] - r.randrange(60, 160))
    jb = io.BytesIO()
    im.save(jb, 'JPEG', quality=62)
    jb.seek(0)
    return jb


def make_multi(sh, seed):
    """합본 — 1쪽 자주 검사(글자 층) + 2쪽 시트 부착(스캔 그림, 90° 누움 → 세로 쪽)"""
    one = io.BytesIO()
    c1 = canvas.Canvas(one, pagesize=(W, H), invariant=1)
    draw_seat(c1, seed + 1)
    c1.showPage()
    c1.save()
    seat_img = raster(one.getvalue(), seed + 2, turn=90)
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(W, H), invariant=1)
    c.setTitle('자주 검사 CHECK SHEET + 시트 부착 (가상 샘플)')
    draw_sheet(c, dict(sh, kind='text'), seed)
    c.showPage()
    c.setPageSize((H, W))
    c.drawImage(ImageReader(seat_img), 0, 0, H, W)
    c.showPage()
    c.save()
    return buf.getvalue()


def make_pdf(sh, seed):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(W, H), invariant=1)
    c.setTitle('자주 검사 CHECK SHEET (가상 샘플)')
    draw_sheet(c, sh, seed)
    c.showPage()
    c.save()
    return buf.getvalue()


def scanify(pdf_bytes, seed):
    """벡터 PDF → 그림 한 장(살짝 기울고 흐린 스캔 흉내) → 글자 층 없는 PDF"""
    r = random.Random(seed)
    tmp = Path(tempfile.mkdtemp())
    try:
        (tmp / 'a.pdf').write_bytes(pdf_bytes)
        subprocess.run(['pdftoppm', '-r', '120', '-gray', '-png', '-singlefile', str(tmp / 'a.pdf'), str(tmp / 'a')], check=True)
        im = Image.open(tmp / 'a.png').convert('L')
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    im = im.rotate(r.uniform(-0.7, 0.7), resample=Image.BICUBIC, fillcolor=255, expand=False)
    im = im.filter(ImageFilter.GaussianBlur(0.5))
    px = im.load()
    for _ in range(2500):   # 스캔 잡티
        x, y = r.randrange(im.width), r.randrange(im.height)
        px[x, y] = max(0, px[x, y] - r.randrange(60, 160))
    jb = io.BytesIO()
    im.save(jb, 'JPEG', quality=62)
    jb.seek(0)
    out = io.BytesIO()
    c = canvas.Canvas(out, pagesize=(W, H), invariant=1)
    c.drawImage(ImageReader(jb), 0, 0, W, H)
    c.showPage()
    c.save()
    return out.getvalue()


# ── 쓰기 ───────────────────────────────────────────────
OUT_PDF.mkdir(parents=True, exist_ok=True)
for old in OUT_PDF.glob('*.pdf'):
    old.unlink()
pdf_b64, file_meta, truth = {}, [], {}
for i, sh in enumerate(SHEETS):
    SIGN_BOXES.clear()
    if sh['kind'] == 'multi':
        data = make_multi(sh, 7000 + i)
        seat_box = [b for b in SIGN_BOXES if b[0] == '시트 결재 칸'][0]
        SIGN_BOXES[:] = [b for b in SIGN_BOXES if b[0] != '시트 결재 칸']
    else:
        data = make_pdf(sh, 7000 + i)
    if sh['kind'] == 'scan':
        data = scanify(data, 8000 + i)
    (OUT_PDF / sh['name']).write_bytes(data)
    pdf_b64[sh['name']] = base64.b64encode(data).decode('ascii')
    file_meta.append({'name': sh['name'], 'size': len(data), 'kind': sh['kind']} if sh['kind'] != 'multi'
                     else {'name': sh['name'], 'size': len(data), 'kind': 'multi', 'pages': ['text', 'scan'], 'rot': {'2': 90}})
    if sh['kind'] == 'multi':
        truth[sh['name'] + '#2'] = {
            'meta': {'date': SEAT['date'], 'process': SEAT['process'], 'model': SEAT['model'], 'part': SEAT['part'], 'name': SEAT['name'], 'mgmt': SEAT['mgmt']},
            'items': [{'no': '1', 'label': '시트 부착 확인(모델링 뷰 풍선 체크)', 'tool': '육안', 'subs': [], 'kind': 'attach'}],
            'hogis': [{'no': SEAT['hogi'], 'values': {'1': '%d/%d' % (SEAT['balloons'] - 1, SEAT['balloons'])}}],
            'problems': {'1': '%s %d번째 풍선 체크 없음' % (SEAT['missing'][0], SEAT['missing'][1] + 1)},
        }
    m = MODELS[sh['m']]
    truth[sh['name']] = {
        'meta': {'date': sh['date'], 'process': m['process'], 'model': m['model'], 'part': m['part'], 'name': m['name'], 'mgmt': m['mgmt']},
        'items': [{'no': it['no'], 'label': label_text(it), 'tool': it['tool'], 'subs': it['subs']} for it in m['items']],
        'hogis': [{'no': hno, 'values': v} for hno, v in sh['hogis']],
        'problems': sh['problems'],
    }
masks = [{'x': round(x0 / W, 4), 'y': round(t0 / H, 4), 'w': round((x1 - x0) / W, 4), 'h': round((t1 - t0) / H, 4), 'label': lab} for lab, x0, t0, x1, t1 in SIGN_BOXES]
# 공정별 기본 가릴 영역 — 시트 양식은 결재 칸 자리가 달라 따로(쪽을 바로 세운 가로 기준 비율). 열쇠는 공정명의 비교 낱말(공백 뺀 소문자)
lab, x0, t0, x1, t1 = seat_box
mask_procs = {SEAT['process'].replace(' ', '').lower(): [{'x': round(x0 / W, 4), 'y': round(t0 / H, 4), 'w': round((x1 - x0) / W, 4), 'h': round((t1 - t0) / H, 4), 'label': lab}]}

# 출하일정 — 출하일 · 고객사 · 모델 · 품번 · 품명 · 호기 · 수량
A, B = MODELS['A'], MODELS['B']
ships = [
    ('2026-10-26', 'MODEL-A100', A['part'], A['name'], '314501', 1, ''),
    ('2026-10-26', 'MODEL-A100', A['part'], A['name'], '314502', 1, ''),
    ('2026-10-27', 'MODEL-A100', A['part'], A['name'], '314503', 1, ''),
    ('2026-10-28', 'MODEL-A100', A['part'], A['name'], '314504', 1, '스캔본'),
    ('2026-10-29', 'MODEL-B200', B['part'], B['name'], '315010', 1, ''),
    ('2026-10-29', 'MODEL-B200', B['part'], B['name'], '315011', 1, ''),
    ('2026-10-30', 'MODEL-B200', B['part'], B['name'], '315012', 1, ''),
    ('2026-10-30', 'MODEL-B200', B['part'], B['name'], '315013', 1, ''),
    ('2026-11-02', 'MODEL-A100', None, A['name'], '314505', 1, '품번 미기재'),
    ('2026-11-03', 'MODEL-A100', A['part'], A['name'], '314506', 1, ''),
    ('2026-11-04', 'MODEL-B200', B['part'], B['name'], 315014, 1, ''),       # 엑셀 숫자 칸
    ('2026-11-05', 'MODEL-B200', B['part'], B['name'], '315015~315016', 2, '2대'),
]
wb = Workbook()
ws = wb.active
ws.title = '출하일정'
ws.append(['출하일정 (가상 샘플)'])
ws['A1'].font = Font(bold=True, size=13)
ws.append([])
ws.append(['출하일', '고객사', '모델', '품번', '품명', '호기', '수량', '비고'])
for cell in ws[3]:
    cell.font = Font(bold=True)
    cell.fill = PatternFill('solid', fgColor='DDE6F2')
for d, mo, pt, nm, hg, q, note in ships:
    ws.append([dt.date.fromisoformat(d), '가상고객', mo, pt, nm, hg, q, note])
    ws.cell(ws.max_row, 1).number_format = 'yyyy-mm-dd'
for col, wdt in zip('ABCDEFGH', [12, 10, 13, 11, 14, 16, 6, 12]):
    ws.column_dimensions[col].width = wdt
wb.save(ROOT / 'samples' / '출하일정_샘플.xlsx')

# 초대형 출하일정 형식(2026-11-10 답 ③ 사진의 칸 모양) — 제목 줄 · 두 줄 머리(하부 아래 Center/Side) · 「모   델」 띄어 쓴 머리 ·
# 품번 칸 없음 · 출하예정은 연도 없는 「월/일」 글자. 브랜드 · 회사 · 실제 모델 · 호기는 넣지 않고 지어낸 값만.
wb2 = Workbook()
w2 = wb2.active
w2.title = '초대형'
w2.append(['초대형 ALC 출하일정 (가상 샘플)'])
w2['A1'].font = Font(bold=True, size=14)
w2.append([])
w2.append(['SS', None, '순번', '연번', '모   델', '호기', '출고처', '참고 호기수', '출하예정', '상부', '하부', None, 'Line', '비   고'])
w2.append([None, None, None, None, None, None, None, None, None, None, 'Center', 'Side', None, None])
for row in (3, 4):
    for cell in w2[row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill('solid', fgColor='D9D9D9')
big = [('3', '11/11', 2, 19, 'MODEL-C300L', '520101', 'EA', '', '11/10'),
       ('4', '11/11', 3, 54, 'MODEL-D500', '610907', 'KR', '30호기(센터)', '11/10'),
       ('5', '11/12', 4, 55, 'MODEL-A100L', '314501', 'EA', '', '11/12'),
       ('6', '11/13', 5, 56, 'MODEL-C300L', '520102~520103', 'EA', '', '11/14')]
for a, b2, sn, yn, mo, hg, dst, ref, due in big:
    w2.append([a, b2, sn, yn, mo, hg, dst, ref or None, due, None, None, None, None, None])
wb2.save(ROOT / 'samples' / '출하일정_초대형형식_샘플.xlsx')

# 소형 · 대형 출하일정 형식(2026-11-10 4차 ③) — 파일마다 칸 구성이 달라도 「모델 · 호기 · 출하예정」 세 칸만 찾아 씀을 보이려는 가상 표.
# 소형: 머리 한 줄 · 엑셀 날짜 · 출하예정이 아직 빈 줄. 대형: 제목 · 작성일 줄 · 두 줄 머리(BOOM 아래 반입/출하) · 「호기수(참고)」 가 「호기」 앞 ·
# 품번 · 수량 칸도 있지만 세 칸만 읽음 · 「11/10(화)」 처럼 요일이 붙은 날짜.
wb3 = Workbook()
w3 = wb3.active
w3.title = '소형'
w3.append(['No', '모델', '호기', '출하예정', '상부', '하부', 'Line', '비고'])
for i, (mo, hg, due, up, lo, ln, note) in enumerate([
        ('MODEL-A100', '314501', '2026-11-12', 'O', 'O', 'L1', ''),
        ('MODEL-A100', '314502', '2026-11-12', 'O', '', 'L1', ''),
        ('MODEL-B200', '315010', '2026-11-13', '', 'O', 'L2', ''),
        ('MODEL-B200', '315011', None, '', '', 'L2', '일정 미정')]):
    w3.append([i + 1, mo, hg, dt.date.fromisoformat(due) if due else None, up or None, lo or None, ln, note or None])
    if due:
        w3.cell(w3.max_row, 4).number_format = 'yyyy-mm-dd'
for cell in w3[1]:
    cell.font = Font(bold=True)
wb3.save(ROOT / 'samples' / '출하일정_소형형식_샘플.xlsx')

wb4 = Workbook()
w4 = wb4.active
w4.title = '대형'
w4.append(['대형 출하일정 (가상 샘플)'])
w4.append([None, None, None, None, None, None, None, None, '작성 2026-11-10'])
w4.append(['연번', '모 델 명', '호기수(참고)', '호  기', '품번', '수량', '출하 예정', '상부', 'BOOM', None, 'Line', '비 고'])
w4.append([None, None, None, None, None, None, None, None, '반입', '출하', None, None])
for row in (3, 4):
    for cell in w4[row]:
        cell.font = Font(bold=True)
for yn, mo, ref, hg, pt, q, due, note in [
        (1, 'MODEL-A100', '12호기', '314503', 'QA-70011', 1, '11/10(화)', None),
        (2, 'MODEL-B200', '3호기', '315012~315013', 'QA-80022', 2, '11월 12일', '2대'),
        (3, 'MODEL-C300', '7호기', '520104', None, 1, '11/14(금)', None)]:
    w4.append([yn, mo, ref, hg, pt, q, due, None, None, None, None, note])
wb4.save(ROOT / 'samples' / '출하일정_대형형식_샘플.xlsx')

ship_rows = []   # 앱이 엑셀을 읽은 것과 같은 모양(호기 범위는 펼침) — 테스트가 parseShip 결과와 맞춰 봄
for d, mo, pt, nm, hg, q, note in ships:
    hs = [str(hg)] if '~' not in str(hg) else [str(x) for x in range(int(hg.split('~')[0]), int(hg.split('~')[1]) + 1)]
    for h1 in hs:
        ship_rows.append({'date': d, 'model': mo, 'part': pt or '', 'hogi': h1, 'qty': 1 if len(hs) > 1 else q, 'name': nm})
sample = {
    'note': '가상 샘플(2차) — 모델 · 품번 · 치수 · 값은 모두 지어낸 것입니다. scripts/make-samples.py 로 다시 만드세요(직접 고치지 마세요).',
    'refDate': '2026-10-28',
    'files': file_meta,
    'truth': truth,
    'masks': masks,
    'maskProcs': mask_procs,
    'shipFile': '출하일정_샘플.xlsx',
    'ships': ship_rows,
}
(ROOT / 'js' / 'sample-data.js').write_text(
    '/* 가상 샘플 — scripts/make-samples.py 가 만든 파일입니다. 직접 고치지 마세요. */\n'
    '(function (root) {\n  root.QCSample = ' + json.dumps(sample, ensure_ascii=False) + ';\n'
    "  if (typeof module === 'object' && module.exports) module.exports = root.QCSample;\n"
    "})(typeof self !== 'undefined' ? self : this);\n", encoding='utf-8')
(ROOT / 'js' / 'sample-pdfs.js').write_text(
    '/* 가상 샘플 검사표 PDF 의 base64 사본 — file:// 로 열어도 샘플이 열리도록. scripts/make-samples.py 가 만듭니다. 처음 샘플 PDF 를 열 때만 불러옵니다. */\n'
    '(function (root) {\n  root.QCSamplePdfs = ' + json.dumps(pdf_b64, ensure_ascii=False) + ';\n'
    "  if (typeof module === 'object' && module.exports) module.exports = root.QCSamplePdfs;\n"
    "})(typeof self !== 'undefined' ? self : this);\n", encoding='utf-8')
print('검사표 %d장 (%s) · 합 %.0fKB · 출하 %d줄 · 가릴 영역 %d' % (len(file_meta), ', '.join(f['kind'] for f in file_meta), sum(f['size'] for f in file_meta) / 1024, len(ships), len(masks)))
