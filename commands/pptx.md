Create a PPTX presentation based on the following description:

$ARGUMENTS

---

## How to execute this skill

1. Write a Python script (e.g. `/tmp/make_pptx.py`) using python-pptx following all design rules below
2. Run it: `uv run --with python-pptx python3 /tmp/make_pptx.py`
3. Preview: convert to PDF then PNG via LibreOffice, then Read each slide image
4. If layout issues are found, fix and re-run. Always confirm visually before reporting done.

> **브랜드/발표 형식 모드 요청 시** (예: 특정 회사 "○○ 형식으로") — §1~7 범용 규칙 위에 **§8 형식 모드 추가 패턴**을 따라 디자인 토큰 · 번호 배지 · 아이콘(로컬 PNG) · 카드 레이아웃을 입힌다. 구체 브랜드 모드와 아이콘 자산은 **각자 로컬에만** 두고(공유 repo엔 패턴만), 발표자의 기존 구조(머리글·셰브론·캡션·"핵심 원칙" 박스)는 보존한다. 평소(기본)에는 §1~7만 사용.

Preview pipeline:
```bash
libreoffice --headless --convert-to pdf OUTPUT.pptx --outdir /tmp/
pdftoppm -r 150 -png /tmp/OUTPUT.pdf /tmp/preview/slide
```

---

## Design Rules (follow strictly)

### 1. Color palette — ONE accent color only

```python
from pptx.dml.color import RGBColor

NAVY  = RGBColor(0x1a, 0x4a, 0x8a)   # slide headers, table headers
BLUE  = RGBColor(0x1f, 0x75, 0xc4)   # section titles, phase headers (accent)
DARK  = RGBColor(0x33, 0x33, 0x33)   # body text
WHITE = RGBColor(0xff, 0xff, 0xff)   # text on dark backgrounds
# Pick ONE additional accent for highlights (e.g. GREEN_LIGHT for active cells)
# Never use more than one highlight color in a single presentation
```

### 2. Text indentation — paragraph level, NOT spaces

Do NOT simulate indentation with spaces or tab characters.
Use python-pptx paragraph properties (`marL`, `indent`, `buChar`).

```python
from lxml import etree

_NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

def set_para_indent(para, level):
    """
    level = -1 : section header  — no bullet, marL=88900
    level =  0 : 1st-level item  — '-' bullet, marL=457200, indent=-317500 (hanging)
    level =  1 : 2nd-level item  — '-' bullet, marL=914400, indent=-317500 (hanging)
    """
    pPr = para._p.get_or_add_pPr()
    for tag in ('buNone', 'buChar', 'buAutoNum', 'buFont', 'buFontTx'):
        for el in pPr.findall(f'{{{_NS_A}}}{tag}'):
            pPr.remove(el)
    if level < 0:
        pPr.set('marL', '88900')
        pPr.set('indent', '0')
        etree.SubElement(pPr, f'{{{_NS_A}}}buNone')
    else:
        pPr.set('marL', '457200' if level == 0 else '914400')
        pPr.set('indent', '-317500')
        etree.SubElement(pPr, f'{{{_NS_A}}}buChar').set('char', '-')
```

### 3. Spacing — use space_before, not empty paragraphs

```python
from pptx.util import Pt

# Before a section header (except the first): 8pt
p.space_before = Pt(8)

# Between bullet items: 1pt
p.space_before = Pt(1)
```

Never insert an empty `("")` paragraph just to add vertical space. Use `space_before`.

### 4. Textbox helper

```python
from pptx.util import Pt, Emu
from pptx.enum.text import PP_ALIGN

def add_tb(slide, l, t, w, h, lines, align=PP_ALIGN.LEFT):
    """
    lines: list of tuples
        (text, size_pt, bold, RGBColor)
        (text, size_pt, bold, RGBColor, space_before_pt)
        (text, size_pt, bold, RGBColor, space_before_pt, indent_level)

    indent_level: -1=header  0=bullet  1=indented bullet  None=no change
    """
    box = slide.shapes.add_textbox(Emu(l), Emu(t), Emu(w), Emu(h))
    tf  = box.text_frame
    tf.word_wrap = True
    first = True
    for item in lines:
        text, size, bold, color = item[:4]
        sp  = item[4] if len(item) > 4 else 0
        lvl = item[5] if len(item) > 5 else None
        p   = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        if sp:
            p.space_before = Pt(sp)
        if lvl is not None:
            set_para_indent(p, lvl)
        r = p.add_run()
        r.text, r.font.size, r.font.bold, r.font.color.rgb = text, Pt(size), bold, color
    return box
```

Usage pattern:
```python
lines = [
    ('Section Title',   12, True,  NAVY, 0,  -1),   # header, no bullet
    ('First point',     10, False, DARK, 1,   0),   # '-' bullet
    ('Sub point',       10, False, DARK, 1,   1),   # '-' indented bullet
    ('Next Section',    12, True,  NAVY, 8,  -1),   # 8pt gap before new section
    ('Another point',   10, False, DARK, 1,   0),
]
add_tb(slide, left, top, width, height, lines)
```

### 5. Table style — full black grid, white body cells, ONE accent color

**모든 셀은 4면(L/R/T/B) 검은색 테두리(full grid)** 를 두른다. **헤더(컬럼) 행만 accent 배경(NAVY)** 이고, **헤더를 제외한 모든 셀은 흰색 배경**이다. python-pptx엔 셀 테두리 고수준 API가 없어 XML 헬퍼로 설정한다.

```python
from pptx.oxml.ns import qn
from pptx.util import Pt

def set_cell_border(cell, color="000000", width=Pt(0.75)):
    """모든 셀에 4면 검은 테두리. ln* 은 fill보다 앞서야 하므로 tcPr 맨 앞에 삽입."""
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    for tag in ("a:lnB", "a:lnT", "a:lnR", "a:lnL"):   # 역순 삽입 → 최종 L,R,T,B 순
        ln = tcPr.makeelement(qn(tag), {"w": str(int(width)), "cap": "flat",
                                        "cmpd": "sng", "algn": "ctr"})
        sf = tcPr.makeelement(qn("a:solidFill"), {})
        c  = tcPr.makeelement(qn("a:srgbClr"), {"val": color})
        sf.append(c); ln.append(sf); tcPr.insert(0, ln)

def set_cell_fmt(cell, text, bg=WHITE, fg=DARK, size=10, bold=False,
                 align=PP_ALIGN.CENTER, wrap=True, border=True):
    cell.fill.solid()
    cell.fill.fore_color.rgb = bg if bg else WHITE     # 헤더 외 전부 흰색
    if border:
        set_cell_border(cell)                          # 모든 셀 검은 4면 테두리
    tf = cell.text_frame
    tf.word_wrap = wrap
    para = tf.paragraphs[0]
    para.alignment = align
    run = para.runs[0] if para.runs else para.add_run()
    run.text, run.font.size, run.font.bold, run.font.color.rgb = text, Pt(size), bold, fg
```

Rules:
- **모든 셀: 4면 검은색 테두리**(`set_cell_border`) — full grid.
- **헤더(컬럼) 행:** `bg=NAVY, fg=WHITE, bold=True`.
- **헤더를 제외한 모든 셀: 흰색 배경**(`set_cell_fmt` 기본값 = `WHITE`).
- 강조/활성 셀: accent **1색만** — 다색 금지.

```python
# Header row (accent) + 모든 본문 셀 = 흰 배경 + 검은 grid (기본)
for ci, h in enumerate(headers):
    set_cell_fmt(tbl.cell(0, ci), h, bg=NAVY, fg=WHITE, bold=True)
for ri in range(1, rows):
    for ci in range(cols):
        set_cell_fmt(tbl.cell(ri, ci), data[ri-1][ci])      # 흰 배경 + 검은 테두리
set_cell_fmt(tbl.cell(2, 3), 'Active', bg=BLUE, fg=WHITE)   # accent 1색만

# Bad: 한 표에 강조색 2개 이상
# set_cell_fmt(tbl.cell(1,1), '...', bg=GREEN)   ← BLUE 이미 썼으면 금지
```

### 6. Shape deletion helper

```python
def del_shape(shape):
    shape._element.getparent().remove(shape._element)
```

### 7. Slide size

Standard widescreen: `prs.slide_width=9144000 EMU (10 in)`, `prs.slide_height=5143500 EMU (5.63 in)`
Wide (임재환 style): `W=9887040 (10.83 in)`, `H=6858000 (7.50 in)` — read from base file.

Always use `W, H = prs.slide_width, prs.slide_height` instead of hardcoded values.
Use fractional positioning: `Emu(int(W * 0.08))` rather than absolute EMU values.

---

## 8. 형식 모드 추가 (브랜드/발표 형식 확장 패턴)

범용 규칙(§1~7) 위에 **특정 브랜드/발표 형식을 일관 적용하는 "모드"** 를 얹는 방법. 핵심 원칙: 발표자가 만든 기존 구조(머리글·셰브론·캡션·"핵심 원칙" 박스 등)는 **보존**하고 그 위에 디자인 언어만 더한다. **구체 브랜드 토큰·아이콘 자산은 공유 repo에 두지 않고 각자 로컬**(예: `~/.claude/commands/pptx.md` 하단의 모드 섹션 + `assets/<brand>-icons/`)에 둔다 — 이 §8은 그 *패턴*만 정의한다.

### 8.1 모드 구성요소
1. **디자인 토큰 블록** — 브랜드 색을 상수로(BRAND/ACCENT/BADGE/CARD/BODY/CAPTION…). 모드에선 §1 범용 팔레트를 대체. "강조 1색" 규칙은 브랜드가 의도적으로 2색 대비(예: 블루+주황)를 쓰면 *그 모드 한정* 예외.
2. **타이포그래피** — run마다 latin/ea/cs typeface 3종 지정(한글 깨짐·폰트 혼용 방지). 제목/본문/배지의 크기·굵기·색 규정.
3. **번호 배지** — 작은 원형 도형 + 흰 숫자(단계 표시).
4. **아이콘** — 이 환경엔 아이콘 MCP가 없으므로 **로컬 PNG**를 `add_picture`로 삽입. 색 변형이 필요하면 색별 PNG를 따로 둔다(python-pptx 리컬러는 제한적).
5. **카드/강조 도형** — 라운드 사각 카드, 강조 박스, 셰브론/화살표 헬퍼.
6. **레이아웃 패턴** — N단계 가로 카드 / 2단 비교 / 세로 스택 등.
7. **보존 규칙** — 기존 머리글·로고·캡션·의도된 색 대비 유지.
8. **검증** — LibreOffice 미리보기로 모든 슬라이드 시각 확인.

### 8.2 재사용 헬퍼 (브랜드 무관 — 토큰/경로만 바꿔 사용)

```python
import os
from pptx.util import Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

PTX = 12700  # 1pt = 12700 EMU

def set_run_font(run, latin="Arial", ea="맑은 고딕"):
    """run마다 latin/ea/cs typeface 지정 — 한글·라틴 혼용 시 필수."""
    rPr = run._r.get_or_add_rPr()
    for tag, face in (('a:latin', latin), ('a:ea', ea), ('a:cs', ea)):
        el = rPr.find(qn(tag))
        if el is None:
            el = etree.SubElement(rPr, qn(tag))
        el.set('typeface', face)

def num_badge(slide, n, left, top, d=Emu(254000), fill=None, line=WHITE, fontc=WHITE):
    """원형 숫자 배지(단계 표시). d 기본=20pt. fill/line/fontc는 모드 토큰으로 전달."""
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, d, d)
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.color.rgb = line; s.line.width = Pt(1.5); s.shadow.inherit = False
    tf = s.text_frame; tf.word_wrap = False
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Emu(0)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = str(n)
    r.font.size = Pt(10); r.font.bold = True; r.font.color.rgb = fontc
    set_run_font(r); return s

def png_icon(slide, icon_dir, name, left, top, size=Emu(304800), variant="blue"):
    """로컬 단색 PNG 아이콘 삽입. 색은 PNG 자체 색(변형은 색별 파일: name_blue.png / name_white.png)."""
    return slide.shapes.add_picture(os.path.join(icon_dir, f"{name}_{variant}.png"),
                                    left, top, height=size)

def card(slide, left, top, w, h, fill, line, rounded=True):
    shp = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, left, top, w, h)
    shp.fill.solid(); shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = line; shp.line.width = Pt(1); shp.shadow.inherit = False
    return shp
```

### 8.3 배치 관례
- **가로 카드**: 배지를 카드 상단 중앙에 걸치게(top ≈ 카드top−11pt, left ≈ 카드중심−10pt), 아이콘(≈24pt)을 그 아래(top ≈ 카드top+13pt), 카드 텍스트는 세로정렬 Top + 상단여백 ~44pt로 내려 배지·아이콘 공간 확보.
- **세로 스택**: 배지를 박스 왼쪽 세로 중앙(left ≈ 박스left+8pt, top ≈ 박스top + 높이/2 − 10pt), 박스 txBody `lIns`를 ~26pt로 늘려 텍스트가 배지를 피하게.
- **2단 비교**: 헤더 펄(좌=회색/우=브랜드색 관례)에 흰 아이콘 ~14pt, 컬럼 제목은 브랜드색.

### 8.4 모드 분리 (중요)
- 공유 repo엔 **이 §8 패턴만** 둔다.
- 구체 브랜드 모드(실제 토큰 hex · 아이콘 파일명 · 정확한 좌표)와 **아이콘 PNG는 각자 로컬**에 둔다 → 브랜드 자산이 공개 repo에 올라가지 않게.

### 8.5 검증
- 모드 적용 후 LibreOffice 미리보기로 전 슬라이드 확인: 제목이 브랜드색인지, 배지·아이콘이 텍스트 위 올바른 위치인지, 헤더 펄 아이콘 색이 맞는지, 보존 요소(머리글·셰브론·캡션·강조 박스)가 유지됐는지.

---

## Checklist before finishing

- [ ] Ran the script successfully (no errors)
- [ ] Previewed every slide as PNG
- [ ] No text overlap with other shapes
- [ ] Table has at most ONE highlight color
- [ ] Indentation uses `set_para_indent()`, not spaces
- [ ] Spacing uses `space_before`, not empty paragraphs
- [ ] Output file saved to the requested path

**형식 모드(§8) 적용 시 추가:**
- [ ] 제목/단계 제목이 브랜드색
- [ ] 번호 배지가 단계 수만큼 정렬되어 얹힘
- [ ] 아이콘(로컬 PNG)이 카드/헤더에 일관 크기·색으로 배치
- [ ] run마다 typeface(latin/ea/cs) 지정
- [ ] 기존 구조(머리글·셰브론·캡션·강조 박스) 보존
