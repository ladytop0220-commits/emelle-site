#!/usr/bin/env python3
"""
에멜 홈페이지 빌더 — Excel 샘플관리 파일로 사이트 자동 업데이트
================================================================

사용법:
  python3 build_emelle.py 샘플관리_20261026.xlsx emelle_template.html emelle.html

  1번째 인자: 쥬쥬님 샘플관리 Excel 파일
  2번째 인자: HTML 템플릿 (현재 사이트 HTML)
  3번째 인자: 저장할 새 HTML 파일

새 상품이 생기면:
  1. Excel 파일에 상품 행 추가 (이미지는 B열에 삽입)
  2. 섹터 컬럼(W열)에 분류 입력: maison / electro / essentiels / collection
  3. 이 스크립트 실행
  4. 생성된 emelle.html을 Vercel에 재업로드

Author: for 쥬쥬 (기펠인터내셔날 · 에멜)
"""

import sys
import os
import re
import json
import base64
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from PIL import Image
import io

# =============================================================
# 섹터 정의 — 색상 및 메타데이터
# =============================================================
SECTORS = {
    'maison':     {'no': 'N° 01', 'name': 'Maison',     'kr': '에멜 · 홈데코',
                   'tagline': '조명과 모빌 — 공간을 완성하는 작은 것들', 'tag': 'MAISON'},
    'electro':    {'no': 'N° 02', 'name': 'Électro',    'kr': '에멜 · 소형가전',
                   'tagline': '매일의 주방을 다정하게 만드는 기기들', 'tag': 'ÉLECTRO'},
    'essentiels': {'no': 'N° 03', 'name': 'Essentiels', 'kr': '에멜 · 생활용품',
                   'tagline': '매일 쓰는, 가장 중요한 것들', 'tag': 'ESSENTIELS'},
    'collection': {'no': 'N° 04', 'name': 'Collection', 'kr': '에멜 · 스페셜',
                   'tagline': '한정과 선물, 특별한 셀렉션', 'tag': 'COLLECTION'},
}

# 상품명 자동 분류 규칙 (섹터 컬럼이 비어있을 때 사용)
AUTO_SECTOR_RULES = [
    ('maison',     ['조명', '모빌', '램프', '캔들', '화병', '베이스', '린넨', '테이블', '러너', '데코', '오브제']),
    ('electro',    ['토스터', '토스토기', '에그메이커', '에그메이크', '그라인더', '글라인더',
                    '오프너', '오픈어', '다지기', '믹서', '블렌더', '전동', '인덕션', '에어프라이어', 'usb']),
    ('essentiels', ['프라이팬', '후라이팬', '냄비', '솥', '도마', '채칼', '식기', '수저', '포크',
                    '칼', '보관', '밀폐', '텀블러', '보온', '샐러드통', '그릇', '컵']),
    ('collection', ['세트', '선물', '한정', '리미티드', '콜라보', '시즌', '스페셜']),
]

def auto_classify(product_name: str) -> str:
    """상품명으로 섹터 자동 분류"""
    name_lower = product_name.lower()
    for sector, keywords in AUTO_SECTOR_RULES:
        for kw in keywords:
            if kw in name_lower:
                return sector
    return 'essentiels'  # 기본값

# =============================================================
# Excel 파싱 (openpyxl 없이 ZIP + XML 직접 파싱)
# =============================================================
NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_DR = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
NS_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'

def parse_xlsx(xlsx_path):
    """Excel 파일에서 상품 데이터와 이미지 추출"""
    print(f"[1/4] Excel 파일 읽기: {xlsx_path}")
    with zipfile.ZipFile(xlsx_path, 'r') as z:
        # Shared strings
        with z.open('xl/sharedStrings.xml') as f:
            sst = ET.parse(f).getroot()
        strings = []
        for si in sst.findall(NS + 'si'):
            parts = [t.text or '' for t in si.iter(NS + 't')]
            strings.append(''.join(parts))

        # Sheet 2 (샘플 요청 및 검토 현황)
        with z.open('xl/worksheets/sheet2.xml') as f:
            sheet = ET.parse(f).getroot()

        # Drawing (이미지 위치)
        try:
            with z.open('xl/drawings/_rels/drawing1.xml.rels') as f:
                rels_tree = ET.parse(f).getroot()
            drawing_rels = {r.get('Id'): r.get('Target').replace('../media/', '')
                            for r in rels_tree}
            with z.open('xl/drawings/drawing1.xml') as f:
                drawing = ET.parse(f).getroot()
        except KeyError:
            drawing_rels = {}
            drawing = None

        # 이미지 추출
        images = {}
        for name in z.namelist():
            if name.startswith('xl/media/'):
                with z.open(name) as f:
                    images[os.path.basename(name)] = f.read()

    # 행 데이터 파싱
    def cell_val(c):
        t = c.get('t', 'n')
        v = c.find(NS + 'v')
        if v is None: return None
        if t == 's':
            try: return strings[int(v.text)]
            except: return v.text
        return v.text

    rows = {}
    sheet_data = sheet.find(NS + 'sheetData')
    for row in sheet_data.findall(NS + 'row'):
        r_idx = int(row.get('r'))
        cells = {}
        for c in row.findall(NS + 'c'):
            ref = c.get('r')
            col_letter = re.match(r'([A-Z]+)', ref).group(1)
            val = cell_val(c)
            if val is not None:
                cells[col_letter] = val
        if cells:
            rows[r_idx] = cells

    # 이미지-행 매핑
    row_images = {}
    if drawing is not None:
        for anchor in drawing:
            tag = anchor.tag.split('}')[1]
            if tag not in ('twoCellAnchor', 'oneCellAnchor'): continue
            from_el = anchor.find(NS_DR + 'from')
            if from_el is None: continue
            r = int(from_el.find(NS_DR + 'row').text) + 1
            c = int(from_el.find(NS_DR + 'col').text) + 1
            if c != 2: continue  # B열(이미지) 만
            pic = anchor.find('.//' + NS_DR + 'pic')
            if pic is None: continue
            blip = pic.find('.//' + NS_A + 'blip')
            if blip is None: continue
            rid = blip.get(NS_R + 'embed')
            img_file = drawing_rels.get(rid)
            if img_file:
                row_images[r] = img_file

    return rows, row_images, images

# =============================================================
# 이미지 처리 (리사이즈 + base64 인코딩)
# =============================================================
def process_image(raw_bytes, size=(480, 480), quality=72):
    """이미지를 정사각형 480x480으로 센터 크롭하고 JPEG로 압축"""
    img = Image.open(io.BytesIO(raw_bytes))
    img = img.convert('RGB')
    w, h = img.size
    s = min(w, h)
    left, top = (w - s) // 2, (h - s) // 2
    img = img.crop((left, top, left + s, top + s))
    img = img.resize(size, Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, 'JPEG', quality=quality, optimize=True)
    b64 = base64.b64encode(buf.getvalue()).decode('ascii')
    return 'data:image/jpeg;base64,' + b64

# =============================================================
# 상품 데이터 정제
# =============================================================
def build_products(rows, row_images, images):
    """Excel 행에서 상품 데이터 추출"""
    print(f"[2/4] 상품 데이터 추출 중...")
    products = []
    for r_idx, cells in sorted(rows.items()):
        if r_idx < 5: continue  # 헤더 행
        no = cells.get('A')
        name_raw = cells.get('C', '').strip()
        price = cells.get('D', '')
        meta = cells.get('H', '').replace('단가·MOQ: ', '').replace('MOQ ', 'MOQ ')
        done_1st = cells.get('I') == '✔'
        sector_manual = cells.get('W', '').strip().lower()  # 사용자가 섹터 지정한 경우

        if not name_raw: continue

        # 상품명과 설명 분리 (첫 줄 = 이름, 나머지 = 설명)
        parts = name_raw.split('\n')
        name = parts[0].strip()
        desc = ' · '.join(p.strip() for p in parts[1:] if p.strip())

        # 섹터 결정: 수동 지정 > 자동 분류
        sector = sector_manual if sector_manual in SECTORS else auto_classify(name)

        # 이미지
        img_file = row_images.get(r_idx)
        uri = None
        if img_file and img_file in images:
            try:
                uri = process_image(images[img_file])
            except Exception as e:
                print(f"    ! 이미지 변환 실패 (행 {r_idx}): {e}")

        # 뱃지
        badge = '1차 완료' if done_1st else None

        # 가격/MOQ 요약
        price_str = f"{price} USD" if price else ''
        if meta:
            meta_short = meta.split('\n')[0][:40]
            desc_combined = f"{price_str} · {meta_short}" if price_str and meta_short else (price_str or meta_short)
        else:
            desc_combined = price_str

        products.append({
            'no': no, 'name': name, 'desc': desc_combined,
            'sector': sector, 'badge': badge, 'uri': uri,
        })
    print(f"    총 {len(products)}개 상품 로드")
    return products

# =============================================================
# 섹터별 큐레이션 (각 섹터 상위 N개만 홈페이지에 노출)
# =============================================================
def curate(products, per_sector=5):
    """섹터별로 완료(✔)된 상품 우선, 그다음 최신순으로 N개 선택"""
    print(f"[3/4] 섹터별 상위 {per_sector}개 큐레이션...")
    buckets = {k: [] for k in SECTORS}
    for p in products:
        if p['sector'] in buckets and p['uri']:
            buckets[p['sector']].append(p)

    curated = []
    for sector_key in SECTORS:
        items = buckets[sector_key]
        # 1차 완료 상품 우선, 그다음 번호 순
        def sort_key(x):
            try: n = int(x['no']) if x['no'] else 999
            except: n = 999
            return (not x['badge'], n)
        items.sort(key=sort_key)
        curated.extend(items[:per_sector])
        print(f"    · {sector_key}: {len(items)}개 중 {min(per_sector, len(items))}개 선택")
    return curated

# =============================================================
# HTML 생성
# =============================================================
def esc(s):
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')

def build_html(curated_products, template_html):
    """템플릿 HTML에 큐레이션된 상품 삽입"""
    print(f"[4/4] HTML 생성 중...")
    strips_html = []
    for key, meta in SECTORS.items():
        items = [p for p in curated_products if p['sector'] == key]
        cards = []
        for p in items:
            badge = f'<span class="product-card-badge">{esc(p["badge"])}</span>' if p.get('badge') else ''
            cards.append(f'''<a href="#{key}" class="product-card">
          <div class="product-card-img">
            {badge}
            <img src="{p['uri']}" alt="{esc(p['name'])}" loading="lazy">
          </div>
          <p class="product-card-sector">{meta['tag']}</p>
          <h4 class="product-card-name">{esc(p['name'])}</h4>
          <p class="product-card-meta">{esc(p['desc'])}</p>
        </a>''')
        cards_joined = '\n          '.join(cards)
        strips_html.append(f'''      <div class="sector-strip">
        <div class="sector-head sector-head-{key}">
          <div>
            <p class="sh-no">{meta['no']}</p>
            <h3 class="sh-name">{meta['name']}</h3>
            <p class="sh-kr">{meta['kr']}</p>
          </div>
          <div>
            <p class="sh-tagline">{meta['tagline']}</p>
            <a href="#{key}" class="sh-link">Discover <span>→</span></a>
          </div>
        </div>
        <div class="sector-products">
          {cards_joined}
        </div>
      </div>''')

    new_section = '''  <section class="collection">
    <div class="container">
      <div class="collection-header">
        <p class="section-eyebrow">PREMIÈRE COLLECTION · 첫 컬렉션</p>
        <h2 class="section-title"><em>Four Houses</em>, One Maison</h2>
        <p class="collection-sub">네 개의 하우스, 하나의 메종 — 광저우에서 소싱 중인 신제품들을 미리 만나보세요.</p>
      </div>
''' + '\n'.join(strips_html) + '\n    </div>\n  </section>'

    pattern = re.compile(r'  <section class="collection">.*?</section>', re.DOTALL)
    new_html = pattern.sub(new_section, template_html, count=1)
    return new_html

# =============================================================
# 메인
# =============================================================
def main():
    if len(sys.argv) < 4:
        print(__doc__)
        print("\n필수 인자가 부족합니다. 사용법을 확인하세요.")
        sys.exit(1)

    xlsx_path = sys.argv[1]
    tmpl_path = sys.argv[2]
    out_path = sys.argv[3]

    if not os.path.exists(xlsx_path):
        print(f"❌ Excel 파일을 찾을 수 없음: {xlsx_path}")
        sys.exit(1)
    if not os.path.exists(tmpl_path):
        print(f"❌ HTML 템플릿을 찾을 수 없음: {tmpl_path}")
        sys.exit(1)

    rows, row_images, images = parse_xlsx(xlsx_path)
    products = build_products(rows, row_images, images)
    curated = curate(products)

    with open(tmpl_path, 'r', encoding='utf-8') as f:
        template = f.read()
    new_html = build_html(curated, template)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(new_html)

    size_kb = os.path.getsize(out_path) / 1024
    print(f"\n✅ 완료! {out_path} ({size_kb:.0f}KB)")
    print(f"   총 {len(curated)}개 상품이 홈페이지에 배치됨")
    print(f"\n   다음 단계: 이 HTML 파일을 Vercel에 재업로드하면 사이트가 업데이트됩니다.")

if __name__ == '__main__':
    main()
