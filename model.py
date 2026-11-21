import json, io
import numpy as np
import openpyxl
book = None

def emit(x):
    return json.dumps(x, ensure_ascii=False, allow_nan=False)

def key(x):
    return str(x).strip()

DEMAND_SHORT = {'9111': '민간소비', '9112': '정부소비', '9121': '민간고정자본형성', '9122': '정부고정자본형성',
                '9131': '재고증감', '9132': '귀중품 순취득', '9140': '수출', '9190': '최종수요계'}
DEMAND_GROUP = {'9111': '소비', '9112': '소비', '9121': '투자', '9122': '투자', '9131': '투자', '9132': '투자', '9140': '수출'}
DEMAND_SHEETS = {'production': '최종수요 항목별 생산유발액', 'imports': '최종수요 항목별 수입유발액',
                 'valueAdded': '최종수요 항목별 부가가치유발액'}

def build_demand():
    """최종수요 항목별 유발액 시트(생산·수입·부가가치)를 읽어 항목별 합계와 산업별 값을 만든다.
    선택 기능이라 시트가 없거나 형식이 달라도 전체 초기화는 실패하지 않는다."""
    try:
        # 연도에 따라 시트 이름의 공백이 다르다(예: 2022년은 '최종수요항목별'), 공백을 무시하고 찾는다.
        found = {}
        for k, target in DEMAND_SHEETS.items():
            hit = [s for s in book if s.replace(' ', '') == target.replace(' ', '')]
            if not hit:
                return None
            found[k] = hit[0]
        items = None; total = {}; per = {c: {} for c in codes}
        for k, s in found.items():
            rows = book[s]
            cols = [(i, key(c)) for i, c in enumerate(rows[4]) if i >= 2 and c is not None]
            if items is None:
                items = [{'code': c, 'name': DEMAND_SHORT.get(c, key(rows[5][i] or c).replace('\n', ' ')),
                          'group': DEMAND_GROUP.get(c, '합계')} for i, c in cols]
            elif [c for _, c in cols] != [it['code'] for it in items]:
                return None
            rr = {key(r[0]): r for r in rows[6:] if r and r[0] is not None}
            for c in codes:
                per[c][k] = [float(rr[c][i] or 0) for i, _ in cols]
            total[k] = [float(rr['9819'][i] or 0) for i, _ in cols]
        return {'items': items, 'total': total, 'sectors': per}
    except Exception:
        return None

def initialize(raw):
    global book, codes, names, L, M, V, sheets, valid
    wb = openpyxl.load_workbook(io.BytesIO(bytes(raw)), read_only=True, data_only=True)
    book = {s: list(wb[s].values) for s in wb.sheetnames}
    wb.close()
    sheets = list(book)
    need = ['국산투입계수(Ad)', '수입투입계수(Am)', '총투입계수(A)', '생산유발계수', '수입유발계수', '부가가치유발계수']
    miss = [s for s in need if s not in book]
    if miss:
        raise ValueError('필요한 시트가 없습니다: ' + ', '.join(miss))
    base = book['국산투입계수(Ad)']
    # 부문 코드: 열 머리글 중 합계·지표 코드(9000 이상 숫자)를 제외한 것.
    # 기본부문·소분류·중분류는 숫자 코드, 대분류는 영문 코드(A, C01 ...)
    codes = []
    for c in base[4][2:]:
        if c is None:
            continue
        s = key(c)
        if s.isdigit() and int(s) >= 9000:
            continue
        codes.append(s)
    if len(codes) < 2 or len(set(codes)) != len(codes):
        raise ValueError('부문 코드를 읽을 수 없습니다. 투입산출표 형식을 확인하세요.')
    nm = {key(r[0]): key(r[1]) for r in base[6:] if r and r[0] is not None and r[1] is not None}
    missing = [c for c in codes if c not in nm]
    if missing:
        raise ValueError('부문명을 찾을 수 없는 코드: ' + ', '.join(missing[:5]))
    names = [nm[c] for c in codes]
    def matrix(s):
        rows = book[s]
        ri = {key(r[0]): i for i, r in enumerate(rows) if r and r[0] is not None}
        ci = {key(c): i for i, c in enumerate(rows[4]) if c is not None}
        a = np.array([[rows[ri[c]][ci[k]] for k in codes] for c in codes], dtype=float)
        if not np.isfinite(a).all():
            raise ValueError(s + ': 비수치 또는 결측 데이터')
        return a
    L = matrix('생산유발계수'); M = matrix('수입유발계수'); V = matrix('부가가치유발계수')
    ad = matrix('국산투입계수(Ad)'); am = matrix('수입투입계수(Am)')
    rows = book['총투입계수(A)']
    ci = {key(c): i for i, c in enumerate(rows[4]) if c is not None}
    vr = next(r for r in rows if r and r[0] is not None and key(r[0]) == '9690')
    v = np.array([vr[ci[c]] for c in codes], dtype=float)
    n = len(codes)
    residuals = {'생산': float(np.max(np.abs((np.eye(n) - ad) @ L - np.eye(n)))),
                 '수입': float(np.max(np.abs(am @ L - M))),
                 '부가가치': float(np.max(np.abs(v[:, None] * L - V)))}
    valid = max(residuals.values()) < 1e-8
    if not valid:
        raise ValueError('계수 검증 실패: ' + str(residuals))
    # 산업 연쇄효과: 영향력계수(후방, 열합 기준) · 감응도계수(전방, 행합 기준)
    col = L.sum(axis=0); row = L.sum(axis=1)
    back = col / col.mean(); fwd = row / row.mean()
    linkage = []
    for c, nme, b, f in zip(codes, names, back, fwd):
        if b > 1 and f > 1: kind = '전·후방 모두 큼'
        elif b > 1: kind = '후방연쇄 큼'
        elif f > 1: kind = '전방연쇄 큼'
        else: kind = '연쇄 작음'
        linkage.append({'code': c, 'name': nme, 'backward': float(b), 'forward': float(f), 'type': kind})
    return emit({'sectors': [{'code': c, 'name': n_} for c, n_ in zip(codes, names)],
                 'sheets': sheets, 'residuals': residuals, 'linkage': linkage, 'demand': build_demand()})

def explore(sheet, rowcode):
    rows = book[sheet]
    choices = [{'code': str(r[0]), 'name': str(r[1])} for r in rows[6:] if len(r) > 1 and r[0] is not None and r[1] is not None]
    chosen = next((r for r in rows[6:] if str(r[0]) == rowcode), None)
    if chosen is None: chosen = next(r for r in rows[6:] if r[0] is not None and r[1] is not None)
    cells = []
    for i in range(2, len(rows[4])):
        if rows[4][i] is None: continue
        val = chosen[i] if i < len(chosen) else None
        if isinstance(val, (float, int)): val = float(val)
        elif val is not None: val = str(val)
        cells.append({'code': str(rows[4][i]), 'name': str(rows[5][i] or ''), 'value': val})
    return emit({'rows': choices, 'selected': str(chosen[0]), 'name': str(chosen[1]),
                 'unit': str(rows[3][0] or '係数').replace('係数', '계수 / 비율: 행·열 명칭 확인'), 'cells': cells})

def simulate(payload):
    if not valid: raise ValueError('검증된 데이터가 없습니다.')
    shocks = json.loads(payload); f = np.zeros(len(codes))
    for s in shocks:
        a = float(s['amount'])
        if not np.isfinite(a) or a < 0: raise ValueError('수요 증가액은 유한한 0 이상의 수여야 합니다.')
        f[codes.index(s['code'])] += a
    if not f.sum() > 0: raise ValueError('한 산업 이상의 양의 수요 증가액을 입력하세요.')
    x = L @ f; m = M @ f; v = V @ f
    result = [{'code': c, 'name': n, 'production': float(a), 'imports': float(b), 'valueAdded': float(d)}
              for c, n, a, b, d in zip(codes, names, x, m, v)]
    result.sort(key=lambda r: r['production'], reverse=True)
    return emit({'demand': float(f.sum()), 'production': float(x.sum()), 'imports': float(m.sum()),
                 'valueAdded': float(v.sum()), 'multiplier': float(x.sum() / f.sum()), 'rows': result})
