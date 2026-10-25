#!/usr/bin/env python3
"""kexcel.py — 한국식 엑셀 도구 (Claude가 한국 사무 엑셀을 정리·검사·가리기 하는 도구)

쓰는 법 (파일은 .xlsx / .xlsm / .csv. 원본은 건드리지 않고 새 파일로 저장)
  python3 kexcel.py info     명단.xlsx                       # 시트·칸·칸마다 무엇이 든 칸인지(주민번호·전화·날짜…)
  python3 kexcel.py check    명단.xlsx [--out 검사.xlsx]       # 주민번호·사업자번호·전화·이메일·날짜가 맞는지, 틀린 칸 노랗게
  python3 kexcel.py clean    명단.xlsx [--what all] [--dry]    # 날짜·전화·사업자번호·금액·띄어쓰기 한 모양으로
  python3 kexcel.py mask     명단.xlsx [--words 홍길동] [--dry] # 주민번호·전화·이메일·계좌·이름 가리기
  python3 kexcel.py addr     명단.xlsx --col 주소               # 주소를 시도·시군구·읍면동·나머지로 나누기
  python3 kexcel.py age      명단.xlsx --col 주민번호 [--on 2026-10-25]  # 생년월일·만 나이·성별 칸 더하기
  python3 kexcel.py workdays 2026-10-01 2026-10-31           # 영업일 수(주말·공휴일 빼고)
  python3 kexcel.py workdays 명단.xlsx --start 시작일 --end 종료일   # 줄마다 영업일 수 칸 더하기
  python3 kexcel.py workdays 2026-10-25 --add 10             # 10 영업일 뒤 날짜
  python3 kexcel.py holidays 2026                             # 그해 공휴일 목록
  python3 kexcel.py money    1234500                          # 일금 일백이십삼만사천오백원정
  python3 kexcel.py money    견적.xlsx --col 금액              # 한글 금액 칸 더하기
  python3 kexcel.py combine  a.xlsx b.xlsx c.csv --out 합본.xlsx [--same 성명=이름]  # 여러 파일·시트를 한 표로(칸 이름 맞춰)
  python3 kexcel.py split    명단.xlsx --by 부서 [--files]      # 칸 값마다 시트(또는 파일)로 나누기
  python3 kexcel.py dupes    명단.xlsx --cols 이름,전화 [--remove]  # 중복 줄 찾기(표시) 또는 지우기
  python3 kexcel.py tocsv    명단.xlsx [--encoding cp949]       # 관공서·은행 업로드용 CSV

공통: --sheet 시트이름, --header 칸이름줄번호(안 주면 알아서 찾음), --out 저장할 곳
준비물: pip install --break-system-packages openpyxl holidays
"""
import argparse
import csv
import datetime as dt
import os
import re
import sys
import unicodedata
import warnings

warnings.filterwarnings("ignore")

TODAY = dt.date.today()


# ───────────────────────── 파일 열기·저장 ─────────────────────────

def _need(mod, pkg=None):
    try:
        return __import__(mod)
    except ImportError:
        print(f"{pkg or mod} 가 없습니다. 먼저 이것을 실행하세요:\n"
              f"  pip install --break-system-packages {pkg or mod}", file=sys.stderr)
        sys.exit(2)


def _read_csv_text(path):
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-16"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            continue
    raise SystemExit("CSV 글자 모양(인코딩)을 알 수 없습니다. 엑셀에서 'CSV UTF-8'로 다시 저장해 주세요.")


def load(path):
    """(통합문서, 원래 형식) — csv 도 openpyxl 통합문서로 바꿔서 다룬다(값은 글자 그대로)."""
    openpyxl = _need("openpyxl")
    if not os.path.exists(path):
        raise SystemExit(f"파일이 없습니다: {path}")
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm"):
        try:
            wb = openpyxl.load_workbook(path, keep_vba=(ext == ".xlsm"))
        except Exception as e:  # noqa: BLE001
            if "encrypt" in str(e).lower() or "zip" in str(e).lower():
                raise SystemExit("암호가 걸렸거나 깨진 엑셀 파일이라 열 수 없습니다. 암호를 푼 파일을 올려 주세요.")
            raise SystemExit(f"엑셀 파일을 열 수 없습니다: {e}")
        return wb, ext
    if ext in (".csv", ".txt", ".tsv"):
        text, enc = _read_csv_text(path)
        sample = text[:4096]
        delim = "\t" if (ext == ".tsv" or sample.count("\t") > sample.count(",")) else ","
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        import io
        for row in csv.reader(io.StringIO(text, newline=""), delimiter=delim):
            ws.append(row)
        wb._kx_csv_encoding = enc
        return wb, ".csv"
    if ext == ".xls":
        raise SystemExit("옛날 엑셀(.xls) 파일입니다. 엑셀에서 '다른 이름으로 저장 → Excel 통합 문서(.xlsx)'로 바꾼 파일을 올려 주세요.")
    raise SystemExit(f"엑셀(.xlsx, .xlsm)이나 CSV 파일이 아닙니다: {path}")


def save(wb, out, kind_hint=".xlsx", encoding="utf-8-sig", sheet=None):
    ext = os.path.splitext(out)[1].lower() or kind_hint
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    if ext in (".csv", ".txt"):
        ws = wb[sheet] if sheet else wb.active
        with open(out, "w", newline="", encoding=encoding, errors="replace") as fh:
            w = csv.writer(fh)
            for row in ws.iter_rows(values_only=True):
                w.writerow([_csv_val(v) for v in row])
    else:
        wb.save(out)
    return out


def _csv_val(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d") if v.time() == dt.time(0) else v.strftime("%Y-%m-%d %H:%M")
    if isinstance(v, dt.date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def default_out(path, tag, ext=None):
    base, e = os.path.splitext(os.path.basename(path))
    e = ext or (".xlsx" if e.lower() in (".csv", ".txt", ".tsv") else e)
    return os.path.join(os.getcwd(), f"{base}_{tag}{e}")


def pick_sheet(wb, name):
    if not name:
        return wb.active
    if name in wb.sheetnames:
        return wb[name]
    if name.isdigit() and 1 <= int(name) <= len(wb.sheetnames):
        return wb[wb.sheetnames[int(name) - 1]]
    raise SystemExit(f"'{name}' 시트가 없습니다. 있는 시트: {', '.join(wb.sheetnames)}")


def _hint_hits(vals):
    n = 0
    for v in vals:
        if isinstance(v, str) and len(v.strip()) <= 20 and not v.startswith("="):
            h = v.lower().replace(" ", "")
            if any(k in h for _, keys in HINTS for k in keys) or h in ("번호", "no", "순번", "부서", "비고", "구분", "직급"):
                n += 1
    return n


def find_header(ws, given=None):
    """칸 이름 줄 번호(1부터). 위 15줄 안에서 칸 이름처럼 보이는 낱말(성명·연락처·주민번호…)이 가장 많은 줄,
    없으면 글자 칸이 둘 이상인 첫 줄."""
    if given:
        return int(given)
    scored = []
    for r in range(1, min(ws.max_row, 15) + 1):
        vals = [c.value for c in ws[r]]
        hits = _hint_hits(vals)
        if hits >= 2:
            scored.append((hits, -r))
    if scored:
        return -max(scored)[1]
    best = None
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = [c.value for c in ws[r]]
        filled = [v for v in vals if v not in (None, "")]
        texts = [v for v in filled if isinstance(v, str) and not v.startswith("=")]
        if len(filled) >= 2 and len(texts) >= max(2, int(len(filled) * 0.6)):
            return r
        if best is None and filled:
            best = r
    return best or 1


def headers(ws, hrow):
    out = {}
    for c in ws[hrow]:
        v = c.value
        if v not in (None, ""):
            name = str(v).strip().replace("\n", " ")
            k, n = name, 2
            while k in out:
                k, n = f"{name}_{n}", n + 1
            out[k] = c.column
    return out


def col_of(ws, hrow, name):
    """칸 이름 또는 열 글자(A, B…)를 열 번호로."""
    from openpyxl.utils import column_index_from_string
    hs = headers(ws, hrow)
    if name in hs:
        return hs[name]
    for k, v in hs.items():
        if k.replace(" ", "") == name.replace(" ", ""):
            return v
    if re.fullmatch(r"[A-Za-z]{1,3}", name):
        return column_index_from_string(name.upper())
    raise SystemExit(f"'{name}' 칸이 없습니다. 있는 칸: {', '.join(hs) or '(칸 이름 줄을 못 찾음 — --header 로 알려 주세요)'}")


def data_rows(ws, hrow):
    for r in range(hrow + 1, ws.max_row + 1):
        if any(c.value not in (None, "") for c in ws[r]):
            yield r


def writable(cell):
    from openpyxl.cell.cell import MergedCell
    if isinstance(cell, MergedCell):
        return False
    v = cell.value
    return not (isinstance(v, str) and v.startswith("="))


def widen(ws, cols, rows):
    """바뀐 칸이 ### 로 안 보이지 않게 열 너비를 넓힌다(좁히지는 않음)."""
    from openpyxl.utils import get_column_letter
    for c in cols:
        need = 8
        for r in list(rows)[:500]:
            v = ws.cell(r, c).value
            if isinstance(v, dt.datetime):
                n = 17 if v.time() != dt.time(0) else 11
            elif isinstance(v, (int, float)):
                n = len(f"{v:,.0f}") + 2
            else:
                n = sum(2 if ord(ch) > 0x2E80 else 1 for ch in _csv_val(v)) + 1
            need = max(need, n)
        letter = get_column_letter(c)
        cur = ws.column_dimensions[letter].width or 8.43
        if cur < min(need, 45):
            ws.column_dimensions[letter].width = min(need, 45)


def add_column(ws, hrow, title, after=None):
    """맨 끝(또는 after 열 바로 뒤가 아니라 맨 끝)에 새 칸. 머리 칸 모양은 옆 칸을 따른다."""
    from copy import copy
    col = ws.max_column + 1
    hs = headers(ws, hrow)
    t, n = title, 2
    while t in hs:
        t, n = f"{title}_{n}", n + 1
    src = ws.cell(hrow, col - 1)
    dst = ws.cell(hrow, col, t)
    if src.has_style:
        dst.font, dst.fill, dst.border, dst.alignment = copy(src.font), copy(src.fill), copy(src.border), copy(src.alignment)
    from openpyxl.utils import get_column_letter
    ws.column_dimensions[get_column_letter(col)].width = max(10, min(40, len(t) * 2 + 4))
    return col


# ───────────────────────── 한국 번호·값 알아보기 ─────────────────────────

def digits(s):
    return re.sub(r"\D", "", str(s))


def norm_text(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    return s.replace(" ", " ").strip()


RRN_RE = re.compile(r"^(\d{6})\s*-?\s*([1-8]\d{6})$")
RRN_IN = re.compile(r"(?<!\d)(\d{2}(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01]))\s*-?\s*([1-8])(\d{6})(?!\d)")
MOBILE_RE = re.compile(r"^(01[016789])\s*[-.)]?\s*(\d{3,4})\s*[-.]?\s*(\d{4})$")
PHONE_RE = re.compile(r"^(02|0[3-6][1-5]|070|050\d?|080)\s*[-.)]?\s*(\d{3,4})\s*[-.]?\s*(\d{4})$")
REP_RE = re.compile(r"^(1[5-9]\d{2})\s*-?\s*(\d{4})$")
PHONE_IN = re.compile(r"(?<!\d)(01[016789]|02|0[3-6][1-5]|070)[-. ](\d{3,4})[-. ](\d{4})(?!\d)|(?<!\d)(01[016789])(\d{3,4})(\d{4})(?!\d)")
BIZ_RE = re.compile(r"^(\d{3})\s*-?\s*(\d{2})\s*-?\s*(\d{5})$")
EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$")
EMAIL_IN = re.compile(r"([\w.+-]{1,64})@([\w-]+(?:\.[\w-]+)+)")
PHONE_IN2 = re.compile(r"(?<![\d+])(\+82[-. ]?1[016789]|\+82[-. ]?2|050\d?|1[568]\d{2})[-. ](\d{3,4})(?:[-. ](\d{4}))?(?!\d)")
BANK_ACCT = re.compile(r"((?:은행|뱅크|계좌|국민|신한|우리|하나|농협|기업|카카오|토스|케이|새마을|우체국|수협|신협|SC|씨티|부산|대구|경남|광주|전북|제주|산업)\D{0,6}?)(?<!\d)(\d{10,16})(?!\d)")
LABEL_NAME = re.compile(r"((?:작성자|담당자|성명|이름|성함|대표자|신청인|보호자|예금주|수령인|받는 ?분)\s*[:：]\s*)([가-힣]{2,4})(?![가-힣●])")
ACCT_IN = re.compile(r"(?<![\d-])(\d{2,6}-\d{2,6}-\d{2,8}(?:-\d{1,4})?)(?![\d-])")

SIDO = [
    ("서울특별시", ["서울특별시", "서울시", "서울"]),
    ("부산광역시", ["부산광역시", "부산시", "부산"]),
    ("대구광역시", ["대구광역시", "대구시", "대구"]),
    ("인천광역시", ["인천광역시", "인천시", "인천"]),
    ("광주광역시", ["광주광역시"]),
    ("대전광역시", ["대전광역시", "대전시", "대전"]),
    ("울산광역시", ["울산광역시", "울산시", "울산"]),
    ("세종특별자치시", ["세종특별자치시", "세종시", "세종"]),
    ("경기도", ["경기도", "경기"]),
    ("강원특별자치도", ["강원특별자치도", "강원도", "강원"]),
    ("충청북도", ["충청북도", "충북"]),
    ("충청남도", ["충청남도", "충남"]),
    ("전북특별자치도", ["전북특별자치도", "전라북도", "전북"]),
    ("전라남도", ["전라남도", "전남"]),
    ("경상북도", ["경상북도", "경북"]),
    ("경상남도", ["경상남도", "경남"]),
    ("제주특별자치도", ["제주특별자치도", "제주도", "제주"]),
]
SIDO_ALIAS = {a: full for full, al in SIDO for a in al}
GWANGJU_GU = {"동구", "서구", "남구", "북구", "광산구"}


def rrn_info(s):
    """주민(외국인)등록번호 → dict(ok, birth, sex, why). 형식이 아니면 None."""
    m = RRN_RE.match(norm_text(s).replace(" ", ""))
    if not m:
        return None
    a, b = m.group(1), m.group(2)
    g = int(b[0])
    century = {1: 1900, 2: 1900, 5: 1900, 6: 1900, 3: 2000, 4: 2000, 7: 2000, 8: 2000}[g]
    try:
        birth = dt.date(century + int(a[:2]), int(a[2:4]), int(a[4:6]))
    except ValueError:
        return {"ok": False, "why": "앞 6자리가 없는 날짜", "birth": None, "sex": None}
    sex = "남" if g % 2 == 1 else "여"
    foreign = g >= 5
    ds = [int(x) for x in a + b]
    w = [2, 3, 4, 5, 6, 7, 8, 9, 2, 3, 4, 5]
    chk = (11 - sum(d * k for d, k in zip(ds[:12], w)) % 11) % 10
    if foreign:
        chk = (chk + 2) % 10
    info = {"ok": True, "why": "", "birth": birth, "sex": sex, "foreign": foreign}
    if birth > TODAY:
        info.update(ok=False, why="생일이 오늘보다 뒤")
    elif chk != ds[12]:
        # 2020-10 이후 새로 받은 번호는 검사 숫자가 없다
        info.update(ok=True, why="검사 숫자 불일치(2020년 10월 이후 발급 번호면 정상)", soft=True)
    return info


def biz_ok(s):
    d = digits(s)
    if len(d) != 10:
        return False
    w = [1, 3, 7, 1, 3, 7, 1, 3, 5]
    n = [int(x) for x in d]
    t = sum(a * b for a, b in zip(n[:9], w)) + (n[8] * 5) // 10
    return (10 - t % 10) % 10 == n[9]


def phone_fmt(v):
    """전화번호를 010-1234-5678 모양으로. 못 알아보면 None."""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        v = str(int(v))
        if re.fullmatch(r"1[016789]\d{7,8}", v):  # 엑셀이 앞 0을 먹은 휴대폰
            v = "0" + v
        elif re.fullmatch(r"[2-6]\d{7,9}", v):
            v = "0" + v
    s = norm_text(v)
    if not re.fullmatch(r"[\d\-.\s()+]+", s):  # 숫자·띄어쓰기·하이픈 말고 다른 글자가 섞이면 전화번호 칸 값이 아님
        return None
    s = s.replace(" ", "")
    s = re.sub(r"^\+?82-?\(?0?\)?", "0", s)
    d0 = digits(s)
    if d0.startswith("050") and len(d0) in (11, 12):
        return f"050-{d0[3:7]}-{d0[7:]}" if len(d0) == 11 else f"{d0[:4]}-{d0[4:8]}-{d0[8:]}"
    s = s.replace("(", "").replace(")", "-") if s.startswith("(") else s
    s = s.replace("0-", "0", 1) if s.startswith("00") else s
    for rx in (MOBILE_RE, PHONE_RE, REP_RE):
        m = rx.match(s)
        if m:
            return "-".join(m.groups())
    d = digits(s)
    if re.fullmatch(r"01[016789]\d{7,8}", d):
        return f"{d[:3]}-{d[3:-4]}-{d[-4:]}"
    if re.fullmatch(r"02\d{7,8}", d):
        return f"02-{d[2:-4]}-{d[-4:]}"
    if re.fullmatch(r"0[3-6][1-5]\d{7,8}|070\d{8}", d):
        return f"{d[:3]}-{d[3:-4]}-{d[-4:]}"
    return None


def biz_fmt(v):
    d = digits(v)
    return f"{d[:3]}-{d[3:5]}-{d[5:]}" if len(d) == 10 else None


DATE_PATTERNS = [
    re.compile(r"^(\d{4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\.?(?:\s*\(?[월화수목금토일]\)?)?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$"),
    re.compile(r"^(\d{4})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일(?:\s*\(?[월화수목금토일]요?일?\)?)?(?:\s*(\d{1,2})\s*[:시]\s*(\d{1,2})분?)?$"),
    re.compile(r"^(\d{4})(\d{2})(\d{2})$"),
    re.compile(r"^(\d{2})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})\.?$"),
    re.compile(r"^(\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일$"),
]
DATE6 = re.compile(r"^(\d{2})(\d{2})(\d{2})$")  # 900101 (생년월일 6자리) — 날짜 칸에서만


def parse_date(v, serial_ok=False, six=False):
    """여러 한국식 날짜 모양 → date/datetime. 못 알아보면 None."""
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        if serial_ok and 20000 <= v <= 80000:
            return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date() if float(v).is_integer() \
                else dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))
        if 19000101 <= v <= 21001231 and float(v).is_integer():
            v = str(int(v))
        else:
            return None
    s = norm_text(v)
    s = re.sub(r"\s*(오전|오후|AM|PM|am|pm)\s*(\d{1,2}):(\d{2})", lambda m: " %d:%s" % (
        (int(m.group(2)) % 12) + (12 if m.group(1) in ("오후", "PM", "pm") else 0), m.group(3)), s)
    pats = DATE_PATTERNS + ([DATE6] if six else [])
    for i, rx in enumerate(pats):
        m = rx.match(s)
        if not m:
            continue
        g = [x for x in m.groups()]
        y = int(g[0])
        if i >= 3:
            y += 2000 if y <= TODAY.year % 100 + 1 else 1900
            if rx is DATE6:
                try:
                    if dt.date(y, int(g[1]), int(g[2])) > TODAY:
                        y -= 100
                except ValueError:
                    return None
        try:
            d = dt.date(y, int(g[1]), int(g[2]))
        except ValueError:
            return None
        if len(g) > 3 and g[3]:
            try:
                return dt.datetime(d.year, d.month, d.day, int(g[3]), int(g[4] or 0), int((g[5] if len(g) > 5 else 0) or 0))
            except ValueError:
                return d
        return d
    return None


HAN_DIGIT = {"영": 0, "일": 1, "이": 2, "삼": 3, "사": 4, "오": 5, "육": 6, "칠": 7, "팔": 8, "구": 9}
UNIT_SMALL = {"십": 10, "백": 100, "천": 1000}
UNIT_BIG = {"만": 10 ** 4, "억": 10 ** 8, "조": 10 ** 12}


def parse_money(v):
    """'1,234원', '₩1,234', '123만원', '1억 2,000만원', '(1,000)', '일금 일백만원정' → 숫자. 못 알아보면 None."""
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    s = norm_text(v)
    if not s:
        return None
    neg = False
    if re.fullmatch(r"\(.*\)", s) or s.startswith("-") or s.startswith("△") or s.startswith("▲"):
        neg = True
        s = s.strip("()-△▲ ")
    s = re.sub(r"^(일금|금)\s*", "", s)
    s = re.sub(r"(원정|원|₩|\\|KRW|정)", "", s).replace(",", "").replace(" ", "")
    if not s:
        return None
    if re.fullmatch(r"\d+(\.\d+)?", s):
        n = float(s)
    elif re.fullmatch(r"[\d.천백십만억조영일이삼사오육칠팔구]+", s):
        n = _han_number(s)
        if n is None:
            return None
    else:
        return None
    n = -n if neg else n
    return int(n) if float(n).is_integer() else n


def _han_number(s):
    total, section, num = 0, 0, None
    i = 0
    try:
        while i < len(s):
            ch = s[i]
            if ch.isdigit() or ch == ".":
                j = i
                while j < len(s) and (s[j].isdigit() or s[j] == "."):
                    j += 1
                num = float(s[i:j])
                i = j
                continue
            if ch in HAN_DIGIT:
                num = HAN_DIGIT[ch]
            elif ch in UNIT_SMALL:
                section += (num if num is not None else 1) * UNIT_SMALL[ch]
                num = None
            elif ch in UNIT_BIG:
                section += num or 0
                total += (section or 1) * UNIT_BIG[ch]
                section, num = 0, None
            else:
                return None
            i += 1
        return total + section + (num or 0)
    except Exception:  # noqa: BLE001
        return None


def han_money(n, plain=False):
    """1234500 → '일금 일백이십삼만사천오백원정'. plain=True 면 '금 백이십삼만사천오백 원'."""
    n = int(round(n))
    if n == 0:
        return "일금 영원정" if not plain else "영 원"
    neg = n < 0
    n = abs(n)
    nums = "영일이삼사오육칠팔구"
    small = ["", "십", "백", "천"]
    big = ["", "만", "억", "조", "경"]
    parts = []
    k = 0
    while n > 0:
        chunk = n % 10000
        n //= 10000
        if chunk:
            s = ""
            for p in range(3, -1, -1):
                d = (chunk // 10 ** p) % 10
                if d:
                    s += ("" if (plain and d == 1 and p > 0) else nums[d]) + small[p]
            parts.append(s + big[k])
        k += 1
    word = "".join(reversed(parts))
    if neg:
        word = "마이너스 " + word
    return f"{word} 원" if plain else f"일금 {word}원정"


HINTS = [
    ("biz", ("사업자",)),
    ("rrn", ("주민", "외국인등록")),
    ("phone", ("전화", "연락처", "휴대", "핸드폰", "휴대폰", "핸드폰번호", "tel", "phone", "hp", "fax", "팩스")),
    ("email", ("이메일", "메일", "e-mail", "email")),
    ("account", ("계좌",)),
    ("addr", ("주소", "소재지", "거주지", "address")),
    ("date", ("일자", "날짜", "생년월일", "생일", "입사일", "퇴사일", "시작일", "종료일", "등록일", "접수일", "date", "기간")),
    ("money", ("금액", "급여", "가격", "단가", "합계", "비용", "매출", "공급가", "세액", "부가세", "원)", "(원", "price", "amount", "총액", "지급액", "수당", "월급", "연봉")),
    ("name", ("성명", "이름", "성함", "대표자", "담당자", "name", "예금주", "보호자")),
    ("text", ("번호", "코드", "code", "id", "no.", "주문", "품번", "사번", "학번", "우편")),
]

NOT_NAMES = {"메모", "비고", "합계", "소계", "총계", "없음", "미정", "공란", "해당없음", "기타", "참고", "확인", "담당", "본인", "대표", "미상"}
KIND_KO = {"rrn": "주민등록번호", "biz": "사업자등록번호", "phone": "전화번호", "email": "이메일", "account": "계좌번호",
           "addr": "주소", "date": "날짜", "money": "금액", "name": "사람 이름", "text": "글", "number": "숫자", "empty": "빈 칸"}


def guess_kind(header, values):
    h = (header or "").lower().replace(" ", "")
    for kind, keys in HINTS:
        if any(k in h for k in keys):
            if kind == "date" and "기간" in h:
                continue
            if kind == "account" and "은행" in h and "번호" not in h:
                continue
            return kind
    vals = [v for v in values if v not in (None, "")]
    if not vals:
        return "empty"
    score = {k: 0 for k in ("rrn", "phone", "biz", "email", "date", "money", "addr")}
    for v in vals:
        s = norm_text(v)
        if rrn_info(s):
            score["rrn"] += 1
        elif isinstance(v, str) and EMAIL_RE.match(s):
            score["email"] += 1
        elif isinstance(v, str) and BIZ_RE.match(s) and "-" in s:
            score["biz"] += 1
        elif isinstance(v, str) and phone_fmt(v) and re.search(r"[-.\s)]", s):
            score["phone"] += 1
        elif parse_date(v) is not None and not (isinstance(v, (int, float)) or re.fullmatch(r"\d{1,7}", s)):
            score["date"] += 1
        elif isinstance(v, str) and (s.split(" ")[0] in SIDO_ALIAS or s.split(" ")[0] == "광주광역시"):
            score["addr"] += 1
        elif isinstance(v, str) and re.search(r"원|₩|만|억", s) and parse_money(s) is not None:
            score["money"] += 1
    kind, n = max(score.items(), key=lambda kv: kv[1])
    if n >= max(1, int(len(vals) * 0.6)):
        return kind
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
        return "number"
    return "text"


def col_kinds(ws, hrow, only=None):
    hs = headers(ws, hrow)
    rows = list(data_rows(ws, hrow))
    out = []
    for name, c in hs.items():
        if only and name not in only:
            continue
        vals = [ws.cell(r, c).value for r in rows[:300]]
        out.append((name, c, guess_kind(name, vals)))
    return out, rows


def pick_cols(ws, hrow, cols_arg, kinds_wanted):
    kinds, rows = col_kinds(ws, hrow)
    if cols_arg:
        want = [x.strip() for x in cols_arg.split(",") if x.strip()]
        cnums = {col_of(ws, hrow, w) for w in want}
        kinds = [k for k in kinds if k[1] in cnums]
    return [k for k in kinds if k[2] in kinds_wanted], rows


# ───────────────────────── 명령들 ─────────────────────────

def cmd_info(a):
    wb, _ = load(a.file)
    print(f"파일: {os.path.basename(a.file)}  |  시트 {len(wb.sheetnames)}개: {', '.join(wb.sheetnames)}")
    sheets = [pick_sheet(wb, a.sheet)] if a.sheet else wb.worksheets
    for ws in sheets:
        if ws.max_row <= 1 and ws.max_column <= 1 and ws.cell(1, 1).value is None:
            print(f"\n[{ws.title}] 빈 시트")
            continue
        hrow = find_header(ws, a.header)
        kinds, rows = col_kinds(ws, hrow)
        print(f"\n[{ws.title}] 칸 이름 줄: {hrow}번째 줄 | 자료 {len(rows)}줄 × 칸 {len(kinds)}개"
              + (f" | 합친 칸 {len(ws.merged_cells.ranges)}곳" if ws.merged_cells.ranges else ""))
        formulas = sum(1 for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
        if formulas:
            print(f"  계산식 칸 {formulas}개 (계산식 칸은 고치지 않습니다)")
        print("  | 칸 | 들어 있는 것 | 채워진 줄 | 보기 |")
        print("  |---|---|---|---|")
        for name, c, kind in kinds:
            vals = [ws.cell(r, c).value for r in rows]
            filled = [v for v in vals if v not in (None, "")]
            ex = next((v for v in filled), "")
            ex = _peek(ex, kind)
            print(f"  | {name} | {KIND_KO.get(kind, kind)} | {len(filled)}/{len(rows)} | {ex} |")
    if any(k for ws in sheets for k in []):
        pass


def _peek(v, kind):
    s = _csv_val(v)
    if kind in ("rrn", "phone", "account", "email", "name"):
        m = mask_value(s, kind)
        if m is None:
            t, n = mask_text(s)
            m = t if n else (s[:3] + "●" * min(8, max(0, len(s) - 3)))
        s = m
    else:
        s = mask_text(s)[0]
    s = s.replace("|", "/").replace("\n", " ")
    return s[:30] + ("…" if len(s) > 30 else "")


def _fill(color):
    from openpyxl.styles import PatternFill
    return PatternFill("solid", start_color=color, end_color=color)


def cmd_check(a):
    from openpyxl.comments import Comment
    wb, kind0 = load(a.file)
    ws = pick_sheet(wb, a.sheet)
    hrow = find_header(ws, a.header)
    targets, rows = pick_cols(ws, hrow, a.cols, {"rrn", "biz", "phone", "email", "date", "money"})
    if not targets:
        print("검사할 칸(주민번호·사업자번호·전화·이메일·날짜·금액)을 못 찾았습니다. --cols 칸이름 으로 알려 주세요.")
        return
    bad_fill, warn_fill = _fill("FFFF00"), _fill("FFE5B4")
    total_bad = total_warn = 0
    seen = {}
    print(f"[{ws.title}] 자료 {len(rows)}줄 검사")
    for name, c, kind in targets:
        bad, warn = [], []
        for r in rows:
            cell = ws.cell(r, c)
            v = cell.value
            if v in (None, ""):
                continue
            if isinstance(v, str) and v.startswith("="):
                continue
            problem, soft = _check_one(kind, v)
            if kind in ("rrn", "email", "phone") and not problem:
                key = (name, digits(v) if kind != "email" else norm_text(v).lower())
                if key in seen:
                    problem, soft = f"{seen[key]}번째 줄과 같은 값(같은 사람이 두 번?)", True
                else:
                    seen[key] = r
            if problem:
                (warn if soft else bad).append((r, problem))
                if writable(cell) or True:
                    try:
                        cell.fill = warn_fill if soft else bad_fill
                        note = ("확인 필요: " if soft else "틀림: ") + problem
                        old = cell.comment.text if cell.comment is not None else ""
                        cell.comment = Comment((old + "\n" + note) if old else note, "kexcel")
                    except Exception:  # noqa: BLE001
                        pass
        total_bad += len(bad)
        total_warn += len(warn)
        print(f"\n■ {name} ({KIND_KO[kind]}) — 틀림 {len(bad)}, 확인 필요 {len(warn)}")
        for r, p in (bad + warn)[: a.show]:
            print(f"  {r}번째 줄: {_peek(ws.cell(r, c).value, kind)} → {p}")
        if len(bad) + len(warn) > a.show:
            print(f"  … 그 밖에 {len(bad) + len(warn) - a.show}곳")
    out = a.out or default_out(a.file, "검사")
    if os.path.splitext(out)[1].lower() == ".csv":
        out = os.path.splitext(out)[0] + ".xlsx"
    save(wb, out)
    print(f"\n합계: 틀림 {total_bad}곳(노란색), 확인 필요 {total_warn}곳(살구색) → {out}")


def _check_one(kind, v):
    """(문제 설명 또는 '', 가벼운지)"""
    s = norm_text(v)
    if kind == "rrn":
        if "*" in s or "●" in s:
            return "", False
        info = rrn_info(s)
        if not info:
            n = len(digits(s))
            return (f"자리 수가 {n}자리(13자리여야 함)" if n != 13 else "주민번호 모양이 아님"), False
        if not info["ok"]:
            return info["why"], False
        if info.get("soft"):
            return info["why"], True
        return "", False
    if kind == "biz":
        d = digits(s)
        if len(d) != 10:
            return f"자리 수가 {len(d)}자리(10자리여야 함)", False
        if not biz_ok(d):
            return "검사 숫자가 안 맞음(번호 오타일 가능성 큼)", False
        return "", False
    if kind == "phone":
        if phone_fmt(v):
            if isinstance(v, (int, float)):
                if phone_fmt(v).startswith("0"):
                    return "숫자로 저장돼 맨 앞 0이 빠짐(정리하면 고쳐짐)", True
                return "숫자로 저장됨(정리하면 글자 모양 번호가 됨)", True
            return "", False
        return "전화번호 모양이 아님", False
    if kind == "email":
        if EMAIL_RE.match(s.replace(" ", "")):
            if " " in s:
                return "중간에 띄어쓰기", True
            return "", False
        return "이메일 모양이 아님", False
    if kind == "date":
        if isinstance(v, (int, float)) and not isinstance(v, bool) and 20000 <= v <= 80000:
            return "엑셀 날짜 숫자로 보임(정리하면 날짜 모양이 됨)", True
        d = parse_date(v, six=True)
        if d is None:
            return "날짜로 못 읽음", False
        if not isinstance(v, (dt.date, dt.datetime)):
            return "글자로 적힌 날짜(정리하면 진짜 날짜가 됨)", True
        dd = d.date() if isinstance(d, dt.datetime) else d
        if dd.year < 1900 or dd.year > TODAY.year + 30:
            return f"{dd.year}년 — 너무 옛날이거나 너무 먼 날", True
        return "", False
    if kind == "money":
        if isinstance(v, (int, float)):
            return "", False
        if parse_money(v) is None:
            return "금액으로 못 읽음", False
        return "글자로 적힌 금액(정리하면 계산되는 숫자가 됨)", True
    return "", False


def cmd_clean(a):
    wb, kind0 = load(a.file)
    ws = pick_sheet(wb, a.sheet)
    hrow = find_header(ws, a.header)
    what = set(a.what.split(",")) if a.what != "all" else {"date", "phone", "biz", "money", "space"}
    kinds_map = {"date": "date", "phone": "phone", "biz": "biz", "money": "money"}
    want = {kinds_map[w] for w in what if w in kinds_map}
    targets, rows = pick_cols(ws, hrow, a.cols, want | ({"text", "name", "addr", "email", "rrn", "account", "number"} if "space" in what else set()))
    changes, fails = [], []
    for name, c, kind in targets:
        for r in rows:
            cell = ws.cell(r, c)
            v = cell.value
            if v in (None, "") or not writable(cell):
                continue
            new, fmt = v, None
            if kind == "date" and "date" in what:
                d = parse_date(v, serial_ok=True, six=True)
                if d is None:
                    fails.append((r, name, v, "날짜로 못 읽음"))
                    continue
                new = d if isinstance(d, dt.datetime) else dt.datetime(d.year, d.month, d.day)
                fmt = "yyyy-mm-dd hh:mm" if isinstance(d, dt.datetime) and d.time() != dt.time(0) else "yyyy-mm-dd"
            elif kind == "phone" and "phone" in what:
                p = phone_fmt(v)
                if not p:
                    fails.append((r, name, v, "전화번호 모양이 아님"))
                    continue
                new, fmt = p, "@"
            elif kind == "biz" and "biz" in what:
                p = biz_fmt(v)
                if not p:
                    fails.append((r, name, v, "10자리가 아님"))
                    continue
                new, fmt = p, "@"
            elif kind == "money" and "money" in what:
                m = parse_money(v)
                if m is None:
                    fails.append((r, name, v, "금액으로 못 읽음"))
                    continue
                new, fmt = m, "#,##0" if isinstance(m, int) else "#,##0.##"
            elif "space" in what and isinstance(v, str):
                new = re.sub(r"[ \t]+", " ", unicodedata.normalize("NFKC", v).replace(" ", " ")).strip()
                new = re.sub(r"\s*\n\s*", "\n", new)
            if new != v or (fmt and cell.number_format != fmt and not isinstance(v, str)):
                if new != v:
                    changes.append((r, name, v, new))
                if not a.dry:
                    cell.value = new
                    if fmt:
                        cell.number_format = fmt
    print(f"[{ws.title}] 바꿀 칸 {len(changes)}곳, 못 바꾼 칸 {len(fails)}곳")
    kind_of_name = {n: k for n, _, k in targets}
    for r, n, old, new in changes[: a.show]:
        k = kind_of_name.get(n, "")
        print(f"  {r}번째 줄 {n}: {_peek(old, k)} → {_peek(new, k)}")
    if len(changes) > a.show:
        print(f"  … 그 밖에 {len(changes) - a.show}곳")
    for r, n, old, why in fails[: a.show]:
        print(f"  ✗ {r}번째 줄 {n}: {_peek(old, kind_of_name.get(n, ''))} ({why})")
    if a.dry:
        print("(미리 보기만 했습니다. 바꾸려면 --dry 를 빼고 다시)")
        return
    widen(ws, {ws.cell(hrow, col_of(ws, hrow, n)).column for _, n, _, _ in changes}, rows)
    out = a.out or default_out(a.file, "정리")
    save(wb, out, encoding=getattr(wb, "_kx_csv_encoding", "utf-8-sig"))
    print(f"저장: {out}")


def mask_value(s, kind):
    s = str(s)
    if kind == "rrn":
        info = RRN_RE.match(s.replace(" ", ""))
        if info:
            return f"{info.group(1)}-{info.group(2)[0]}●●●●●●"
        d = digits(s)
        if len(d) >= 7:  # 자리 수가 틀린 번호도 가린다
            return f"{d[:6]}-{d[6]}" + "●" * max(1, len(d) - 7)
        return None
    if kind == "phone":
        p = phone_fmt(s)
        if p and p.count("-") >= 1:
            a, b, c = p.split("-") if p.count("-") == 2 else (p.split("-")[0], "", p.split("-")[-1])
            return f"{a}-{'●' * len(b)}-{c}" if b else f"{a}-●●●●"
        return None
    if kind == "email":
        m = EMAIL_RE.match(s.strip())
        if m:
            user, dom = s.strip().split("@", 1)
            keep = user[:2] if len(user) > 3 else user[:1]
            return keep + "●" * max(3, len(user) - len(keep)) + "@" + dom
        return None
    if kind == "account":
        d = re.sub(r"\d", "#", s)
        nums = digits(s)
        if len(nums) < 8:
            return None
        keep_head, keep_tail = 3, 3
        out, i = [], 0
        for ch in s:
            if ch.isdigit():
                out.append(ch if i < keep_head or i >= len(nums) - keep_tail else "●")
                i += 1
            else:
                out.append(ch)
        del d
        return "".join(out)
    if kind == "name":
        n = s.strip()
        if not n or len(n) > 12 or n in NOT_NAMES or TOTAL_RE.match(n):
            return None
        if len(n) == 1:
            return n
        if len(n) == 2:
            return n[0] + "●"
        if re.fullmatch(r"[가-힣]+", n):
            return n[0] + "●" * (len(n) - 2) + n[-1]
        parts = n.split()
        if len(parts) >= 2:
            return " ".join(p[0] + "●" * (len(p) - 1) for p in parts)
        return n[0] + "●" * (len(n) - 1)
    return None


def _rough_mask(s):
    d = digits(s)
    keep_head, keep_tail = 3, 2
    out, i = [], 0
    for ch in s:
        if ch.isdigit():
            out.append(ch if i < keep_head or i >= len(d) - keep_tail else "●")
            i += 1
        else:
            out.append(ch)
    return "".join(out)


def mask_text(s, words=()):
    """글 속에 섞인 개인정보 가리기 (가린 수 함께)."""
    n = 0

    def rr(m):
        nonlocal n
        n += 1
        return f"{m.group(1)}-{m.group(2)}●●●●●●"
    s2 = RRN_IN.sub(rr, s)

    def ph(m):
        nonlocal n
        n += 1
        if m.group(1):
            return f"{m.group(1)}-{'●' * len(m.group(2))}-{m.group(3)}"
        return f"{m.group(4)}-{'●' * len(m.group(5))}-{m.group(6)}"
    s2 = PHONE_IN.sub(ph, s2)

    def em(m):
        nonlocal n
        n += 1
        u = m.group(1)
        return (u[:2] if len(u) > 3 else u[:1]) + "●●●" + "@" + m.group(2)
    s2 = EMAIL_IN.sub(em, s2)

    def ac(m):
        nonlocal n
        v = m.group(1)
        if parse_date(v) or len(digits(v)) < 10 or (BIZ_RE.match(v) and biz_ok(v)):
            return v
        n += 1
        return mask_value(v, "account") or v

    def ph2(m):
        nonlocal n
        n += 1
        return f"{m.group(1)}-{'●' * len(m.group(2))}" + (f"-{m.group(3)}" if m.group(3) else "")
    s2 = PHONE_IN2.sub(ph2, s2)
    s2 = ACCT_IN.sub(ac, s2)

    def ba(m):
        nonlocal n
        n += 1
        d = m.group(2)
        return m.group(1) + d[:3] + "●" * (len(d) - 6) + d[-3:]
    s2 = BANK_ACCT.sub(ba, s2)
    def lb(m):
        nonlocal n
        n += 1
        return m.group(1) + (mask_value(m.group(2), "name") or m.group(2))
    s2 = LABEL_NAME.sub(lb, s2)
    for w in words:
        if w and w in s2:
            n += s2.count(w)
            s2 = s2.replace(w, mask_value(w, "name") or "●" * len(w))
    return s2, n


def cmd_mask(a):
    wb, kind0 = load(a.file)
    words = [w.strip() for w in (a.words or "").split(",") if w.strip()]
    sheets = [pick_sheet(wb, a.sheet)] if a.sheet else wb.worksheets
    total = 0
    report = []
    all_names = set()
    if not a.keep_names:
        for ws in wb.worksheets:  # 이름 칸의 이름은 다른 시트의 글 속에서도 가린다
            hr = find_header(ws, a.header if (a.sheet and ws.title == pick_sheet(wb, a.sheet).title) else None)
            ks, rs = col_kinds(ws, hr)
            for _, c, k in ks:
                if k == "name":
                    for r in rs:
                        v = ws.cell(r, c).value
                        if isinstance(v, str) and re.fullmatch(r"[가-힣]{2,4}", v.strip()) and v.strip() not in NOT_NAMES \
                                and not TOTAL_RE.match(v):
                            all_names.add(v.strip())
    for ws in sheets:
        hrow = find_header(ws, a.header if a.sheet else None)
        kinds, rows = col_kinds(ws, hrow)
        want = {"rrn", "phone", "email", "account"} | ({"name"} if not a.keep_names else set())
        if a.cols:
            cn = {col_of(ws, hrow, x.strip()) for x in a.cols.split(",")}
            kinds = [k for k in kinds if k[1] in cn]
            kinds = [(n, c, k if k in want else "text") for n, c, k in kinds]
        kind_by_col = {c: k for _, c, k in kinds if k in want}
        kind_all = {c: k for _, c, k in kinds}
        names_seen = all_names
        cnt = {}
        allwords = words + sorted(names_seen, key=len, reverse=True)
        for row in ws.iter_rows(min_row=1):  # 칸 이름 줄 위(제목·작성자 줄)도 본다
            for cell in row:
                if cell.comment is not None and not a.dry:  # 셀 메모 속 개인정보
                    t, m = mask_text(cell.comment.text or "", allwords)
                    if m:
                        from openpyxl.comments import Comment
                        cell.comment = Comment(t, cell.comment.author or "")
                        cnt["셀 메모"] = cnt.get("셀 메모", 0) + 1
                v = cell.value
                if v in (None, "") or not writable(cell):
                    continue
                above = cell.row <= hrow
                k = None if above else kind_by_col.get(cell.column)
                new = None
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    sv = str(int(v)) if float(v).is_integer() else str(v)
                    if len(sv) == 13 and rrn_info(sv):  # 숫자로 저장된 주민번호
                        new, k = mask_value(sv, "rrn"), k or "rrn"
                    elif k == "phone" and phone_fmt(v):
                        new = mask_value(phone_fmt(v), "phone")
                    elif k == "rrn":
                        new = mask_value(sv, "rrn")
                    elif k == "account":
                        new = mask_value(sv, "account")
                elif k:
                    new = mask_value(v, k)
                    if new is None and isinstance(v, str):
                        t, m = mask_text(v, allwords)
                        new = t if m else None
                        if new is None and k in ("rrn", "phone", "account") and len(digits(v)) >= 6:
                            new = _rough_mask(v)  # 모양이 틀린 번호도 가운데를 가린다
                        elif new is None and k == "email" and len(v.strip()) > 3:
                            u, _, rest = v.strip().partition("@")
                            new = u[:2] + "●" * max(3, len(u) - 2) + ("@" + rest if rest else "")
                elif isinstance(v, str) and (above or kind_all.get(cell.column) not in ("biz", "date", "money")):
                    t, m = mask_text(v, allwords)
                    new = t if m else None
                if new is not None and new != v:
                    key = KIND_KO.get(k, "글 속 개인정보")
                    cnt[key] = cnt.get(key, 0) + 1
                    if len(report) < a.show:
                        report.append(f"  [{ws.title}] {cell.coordinate}: {_peek(new, '')}")
                    if not a.dry:
                        cell.value = new
                        if k in ("phone", "rrn"):
                            cell.number_format = "@"
        if cnt:
            print(f"[{ws.title}] " + ", ".join(f"{k} {v}곳" for k, v in cnt.items()))
            total += sum(cnt.values())
    if not total:
        print("가릴 개인정보를 못 찾았습니다. 이름 칸이 따로 없으면 --words 홍길동,김철수 로 알려 주세요.")
        return
    print("가린 모양 보기:")
    print("\n".join(report))
    if a.dry:
        print(f"(미리 보기: 모두 {total}곳. 가리려면 --dry 를 빼고 다시)")
        return
    out = a.out or default_out(a.file, "가림")
    save(wb, out, encoding=getattr(wb, "_kx_csv_encoding", "utf-8-sig"))
    print(f"모두 {total}곳 가림(숨긴 시트·셀 메모 포함) → {out}\n※ 그림 속 글자, 시트 이름, 계산식 칸, 파일 정보(작성자)는 가리지 못했습니다.")


def split_addr(s):
    s0 = norm_text(s)
    zipc = ""
    m = re.match(r"^[\(\[]?(\d{5})[\)\]]?\s+(.*)$", s0)
    if m:
        zipc, s0 = m.group(1), m.group(2)
    s0 = re.sub(r"\s+", " ", s0)
    note = ""
    m = re.search(r"\s*(\([^()]*\))\s*$", s0)
    if m:
        note, s0 = m.group(1), s0[: m.start()]
    toks = s0.split(" ")
    if not toks or not toks[0]:
        return None
    sido = SIDO_ALIAS.get(toks[0])
    i = 1
    if toks[0] == "광주" or toks[0] == "광주시":
        if len(toks) > 1 and toks[1] in GWANGJU_GU:
            sido = "광주광역시"
        elif toks[0] == "광주시":
            sido, i = "경기도", 0
        else:
            sido = "광주광역시"
    if not sido:
        if toks[0].endswith("시") or toks[0].endswith("군"):
            sido, i = "", 0
        else:
            return None
    sgg = []
    while i < len(toks) and re.search(r"[시군구]$", toks[i]) and len(sgg) < 2 and not re.search(r"[로길]$", toks[i]):
        if sgg and not toks[i].endswith("구"):
            break
        sgg.append(toks[i])
        i += 1
    emd, kind = "", ""
    if i < len(toks):
        t = toks[i]
        if re.search(r"(로|길|대로)(\d+(번?길)?)?$", t) or re.search(r"\d+(번)?길$", t):
            emd, kind = t, "도로명"
            i += 1
        elif re.search(r"[읍면]$", t):
            emd = t
            i += 1
            if i < len(toks) and re.search(r"(로|길)(\d+(번?길)?)?$", toks[i]):
                emd, kind = emd + " " + toks[i], "도로명"
                i += 1
            elif i < len(toks) and re.search(r"리$", toks[i]):
                emd, kind = emd + " " + toks[i], "지번"
                i += 1
        elif re.search(r"(동|가|리)\d*$", t):
            emd, kind = t, "지번"
            i += 1
    rest = " ".join(toks[i:])
    if not kind and re.match(r"^\d", rest):
        kind = "지번"
    return {"우편번호": zipc, "시도": sido, "시군구": " ".join(sgg), "읍면동·도로명": emd,
            "나머지 주소": (rest + (" " + note if note else "")).strip(), "구분": kind}


def cmd_addr(a):
    wb, _ = load(a.file)
    ws = pick_sheet(wb, a.sheet)
    hrow = find_header(ws, a.header)
    if not a.col:
        cands = [n for n, c, k in col_kinds(ws, hrow)[0] if k == "addr"]
        if not cands:
            raise SystemExit("주소 칸을 못 찾았습니다. --col 칸이름 으로 알려 주세요.")
        a.col = cands[0]
    c = col_of(ws, hrow, a.col)
    keys = ["시도", "시군구", "읍면동·도로명", "나머지 주소", "구분"]
    rows = list(data_rows(ws, hrow))
    res = [split_addr(ws.cell(r, c).value) if ws.cell(r, c).value else None for r in rows]
    if any(x and x["우편번호"] for x in res):
        keys = ["우편번호"] + keys
    cols = {k: add_column(ws, hrow, k) for k in keys}
    bad = 0
    for r, x in zip(rows, res):
        if not x:
            if ws.cell(r, c).value:
                bad += 1
            continue
        for k in keys:
            ws.cell(r, cols[k], x[k] or None)
    out = a.out or default_out(a.file, "주소")
    save(wb, out)
    ok = sum(1 for x in res if x)
    print(f"주소 {ok}개를 나눔, 못 나눈 것 {bad}개 → {out}")
    for x in [x for x in res if x][:5]:
        print("  " + " | ".join(f"{k}:{x[k]}" for k in keys if x[k]))
    if not any(x and x["시도"] for x in res):
        print("※ 시도 이름이 없는 주소는 시도 칸이 비어 있습니다.")


def birth_of(v):
    info = rrn_info(v) if isinstance(v, str) else None
    if info and info.get("birth"):
        return info["birth"], info["sex"]
    d = parse_date(v, six=True)
    if d:
        d = d.date() if isinstance(d, dt.datetime) else d
        if d > TODAY and d.year >= 2000 and len(digits(v)) == 6:
            d = d.replace(year=d.year - 100)
        return d, None
    s = digits(v)
    if len(s) == 6:  # 900101 같은 생년월일 6자리
        y = int(s[:2])
        y += 2000 if y <= TODAY.year % 100 else 1900
        try:
            return dt.date(y, int(s[2:4]), int(s[4:6])), None
        except ValueError:
            return None, None
    return None, None


def man_age(b, on):
    return on.year - b.year - ((on.month, on.day) < (b.month, b.day))


def cmd_age(a):
    wb, _ = load(a.file)
    ws = pick_sheet(wb, a.sheet)
    hrow = find_header(ws, a.header)
    on = parse_date(a.on) if a.on else TODAY
    if on is None:
        raise SystemExit(f"기준 날짜로 못 읽었습니다: {a.on}")
    on = on.date() if isinstance(on, dt.datetime) else on
    if not a.col:
        cands = [n for n, c, k in col_kinds(ws, hrow)[0] if k == "rrn"] or \
                [n for n, c, k in col_kinds(ws, hrow)[0] if k == "date" and "생" in n]
        if not cands:
            raise SystemExit("주민번호나 생년월일 칸을 못 찾았습니다. --col 칸이름 으로 알려 주세요.")
        a.col = cands[0]
    c = col_of(ws, hrow, a.col)
    rows = list(data_rows(ws, hrow))
    got = [birth_of(ws.cell(r, c).value) for r in rows]
    has_sex = any(s for _, s in got)
    cb = add_column(ws, hrow, "생년월일")
    ca = add_column(ws, hrow, f"만 나이({on:%Y-%m-%d} 기준)")
    cs = add_column(ws, hrow, "성별") if has_sex else None
    bad = 0
    for r, (b, sx) in zip(rows, got):
        if not b:
            if ws.cell(r, c).value not in (None, ""):
                bad += 1
            continue
        cell = ws.cell(r, cb, dt.datetime(b.year, b.month, b.day))
        cell.number_format = "yyyy-mm-dd"
        ws.cell(r, ca, man_age(b, on))
        if cs and sx:
            ws.cell(r, cs, sx)
    out = a.out or default_out(a.file, "나이")
    save(wb, out)
    print(f"{len(rows) - bad}명 계산, 못 읽은 칸 {bad}개 ({on:%Y-%m-%d} 기준 만 나이) → {out}")


def kr_holidays(years, extra=()):
    hol = _need("holidays")
    h = {}
    for y in years:
        for d, n in hol.KR(years=y).items():
            h[d] = n
    for e in extra:
        d = parse_date(e)
        if d:
            h[d if isinstance(d, dt.date) and not isinstance(d, dt.datetime) else d.date()] = "임시공휴일(직접 넣음)"
    return h


def workdays_between(s, e, hol, sat=False):
    if e < s:
        return -workdays_between(e, s, hol, sat)
    n, d = 0, s
    while d <= e:
        if d.weekday() < (6 if sat else 5) and d not in hol:
            n += 1
        d += dt.timedelta(days=1)
    return n


def add_workdays(s, k, hol, sat=False):
    d, step = s, (1 if k >= 0 else -1)
    left = abs(k)
    while left:
        d += dt.timedelta(days=step)
        if d.weekday() < (6 if sat else 5) and d not in hol:
            left -= 1
    return d


def _as_date(x):
    d = parse_date(x, serial_ok=True)
    if d is None:
        return None
    return d.date() if isinstance(d, dt.datetime) else d


def cmd_workdays(a):
    extra = [x for x in (a.extra or "").split(",") if x.strip()]
    if os.path.exists(a.first) and os.path.splitext(a.first)[1].lower() in (".xlsx", ".xlsm", ".csv"):
        wb, _ = load(a.first)
        ws = pick_sheet(wb, a.sheet)
        hrow = find_header(ws, a.header)
        if not (a.start and a.end):
            raise SystemExit("--start 시작일칸 --end 종료일칸 을 알려 주세요.")
        cs, ce = col_of(ws, hrow, a.start), col_of(ws, hrow, a.end)
        rows = list(data_rows(ws, hrow))
        ds = [(_as_date(ws.cell(r, cs).value), _as_date(ws.cell(r, ce).value)) for r in rows]
        yrs = {d.year for p in ds for d in p if d}
        hol = kr_holidays(range(min(yrs or {TODAY.year}) - 1, max(yrs or {TODAY.year}) + 2), extra)
        cw = add_column(ws, hrow, "영업일 수")
        cc = add_column(ws, hrow, "달력 일수")
        bad = 0
        for r, (s, e) in zip(rows, ds):
            if not (s and e):
                bad += 1
                continue
            ws.cell(r, cw, workdays_between(s, e, hol, a.sat))
            ws.cell(r, cc, (e - s).days + 1)
        out = a.out or default_out(a.first, "영업일")
        save(wb, out)
        print(f"{len(rows) - bad}줄 계산(시작일·종료일 모두 넣어 셈, 토·일·공휴일 뺌{' — 토요일은 일하는 날로' if a.sat else ''}), 못 읽은 줄 {bad} → {out}")
        return
    s = _as_date(a.first)
    if not s:
        raise SystemExit(f"날짜로 못 읽었습니다: {a.first}")
    if a.add is not None:
        hol = kr_holidays(range(s.year - 1, s.year + 3), extra)
        d = add_workdays(s, a.add, hol, a.sat)
        print(f"{s:%Y-%m-%d}({'월화수목금토일'[s.weekday()]})에서 {a.add} 영업일 {'뒤' if a.add >= 0 else '앞'}: "
              f"{d:%Y-%m-%d}({'월화수목금토일'[d.weekday()]})")
        return
    if not a.second:
        raise SystemExit("끝 날짜를 주거나 --add 일수 를 주세요.")
    e = _as_date(a.second)
    if not e:
        raise SystemExit(f"날짜로 못 읽었습니다: {a.second}")
    hol = kr_holidays(range(min(s.year, e.year), max(s.year, e.year) + 1), extra)
    n = workdays_between(s, e, hol, a.sat)
    within = sorted(d for d in hol if min(s, e) <= d <= max(s, e) and d.weekday() < 5)
    print(f"{s:%Y-%m-%d} ~ {e:%Y-%m-%d}: 달력 {abs((e - s).days) + 1}일, 영업일 {n}일 (양 끝 포함)")
    if within:
        print("  빠진 평일 공휴일: " + ", ".join(f"{d:%m-%d} {hol[d]}" for d in within))


def cmd_holidays(a):
    if a.year and not (a.year.isdigit() and 1950 <= int(a.year) <= 2100):
        raise SystemExit(f"연도를 숫자로 주세요(예: 2026): {a.year}")
    y = int(a.year or TODAY.year)
    hol = kr_holidays([y])
    print(f"{y}년 공휴일 {len(hol)}일 (일요일과 겹친 날 포함)")
    for d in sorted(hol):
        print(f"  {d:%Y-%m-%d}({'월화수목금토일'[d.weekday()]}) {hol[d]}")
    print("※ 나라에서 갑자기 정하는 임시공휴일은 빠져 있을 수 있습니다. 있으면 --extra 2026-10-02 처럼 넣습니다.")


def cmd_money(a):
    if not os.path.exists(a.first):
        n = parse_money(a.first)
        if n is None:
            raise SystemExit(f"금액으로 못 읽었습니다: {a.first}")
        if abs(n) >= 10 ** 20:
            raise SystemExit("너무 큰 금액입니다(1해 원 미만까지).")
        print(han_money(n, a.plain))
        print(f"금 {int(n):,}원")
        if not float(n).is_integer():
            print(f"※ 원 아래({n - int(n):.2f})는 한글 금액에서 뺐습니다.")
        return
    wb, _ = load(a.first)
    ws = pick_sheet(wb, a.sheet)
    hrow = find_header(ws, a.header)
    if not a.col:
        cands = [n for n, c, k in col_kinds(ws, hrow)[0] if k in ("money", "number")]
        if not cands:
            raise SystemExit("금액 칸을 못 찾았습니다. --col 칸이름 으로 알려 주세요.")
        a.col = cands[0]
    c = col_of(ws, hrow, a.col)
    nc = add_column(ws, hrow, f"{a.col}(한글)")
    rows, bad = list(data_rows(ws, hrow)), 0
    for r in rows:
        v = ws.cell(r, c).value
        if isinstance(v, str) and v.startswith("="):
            bad += 1
            continue
        n = parse_money(v)
        if n is None:
            bad += 1
            continue
        ws.cell(r, nc, han_money(n, a.plain))
    out = a.out or default_out(a.first, "한글금액")
    save(wb, out)
    print(f"{len(rows) - bad}줄에 한글 금액을 넣음, 못 읽은 칸 {bad}개(계산식 칸 포함) → {out}")


TOTAL_RE = re.compile(r"^\s*(합\s*계|소\s*계|총\s*계|누\s*계|계|총합|total|sum|subtotal)\s*$", re.I)


def _table(ws, hrow, vws=None):
    """칸 이름 → 값 줄 목록. 계산식 칸은 엑셀이 저장해 둔 결과 값으로 옮기고(없으면 빈칸),
    '합계·소계' 줄은 합친 표에서 틀린 숫자가 되므로 뺀다."""
    hs = headers(ws, hrow)
    rows = []
    for r in data_rows(ws, hrow):
        row = {}
        is_total = False
        for k, c in hs.items():
            v = ws.cell(r, c).value
            if isinstance(v, str) and TOTAL_RE.match(v):
                is_total = True
            if isinstance(v, str) and v.startswith("="):
                cached = vws.cell(r, c).value if vws is not None else None
                if cached is None:
                    ws._kx_formulas = getattr(ws, "_kx_formulas", 0) + 1
                v = cached
            row[k] = v
        if is_total:
            ws._kx_totals = getattr(ws, "_kx_totals", 0) + 1
            continue
        if any(v not in (None, "") for v in row.values()):
            rows.append(row)
    return list(hs), rows


def _values_wb(path):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        try:
            return openpyxl.load_workbook(path, data_only=True)
        except Exception:  # noqa: BLE001
            return None
    return None


def _write_table(ws, cols, rows, fmt_src=None):
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    ws.append(cols)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", start_color="DDEBF7", end_color="DDEBF7")
    for r in rows:
        ws.append([r.get(c) for c in cols])
    for i, c in enumerate(cols, 1):
        width = max([len(str(c))] + [len(_csv_val(r.get(c))) for r in rows[:200]] + [4])
        ws.column_dimensions[get_column_letter(i)].width = min(45, width * 1.6 + 2)
        for row in ws.iter_rows(min_row=2, min_col=i, max_col=i):
            for cell in row:
                if isinstance(cell.value, (dt.date, dt.datetime)):
                    cell.number_format = "yyyy-mm-dd"
    ws.freeze_panes = "A2"


def cmd_combine(a):
    openpyxl = _need("openpyxl")
    allcols, allrows, srcs = [], [], []
    same = {}
    for pair in (a.same or "").split(","):  # "성명=이름" → '이름' 칸을 '성명' 칸으로
        if "=" in pair:
            keep, *others = [x.strip() for x in pair.split("=")]
            for o in others:
                same[o] = keep
    for f in a.files:
        wb, _ = load(f)
        vwb = _values_wb(f)
        if a.sheet:
            sheets = [wb[a.sheet]] if a.sheet in wb.sheetnames else [wb.active]
        else:
            sheets = wb.worksheets
        for ws in sheets:
            if ws.max_row < 2:
                continue
            hrow = find_header(ws, a.header)
            cols, rows = _table(ws, hrow, vwb[ws.title] if vwb and ws.title in vwb.sheetnames else None)
            if same:
                cols = [same.get(c, c) for c in cols]
                rows = [{same.get(k, k): v for k, v in r.items()} for r in rows]
            for c in cols:
                if c not in allcols:
                    allcols.append(c)
            src = os.path.basename(f) + ("" if len(wb.sheetnames) == 1 else f" / {ws.title}")
            for r in rows:
                r["출처"] = src
            allrows += rows
            fx, tt = getattr(ws, "_kx_formulas", 0), getattr(ws, "_kx_totals", 0)
            srcs.append(f"{src}: {len(rows)}줄, 칸 {len(cols)}개" + (f" (합계·소계 줄 {tt}개 뺌)" if tt else "")
                        + (f" (값이 저장 안 된 계산식 칸 {fx}개는 빈칸)" if fx else ""))
    out_wb = openpyxl.Workbook()
    ws = out_wb.active
    ws.title = "합본"
    cols = allcols + (["출처"] if not a.no_source else [])
    _write_table(ws, cols, allrows)
    out = a.out or os.path.join(os.getcwd(), "합본.xlsx")
    save(out_wb, out)
    print("\n".join(srcs))
    print(f"모두 {len(allrows)}줄, 칸 {len(allcols)}개 → {out}")
    odd = [c for c in allcols if sum(1 for r in allrows if r.get(c) not in (None, "")) < len(allrows) * 0.5]
    if odd:
        print("※ 일부 파일에만 있는 칸(이름이 조금 다른 같은 칸일 수 있음): " + ", ".join(odd[:15]))


def _safe_sheet(name, used):
    n = re.sub(r"[\\/*?:\[\]]", "_", str(name if name not in (None, "") else "(빈 값)"))[:31].strip("'") or "_"
    if n.lower() == "history":
        n += "_"
    base, k = n, 2
    while n.lower() in used:
        n = f"{base[:28]}_{k}"
        k += 1
    used.add(n.lower())
    return n


def cmd_split(a):
    openpyxl = _need("openpyxl")
    wb, _ = load(a.file)
    ws = pick_sheet(wb, a.sheet)
    hrow = find_header(ws, a.header)
    col_of(ws, hrow, a.by)
    vwb = _values_wb(a.file)
    cols, rows = _table(ws, hrow, vwb[ws.title] if vwb and ws.title in vwb.sheetnames else None)
    key = next(c for c in cols if c.replace(" ", "") == a.by.replace(" ", "")) if a.by not in cols else a.by
    groups = {}
    for r in rows:
        groups.setdefault(_csv_val(r.get(key)).strip(), []).append(r)
    used = set()
    if a.files:
        outdir = a.out or os.path.join(os.getcwd(), f"{os.path.splitext(os.path.basename(a.file))[0]}_나눔")
        os.makedirs(outdir, exist_ok=True)
        for g, rs in groups.items():
            w2 = openpyxl.Workbook()
            _write_table(w2.active, cols, rs)
            w2.active.title = _safe_sheet(g, set())
            fn = re.sub(r'[\\/:*?"<>|]', "_", g or "(빈 값)")[:60].rstrip(". ")
            base, k = fn, 2
            while fn.lower() in used:  # 이름이 겹치면(A/B와 A_B) 번호를 붙인다
                fn, k = f"{base}_{k}", k + 1
            used.add(fn.lower())
            save(w2, os.path.join(outdir, fn + ".xlsx"))
        print(f"'{key}' 값 {len(groups)}가지로 파일 {len(groups)}개 → {outdir}")
    else:
        w2 = openpyxl.Workbook()
        w2.remove(w2.active)
        for g, rs in groups.items():
            _write_table(w2.create_sheet(_safe_sheet(g, used)), cols, rs)
        out = a.out or default_out(a.file, "나눔", ".xlsx")
        save(w2, out)
        print(f"'{key}' 값 {len(groups)}가지로 시트 {len(groups)}개 → {out}")
    for g, rs in sorted(groups.items(), key=lambda kv: -len(kv[1]))[:15]:
        print(f"  {g or '(빈 값)'}: {len(rs)}줄")


def _key_norm(v):
    if v is None:
        return ""
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        p = phone_fmt(v)
        return digits(p) if p else _csv_val(v)
    s = norm_text(v).lower()
    if phone_fmt(s) or rrn_info(s) or BIZ_RE.match(s):
        return digits(s)
    return re.sub(r"\s+", "", s)


def cmd_dupes(a):
    wb, _ = load(a.file)
    ws = pick_sheet(wb, a.sheet)
    hrow = find_header(ws, a.header)
    keys = [col_of(ws, hrow, x.strip()) for x in a.cols.split(",")] if a.cols else list(headers(ws, hrow).values())
    rows = list(data_rows(ws, hrow))
    groups = {}
    for r in rows:
        k = tuple(_key_norm(ws.cell(r, c).value) for c in keys)
        if all(x == "" for x in k):
            continue
        groups.setdefault(k, []).append(r)
    dup = [g for g in groups.values() if len(g) > 1]
    extra = sum(len(g) - 1 for g in dup)
    print(f"[{ws.title}] {len(rows)}줄 중 겹치는 묶음 {len(dup)}개 (지울 수 있는 줄 {extra}개)")
    for g in dup[: a.show]:
        print(f"  {', '.join(str(r) for r in g)}번째 줄: " + " / ".join(_peek(ws.cell(g[0], c).value, '') for c in keys))
    if not dup:
        return
    if a.remove and any(m.max_row > hrow for m in ws.merged_cells.ranges):
        print("※ 자료 칸에 합친 칸이 있어서 줄을 지우면 표가 틀어집니다. 지우지 않고 색 표시만 합니다.")
        a.remove = False
    if a.remove:
        drop = sorted((r for g in dup for r in g[1:]), reverse=True)
        for r in drop:
            ws.delete_rows(r)
        if any(isinstance(c.value, str) and c.value.startswith("=") for row in ws.iter_rows() for c in row):
            print("※ 이 시트에 계산식이 있습니다. 줄을 지우면 합계 같은 계산식의 범위가 어긋날 수 있으니 엑셀에서 확인하세요.")
        out = a.out or default_out(a.file, "중복지움")
        save(wb, out, encoding=getattr(wb, "_kx_csv_encoding", "utf-8-sig"))
        print(f"뒤에 나온 겹친 줄 {len(drop)}개를 지움(처음 나온 줄은 남김) → {out}")
    else:
        mc = add_column(ws, hrow, "중복")
        fills = ["FFF2CC", "E2EFDA", "FCE4D6", "DDEBF7", "EDE7F6"]
        for i, g in enumerate(dup, 1):
            for j, r in enumerate(g):
                ws.cell(r, mc, f"묶음{i}" + (" (처음)" if j == 0 else ""))
                for c in keys:
                    try:
                        ws.cell(r, c).fill = _fill(fills[(i - 1) % len(fills)])
                    except Exception:  # noqa: BLE001
                        pass
        out = a.out or default_out(a.file, "중복표시")
        if os.path.splitext(out)[1].lower() == ".csv":
            out = os.path.splitext(out)[0] + ".xlsx"
        save(wb, out)
        print(f"겹친 줄에 색과 '중복' 칸을 표시함 → {out}")


def cmd_tocsv(a):
    wb, _ = load(a.file)
    ws = pick_sheet(wb, a.sheet)
    enc = {"cp949": "cp949", "euc-kr": "cp949", "utf8": "utf-8-sig", "utf-8": "utf-8-sig"}.get(a.encoding.lower(), a.encoding)
    out = a.out or default_out(a.file, "csv", ".csv")
    if os.path.exists(a.file) and a.file.lower().endswith((".xlsx", ".xlsm")):
        # 계산식 칸은 저장된 계산 결과로
        import openpyxl
        wbv = openpyxl.load_workbook(a.file, data_only=True)
        wsv = wbv[ws.title]
        empty_formula = sum(1 for row in ws.iter_rows() for c in row
                            if isinstance(c.value, str) and c.value.startswith("=") and wsv[c.coordinate].value is None)
        ws = wsv
        if empty_formula:
            print(f"※ 계산식 칸 {empty_formula}개는 엑셀에서 한 번도 계산·저장되지 않아 값이 비어 있습니다. 엑셀에서 열어 저장한 뒤 다시 하면 채워집니다.")
    with open(out, "w", newline="", encoding=enc, errors="replace") as fh:
        w = csv.writer(fh)
        n = 0
        for row in ws.iter_rows(values_only=True):
            w.writerow([_csv_val(v) for v in row])
            n += 1
    bad = 0
    if enc == "cp949":
        txt = "\n".join(",".join(_csv_val(v) for v in row) for row in ws.iter_rows(values_only=True))
        bad = sum(1 for ch in txt if not _cp949_ok(ch))
    print(f"[{ws.title}] {n}줄 → {out} (글자 모양 {'CP949(관공서·은행용)' if enc == 'cp949' else 'UTF-8'})")
    if bad:
        print(f"※ CP949로 못 쓰는 글자 {bad}개(이모지·특수 한자 등)는 ?로 바뀌었습니다.")


def _cp949_ok(ch):
    try:
        ch.encode("cp949")
        return True
    except UnicodeEncodeError:
        return False


# ───────────────────────── 명령 줄 ─────────────────────────

def main(argv=None):
    p = argparse.ArgumentParser(prog="kexcel", description="한국식 엑셀 도구")
    sub = p.add_subparsers(dest="cmd", required=True)

    def common(sp, out=True):
        sp.add_argument("--sheet")
        sp.add_argument("--header", type=int)
        if out:
            sp.add_argument("--out")
        sp.add_argument("--show", type=int, default=15)

    sp = sub.add_parser("info"); sp.add_argument("file"); common(sp, False); sp.set_defaults(fn=cmd_info)
    sp = sub.add_parser("check"); sp.add_argument("file"); sp.add_argument("--cols"); common(sp); sp.set_defaults(fn=cmd_check)
    sp = sub.add_parser("clean"); sp.add_argument("file"); sp.add_argument("--what", default="all")
    sp.add_argument("--cols"); sp.add_argument("--dry", action="store_true"); common(sp); sp.set_defaults(fn=cmd_clean)
    sp = sub.add_parser("mask"); sp.add_argument("file"); sp.add_argument("--words"); sp.add_argument("--cols")
    sp.add_argument("--keep-names", action="store_true"); sp.add_argument("--dry", action="store_true"); common(sp)
    sp.set_defaults(fn=cmd_mask)
    sp = sub.add_parser("addr"); sp.add_argument("file"); sp.add_argument("--col"); common(sp); sp.set_defaults(fn=cmd_addr)
    sp = sub.add_parser("age"); sp.add_argument("file"); sp.add_argument("--col"); sp.add_argument("--on"); common(sp)
    sp.set_defaults(fn=cmd_age)
    sp = sub.add_parser("workdays"); sp.add_argument("first"); sp.add_argument("second", nargs="?")
    sp.add_argument("--add", type=int); sp.add_argument("--start"); sp.add_argument("--end")
    sp.add_argument("--sat", action="store_true"); sp.add_argument("--extra"); common(sp); sp.set_defaults(fn=cmd_workdays)
    sp = sub.add_parser("holidays"); sp.add_argument("year", nargs="?"); sp.set_defaults(fn=cmd_holidays)
    sp = sub.add_parser("money"); sp.add_argument("first"); sp.add_argument("--col"); sp.add_argument("--plain", action="store_true")
    common(sp); sp.set_defaults(fn=cmd_money)
    sp = sub.add_parser("combine"); sp.add_argument("files", nargs="+"); sp.add_argument("--no-source", action="store_true")
    sp.add_argument("--same", help="이름이 다른 같은 칸: 성명=이름,연락처=휴대폰")
    common(sp); sp.set_defaults(fn=cmd_combine)
    sp = sub.add_parser("split"); sp.add_argument("file"); sp.add_argument("--by", required=True)
    sp.add_argument("--files", action="store_true"); common(sp); sp.set_defaults(fn=cmd_split)
    sp = sub.add_parser("dupes"); sp.add_argument("file"); sp.add_argument("--cols"); sp.add_argument("--remove", action="store_true")
    common(sp); sp.set_defaults(fn=cmd_dupes)
    sp = sub.add_parser("tocsv"); sp.add_argument("file"); sp.add_argument("--encoding", default="utf-8")
    common(sp); sp.set_defaults(fn=cmd_tocsv)

    a = p.parse_args(argv)
    try:
        a.fn(a)
    except BrokenPipeError:
        pass


if __name__ == "__main__":
    main()
