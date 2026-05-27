"""
Rebuild BM25 index tu file NhatGo_TOAN_BO_QUY_DINH_2015_2026.docx
Chay: python rebuild_index.py
"""
import os, re, pickle

try:
    from docx import Document as DocxDocument
except ImportError:
    print("LOI: Chua cai python-docx. Chay: pip install python-docx")
    exit(1)

try:
    from rank_bm25 import BM25Okapi
except ImportError:
    print("LOI: Chua cai rank-bm25. Chay: pip install rank-bm25")
    exit(1)

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
DOCX_KB_FILE = os.path.join(BASE_DIR, "NhatGo_TOAN_BO_QUY_DINH_2015_2026.docx")
INDEX_CACHE  = os.path.join(BASE_DIR, "bm25_index.pkl")

STOPWORDS = {
    'toi','ban','anh','chi','em','minh','ho','chung','no',
    'va','hay','hoac','nhung','ma','vi','nen','thi','la','cua',
    'co','khong','duoc','cho','voi','ve','trong','tren','duoi',
    'khi','neu','sao','gi','nao','day','do','nay','kia',
    'bao','nhieu','it','lam','qua','rat','cung','deu','da',
    'dang','se','van','con','toi','den','nhu','theo','sau',
    'truoc','luc','ngay','luon','moi','cac','nhung','mot','hai',
    'lam','noi','hoi','xin','muon','can','phai','biet',
    'oi','the','vay','nha','nhe','bi','do','tu','ra','vo',
    'ty','cong','dieu','khoan','lao',
}

def tokenize_vi(text):
    text = text.lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    tokens = text.split()
    tokens = [t for t in tokens if t not in STOPWORDS and len(t) >= 2]
    return tokens if tokens else text.split()

_SKIP_TITLES = {'TỔNG HỢP TOÀN BỘ QUY ĐỊNH', 'CÔNG TY TNHH', 'MỤC LỤC', 'NHẤT GỖ'}

def _para_bold_size(p):
    bold, size = False, 11.0
    for r in p.runs:
        if r.text.strip():
            if r.bold:
                bold = True
            if r.font.size:
                size = max(size, r.font.size.pt)
    return bold, size

def _is_separator(text):
    stripped = text.replace(' ', '')
    return len(stripped) > 5 and set(stripped) <= {'─', '═', '-', '=', '|'}

def extract_chunks(filepath):
    chunks = []
    if not os.path.exists(filepath):
        print(f"LOI: Khong tim thay file {filepath}")
        return chunks

    def _flush(chapter, year, body):
        body_text = '\n'.join(body).strip()
        if not body_text or len(body_text) < 80:
            return
        title = f"{chapter} (nam {year})" if year != '?' else chapter
        full = f"CHU DE: {chapter}\nNAM: {year}\n\n{body_text}"
        chunks.append({
            'title': title[:120],
            'text': full,
            'source': chapter[:80],
            'year': year,
            'tokens': tokenize_vi(full),
        })

    doc = DocxDocument(filepath)
    current_chapter = ''
    current_year = '?'
    current_body = []
    in_content = False

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        bold, size = _para_bold_size(p)

        if not in_content:
            if bold and size >= 12 and 'CHƯƠNG' in text.upper():
                in_content = True
            continue

        if _is_separator(text) or size <= 9:
            _flush(current_chapter, current_year, current_body)
            current_body = []
            current_year = '?'
            continue

        if bold and size >= 16:
            if any(kw in text.upper() for kw in _SKIP_TITLES):
                continue
            if 'CHƯƠNG' in text.upper():
                continue
            _flush(current_chapter, current_year, current_body)
            current_chapter = text
            current_year = '?'
            current_body = []
            continue

        if bold and 12 <= size <= 14:
            if 'CHƯƠNG' in text.upper():
                continue
            m = re.search(r'[Nn]ăm\s+(\d{4})', text)
            if m:
                _flush(current_chapter, current_year, current_body)
                current_year = m.group(1)
                current_body = []
                continue

        if '📁' in text or 'File gốc' in text or 'File goc' in text:
            continue

        current_body.append(text)

    _flush(current_chapter, current_year, current_body)
    return chunks


if __name__ == '__main__':
    print(f"Dang doc: {DOCX_KB_FILE}")
    chunks = extract_chunks(DOCX_KB_FILE)
    if not chunks:
        print("CANH BAO: Khong co noi dung nao duoc doc!")
        exit(1)
    print(f"So chunks: {len(chunks)}")

    from collections import Counter
    topics = Counter(c['source'] for c in chunks)
    print("\nThong ke chu de:")
    for topic, cnt in topics.most_common():
        print(f"  {topic[:50]}: {cnt} van ban")

    tokenized = [c['tokens'] for c in chunks]
    bm25 = BM25Okapi(tokenized)
    with open(INDEX_CACHE, 'wb') as f:
        pickle.dump({'version': '5.0', 'bm25': bm25, 'chunks': chunks}, f)
    size_kb = os.path.getsize(INDEX_CACHE) // 1024
    print(f"\nDa luu index v5.0: bm25_index.pkl ({size_kb}KB)")
    print("Rebuild thanh cong!")
