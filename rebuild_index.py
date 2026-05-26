"""
Rebuild BM25 index tu file ToanBoQuyDinh_NhatGo_2026.docx
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
DOCX_KB_FILE = os.path.join(BASE_DIR, "ToanBoQuyDinh_NhatGo_2026.docx")
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

def _extract_year(meta_line):
    m = re.search(r'(201[89]|202\d)', meta_line)
    return m.group(1) if m else '?'

def extract_chunks(filepath):
    chunks = []
    if not os.path.exists(filepath):
        print(f"LOI: Khong tim thay file {filepath}")
        return chunks
    doc = DocxDocument(filepath)
    current_title, current_meta, current_body = '', '', []
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        style = para.style.name if para.style else ''
        if style.startswith('Heading 2'):
            if current_body and len('\n'.join(current_body)) > 150:
                full = current_title + '\n' + current_meta + '\n' + '\n'.join(current_body)
                chunks.append({
                    'title': current_title[:120],
                    'text': full,
                    'source': current_title[:80],
                    'year': _extract_year(current_meta),
                    'tokens': tokenize_vi(full),
                })
            current_title, current_meta, current_body = text, '', []
        elif style.startswith('Heading 1'):
            current_title, current_body = '', []
        else:
            if 'Nam ban hanh' in text or 'File goc' in text:
                current_meta = text
            elif text != '[File .doc cu - can mo truc tiep de xem noi dung]':
                current_body.append(text)
    if current_body and len('\n'.join(current_body)) > 150:
        full = current_title + '\n' + current_meta + '\n' + '\n'.join(current_body)
        chunks.append({
            'title': current_title[:120],
            'text': full,
            'source': current_title[:80],
            'year': _extract_year(current_meta),
            'tokens': tokenize_vi(full),
        })
    return chunks

if __name__ == '__main__':
    print(f"Dang doc: {DOCX_KB_FILE}")
    chunks = extract_chunks(DOCX_KB_FILE)
    if not chunks:
        print("CANH BAO: Khong co noi dung nao duoc doc!")
        exit(1)
    print(f"So muc quy dinh: {len(chunks)}")
    tokenized = [c['tokens'] for c in chunks]
    bm25 = BM25Okapi(tokenized)
    with open(INDEX_CACHE, 'wb') as f:
        pickle.dump({'version': '4.0', 'bm25': bm25, 'chunks': chunks}, f)
    size_kb = os.path.getsize(INDEX_CACHE) // 1024
    print(f"Da luu index: bm25_index.pkl ({size_kb}KB)")
    print("Rebuild thanh cong!")
