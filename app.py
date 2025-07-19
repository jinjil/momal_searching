import os, tempfile
import gradio as gr
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
import xlsxwriter

try:
    import ujson as json
except ImportError:
    import json

def process_json_folder(folder, keywords, window):
    all_rows = []
    for fn in os.listdir(folder):
        if not fn.lower().endswith(".json"):
            continue
        with open(os.path.join(folder, fn), encoding="utf-8") as f:
            data = json.load(f)
        utts = data.get('document',[{}])[0].get('utterance',[])
        for idx, utt in enumerate(utts):
            text = utt.get('form','')
            found = [k for k in keywords if k in text]
            if found:
                for off in range(-window, window+1):
                    j = idx + off
                    if 0 <= j < len(utts):
                        u = utts[j]
                        all_rows.append({
                            "파일명": fn,
                            "발화자 id": u.get('speaker_id'),
                            "발화 내용": u.get('form',''),
                            "포함된 어휘": ",".join(found),
                            "순서": off
                        })
    return pd.DataFrame(all_rows)

def make_excel(df):
    df = df[df['발화 내용'].str.strip() != ""].reset_index(drop=True)
    df['순서'] = pd.to_numeric(df['순서']).fillna(0).astype(int)
    mask = df['순서'].diff().le(0).fillna(True)
    df['Group_Index'] = mask.cumsum().astype(int)

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb = xlsxwriter.Workbook(tmp.name)
    ws = wb.add_worksheet()
    border = wb.add_format({'bottom':6})
    bold   = wb.add_format({'bold':True})

    # 헤더
    for c,col in enumerate(df.columns):
        ws.write(0,c,col,bold)
    # 내용
    for r,row in df.iterrows():
        fmt = border if (r+1==len(df) or df.at[r+1,'Group_Index']!=row['Group_Index']) else None
        for c,col in enumerate(df.columns):
            ws.write(r+1, c, row[col], fmt)
    wb.close()
    return tmp.name

def run(json_zip, kw_text, window):
    # 1) ZIP 풀기
    tmpdir = tempfile.mkdtemp()
    import zipfile
    with zipfile.ZipFile(json_zip.name, 'r') as z:
        z.extractall(tmpdir)
    # 2) DataFrame 생성
    keys = [k.strip() for k in kw_text.split(',') if k.strip()]
    df = process_json_folder(tmpdir, keys, window)
    if df.empty:
        return None, "🔍 검색 결과가 없습니다."
    # 3) 엑셀 생성
    path = make_excel(df)
    return path, "✅ 엑셀 생성 완료!"

iface = gr.Interface(
    fn=run,
    inputs=[
        gr.File(label="JSON(.zip) 업로드", type="file"),
        gr.Textbox(label="검색 키워드 (콤마 구분)", value="어떻,어떡"),
        gr.Slider(minimum=0, maximum=5, value=2, step=1, label="Window")
    ],
    outputs=[
        gr.File(label="다운로드할 엑셀 파일"),
        gr.Textbox(label="")
    ],
    title="대화 검색 → 엑셀 생성"
)

if __name__=="__main__":
    iface.launch()
