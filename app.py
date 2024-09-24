import re
import fitz
import time
import easyocr
import pandas as pd
import streamlit as st
from streamlit import session_state as ss
from streamlit_pdf_viewer import pdf_viewer
from streamlit_gsheets import GSheetsConnection

st.set_page_config(layout="wide")

def extract_pdf_data(doc, progress_bar):
    df = pd.DataFrame(columns=['name', 'address', 'tel', 'product', 'quantity', 'order_no', 'order_date', 'post_code'])
    # doc = fitz.open(file)
    for i in range(len(doc)):
        time.sleep(0.01)
        progress_bar.progress(100//len(doc)*(i + 1), text=progress_text)
        if doc[0].get_text()[:4] == 'PICK':
            name_address = []
            page = doc[i]
            image_list = page.get_images()
            for idx, img in enumerate(image_list, start=1):
                if idx in [3,5]:
                    data = doc.extract_image(img[0])
                    image_stream = data.get('image')
                    reader = easyocr.Reader(['th','en'])
                    result = reader.readtext(image_stream)
                    text = [item[1] for item in result]
                    name_address.append("".join(text))
            prd_chk, ord_chk = 0,0
            prd, prd_tmp , qty, ord_no = [],'',[],''
            words = page.get_text('dict')
            for b in words["blocks"]:
                if b['type'] == 1:
                    continue
                else:
                    for l in b['lines']:
                        for s in l['spans']:
                            if s['text'].startswith('Shopee Order No'):
                                ord_chk = 1
                            elif ord_chk == 1:
                                ord_no = s['text']
                                ord_chk = 0
                            elif s['text'] == 'Qty':
                                prd_chk = 1
                            elif prd_chk == 1 and s['text'].isdigit():
                                prd_chk = 2
                            elif prd_chk == 2:
                                if prd_tmp == 'Total:':
                                    prd_tmp = ''
                                    prd_chk = 0 
                                elif s['text'].isdigit():
                                    prd.append(prd_tmp)
                                    qty.append(int(s['text']))
                                    prd_tmp = ''
                                    prd_chk = 1           
                                else:
                                    prd_tmp += s['text'].strip()
            new_rows = [['', '','',prd[i],qty[i],'','',''] for i in range(len(prd))]
            temp = pd.DataFrame(new_rows, columns=df.columns)
            temp['name'] = name_address[0]
            temp['address'] = name_address[1]
            temp['order_no'] = ord_no
            df = pd.concat([df,temp], ignore_index=True, axis=0) 
        else:
            words = doc[0].get_text('dict')
            adr_chk, name_chk, ord_chk, prd_chk = 1,0,0,0
            adr, name, ord_no, prd, prd_tmp , qty = '','','',[],'',[]
            for b in words["blocks"]:
                if b['type'] == 1:
                    continue
                else:
                    for l in b['lines']:
                        for s in l['spans']:
                            if 'ชําระโดย' in s['text']:
                                adr_chk = 0
                            elif adr_chk == 1:
                                adr += s['text']
                            elif s['text'] == 'ถึง':
                                name_chk = 1
                            elif name_chk == 1:
                                name = s['text']
                                name_chk = 0
                            elif s['text'] == 'Order ID':
                                ord_chk = 1
                            elif ord_chk == 1:
                                ord_no = s['text']
                                ord_chk = 0
                            elif s['text'] == 'Qty':
                                prd_chk = 1
                            elif prd_chk == 1 and not s['text'].isdigit():
                                prd_chk = 0
                            elif prd_chk == 1 and s['text'].isdigit():
                                prd_chk = 2
                            elif prd_chk == 2:
                                if prd_tmp == 'Total:':
                                    prd_tmp = ''
                                    prd_chk = 0 
                                elif s['text'].isdigit():
                                    prd.append(prd_tmp)
                                    qty.append(int(s['text']))
                                    prd_tmp = ''
                                    prd_chk = 1
                                else:
                                    prd_tmp += s['text'].strip()
            new_rows = [['', '','',prd[i],qty[i],'','',''] for i in range(len(prd))]
            temp = pd.DataFrame(new_rows, columns=df.columns)
            temp['name'] = name
            temp['address'] = adr
            temp['order_no'] = ord_no
            df = pd.concat([df,temp], ignore_index=True, axis=0)
    df['post_code'] = df['address'].apply(lambda x: re.findall(r"\d{5}", x)[0])
    time.sleep(1)
    progress_bar.empty()
    return df

col1, col2 = st.columns(2)

with col1:
    st.header("Upload file")
    # Declare variable.
    if 'pdf_ref' not in ss:
        ss.pdf_ref = None

    # Access the uploaded ref via a key.
    file = st.file_uploader("Upload PDF file", type=('pdf'), key='pdf')

    if ss.pdf:
        ss.pdf_ref = ss.pdf  # backup

    container = st.container(height=400, border=True)

    # Now you can access "pdf_ref" anywhere in your app.
    if ss.pdf_ref:
        binary_data = ss.pdf_ref.getvalue()
        with container:    
            pdf_viewer(input=binary_data, width=500)


with col2:
    pdf_df = None
    st.header("Extract text")
    if file is not None:
        pdf_bytes = file.getvalue()
        # Open the PDF with PyMuPDF (fitz) using the bytes
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        progress_text = "Please wait."
        my_bar = st.progress(0, text=progress_text)
        pdf_df = extract_pdf_data(doc, my_bar)
        editor_df = st.data_editor(pdf_df)
        append_bt = st.button('Append to Google Sheet')

    if append_bt:
        conn = st.connection("gsheets", type=GSheetsConnection)
        url = "https://docs.google.com/spreadsheets/d/1CpM_k15yxdB0r9QQJA2esNsRLwU3Lq0LGEIxoR6JkuI/edit?usp=sharing"
        df = conn.read(spreadsheet=url, worksheet="ST")
        new_df = pd.concat([df, editor_df], axis=0)
        try:
            conn.update(spreadsheet=url, worksheet="ST", data=new_df)
            st.success("Done!")
        except Exception as e:
            st.success("Got Error!")
    