import streamlit as st
from ppt_lib import *
import base64

def create_download_link(data, filename):
    b64 = base64.b64encode(data).decode()  # 將檔案數據轉換為 base64 編碼
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}" download="{filename}">點此下載</a>'
    return href

def split_list(input_list, chunk_size):
    return [input_list[i:i + chunk_size] for i in range(0, len(input_list), chunk_size)]

if __name__ == '__main__':
    st.set_page_config(layout = 'wide')
    
    st.title("ChatPPT")
    
    title = st.text_input("請輸入標題:", value = '要如何學習機器學習')
    sub_title = st.text_input("請輸入副標題:", value = '李御璽/銘傳大學資工系')
    
    context = st.text_area("請輸入內文:", value = '', height=170)
    
    if st.button("產生PPT"):
        prs = Presentation()
        
        sub_title = sub_title.replace('/', '\n')
        
        create_title(prs, title, sub_title)
        
        item_list = []
        ctn_title = '步驟'
        
        for line in context.splitlines():
            if len(line) >= 2 and line[0].isdigit() and line[1] == '.':
                item_list.append(line[2:].strip() + '\n')
            elif len(line) >= 2:
                if line.find('以下') >= 0:
                    ctn_title = line[line.find('以下'):].strip()
        
        chunked_lists = split_list(item_list, 4)

        slides = []

        # Print the chunked lists
        for sublist in chunked_lists:
            slides.append({'title': ctn_title, 'content': sublist})
        
        create_body(prs, slides)
    
        filename = title + '.pptx'
        prs.save(filename)   

        with open(filename, 'rb') as f:
            data = f.read()

        #顯示下載連結
        st.markdown(create_download_link(data, filename), unsafe_allow_html=True)