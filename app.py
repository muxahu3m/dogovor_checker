import streamlit as st
import os 
import json
import io
import docx2txt
import re
import nltk
import pymorphy2
from dotenv import dotenv_values
from fuzzywuzzy import fuzz
from fuzzywuzzy import process 
import pandas as pd

morph = pymorphy2.MorphAnalyzer()

def get_doc_text(filepath, file):
    if True:#file.endswith('.docx'):
        text = docx2txt.process(filepath+file)
        return text
    elif file.endswith('.doc'):
        # converting .doc to .docx
        doc_file = filepath + file
        docx_file = filepath + file + 'x'
        if not os.path.exists(docx_file):
            os.system('antiword ' + doc_file + ' > ' + docx_file)
            with open(docx_file) as f:
                text = f.read()
            os.remove(docx_file) #docx_file was just to read, so deleting
        else:
          # already a file with same name as doc exists having docx extension, 
          # which means it is a different file, so we cant read it
            print('Info : file with same name of doc exists having docx extension, so we cant read it')
            text = ''
        return text
    
def clean_text(text, clean_PD=False):
    text = re.sub('[\\n]{2,99}','\n',text)
    text = re.sub(r'[\s]{2,99}',' ',text)
    if clean_PD:
        text = re.sub('[#№:\s]{1}[\d]{15}\s',' 555555555555555 ',text) #Егрип
        text = re.sub('[#№:\s]{1}[\d]{13}\s',' 5555555555555 ',text) #ОГРН
        text = re.sub('[#№:\s]{1}[\d]{12}\s',' 555555555555 ',text) #ИНН физ лиц
        text = re.sub('[#№:\s]{1}[\d]{10}\s',' 5555555555 ',text) #ИНН юр лиц
        text = re.sub('[#№:\s]{1}[\d]{9}\s', ' 555555555 ',text) #БИК
        text = re.sub('[#№:\s]{1}[\d]{20}\s',' 55555555555555555555 ',text) #счета
        text = re.sub('((8|\+7)[\- ]?){1}(\(?\d{3}\)?[\- ]?)?[\d\- ]{7,10}',' +74955555555 ',text)
        text = re.sub('[A-Za-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,4}',' admin@mail.ru ',text)

        prob_thresh = 0.07
        fio = {}
        for word in nltk.word_tokenize(text):
            if word[-1]=='.':
                word = word[:-1]
            for p in morph.parse(word):
                if word in ('ЕГРИП', 'ИНН', 'ОГРНИП'):
                    continue
                if 'Name' in p.tag and p.score >= prob_thresh  and word not in fio.keys():
                    print('Name', word, p.score)
                    fio[word] = morph.parse('Иван')[0].inflect({p.tag.case}).word.title()
                elif 'Patr' in p.tag and p.score >= prob_thresh  and word not in fio.keys():
                    print('Patr', word, p.score)
                    fio[word] = morph.parse('Иванович')[0].inflect({p.tag.case}).word.title()
                elif 'Surn' in p.tag and p.score >= prob_thresh and word not in fio.keys():
                    print('Surn', word, p.score)
                    fio[word] = morph.parse('Иванов')[0].inflect({p.tag.case}).word.title()
        #         elif 'Geox' in p.tag and p.score >= prob_thresh and word not in fio.keys():
        #             print('Geos', word, p.score)
        #             fio[word] = morph.parse('Бобруйск')[0].inflect({p.tag.case}).word.title()

        for w ,n in fio.items():
            text = text.replace(w, n)
    return text

st.set_page_config(page_title="Проверка договоров по чеклисту", page_icon="📝")
st.title('Проверка договоров по чеклисту')
clean_PD = st.sidebar.checkbox('Скрывать персональные данные', False)
uploaded_file = st.sidebar.file_uploader('Загрузите файл с договором для оценки...', type={ "docx"})


if uploaded_file:
    # uploaded_file = StringIO(uploaded_file.getvalue().decode("utf-8"))
    # with open(uploaded_file, 'r') as f:
    #     dialog = f.read()
    bytes_data = uploaded_file.getvalue()
    raw_text = docx2txt.process(uploaded_file)
    # st.write("dogovor.docx", uploaded_file.name)
    # st.write(bytes_data)
    # text = get_doc_text("", "dogovor.docx")
    text = clean_text(raw_text, clean_PD=clean_PD)
    # file = uploaded_file.read()#.decode('utf-8')
    st.text_area('Предпросмотр Договора',value=text)
    # st.write('ПД не удаляются')
    df = pd.read_excel('dogovor_checklist.xlsx')
    df['Описание'] = df['Описание'].ffill()
    result = []
    output = ""
    for p in df['Описание'].unique():
        output = output + f"\n{p},(количество проверяемых пунктов {len(df[df['Описание']==p]['Формулировка '].tolist())}) \n"
        for i, f in enumerate(df[df['Описание']==p]['Формулировка '].tolist()):
            if pd.isnull(f):
                continue
            r = fuzz.partial_ratio(f,text[6:])    # не понимаю что за косяк с полной шапкой, кажется какой-то спецсимвол проскакивает
            if r==45:
                st.write(f)
            if r>=90:
                givenOpt = text.split('\n') 
                ratios = process.extract(f,givenOpt) 
                output = output + f' - Пункт {i+1} присутствует: {ratios[0][0]}\n'
                res = 'Присутствует'
            else:
                output = output + f' - Пункт {i+1} отсутствует. Ближайший: {r} %\n'
                givenOpt = text.split('\n') 
                ratios = process.extract(f,givenOpt) 
                res = 'Отсутствует'
            result.append([p,f,res, ratios[0][0]])    
    dr = pd.DataFrame(result, columns=['Раздел', 'Искомый фрагмент', 'Результат', 'Найденный фрагмент'])
    dr['Найденный фрагмент'] = dr.apply(lambda x: None if x['Результат']=='Отсутствует' else x['Найденный фрагмент'],axis=1)

    st.write(f"""Проверено пунктов: {len(dr)}  
Найдено фрагментов: {len(dr[dr['Результат']=='Присутствует'])}  
Отсутствует фрагментов: {len(dr[dr['Результат']=='Отсутствует'])}  
Полный отчет в Excel""")
    buffer = io.BytesIO()
    # dr.to_excel(buffer, sheet_name='Sheet1', index=False)
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    # Write each dataframe to a different worksheet.
        dr.to_excel(writer, sheet_name='Sheet1', index=False)
        worksheet = writer.sheets['Sheet1']  # pull worksheet object

        for idx, col in enumerate(dr):  # loop through all columns
            series = dr[col]
            max_len = max((
                series.astype(str).map(len).max(),  # len of largest item
                len(str(series.name))  # len of column name/header
                )) + 1  # adding a little extra space
            print(max_len)
            if idx==0:
                worksheet.set_column(idx, idx, 25)
            elif max_len>100:
                worksheet.set_column(idx, idx, 100)
            else: 
                worksheet.set_column(idx, idx, max_len)  # set column width
        # writer.close()

    download2 = st.download_button(
        label="Скачать отчёт о проверке в Excel",
        data=buffer,
        key=1,
        file_name=f'{uploaded_file.name}_checkResults.xlsx',
        mime='application/vnd.ms-excel'
    )
    # print(output)

    st.text_area('Результаты проверки: ', output, height=400)
#             print(ratios[0][0])
    # st.text_area('Содержимое диалога',value=dialog)
    # with st.spinner(text="Ждём ответа от YandexGPT, немного терпения..."):
    #     ans, answer, json_valid = ask_gpt(dialog)

    # try:
    #     df = pd.DataFrame(pd.Series(answer), columns=['Результат'])
    #     df.index.name = 'Вопросы'
    #     st.dataframe(df.reset_index(), hide_index=True)
    # except Exception as e:
    #     st.text_area('Не распарсилось',value=f'{e}\n\n{e.args}')

    # st.text_area('Ответ Яндекса',value=f'{json_valid}\n\n{answer}', height=300)
    # st.text_area('Ответ Яндекса полный',value=f'{ans}\n\n{json_valid}', height=300)

