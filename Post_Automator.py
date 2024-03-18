#For Process
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from xlsxwriter import Workbook


from PIL import Image
import io
import sys
import easyocr
import zipfile

from base64 import b64decode
import pandas as pd
import requests
import os
from time import sleep,time
import datetime

def start(df,i,l,sleep_,pdf_opt):
    reader = easyocr.Reader(['en'])
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    driver = webdriver.Chrome(options=chrome_options)
    
    driver.get("https://www.indiapost.gov.in/_layouts/15/DOP.Portal.Tracking/TrackConsignment.aspx")
    def captcha_solve():
        link = driver.find_element(By.XPATH,"//div[@class = 'input-group']//img").get_attribute('src')
        response = requests.get(link)
        sleep(4)
        image = Image.open(io.BytesIO(response.content))
        image = image.convert('RGB')
        image.save('captcha.jpg', 'JPEG')
        
        try:
            result = reader.readtext('captcha.jpg')[0][1].replace(' ','')
        except :
            result=''
        os.remove('captcha.jpg')
        return result
    
    def captcha_context():
        cap = captcha_solve()
        if cap == '':
            driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_imgbtnCaptcha').click()
            return ''
        context = driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_lblCaptcha').text
        if context == 'Enter the First number':
            return cap[0] if len(cap) >= 1 else ''
        if context == 'Enter the Second number':
            return cap[1] if len(cap) >= 2 else ''
        if context == 'Enter the Third number':
            return cap[2] if len(cap) >= 3 else ''
        if context == 'Enter the Fourth number':
            return cap[3] if len(cap) >= 4 else ''
        if context == 'Enter the Fifth number':
            return cap[4] if len(cap) == 5 else ''
        return ''
        
    
    pdfs = []
    df = df[i-1:l]
    df.index = range(i,l+1)
    df_view = st.empty()
    df_view.dataframe(df)
    cols = st.columns(4)
    with st.spinner('Please wait..'):
        sleep(1)
    with st.status("Processing.....",expanded=True):
        try:
            ot = time()
            rt = 0
            while i<=l:
                ref = df.loc[i,'RPAD Barcode No ']
                if str(ref)=='nan':
                    i += 1
                    continue
                if rt == 0:
                    rt = time()
                ip = driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_txtOrignlPgTranNo')
                ip.clear()
                ip.send_keys(ref)
                while 'number' not in driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_lblCaptcha').text:
                    driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_imgbtnCaptcha').click()
                    sleep(2)

                cap = ''
                
                t = time()
                flag = False 
                while cap=='':
                    cap = captcha_context()
                    if time()-t > 30 :
                        flag = True 
                        break 
                if flag:
                    driver.get("https://www.indiapost.gov.in/_layouts/15/DOP.Portal.Tracking/TrackConsignment.aspx")
                    continue
                driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_txtCaptcha').send_keys(cap)
                try:
                    driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_btnSearch').click()
                except :
                    pass
                t = time()
                flag = False
                while True:
                    
                    try:
                        btn = driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_btnTrackMore')
                        break
                    except:
                        pass
                    if time()-t > sleep_:
                        flag = True 
                        break
                
                if flag:
                    driver.get("https://www.indiapost.gov.in/_layouts/15/DOP.Portal.Tracking/TrackConsignment.aspx")
                    continue  
                try:
                    df.loc[i,'Delivery Report'] = str(driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[4]").text)
                    df.loc[i,'date']   = str(driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[1]").text)
                    df.loc[i,'time']  = str(driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[2]").text)
                    df.loc[i,'office'] = str(driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[3]").text)
                    if pdf_opt:
                        pdfs.append((driver.execute_cdp_cmd('Page.printToPDF',{})['data'],ref+'.pdf'))
                    btn.click()
                    df_view.dataframe(df)
                    
                    rt = str(datetime.timedelta(seconds=int(time()-rt))).split(':')
                    st.write(str(i)+') Record '+ref+' is Completd  -  '+rt[1]+':'+rt[2])
                    rt = 0
                    i += 1
                    count += 1
                except:
                    i-=1
                i+=1   
                sleep(2)
        except :
            pass
        ot = str(datetime.timedelta(seconds=int(time()-ot)))
        st.write('Total time :- '+ot)
    return df,pdfs
    

import streamlit as st

uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])
cols = st.columns(5)

# Create a file uploader widget
with cols[0]:
    start_ = st.text_input('Start at : ',placeholder='Index')
with cols[1]:
    end = st.text_input('End at : ',placeholder='Index')
with cols[2]:
    sleep_ = st.text_input('Limit : ',placeholder='Secounds')
with cols[3]:
    pdf_opt = st.checkbox("Generate PDF's")
with cols[4]:
    st.write()
    st.write()
    bt = st.button('START',help='Click to start the process')
# Check if a file was uploaded
if bt:
    if  uploaded_file is not None:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(uploaded_file)
        if len(list(df.columns)) != 6:
            st.error('ERROR!!! Invalid Excel Format')
        df.columns = ['Name','RPAD Barcode No ','date','time','office','Delivery Report']
        for i in ['Name','RPAD Barcode No ','date','time','office','Delivery Report']:
            df[i] = df[i].astype(str)

        if start_ == '' or not start_.isdigit():
            start_ = 1
        else:
            start_ = int(start_)
        if end == '' or not end.isdigit():
            end = len(df['RPAD Barcode No '])
        else:
            end = int(end)
        if sleep_ == '' or not sleep_.isdigit():
            sleep_ = 4
        else:
            sleep_ = int(sleep_)
        df,pdfs = start(df,start_,end,sleep_,pdf_opt)
        zip_data = io.BytesIO()
        with zipfile.ZipFile(zip_data, 'w') as zipf:
        # Add Excel file to the zip folder with a custom file name
            excel_file = io.BytesIO()
            with pd.ExcelWriter(excel_file, engine='xlsxwriter', mode='w') as writer:
                df.to_excel(writer, index=False)
            excel_file.seek(0)
            zipf.writestr('output.xlsx', excel_file.read())
    
    # Provide Excel content as binary data to the download_button
#         st.download_button(label="Download Excel", data=excel_content, file_name="output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        if pdf_opt:
            # Create a zip file in memory
            with zipfile.ZipFile(zip_data, 'a') as zipf:
                for pdf_data, pdf_name in pdfs:
                    zipf.writestr(pdf_name+'.pdf', b64decode(pdf_data))
            # Provide a download button for the zip file
        st.download_button(label='Download Files', data=zip_data.getvalue(), file_name='output.zip', mime='application/zip',help="Click to Download Excel File and PDF's")
        
    else:
        st.error('No file Selected!!!')
    