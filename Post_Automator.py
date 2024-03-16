#For Process
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from xlsxwriter import Workbook


from PIL import Image
from io import BytesIO
import easyocr
import sys

from base64 import b64decode
import pandas as pd
import requests
import os
from time import sleep


import zipfile
from io import BytesIO

# Function to generate ZIP file containing PDFs
def generate_zip(pdf_files, zip_file_path):
    with zipfile.ZipFile(zip_file_path, 'w') as zipf:
        for pdf_content, file_name in pdf_files:
            zipf.writestr(file_name, pdf_content.getvalue())

# Function to generate PDFs and return as BytesIO objects
def pdf_gen(pdfs, file_names=None):
    pdfs_zip = []    
    for i, file_name in zip(pdfs,file_names):
        pdf_content = BytesIO() 
        pdf_content.write(i)  
        pdf_content.seek(0) 
        pdfs_zip.append((pdf_content, file_name))  # Append (pdf_content, file_name) tuple to list
    
    generate_zip(pdfs_zip,'output.zip')

# Create a download button for the ZIP file


def start(df,start_,end,sleep_):
    reader = easyocr.Reader(['en'])
    driver = webdriver.Chrome()
    driver.get("https://www.indiapost.gov.in/_layouts/15/DOP.Portal.Tracking/TrackConsignment.aspx")
    
    def captcha_solve():
        link = driver.find_element(By.XPATH,"//div[@class = 'input-group']//img").get_attribute('src')
        response = requests.get(link)
        sleep(4)
        image = Image.open(BytesIO(response.content))
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
        if context == 'Enter the Secound number':
            return cap[1] if len(cap) >= 2 else ''
        if context == 'Enter the Third number':
            return cap[2] if len(cap) >= 3 else ''
        if context == 'Enter the Fourth number':
            return cap[3] if len(cap) >= 4 else ''
        if context == 'Enter the Fifth number':
            return cap[4] if len(cap) == 5 else ''
        if context == 'Evaluate the Expression':
            try:
                return eval(cap[:-1])
            except :
                return ''
        return cap
    
    pdfs = []
    
    refs = list(df['RPAD Barcode No '])
    l = int(end)  if end != '' else len(refs)
    i = int(start_)  if start_ != '' else 1
    sleep_ = int(sleep_) if sleep_ != '' or int(sleep_) > 4 else 4
    
    df = df[i-1:l]
    df.index = range(i,l+1)
    df_view = st.empty()
    df_view.dataframe(df)
    with st.spinner('Please wait..'):
        sleep(1)
    with st.status("Processing.....",expanded=True):
        while i<l:
            if str(refs[i])=='nan':
                i += 1
                continue 
            ip = driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_txtOrignlPgTranNo')
            ip.clear()
            ip.send_keys(refs[i])
            cap = ''
            c = 0
            while cap=='':
                cap = captcha_context()
            driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_txtCaptcha').send_keys(cap)
            try:
                driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_btnSearch').click()
                sleep(sleep_)
            except :
                if c==5:
                    df.loc[i,'Delivery Report'] = None
                    df.loc[i,'date'] = None
                    df.loc[i,'time'] = None
                    df.loc[i,'office'] = None
                    pdfs.append('')
                    i += 1
                    driver.get("https://www.indiapost.gov.in/_layouts/15/DOP.Portal.Tracking/TrackConsignment.aspx")
                    continue 
                i += 1
                sleep(1)
                c += 1
                continue
            try:
                btn = driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_btnTrackMore')                
                df.loc[i,'Delivery Report'] = str(driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[4]").text)
                df.loc[i,'date']   = str(driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[1]").text)
                df.loc[i,'time']  = str(driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[2]").text)
                df.loc[i,'office'] = str(driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[3]").text)
                sleep(4)
                btn.click()
                df_view.dataframe(df)
#                 pdfs.append(driver.execute_cdp_cmd('Page.printToPDF',{})['data'])
                st.write(str(i)+') Record '+refs[i]+' is Completd.')
                i += 1
                count += 1
            except:
                i-=1
            i+=1   
            sleep(1)
#     pdfs = pdf_gen(pdfs,refs)
    return df
    

import streamlit as st

uploaded_file = st.file_uploader("Upload the Excel file", type=["xlsx"])
cols = st.columns(4)

# Create a file uploader widget
with cols[0]:
    start_ = st.text_input('Start at : ',placeholder='Index')
with cols[1]:
    end = st.text_input('End at : ',placeholder='Index')
with cols[2]:
    sleep_ = st.text_input('Sleep Time : ',placeholder='Secounds')
with cols[3]:
    st.write()
    st.write()
    bt = st.button('Start the Process',help='Click to start the process')
# Check if a file was uploaded
if bt:
    if  uploaded_file is not None:
        # Read the Excel file into a pandas DataFrame
        df = pd.read_excel(uploaded_file)
        if len(list(df.columns)) != 6:
            st.error('ERROR!!! Invalid Excel Format')
        df.columns = ['Name','RPAD Barcode No ','date','time','office','Delivery Report']
        df = start(df,start_,end,sleep_)
        excel_file = BytesIO()
        with pd.ExcelWriter(excel_file, engine='xlsxwriter', mode='w',) as writer:
            df.to_excel(writer, index=False)
        excel_content = excel_file.getvalue()

    # Provide Excel content as binary data to the download_button
        st.download_button(
                        label="Download Excel",
                        data=excel_content,
                        file_name="output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    #     st.download_button(
    #                         label="Download PDFs",
    #                         data='output.zip',
    #                         file_name="pdfs.zip",
    #                         mime="application/zip"
    #                     )
    else:
        st.error('No file Selected!!!')
    

