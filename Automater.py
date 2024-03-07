#For UI
from tkinter import Tk,Label,Button,Entry,StringVar,PhotoImage
from tkinter.messagebox import showerror,askyesno,showinfo
from tkinter.filedialog import askopenfilename,askdirectory

#For Processor
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from PIL import Image
from io import BytesIO
import easyocr

from base64 import b64decode

import pandas as pd
import requests
import os
from time import sleep


class Post_Status_Extract_Update:
    def __init__(self):
        self.reader = easyocr.Reader(['en'])
        chrome_options = Options()
#         chrome_options.add_argument('--headless')
#         chrome_options.add_argument('--disable-gpu')
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.get("https://www.indiapost.gov.in/_layouts/15/DOP.Portal.Tracking/TrackConsignment.aspx")
        
    def pdf_gen(self,ref_id):
        pdf = self.driver.execute_cdp_cmd('Page.printToPDF',{})
        with open(self.op_path+'/PDFS/'+ref_id+'.pdf','wb') as f:
            f.write(b64decode(pdf['data']))
        
    def captcha_solve(self):
        link = self.driver.find_element(By.XPATH,"//div[@class = 'input-group']//img").get_attribute('src')
        
        response = requests.get(link)
        sleep(4)
        image = Image.open(BytesIO(response.content))
        image = image.convert('RGB')
        image.save('captcha.jpg', 'JPEG')
        
        try:
            result = self.reader.readtext('captcha.jpg')[0][1].replace(' ','')
        except :
            result=''
        os.remove('captcha.jpg')
        return result
    
    
    
    def captcha_context(self):
        cap = self.captcha_solve()
        if cap == '':
            self.driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_imgbtnCaptcha').click()
            return ''
        context = self.driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_lblCaptcha').text
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
    
    def extract(self,refs):
        d = {'status':[],'dod':[]}
        l = len(refs)
        i = 0
        while i<l:
            if str(refs[i])=='nan':
                i += 1
                d['dod'].append('')
                d['status'].append('')
                continue 
            ip = self.driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_txtOrignlPgTranNo')
            ip.clear()
            ip.send_keys(refs[i])
            cap = ''
            while cap=='':
                cap = self.captcha_context()
            self.driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_ucCaptcha1_txtCaptcha').send_keys(cap)
            try:
                self.driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_btnSearch').click()
                sleep(4)
            except :
                Label(self.w,text=' '*14).grid(row=self.i,column=0)
                Label(self.w,text='Record '+refs[i]+' Failed! Trying Again').grid(row=self.i,column=1)
                self.i += 1
                self.w.update()
                sleep(1)
                continue
            try:
                btn = self.driver.find_element(By.ID,'ctl00_PlaceHolderMain_ucNewLegacyControl_btnTrackMore')
                self.pdf_gen(refs[i])
                d['dod'].append(self.driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[1]").text)
                d['status'].append(self.driver.find_element(By.XPATH,"//table[@class = 'responsivetable MailArticleEvntOER']//tbody//tr[2]//td[4]").text)
                sleep(4)
                btn.click()
                Label(self.w,text=' '*14).grid(row=self.i,column=0)
                Label(self.w,text=str(self.refs)+'. Record  '+refs[i]+' is Completed',font=("Arial", 10)).grid(row=self.i,column=1)
                self.i += 1
                self.refs += 1
                self.w.update()
            except:
                i-=1
            i+=1   
            sleep(1)
        return d
        
        
    def pre_processing(self,fn):
        df = pd.read_excel(fn) 
        df.columns = ['S.No.', 'Ref.No', 'Loan No', 'Customer Name', 'Consignment No.1', 'Status.1', 'Date of delivery.1', 'Co.Applicant', 'Consignment No.2', 'Status.2', 'Date of delivery.2', 'Guarenter', 'Consignment No.3', 'Status.3', 'Date of delivery.3', 'Product', 'Notice Sent Date', 'Notice Received/Returned ', 'Bill', 'Claimed/Not']
        return df
    
    def data_handling(self,df,d1,d2,d3):
        df['Status.1'] = d1['status']
        df['Date of delivery.1'] = d1['dod']
        df['Status.2'] = d2['status']
        df['Date of delivery.2'] = d2['dod']
        df['Status.3'] = d3['status']
        df['Date of delivery.3'] = d3['dod']
        return df
        
    def start(self,file,op_path,opfile):
        self.op_path = op_path
        try:
            os.mkdir(op_path+'/PDFS')
        except :
            pass
        self.w = Tk()
        self.w.state('zoomed')
        self.w.iconphoto(False ,PhotoImage(file='Icon\\download.png'))
        self.w.title('Activity Log')
        Label(self.w,text='Activity Log :-',font=("Arial", 12)).grid(row=0,column=0)
        Label(self.w,text='Process Started',font=("Arial",10)).grid(row=0,column=1)
        self.w.update()
        df = self.pre_processing(file)
        self.i = 1
        self.refs = 1
        d1 = self.extract(list(df['Consignment No.1']))
        d2 = self.extract(list(df['Consignment No.2']))
        d3 = self.extract(list(df['Consignment No.3']))
        df = self.data_handling(df,d1,d2,d3)
        df.to_excel(op_path+'/'+opfile,index=False)
        self.driver.quit()
        showinfo('Info','Proccess Completed Successfully')
        self.w.destroy()
        

class UI:
    def start_process(self,w,d):
        if self.fn[-4:] == 'xlsx':
            self.op_fn = d.get()
            if self.op_fn=='':
                self.op_fn = self.fn.split('/')[-1].replace('.xlsx','')
                
            w.destroy()
            while True :
                try:
                    obj = Post_Status_Extract_Update()
                    obj.start(self.fn,self.op,self.op_fn+'.xlsx')
                    del obj
                    break 
                except:
                    obj.w.destroy()
                    del obj
                    if not askyesno('ERROR','Process Failed!!! Do you want to try Again'):
                        break 
        else:
            showerror('ERROR','Selected File is not an excel file')
            
    def file_select(self,w,s):
        self.fn = askopenfilename()
        s.set(self.fn)
        
    def op_select(self,w,d):
        self.op = askdirectory()
        Label(w,text=' '*10).grid(row=2,column=0)
        
        Button(w,text='Start',width=20,command=lambda:self.start_process(w,d)).grid(row=2,column=1)

    def select_page(self):
        w = Tk()
        w.title('Automater')
        w.iconphoto(False ,PhotoImage(file='Icon\\download.png'))
#         w.geometry('600x35')
        s = StringVar()
        Label(w,text='Select the Excel File :').grid(row=0,column=0)
        e = Entry(w,width=50,textvariable=s,background='#ffffff',fg='#000000')
        e.config(state='disabled')
        e.grid(row=0,column=1)
        Button(w,text=' Choose File ',command=lambda :self.file_select(w,s)).grid(row=0,column=2)
        d = StringVar()
        Label(w,text='Enter output File Name:').grid(row=1,column=0)
        Entry(w,width=50,textvariable=d,background='#ffffff',fg='#000000').grid(row=1,column=1)
        Button(w,text='Choose Folder',command=lambda :self.op_select(w,d)).grid(row=1,column=2)
        w.mainloop()
    
UI().select_page()