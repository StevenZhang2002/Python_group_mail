import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
from tqdm import tqdm


class Email:
    def __init__(self):
        self.smtp = smtplib.SMTP()
        self.smtp.connect('smtp.163.com')
        self.email_body = MIMEMultipart('mixed')
        self.email_address = 'SonyIntern2024@163.com'
        self.email_auth_code = 'YYEOPJVMTMZAXABY'
        self.smtp.login(self.email_address, self.email_auth_code)



    def send_email(self, receiver_email, email_title, email_content, attachment_path, files):

        self.email_body['Subject'] = email_title
        self.email_body['From'] = self.email_address
        self.email_body['To'] = receiver_email

        for file in files:
            file_path = os.path.join(attachment_path,file)
            print(file_path)
            if os.path.isfile(file_path):
                print("is a file")
                att = MIMEText(open(file_path, 'rb').read(), 'base64', 'utf-8')
                att["Content-Type"] = 'application/octet-stream'
                att.add_header("Content-Disposition", "attachment", filename=("gbk", "", file))
                self.email_body.attach(att)

        text_plain = MIMEText(email_content, 'plain', 'utf-8')
        self.email_body.attach(text_plain)
        self.smtp.sendmail(self.email_address, receiver_email, self.email_body.as_string())


class Excel_Processor:
    def __init__(self,file_path):
        self.contact_info = dict()
        self.df = pd.read_excel(file_path,usecols=['导师','邮箱'])
        self.length = self.df.shape[0]
        i = 0
        while i<self.length:
            self.contact_info[self.df.iloc[i].values[0]]=self.df.iloc[i].values[1]
            i+=1
    def get_contact_info(self):
        return self.contact_info



if __name__ == '__main__':
    with open('message.txt','r') as f:
        msg = f.read()
    processor = Excel_Processor(r"test.xlsx")
    try:
        current_date = datetime.now().date()
        for _ in range(1):
            for contact in tqdm(processor.get_contact_info(), desc="Processing contacts"):
                email_send = Email()
                new_msg = msg.replace('PFNAME', contact)
                new_msg = new_msg.replace('TIMETIME', str(current_date))
                email_send.send_email(processor.contact_info.get(contact), f"status update: view your application", new_msg, r"", ["CV Bowen Zhang.pdf"])

    except Exception as e:
        print(f"error is {e}")

