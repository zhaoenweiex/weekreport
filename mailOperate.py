#!/usr/bin/python3
# _*_ coding: utf-8 _*_

import base64
import datetime
import time
import email
import os
import poplib
import smtplib
from datetime import timedelta
from email.header import decode_header,Header
#处理多种形态的邮件主体我们需要 MIMEMultipart 类
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
#发送字符串的邮件
from email.mime.text import MIMEText
from email.parser import Parser
from email.utils import parseaddr
from email.encoders import encode_base64
from email.mime.base import MIMEBase
import dateutil.parser


def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value


def guess_charset(msg):
    # 先从msg对象获取编码:
    charset = msg.get_charset()
    if charset is None:
        # 如果获取不到，再从Content-Type字段获取:
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset



def get_email_headers(msg):
    # 邮件的From, To, Subject存在于根对象上:
    headers = {}
    for header in ['From', 'To', 'Subject', 'Date']:
        value = msg.get(header, '')
        if value:
            if header == 'Date':
                headers['date'] = value
            if header == 'Subject':
                # 需要解码Subject字符串:
                subject = decode_str(value)
                headers['subject'] = subject
            else:
                # 需要解码Email地址:
                hdr, addr = parseaddr(value)
                name = decode_str(hdr)
                value = u'%s <%s>' % (name, addr)
                if header == 'From':
                    from_address = value
                    headers['from'] = from_address
                else:
                    to_address = value
                    headers['to'] = to_address
    content_type = msg.get_content_type()
    return headers


# indent用于缩进显示:
def get_email_content(message, base_save_path,dirPath,keyName):
    j = 0
    content = ''
    attachment_files = []
    for part in message.walk():
        j = j + 1
        file_name = part.get_filename()
        contentType = part.get_content_type()
        # 保存附件
        if file_name:  # Attachment
            # Decode filename
            h = email.header.Header(file_name)
            dh = email.header.decode_header(h)
            filename = decode_str(file_name)
            if filename.find(keyName) != -1:
                data = part.get_payload(decode=True)
                att_file = open(dirPath+'/' + filename, 'wb')
                attachment_files.append(filename)
                att_file.write(data)
                att_file.close()
        elif contentType == 'text/plain' or contentType == 'text/html':
            # 保存正文
            data = part.get_payload(decode=True)
            charset = guess_charset(part)
            if charset:
                charset = charset.strip().split(';')[0]
                print('charset:' + charset)
                data = data.decode(charset)
            content = data
    return content, attachment_files



def downloadReports(emailaddress,password,pop3_server,teamNumber,upTimeBounding,downTimeBounding,flagName,dirPath):
    # 连接到POP3服务器:
    server = poplib.POP3(pop3_server)
    # 可以打开或关闭调试信息:
    # server.set_debuglevel(1)
    # POP3服务器的欢迎文字:
    print(server.getwelcome())
    # 身份认证:
    server.user(emailaddress)
    server.pass_(password)
    # stat()返回邮件数量和占用空间:
    messagesCount, messagesSize = server.stat()
    print('messagesCount:', messagesCount)
    print('messagesSize:', messagesSize)
    # list()返回所有邮件的编号:
    resp, mails, octets = server.list()   
    # 获取最新10封邮件, 注意索引号从1开始:
    if teamNumber<1:
        length =  len(mails)
    else:
        length = teamNumber*2    
    for i in range(length):
        print('---------- 正在处理'+str(i)+'/'+str(length)+' ----------')
        resp, lines, octets = server.retr(len(mails) - i)
        # lines存储了邮件的原始文本的每一行,
        # 可以获得整个邮件的原始文本:
        strLines = []
        for line in lines:
            strInfo = line.decode()
            strLines.append(strInfo)
        msg_content = '\n'.join(strLines)
        # 把邮件内容解析为Message对象：
        msg = Parser().parsestr(msg_content)
        # 但是这个Message对象本身可能是一个MIMEMultipart对象，即包含嵌套的其他MIMEBase对象，
        # 嵌套可能还不止一层。所以我们要递归地打印出Message对象的层次结构：        
        base_save_path = '/media/markliu/Entertainment/email_attachments/'
        msg_headers = get_email_headers(msg)
        dateStr=msg_headers['date']
        if dateStr.find("(GMT")==-1:                
          receiveDate = dateutil.parser.parse(dateStr)
          now = datetime.datetime.now()
          this_week_start = now - timedelta(days=now.weekday())
          this_week_end = now + timedelta(days=6 - now.weekday())          
          beforeDistance=(receiveDate.replace(tzinfo=None) - this_week_start)
          afterDistance=this_week_end-(receiveDate.replace(tzinfo=None))
          if (beforeDistance.days < upTimeBounding) and (beforeDistance.days >downTimeBounding) and (afterDistance.days >downTimeBounding) and (afterDistance.days < upTimeBounding):
            content, attachment_files = get_email_content(msg, base_save_path,dirPath,flagName)
            print('subject:' + msg_headers['subject'])
            print('from_address:' + msg_headers['from'])
            print('to_address:' + msg_headers['to'])
            print('date:' + msg_headers['date'])
            print('content:' + content)
            if len(attachment_files) > 0:
                print('attachment_files: ' + str(attachment_files))
    server.quit()
    return 

def sendResults(fileNameArray,fromaddr,psw,serverAddress):
    timeStr= time.strftime("%Y%m%d", time.localtime())
    topic='软件二组'+timeStr+'周报汇总'
    sendResults(fileNameArray,fromaddr,fromaddr,psw,serverAddress,topic)

def sendResults(fileNameArray,fromaddr,toaddr,psw,serverAddress,topic):
    server = smtplib.SMTP(serverAddress)
    server.login(fromaddr,psw)
    m = MIMEMultipart()
    for file in fileNameArray:
        fileApart =  MIMEBase('application', 'octet-stream')
        fileApart.set_payload(open(file,'rb').read())
        fileApart.add_header('Content-Disposition', 'attachment', filename=Header(file, 'utf-8').encode())
        encode_base64(fileApart)
        m.attach(fileApart)   
    m['Subject'] = topic
    server.sendmail(fromaddr, toaddr, m.as_string())
    print('send success')
    server.quit()

if __name__ == '__main__':
    # 输入邮件地址, 口令和POP3服务器地址:
    emailaddress = '18622939753@163.com'
    # 注意使用开通POP，SMTP等的授权码
    password = '860124Ww'
    pop3_server = 'pop.163.com'
    teamNumber=7
    downloadReports(emailaddress,password,pop3_server,teamNumber)
