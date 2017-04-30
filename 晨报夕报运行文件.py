# -*- coding: utf-8 -*-
"""
Created on Fri Nov  4 11:28:21 2016

@author: JJJ
"""
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart  
from email.mime.application import MIMEApplication 
from WindPy import w 
from dailyReport import dailyReport, dailyReportPart
import time
import datetime
if __name__ == '__main__':
    
    start = time.clock();
    w.start();
    
    a = dailyReportPart(extraction_time = 1).get_data();
    b = dailyReportPart(extraction_time = 2).get_data();    
    c = dailyReport(a,b);        
    c.toExcel();
    end = time.clock();
    print ('run time %fs' % (end - start));
    today = datetime.date.today()
    year,month,day = today.year,today.month,today.day
    today = str(year) + '-' + str(month) + '-' + str(day)
    #send email
#    _user = 'shixisheng1@sinvofund.com'
#    _pwd = 'Jjj19910309'
#    #_to = ['986233620@qq.com']
#    _to = ['wuliang@sinvofund.com, dingping@sinvofund.com, yuli@sinvofund.com, shixisheng1@sinvofund.com']
#    msg = MIMEMultipart(); 
#    msg["Subject"] = "晨报夕报 "  + today;
#    msg["From"]    = _user;
#    msg["To"]      = ','.join(_to);    
#    part = MIMEText("尊敬的各位：\n  附件是今天的晨报夕报,请查收~\n  祝好~\n\n---------------------------------\n新沃基金固定收益部\n实习生   蒋进进\nTel:    18801900348\nEmail:   jiangjinjinyxt@sjtu.edu.cn");  
#    msg.attach(part)  ;
#    part = MIMEApplication(open('输出/晨报夕报 ' + today + '.xlsx','rb').read());  
#    part.add_header('Content-Disposition', 'attachment', filename="晨报夕报" + today + ".xlsx");  
#    msg.attach(part);
#    try:
#        s = smtplib.SMTP_SSL("smtp.exmail.qq.com", 465);
#        s.login(_user, _pwd);
#        s.sendmail(_user, _to, msg.as_string());
#        s.quit();
#        print ("Success!");
#    except:
#        print ("Falied");

    w.close();