import pandas as pd
import numpy as np
import xlrd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date,timedelta
import datetime
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
from email.mime.image import MIMEImage
import os
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib.pyplot import figure
import calendar
from datetime import datetime
import pandas as pd
import pandas as pd
import numpy as np
import xlrd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import io
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import numpy as np
from email.mime.image import MIMEImage
import os
import matplotlib.pyplot as plt
from matplotlib import rc
from matplotlib.pyplot import figure
from datetime import date,timedelta
import datetime
from dateutil.relativedelta import relativedelta
import openpyxl

server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
msg['To']="iggeun.kwon@lge.com, joonseok.ahn@lge.com, soonan.park@lge.com, jinhong.min@lge.com, chris.cole@lge.com, jiyoon1.heo@lge.com, seungjae.cho@lge.com, eunbi1.yoon@lge.com,remoun.abdo@lge.com,min1.park@lge.com, russell.wilson@lge.com, alfonza.hall@lge.com"

#Subject 꾸미기
today=date.today()
today=today.strftime('%m/%d')
msg['Subject']='Daily Service for CK 5.0_'+today

########### New Services 만들기 ##############
# 파일 불러오기
svc_data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/M+3 task/200 Top Loader/1_Daily SVC Report/5_CK 5.0 (WT7150).xlsx',sheet_name='SVC')
    
# 오늘 날짜 추출
today=date.today()
today=today.strftime('%Y-%m-%d') ##today의 형식을 바꾸겠다, string 변환

# 날짜 해당 svc data 뽑기
today_svc_data=svc_data[svc_data['Report_Date'] == today]
print(today_svc_data)

# 인덱스 1부터 맞추기
today_svc_data.index = np.arange(1, len(today_svc_data) + 1)

# New Services
n=1
New_Services=""
while n <= len(today_svc_data):
    New_Services=New_Services + str(n) +") "+str(today_svc_data["Symptoms"][n])+" - "+str(today_svc_data["detail"][n])+"\n   "+str(today_svc_data["RCPT_NO_ORD_NO"][n])+" /// "+str(today_svc_data["SERIAL_NO 1"][n])+"\n"
    n=n+1
# New Service Text
print(New_Services)


########### SVC 만들기 ############## 
today_svc_data=svc_data[svc_data['Report_Date'] == today]
Increase_SVC=len(today_svc_data)
Today_SVC=len(svc_data)
Yesterday_SVC=Today_SVC-Increase_SVC
# SVC Text
SVC="Service Status : "+ str(Yesterday_SVC) + " → " + str(Today_SVC) +" ("+str(Increase_SVC)+"↑)"
print(SVC)

########### Week 추출하기 #######
year, week_num, day_of_week=date.today().isocalendar()
week_num=week_num+11
Weeks='W'+str(week_num)

########### Sales 만들기 #################
##### Today Sales
# 파일 불러오기
today_sales_data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/M+3 task/200 Top Loader/1_Daily SVC Report/5_CK 5.0 (WT7150).xlsx',sheet_name='GQIS')

# 데이터 정리
today_sales_data=today_sales_data.drop(['Unnamed: 0','Week',week_num],axis=1)
today_sales_data=today_sales_data.T
today_sales_data=today_sales_data[[10]]
today_sales_data.loc['col_sum', :] = today_sales_data.sum() # 세로로 합계
Today_Sales=today_sales_data.at["col_sum",10]

today_sales={'0':[Today_Sales]}
today_sales=pd.DataFrame(today_sales)
#today_sales.to_excel('3.xlsx')
today_sales.to_excel('//US-SO11-NA08765/R&D Secrets/M+3 task/200 Top Loader/1_Daily SVC Report/CK 5.0 Semi Sales/'+today+'.xlsx')


#### Yesterday Sales
today=date.today()
today_weekday=datetime.datetime.today().weekday()

if today_weekday==0: # 월요일
    yesterday=today-timedelta(days=3)
else:
    yesterday=today-timedelta(days=1)
    
yesterday=yesterday.strftime('%Y-%m-%d')
yesterday_sales=pd.read_excel('//US-SO11-NA08765/R&D Secrets/M+3 task/200 Top Loader/1_Daily SVC Report/CK 5.0 Semi Sales/'+yesterday+'.xlsx')
yesterday_sales=yesterday_sales.drop("Unnamed: 0",axis=1)
Yesterday_Sales=yesterday_sales.at[0,'0']

#### Save today Sales
today=today.strftime('%Y-%m-%d')
yesterday_sales=pd.read_excel('//US-SO11-NA08765/R&D Secrets/M+3 task/200 Top Loader/1_Daily SVC Report/CK 5.0 Semi Sales/'+today+'.xlsx')

#Sales Text
Sales="Sales Status : "+ str(int(Yesterday_Sales)) + " → " + str(int(Today_Sales)) +" ("+str(int(Today_Sales-Yesterday_Sales))+"↑)"
print(Sales)


####################### FDR 만들기 #####################
Today_FDR=round(Today_SVC*100/Today_Sales,2)
Yesterday_FDR=round(Yesterday_SVC*100/Yesterday_Sales,2)
print(Yesterday_FDR)
print(Today_FDR)

###################### Target 만들기 ###############################
# goal 파일 불러오기
goal_data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/M+3 task/200 Top Loader/1_Daily SVC Report/5_CK 5.0 (WT7150).xlsx',sheet_name='FDR')
goal_data.index=goal_data['Unnamed: 0']
goal_data=goal_data.drop(['Unnamed: 0'],axis=1)
goal_data=goal_data.T
goal_data.index=goal_data['Week']
Target=goal_data.at[Weeks,"Target"]


# FDR status 추출
FDR="FDR Status : "+str(Yesterday_FDR)+  " → " + str(Today_FDR) + " % ("+str(Weeks)+" Target "+str(round(Target,2))+"%)"
print(FDR)

#################### Pivot chart 만들기 ###############################
ea= svc_data["Symptoms"].value_counts(dropna=True,sort=True).to_frame()
ea.columns=["EA"]
total=pd.DataFrame([Today_SVC])
total.columns=["EA"]
total.index=["TOTAL"]
ea=pd.concat([ea,total],axis=0)


ppm=pd.DataFrame()
ppm["PPM"]=round(ea["EA"]*1000000/Today_Sales,0)
ppm=ppm.astype(int)

#table 합
Table=pd.concat([ea,ppm],axis=1)

###########################설정
fig, ax = plt.subplots(1,3)
fig = plt.figure(constrained_layout = True, figsize=(14,4))
gs = fig.add_gridspec(4,14)
ax[0] = fig.add_subplot(gs[0:4,0:3])
ax[1] = fig.add_subplot(gs[0:4,3:10])
ax[2] = fig.add_subplot(gs[0:4,10:14])
ax[2].set_axis_off()
###########################pivot chart matplotlib######################33
ax[0].set_axis_off()
col_labels=Table.columns
#row_colours=['#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E','#E6FA9E']
#'#CDCACF'
SVC_table=ax[0].table(cellText=Table.values, rowLabels=Table.index, colLabels=col_labels, loc='center', colColours=['#A9F1F1','#A9F1F1'])

SVC_table.auto_set_font_size(False)
SVC_table.set_fontsize(10)
SVC_table.auto_set_column_width(col=list(range(len(col_labels))))
ax[0].set_title('Service Overview',  pad=0.05, fontsize=11,x=0.2,y=0.97)


##############################################그래프 만들기##########################################################3
# 달 이름 정하기
# 매달 1일에는 마지막 달 마지막 날이 추출되는 조건 추가
today=date.today()
k=today.strftime('%d')
print(k)
if k=='01':
    today=today-relativedelta(days=1)
else:
    print("not the first day of month")
    
Data0M=today.strftime('%Y-%m')
date0M_name=today.strftime('%y.%m')

date1M_name=today-relativedelta(months=1)
Data1M=date1M_name.strftime('%Y-%m')
date1M_name=date1M_name.strftime('%y.%m')

date2M_name=today-relativedelta(months=2)
Data2M=date2M_name.strftime('%Y-%m')
date2M_name=date2M_name.strftime('%y.%m')

data=pd.read_excel('//US-SO11-NA08765/R&D Secrets/M+3 task/200 Top Loader/1_Daily SVC Report/5_CK 5.0 (WT7150).xlsx',sheet_name='GQIS')
data.index=data['Week']
data=data.T
data=data[['PRODUCT_GROUP','Total Sum of DAILYSVCCNT','Total Sum of DAILYSALESQTY']]
data.columns=['Date','SVC','Sales']
data=data.reset_index()
data=data.drop([0,1,2],axis=0)
data=data.drop(['index'],axis=1)

# 파일 정리하
data=data.dropna()
print(data)
data2M=data[data['Date'].str.contains(Data2M)]
data2M=data2M.T
data2M.columns=data2M.loc['Date'].str[8:10]
data2M=data2M.drop(['Date'],axis=0)
data2M.index=['SVC_'+date2M_name,'Sales_'+date2M_name]

data1M=data[data['Date'].str.contains(Data1M)]
data1M=data1M.T
data1M.columns=data1M.loc['Date'].str[8:10]
data1M=data1M.drop(['Date'],axis=0)
data1M.index=['SVC_'+date1M_name,'Sales_'+date1M_name]


data0M=data[data['Date'].str.contains(Data0M)]
data0M=data0M.T
data0M.columns=data0M.loc['Date'].str[8:10]
data0M=data0M.drop(['Date'],axis=0)
data0M.index=['SVC_'+date0M_name,'Sales_'+date0M_name]


data2M=data2M.T
data2M.index=data2M.index.astype(int)
data2M=data2M.cumsum()
data1M=data1M.T
data1M.index=data1M.index.astype(int)
data1M=data1M.cumsum()
data0M=data0M.T
data0M.index=data0M.index.astype(int)
data0M=data0M.cumsum()


# 그래프 그리기 위해 데이터 합치기
result=pd.concat([data2M,data1M,data0M],axis=1)
result.columns=['SVC_'+date2M_name,'Sales_'+date2M_name,
                'SVC_'+date1M_name,'Sales_'+date1M_name,
                'SVC_'+date0M_name,'Sales_'+date0M_name]
print(result)

# 그래프 데이터 넣기
axF1=result[['Sales_'+date2M_name,'Sales_'+date1M_name,'Sales_'+date0M_name]].plot(kind='bar', use_index=True, color=['#E0C8F7','#5997F9','#F3BA0A'],ax=ax[1]) # Sales 
axF1.set_ylabel('Sales',color='gray')
axF2 = axF1.twinx()
axF2.plot(result[['SVC_'+date2M_name]],linestyle='-', linewidth=1.0,color='green',label='SVC_'+date2M_name) # SVC
axF2.plot(result[['SVC_'+date1M_name]],linestyle='-', linewidth=1.0,color='black',label='SVC_'+date1M_name) # SVC
axF2.plot(result[['SVC_'+date0M_name]], linestyle='-', marker='o', linewidth=2.0,color='red',label='SVC_'+date0M_name) # SVC

axF2.set_ylabel('SVC',color='gray')

#그래프 UI
axF1.set_title("Service & Sales Monitoring",fontsize=11)
axF1.set_xlabel("Date",color='gray')
axF1.set_xticklabels(result.index,rotation=0)
axF1.set_xlim=axF2.set_xlim
axF1.set_ylim(0,20000)
#axF2.set_ylim(0,50)
axF2.set_xlim(0.5,31.5)
axF1.legend(loc='upper left')
axF2.legend(loc='upper left',bbox_to_anchor=(0,0.78))
plt.savefig('ck5.0.png')
 

######################## Body 꾸미기 
## 내용 구성
text0='This is DX activities from LGEUS R&D Team\nPerson in charge: LGEUS R&D Team Eunbi Yoon\n\n'
text1='Dear All,\n\nThis is the report of daily new model monitoring for CK 5.0 Semi Tub Model (WT7150CW).\n* It is based on what happened yesterday.\n\n\n[Service Overview]'
text2=FDR
text3=SVC
text4=Sales
text5='\n\n[Detailed Review]'
#text6='[New Services Review]'
text7=New_Services
textblank='\n\n'

## 본문에 첨부
msg.attach(MIMEText(text0,'plain'))
msg.attach(MIMEText(text1,'plain'))
msg.attach(MIMEText(text2,'plain'))
msg.attach(MIMEText(text3,'plain'))
msg.attach(MIMEText(text4))
msg.attach(MIMEText(text5,'plain'))

#첨부 파일1
with open('ck5.0.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('ck5.0.png'))
msg.attach(image)

#msg.attach(MIMEText(text6,'plain'))
msg.attach(MIMEText(text7,'plain'))

with open('sign.png', 'rb') as f:
        img_data = f.read()
image = MIMEImage(img_data, name=os.path.basename('sign.png'))
msg.attach(image)

#######################메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")
