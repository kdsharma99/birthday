import pandas as pd
import datetime
import smtplib
import json
import urllib.request
import urllib.parse
with open("config.json",'r') as c:
    params=json.load(c)["params"]
def send_email(to,sub,msg):
    s=smtplib.SMTP("smtp.gmail.com",587)
    s.starttls()
    s.login(params["GMAIL_ID"],params["GMAIL_PSWD"])
    s.sendmail(params["GMAIL_ID"],to,f"Subject:{sub}\n\n{msg}")
    s.quit()
def send_sms(apikey, numbers, sender, message):
    paramss = {'apikey':'RZ3Q9fYKc/k-FzMebqVRq1F0i6t7Vxt8r6lvIG9FOp', 'numbers': numbers, 'message' : message, 'sender': sender}
    f = urllib.request.urlopen('https://api.textlocal.in/send/?'+urllib.parse.urlencode(paramss))
    return (f.read(), f.code)
if __name__=="__main__":
    df=pd.read_excel("data.xlsx")
    today=datetime.datetime.now().strftime("%d-%m")
    yearnow=datetime.datetime.now().strftime("%Y")
    writeInd=[]
    for index,item in df.iterrows():
        bday=item["Birthday"].strftime("%d-%m")
        if (today==bday) and yearnow not in str(item["Year"]):
            send_email(item["email"],"Happy Birthday",item["Message"])
            send_sms('apikey',item["phone"],"Kushal Sharma",item["Message"])
            writeInd.append(index)
    for i in writeInd:
        yr=df.loc[i,"Year"]
        df.loc[i,"Year"]=str(yr)+","+str(yearnow)
    df.to_excel("data.xlsx",index=False)