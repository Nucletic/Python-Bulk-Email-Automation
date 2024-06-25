import requests
import time
ReciverFile = ['santosh.aeg@gmail.com', 'kuntalpriyanshu608@gmail.com',
               'sales@yatiglobalsolutions.com', 'sfsf@ddfg.com', 'fsdfgsdf@dfgdf']
for i in ReciverFile:
    response = requests.get('https://isitarealemail.com/api/email/validate',
                            params={'email': i})
    status = response.json()['status']
    time.sleep(2)
    if status == "invalid":
        print("invalid email ", i)
        

