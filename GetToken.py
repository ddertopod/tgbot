import requests
import json
import os
from dotenv import load_dotenv
load_dotenv()
url = 'http://base.metallplace.ru:8080/login'
payload = {"username": os.getenv('P_User'), "password": os.getenv('P_Pass')}
res = requests.post(url = url,  data = json.dumps(payload))
data = json.loads(res.text)
token = data['token']
print(token)
headers = {'Authorization': '{}'.format(token)}
print(headers)