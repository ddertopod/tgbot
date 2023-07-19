import requests
import json
import GetToken
payload = {"material_source_id": '1'}
r = requests.post(url = 'http://base.metallplace.ru:8080/getPropertyList', headers = GetToken.headers, data = json.dumps(payload)) 
data = r.json()
pretty = json.dumps(data, sort_keys=False, indent=4, ensure_ascii= False, separators=(',', ': '))
print(pretty)
