import requests
import json
import GetToken
r = requests.post(url = 'http://base.metallplace.ru:8080/getMaterialList', headers = GetToken.headers, data = json.dumps({})) 
data = r.json()
pretty = json.dumps(data, sort_keys=False, indent=4, ensure_ascii= False, separators=(',', ': '))
print(pretty)
with open('data.json', 'w', encoding= "utf-8") as f:
    json.dump(pretty, f, sort_keys=False, indent=4, ensure_ascii= False, separators=(',', ': '))