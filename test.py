import requests
URL = ""
response = requests.get(URL)
open("instagram.ico", "wb").write(response.content)