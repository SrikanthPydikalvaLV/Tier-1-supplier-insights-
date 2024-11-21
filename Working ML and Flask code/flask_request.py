import requests

url = "http://127.0.0.1:5000/predict"
file_path = "integratedreport2023e.pdf"

with open(file_path, 'rb') as file:
    response = requests.post(url, files={'file': file})

print(response.json())