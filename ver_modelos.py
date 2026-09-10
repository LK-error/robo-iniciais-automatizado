import requests

chave = "AQ.Ab8RN6KEBhtONf32G73Ml3Zf9AnK6W4AVKxIZezJiULDaMJRJQ"
url = f"https://generativelanguage.googleapis.com/v1/models?key={chave}"

resposta = requests.get(url).json()
for modelo in resposta.get("models", []):
    print(modelo["name"])

