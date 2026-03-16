import requests

public_key = "project_public_ce05b6f0aa522688fef673e30f9bc173_VdH2w681cbdc456614c52728185d67824f7cd"

# 1 - criar task
url = "https://api.ilovepdf.com/v1/start/compress"

response = requests.post(url, json={
    "public_key": public_key
})

data = response.json()

server = data["server"]
task = data["task"]

# 2 - enviar arquivo
files = {
    "file": open("*frequencia_secretaria.pdf", "rb")
    # "file": open("*relatorio_rh*.pdf", "rb")
}

upload_url = f"https://{server}/v1/upload"

upload = requests.post(upload_url, data={
    "task": task
}, files=files)

# 3 - processar
process_url = f"https://{server}/v1/process"

process = requests.post(process_url, data={
    "task": task
})

# 4 - baixar resultado
download_url = f"https://{server}/v1/download/{task}"

pdf = requests.get(download_url)

with open("arquivo_comprimido.pdf", "wb") as f:
    f.write(pdf.content)

print("PDF processado com sucesso!")