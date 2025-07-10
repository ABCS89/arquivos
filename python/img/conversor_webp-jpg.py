import os
from PIL import Image

# Pergunte pelo nome do arquivo (sem extensão)
nome_arquivo = input("Digite o nome do arquivo (sem extensão): ")

# Crie o caminho completo do arquivo
webp_path = f"{nome_arquivo}.webp"

# Verifique se o arquivo existe
print(f"Verificando a existência do arquivo: {webp_path}")
if not os.path.isfile(webp_path):
    print(f"O arquivo {webp_path} não foi encontrado.")
else:
    try:
        # Abra a imagem WEBP
        print(f"Abrindo o arquivo: {webp_path}")
        webp_image = Image.open(webp_path)
        print(f"Arquivo {webp_path} aberto com sucesso.")

        # Converta para JPG
        jpg_path = f"{nome_arquivo}.jpg"
        webp_image.convert('RGB').save(jpg_path, 'JPEG')
        print(f"Conversão concluída com sucesso! Arquivo salvo como {jpg_path}.")
    except Exception as e:
        print(f"Ocorreu um erro ao abrir ou converter a imagem: {e}")
