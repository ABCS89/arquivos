from PIL import Image

# Abra a imagem WEBP
webp_image = Image.open('rebel.webp')

# Converta para JPG
webp_image.convert('RGB').save('imagem_convertida.jpg', 'JPEG')

print("Conversão concluída com sucesso!")
