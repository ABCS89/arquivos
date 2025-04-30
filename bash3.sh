#!/bin/bash
# Atualizar o sistema
sudo apt update -y
sudo apt upgrade -y


# Instalar dependências básicas
sudo apt-get install -y python3 python3-pip python3-venv unzip

# Instalar AWS CLI v2
curl "https://awscli.amazonaws.com/awscli-exe-linux-x86_64.zip" -o "awscliv2.zip"
unzip awscliv2.zip
sudo ./aws/install
rm awscliv2.zip
rm -rf aws

sudo apt-get update
sudo apt-get install python3-venv

python3 -m venv myenv
source myenv/bin/activate

# Instalar bibliotecas Python necessárias
pip install pandas odfpy reportlab
