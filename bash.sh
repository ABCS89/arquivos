#!/bin/bash
# Update and upgrade the system
sudo apt update
sudo apt upgrade -y

# Install wget
sudo apt install -y wget

# Download and install Go
GO_VERSION="1.20.3"
wget https://golang.org/dl/go${GO_VERSION}.linux-amd64.tar.gz
sudo tar -C /usr/local -xzf go${GO_VERSION}.linux-amd64.tar.gz

# Add Go to the PATH
echo "export PATH=\$PATH:/usr/local/go/bin" >> ~/.profile
source ~/.profile

# Verify the installation
go version
