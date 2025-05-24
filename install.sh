#!/bin/bash

# 安装脚本：适用于Ubuntu 22.04及更高版本
# 此脚本安装运行Outlook Creator所需的最小依赖集

# 更新软件包
echo "正在更新软件包..."
sudo apt update

# 安装Python和pip
echo "正在安装Python和pip..."
sudo apt install -y python3 python3-pip

# 安装Chrome浏览器
echo "正在安装Chrome浏览器..."
sudo apt install -y wget
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo apt install -y ./google-chrome-stable_current_amd64.deb
rm ./google-chrome-stable_current_amd64.deb

# 安装Xvfb和其他必要依赖
echo "正在安装Xvfb和其他必要依赖..."
sudo apt install -y xvfb

# 安装Python依赖
echo "正在安装Python依赖..."
pip3 install -r requirements.txt

echo "安装完成！"
echo "您可以使用以下命令运行脚本："
echo "python3 outlook_creator.py create"
echo "或使用Xvfb运行："
echo "xvfb-run python3 outlook_creator.py create"
