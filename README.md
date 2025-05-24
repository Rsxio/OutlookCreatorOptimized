# Outlook 自动创建与管理脚本使用说明

## 功能介绍

本脚本提供以下功能：

1. 自动创建outlook.com邮箱，随机生成邮箱名和密码
2. 随机生成大于18岁的个人信息
3. 自动绑定TOTP
4. 支持命令行传入SOCKS5代理并支持代理轮换
5. 将邮箱、密码和TOTP信息保存到文件中，支持查看和导出
6. 添加自动修改outlook密码的功能，支持批量处理已有账号

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 创建新账号

```bash
# 创建1个账号（默认使用无头模式，适合VPS无图形界面环境）
python outlook_creator.py create

# 创建多个账号
python outlook_creator.py create -c 5

# 使用多线程
python outlook_creator.py create -c 5 -t 3

# 使用SOCKS5代理
python outlook_creator.py create -p 127.0.0.1:1080

# 使用代理文件（每行一个代理）
python outlook_creator.py create -P proxies.txt

# 如果需要在有图形界面的环境下运行（不推荐）
python outlook_creator.py create --no-headless
```

### 修改账号密码

```bash
# 修改所有账号密码
python outlook_creator.py change

# 修改指定邮箱密码
python outlook_creator.py change -e example@outlook.com

# 使用多线程
python outlook_creator.py change -t 3

# 使用SOCKS5代理
python outlook_creator.py change -p 127.0.0.1:1080
```

### 导出账号信息

```bash
# 导出为文本格式（邮箱—-密码—-TOTP格式，默认）
python outlook_creator.py export -o accounts_export.txt

# 导出为CSV格式
python outlook_creator.py export -o accounts_export.csv --format csv
```

## 输出文件

- `outlook_accounts.csv`: 保存所有账号信息
- `totp_secrets.json`: 保存TOTP密钥信息
- `outlook_creator.log`: 日志文件
- 自动生成的导出文件: 创建或修改账号后会自动生成导出文件，格式为"邮箱—-密码—-TOTP Key"

## 在无图形界面VPS上使用

1. 脚本默认使用无头模式，适合在无图形界面的VPS上运行
2. 确保安装了Chrome浏览器和必要的依赖：

```bash
# 更新软件包
sudo apt update

# 安装Chrome浏览器
sudo apt install -y wget
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo apt install -y ./google-chrome-stable_current_amd64.deb

# 安装必要的依赖
sudo apt install -y xvfb libxi6 libgconf-2-4
```

3. 使用Xvfb可以在无图形界面的服务器上运行浏览器：

```bash
# 安装Xvfb
sudo apt install -y xvfb

# 使用Xvfb运行脚本
xvfb-run python outlook_creator.py create
```

## 在VSCode终端中使用

1. 在VSCode中打开项目文件夹
2. 确保本地安装了Chrome浏览器
3. 在终端中运行脚本，默认使用无头模式

## 注意事项

1. 脚本需要安装Chrome浏览器和对应版本的ChromeDriver
2. 首次运行时会自动下载ChromeDriver
3. 注册过程中可能需要处理验证码或其他验证步骤
4. 使用代理时请确保代理可用
5. 批量操作时建议使用多线程提高效率
6. 所有操作完成后，账号信息会以"邮箱—-密码—-TOTP Key"格式显示在控制台并保存到文件

## 常见问题

1. **ChromeDriver下载失败**：可以手动下载对应版本的ChromeDriver并放置在适当位置
2. **验证码处理**：当前版本需要手动处理验证码，未来版本可能会添加自动处理功能
3. **代理不可用**：请检查代理格式和可用性，确保代理支持SOCKS5协议
4. **无法在VPS上运行**：确保已安装所有必要依赖，并使用Xvfb或确保脚本使用无头模式
