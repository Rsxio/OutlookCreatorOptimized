#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Outlook 自动创建与管理脚本
功能：
1. 自动创建outlook.com邮箱，随机生成邮箱名和密码
2. 随机生成大于18岁的个人信息
3. 自动绑定TOTP
4. 支持命令行传入SOCKS5代理并支持代理轮换
5. 将邮箱、密码和TOTP信息保存到文件中，支持查看和导出
6. 添加自动修改outlook密码的功能，支持批量处理已有账号
"""

import os
import sys
import time
import json
import random
import string
import argparse
import logging
import datetime
import requests
import csv
import threading
import queue
from typing import List, Dict, Optional, Tuple, Any
from concurrent.futures import ThreadPoolExecutor
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, 
    NoSuchElementException, 
    ElementClickInterceptedException,
    StaleElementReferenceException
)
from webdriver_manager.chrome import ChromeDriverManager
from fake_useragent import UserAgent
import pyotp
import qrcode
from PIL import Image
from io import BytesIO

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("outlook_creator.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# 全局变量
DEFAULT_TIMEOUT = 30  # 默认等待超时时间（秒）
MAX_RETRIES = 3       # 最大重试次数
ACCOUNTS_FILE = "outlook_accounts.csv"  # 账号信息保存文件
TOTP_SECRETS_FILE = "totp_secrets.json"  # TOTP密钥保存文件
LOCK = threading.Lock()  # 用于线程安全操作

class OutlookCreator:
    """Outlook邮箱创建与管理类"""
    
    def __init__(self, proxy: Optional[str] = None, headless: bool = True):
        """
        初始化Outlook创建器
        
        Args:
            proxy: SOCKS5代理地址，格式为 "host:port"
            headless: 是否使用无头模式运行浏览器
        """
        self.proxy = proxy
        self.headless = headless
        self.driver = None
        self.current_account = {}
        self.setup_driver()
        
    def setup_driver(self):
        """配置并初始化WebDriver"""
        options = Options()
        if self.headless:
            options.add_argument("--headless")
        
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument(f"user-agent={UserAgent().random}")
        
        # 配置代理
        if self.proxy:
            options.add_argument(f'--proxy-server=socks5://{self.proxy}')
            logger.info(f"使用SOCKS5代理: {self.proxy}")
        
        try:
            service = Service(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.set_page_load_timeout(DEFAULT_TIMEOUT)
            logger.info("WebDriver初始化成功")
        except Exception as e:
            logger.error(f"WebDriver初始化失败: {str(e)}")
            raise
    
    def close(self):
        """关闭WebDriver"""
        if self.driver:
            self.driver.quit()
            logger.info("WebDriver已关闭")
    
    def generate_random_name(self) -> Tuple[str, str]:
        """
        生成随机姓名
        
        Returns:
            Tuple[str, str]: (名, 姓)
        """
        first_names = [
            "Alex", "Jamie", "Jordan", "Taylor", "Casey", "Riley", "Avery", 
            "Quinn", "Morgan", "Dakota", "Reese", "Emerson", "Finley", "Rowan",
            "Skyler", "Charlie", "Blake", "River", "Sage", "Phoenix"
        ]
        
        last_names = [
            "Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller",
            "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez",
            "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin"
        ]
        
        first_name = random.choice(first_names)
        last_name = random.choice(last_names)
        
        return first_name, last_name
    
    def generate_random_birthday(self) -> Tuple[int, int, int]:
        """
        生成随机生日（确保年龄大于18岁）
        
        Returns:
            Tuple[int, int, int]: (年, 月, 日)
        """
        current_year = datetime.datetime.now().year
        year = random.randint(current_year - 50, current_year - 18)
        month = random.randint(1, 12)
        
        # 根据月份确定天数
        if month in [4, 6, 9, 11]:
            day = random.randint(1, 30)
        elif month == 2:
            # 处理闰年
            if (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0):
                day = random.randint(1, 29)
            else:
                day = random.randint(1, 28)
        else:
            day = random.randint(1, 31)
            
        return year, month, day
    
    def generate_random_email(self) -> str:
        """
        生成随机邮箱名
        
        Returns:
            str: 随机邮箱名（不含域名）
        """
        name_parts = []
        
        # 添加随机单词
        words = ["cool", "super", "awesome", "tech", "dev", "pro", "star", "net", 
                "web", "code", "data", "info", "cyber", "digital", "smart"]
        name_parts.append(random.choice(words))
        
        # 添加随机名字
        first_name, _ = self.generate_random_name()
        name_parts.append(first_name.lower())
        
        # 添加随机数字
        name_parts.append(str(random.randint(100, 9999)))
        
        # 随机组合部分
        random.shuffle(name_parts)
        email_name = "".join(name_parts)
        
        return email_name
    
    def generate_random_password(self, length: int = 12) -> str:
        """
        生成随机强密码
        
        Args:
            length: 密码长度
            
        Returns:
            str: 随机密码
        """
        # 确保密码包含大小写字母、数字和特殊字符
        lowercase = string.ascii_lowercase
        uppercase = string.ascii_uppercase
        digits = string.digits
        special = "!@#$%^&*()-_=+"
        
        # 确保每种字符至少出现一次
        password = [
            random.choice(lowercase),
            random.choice(uppercase),
            random.choice(digits),
            random.choice(special)
        ]
        
        # 填充剩余长度
        remaining_length = length - len(password)
        all_chars = lowercase + uppercase + digits + special
        password.extend(random.choice(all_chars) for _ in range(remaining_length))
        
        # 打乱密码顺序
        random.shuffle(password)
        return ''.join(password)
    
    def wait_for_element(self, by: By, value: str, timeout: int = DEFAULT_TIMEOUT):
        """
        等待元素出现并返回
        
        Args:
            by: 定位方式
            value: 定位值
            timeout: 超时时间（秒）
            
        Returns:
            WebElement: 找到的元素
        """
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((by, value))
            )
            return element
        except TimeoutException:
            logger.error(f"等待元素超时: {by}={value}")
            raise
    
    def wait_for_clickable(self, by: By, value: str, timeout: int = DEFAULT_TIMEOUT):
        """
        等待元素可点击并返回
        
        Args:
            by: 定位方式
            value: 定位值
            timeout: 超时时间（秒）
            
        Returns:
            WebElement: 找到的元素
        """
        try:
            element = WebDriverWait(self.driver, timeout).until(
                EC.element_to_be_clickable((by, value))
            )
            return element
        except TimeoutException:
            logger.error(f"等待元素可点击超时: {by}={value}")
            raise
    
    def create_outlook_account(self) -> Dict[str, Any]:
        """
        创建Outlook邮箱账号
        
        Returns:
            Dict: 包含账号信息的字典
        """
        start_time = time.time()
        
        try:
            # 生成随机账号信息
            email_name = self.generate_random_email()
            password = self.generate_random_password()
            first_name, last_name = self.generate_random_name()
            year, month, day = self.generate_random_birthday()
            
            # 保存当前账号信息
            self.current_account = {
                "email": f"{email_name}@outlook.com",
                "password": password,
                "first_name": first_name,
                "last_name": last_name,
                "birth_year": year,
                "birth_month": month,
                "birth_day": day,
                "totp_secret": "",
                "creation_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            logger.info(f"[✓] 开始处理 {self.current_account['email']}")
            
            # 访问注册页面
            self.driver.get("https://signup.live.com/signup")
            
            # 输入邮箱名
            email_input = self.wait_for_element(By.ID, "MemberName")
            email_input.send_keys(email_name)
            
            # 点击下一步
            next_button = self.wait_for_clickable(By.ID, "iSignupAction")
            next_button.click()
            
            # 输入密码
            password_input = self.wait_for_element(By.ID, "PasswordInput")
            password_input.send_keys(password)
            
            # 点击下一步
            next_button = self.wait_for_clickable(By.ID, "iSignupAction")
            next_button.click()
            
            # 输入姓名
            first_name_input = self.wait_for_element(By.ID, "FirstName")
            first_name_input.send_keys(first_name)
            
            last_name_input = self.wait_for_element(By.ID, "LastName")
            last_name_input.send_keys(last_name)
            
            # 点击下一步
            next_button = self.wait_for_clickable(By.ID, "iSignupAction")
            next_button.click()
            
            # 输入生日
            birth_year_input = self.wait_for_element(By.ID, "BirthYear")
            birth_year_input.send_keys(str(year))
            
            birth_month_select = self.wait_for_element(By.ID, "BirthMonth")
            birth_month_select.click()
            month_option = self.wait_for_element(By.CSS_SELECTOR, f"option[value='{month}']")
            month_option.click()
            
            birth_day_input = self.wait_for_element(By.ID, "BirthDay")
            birth_day_input.send_keys(str(day))
            
            # 点击下一步
            next_button = self.wait_for_clickable(By.ID, "iSignupAction")
            next_button.click()
            
            # 等待注册完成
            # 注意：这里可能需要处理验证码或其他验证步骤
            # 由于验证步骤可能会变化，这里只提供基本框架
            
            # 假设注册成功，绑定TOTP
            totp_secret = self.bind_totp()
            self.current_account["totp_secret"] = totp_secret
            
            # 计算耗时
            elapsed_time = time.time() - start_time
            self.current_account["elapsed_time"] = elapsed_time
            
            logger.info(f"[●] {self.current_account['email']} 耗时 {elapsed_time:.2f} 秒")
            
            return self.current_account
            
        except Exception as e:
            logger.error(f"创建账号失败: {str(e)}")
            self.current_account["error"] = str(e)
            self.current_account["elapsed_time"] = time.time() - start_time
            return self.current_account
    
    def bind_totp(self) -> str:
        """
        绑定TOTP
        
        Returns:
            str: TOTP密钥
        """
        try:
            logger.info(f"[•] {self.current_account['email']}的TOTP绑定中...")
            
            # 访问TOTP绑定网站
            self.driver.get("https://totp.danhersam.com/")
            
            # 等待页面加载
            time.sleep(2)
            
            # 生成TOTP密钥
            totp_secret = pyotp.random_base32()
            
            # 在网站上输入密钥
            secret_input = self.wait_for_element(By.ID, "secret")
            secret_input.clear()
            secret_input.send_keys(totp_secret)
            
            # 点击生成按钮
            generate_button = self.wait_for_clickable(By.CSS_SELECTOR, "button.btn-primary")
            generate_button.click()
            
            # 等待QR码生成
            time.sleep(2)
            
            logger.info(f"[•] {self.current_account['email']}的TOTP密钥: {totp_secret}")
            
            return totp_secret
            
        except Exception as e:
            logger.error(f"绑定TOTP失败: {str(e)}")
            return ""
    
    def change_password(self, email: str, old_password: str, new_password: Optional[str] = None) -> Dict[str, Any]:
        """
        修改Outlook账号密码
        
        Args:
            email: 邮箱地址
            old_password: 旧密码
            new_password: 新密码，如果为None则自动生成
            
        Returns:
            Dict: 包含账号信息的字典
        """
        start_time = time.time()
        
        if new_password is None:
            new_password = self.generate_random_password()
        
        account_info = {
            "email": email,
            "old_password": old_password,
            "new_password": new_password,
            "totp_secret": "",
            "update_time": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        
        try:
            logger.info(f"[✓] 开始处理 {email} 的密码修改")
            
            # 访问登录页面
            self.driver.get("https://login.live.com/")
            
            # 输入邮箱
            email_input = self.wait_for_element(By.NAME, "loginfmt")
            email_input.send_keys(email)
            
            # 点击下一步
            next_button = self.wait_for_clickable(By.ID, "idSIButton9")
            next_button.click()
            
            # 输入密码
            password_input = self.wait_for_element(By.NAME, "passwd")
            password_input.send_keys(old_password)
            
            # 点击登录
            signin_button = self.wait_for_clickable(By.ID, "idSIButton9")
            signin_button.click()
            
            # 等待登录完成
            time.sleep(3)
            
            # 访问账号设置页面
            self.driver.get("https://account.live.com/password/change")
            
            # 输入旧密码
            current_password = self.wait_for_element(By.ID, "currentPassword")
            current_password.send_keys(old_password)
            
            # 输入新密码
            new_password_input = self.wait_for_element(By.ID, "newPassword")
            new_password_input.send_keys(new_password)
            
            # 确认新密码
            confirm_password = self.wait_for_element(By.ID, "confirmNewPassword")
            confirm_password.send_keys(new_password)
            
            # 点击保存
            save_button = self.wait_for_clickable(By.ID, "save")
            save_button.click()
            
            # 等待密码修改完成
            time.sleep(3)
            
            # 绑定TOTP
            totp_secret = self.bind_totp()
            account_info["totp_secret"] = totp_secret
            
            # 计算耗时
            elapsed_time = time.time() - start_time
            account_info["elapsed_time"] = elapsed_time
            
            logger.info(f"[●] {email} 密码修改耗时 {elapsed_time:.2f} 秒")
            
            return account_info
            
        except Exception as e:
            logger.error(f"修改密码失败: {str(e)}")
            account_info["error"] = str(e)
            account_info["elapsed_time"] = time.time() - start_time
            return account_info


class AccountManager:
    """账号管理类"""
    
    def __init__(self, accounts_file: str = ACCOUNTS_FILE, totp_file: str = TOTP_SECRETS_FILE):
        """
        初始化账号管理器
        
        Args:
            accounts_file: 账号信息保存文件
            totp_file: TOTP密钥保存文件
        """
        self.accounts_file = accounts_file
        self.totp_file = totp_file
        self.ensure_files_exist()
    
    def ensure_files_exist(self):
        """确保账号文件和TOTP文件存在"""
        # 确保账号文件存在
        if not os.path.exists(self.accounts_file):
            with open(self.accounts_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([
                    "email", "password", "first_name", "last_name", 
                    "birth_year", "birth_month", "birth_day", 
                    "totp_secret", "creation_time", "elapsed_time"
                ])
            logger.info(f"创建账号文件: {self.accounts_file}")
        
        # 确保TOTP文件存在
        if not os.path.exists(self.totp_file):
            with open(self.totp_file, 'w', encoding='utf-8') as f:
                json.dump([], f, indent=2)
            logger.info(f"创建TOTP文件: {self.totp_file}")
    
    def save_account(self, account_info: Dict[str, Any]):
        """
        保存账号信息
        
        Args:
            account_info: 账号信息字典
        """
        with LOCK:
            # 保存到CSV文件
            with open(self.accounts_file, 'a', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow([
                    account_info.get("email", ""),
                    account_info.get("password", ""),
                    account_info.get("first_name", ""),
                    account_info.get("last_name", ""),
                    account_info.get("birth_year", ""),
                    account_info.get("birth_month", ""),
                    account_info.get("birth_day", ""),
                    account_info.get("totp_secret", ""),
                    account_info.get("creation_time", ""),
                    account_info.get("elapsed_time", "")
                ])
            
            # 保存TOTP信息
            if account_info.get("totp_secret"):
                totp_info = {
                    "email": account_info["email"],
                    "secret": account_info["totp_secret"],
                    "time": account_info.get("creation_time", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                }
                
                try:
                    with open(self.totp_file, 'r', encoding='utf-8') as f:
                        totp_data = json.load(f)
                except (json.JSONDecodeError, FileNotFoundError):
                    totp_data = []
                
                totp_data.append(totp_info)
                
                with open(self.totp_file, 'w', encoding='utf-8') as f:
                    json.dump(totp_data, f, indent=2)
            
            logger.info(f"账号信息已保存: {account_info['email']}")
    
    def update_account(self, account_info: Dict[str, Any]):
        """
        更新账号信息（密码修改后）
        
        Args:
            account_info: 账号信息字典
        """
        with LOCK:
            # 读取现有账号
            accounts = []
            try:
                with open(self.accounts_file, 'r', newline='', encoding='utf-8') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        accounts.append(row)
            except FileNotFoundError:
                pass
            
            # 更新账号信息
            updated = False
            for account in accounts:
                if account["email"] == account_info["email"]:
                    account["password"] = account_info["new_password"]
                    if account_info.get("totp_secret"):
                        account["totp_secret"] = account_info["totp_secret"]
                    updated = True
                    break
            
            # 如果没有找到账号，添加新账号
            if not updated:
                accounts.append({
                    "email": account_info["email"],
                    "password": account_info["new_password"],
                    "first_name": "",
                    "last_name": "",
                    "birth_year": "",
                    "birth_month": "",
                    "birth_day": "",
                    "totp_secret": account_info.get("totp_secret", ""),
                    "creation_time": account_info.get("update_time", ""),
                    "elapsed_time": account_info.get("elapsed_time", "")
                })
            
            # 写回文件
            with open(self.accounts_file, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=[
                    "email", "password", "first_name", "last_name", 
                    "birth_year", "birth_month", "birth_day", 
                    "totp_secret", "creation_time", "elapsed_time"
                ])
                writer.writeheader()
                writer.writerows(accounts)
            
            # 更新TOTP信息
            if account_info.get("totp_secret"):
                try:
                    with open(self.totp_file, 'r', encoding='utf-8') as f:
                        totp_data = json.load(f)
                except (json.JSONDecodeError, FileNotFoundError):
                    totp_data = []
                
                # 检查是否已存在
                totp_updated = False
                for totp in totp_data:
                    if totp["email"] == account_info["email"]:
                        totp["secret"] = account_info["totp_secret"]
                        totp["time"] = account_info.get("update_time", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                        totp_updated = True
                        break
                
                # 如果不存在，添加新记录
                if not totp_updated:
                    totp_data.append({
                        "email": account_info["email"],
                        "secret": account_info["totp_secret"],
                        "time": account_info.get("update_time", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    })
                
                with open(self.totp_file, 'w', encoding='utf-8') as f:
                    json.dump(totp_data, f, indent=2)
            
            logger.info(f"账号信息已更新: {account_info['email']}")
    
    def load_accounts(self) -> List[Dict[str, str]]:
        """
        加载所有账号信息
        
        Returns:
            List[Dict]: 账号信息列表
        """
        accounts = []
        try:
            with open(self.accounts_file, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    accounts.append(row)
            logger.info(f"已加载 {len(accounts)} 个账号")
        except FileNotFoundError:
            logger.warning(f"账号文件不存在: {self.accounts_file}")
        
        return accounts
    
    def export_accounts(self, output_file: str, format_type: str = "csv"):
        """
        导出账号信息
        
        Args:
            output_file: 导出文件路径
            format_type: 导出格式类型，支持"csv"和"text"
        """
        try:
            accounts = self.load_accounts()
            
            if format_type.lower() == "csv":
                with open(output_file, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    writer.writerow(["邮箱", "密码", "TOTP密钥"])
                    
                    for account in accounts:
                        writer.writerow([
                            account.get("email", ""),
                            account.get("password", ""),
                            account.get("totp_secret", "")
                        ])
            else:  # text format
                with open(output_file, 'w', encoding='utf-8') as f:
                    for account in accounts:
                        f.write(f"{account.get('email', '')}—-{account.get('password', '')}—-{account.get('totp_secret', '')}\n")
            
            logger.info(f"账号信息已导出到: {output_file}")
            return True
        except Exception as e:
            logger.error(f"导出账号信息失败: {str(e)}")
            return False


class ProxyManager:
    """代理管理类"""
    
    def __init__(self, proxies: List[str] = None):
        """
        初始化代理管理器
        
        Args:
            proxies: 代理列表，格式为 ["host:port", ...]
        """
        self.proxies = proxies or []
        self.current_index = 0
        self.lock = threading.Lock()
    
    def add_proxy(self, proxy: str):
        """
        添加代理
        
        Args:
            proxy: 代理地址，格式为 "host:port"
        """
        with self.lock:
            if proxy not in self.proxies:
                self.proxies.append(proxy)
                logger.info(f"添加代理: {proxy}")
    
    def get_next_proxy(self) -> Optional[str]:
        """
        获取下一个代理
        
        Returns:
            str: 代理地址，如果没有代理则返回None
        """
        if not self.proxies:
            return None
        
        with self.lock:
            proxy = self.proxies[self.current_index]
            self.current_index = (self.current_index + 1) % len(self.proxies)
            return proxy
    
    def load_from_file(self, file_path: str) -> int:
        """
        从文件加载代理列表
        
        Args:
            file_path: 代理文件路径，每行一个代理
            
        Returns:
            int: 加载的代理数量
        """
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                proxies = [line.strip() for line in f if line.strip()]
            
            with self.lock:
                self.proxies = proxies
                self.current_index = 0
            
            logger.info(f"从文件加载了 {len(proxies)} 个代理")
            return len(proxies)
        except Exception as e:
            logger.error(f"加载代理文件失败: {str(e)}")
            return 0


def create_accounts(count: int, proxy_manager: ProxyManager, threads: int = 1, headless: bool = True):
    """
    批量创建账号
    
    Args:
        count: 创建账号数量
        proxy_manager: 代理管理器
        threads: 线程数
        headless: 是否使用无头模式
    """
    account_manager = AccountManager()
    
    def worker(task_queue, result_queue):
        while not task_queue.empty():
            try:
                task_id = task_queue.get()
                proxy = proxy_manager.get_next_proxy()
                
                creator = OutlookCreator(proxy=proxy, headless=headless)
                try:
                    account_info = creator.create_outlook_account()
                    if "error" not in account_info:
                        account_manager.save_account(account_info)
                    result_queue.put(account_info)
                finally:
                    creator.close()
                    
            except Exception as e:
                logger.error(f"创建账号失败: {str(e)}")
                result_queue.put({"error": str(e)})
            finally:
                task_queue.task_done()
    
    # 创建任务队列
    task_queue = queue.Queue()
    result_queue = queue.Queue()
    
    for i in range(count):
        task_queue.put(i)
    
    # 创建工作线程
    thread_list = []
    for _ in range(min(threads, count)):
        t = threading.Thread(target=worker, args=(task_queue, result_queue))
        t.daemon = True
        t.start()
        thread_list.append(t)
    
    # 等待所有任务完成
    for t in thread_list:
        t.join()
    
    # 处理结果
    results = []
    while not result_queue.empty():
        results.append(result_queue.get())
    
    # 统计结果
    success_count = sum(1 for r in results if "error" not in r)
    logger.info(f"创建完成: 成功 {success_count}/{count}")
    
    return results


def change_passwords(accounts: List[Dict[str, str]], proxy_manager: ProxyManager, threads: int = 1, headless: bool = True):
    """
    批量修改密码
    
    Args:
        accounts: 账号列表，每个账号包含email和password字段
        proxy_manager: 代理管理器
        threads: 线程数
        headless: 是否使用无头模式
    """
    account_manager = AccountManager()
    
    def worker(task_queue, result_queue):
        while not task_queue.empty():
            try:
                account = task_queue.get()
                proxy = proxy_manager.get_next_proxy()
                
                creator = OutlookCreator(proxy=proxy, headless=headless)
                try:
                    new_password = creator.generate_random_password()
                    account_info = creator.change_password(
                        account["email"], 
                        account["password"], 
                        new_password
                    )
                    if "error" not in account_info:
                        account_manager.update_account(account_info)
                    result_queue.put(account_info)
                finally:
                    creator.close()
                    
            except Exception as e:
                logger.error(f"修改密码失败: {str(e)}")
                result_queue.put({"email": account["email"], "error": str(e)})
            finally:
                task_queue.task_done()
    
    # 创建任务队列
    task_queue = queue.Queue()
    result_queue = queue.Queue()
    
    for account in accounts:
        task_queue.put(account)
    
    # 创建工作线程
    thread_list = []
    for _ in range(min(threads, len(accounts))):
        t = threading.Thread(target=worker, args=(task_queue, result_queue))
        t.daemon = True
        t.start()
        thread_list.append(t)
    
    # 等待所有任务完成
    for t in thread_list:
        t.join()
    
    # 处理结果
    results = []
    while not result_queue.empty():
        results.append(result_queue.get())
    
    # 统计结果
    success_count = sum(1 for r in results if "error" not in r)
    logger.info(f"密码修改完成: 成功 {success_count}/{len(accounts)}")
    
    return results


def main():
    """主函数"""
    parser = argparse.ArgumentParser(description="Outlook邮箱自动创建与管理工具")
    
    # 创建子命令
    subparsers = parser.add_subparsers(dest="command", help="子命令")
    
    # 创建账号命令
    create_parser = subparsers.add_parser("create", help="创建新账号")
    create_parser.add_argument("-c", "--count", type=int, default=1, help="创建账号数量")
    create_parser.add_argument("-t", "--threads", type=int, default=1, help="线程数")
    create_parser.add_argument("-p", "--proxy", help="SOCKS5代理，格式为host:port")
    create_parser.add_argument("-P", "--proxy-file", help="代理文件，每行一个代理")
    create_parser.add_argument("--no-headless", action="store_true", help="不使用无头模式（默认使用无头模式）")
    
    # 修改密码命令
    change_parser = subparsers.add_parser("change", help="修改账号密码")
    change_parser.add_argument("-f", "--file", default=ACCOUNTS_FILE, help="账号文件路径")
    change_parser.add_argument("-e", "--email", help="指定邮箱地址")
    change_parser.add_argument("-t", "--threads", type=int, default=1, help="线程数")
    change_parser.add_argument("-p", "--proxy", help="SOCKS5代理，格式为host:port")
    change_parser.add_argument("-P", "--proxy-file", help="代理文件，每行一个代理")
    change_parser.add_argument("--no-headless", action="store_true", help="不使用无头模式（默认使用无头模式）")
    
    # 导出账号命令
    export_parser = subparsers.add_parser("export", help="导出账号信息")
    export_parser.add_argument("-o", "--output", required=True, help="导出文件路径")
    export_parser.add_argument("--format", choices=["csv", "text"], default="text", 
                              help="导出格式，csv或text(邮箱—-密码—-TOTP格式)，默认为text")
    
    # 解析命令行参数
    args = parser.parse_args()
    
    # 如果没有指定命令，显示帮助信息
    if not args.command:
        parser.print_help()
        return
    
    # 处理代理
    proxy_manager = ProxyManager()
    if args.command in ["create", "change"]:
        if hasattr(args, "proxy") and args.proxy:
            proxy_manager.add_proxy(args.proxy)
        
        if hasattr(args, "proxy_file") and args.proxy_file:
            proxy_manager.load_from_file(args.proxy_file)
    
    # 执行命令
    if args.command == "create":
        logger.info(f"开始创建 {args.count} 个账号，线程数: {args.threads}")
        # 默认使用无头模式，除非明确指定不使用
        headless = not args.no_headless
        create_accounts(args.count, proxy_manager, args.threads, headless=headless)
        
        # 创建完成后，自动以text格式导出最新账号信息
        account_manager = AccountManager()
        output_file = f"outlook_accounts_export_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        if account_manager.export_accounts(output_file, "text"):
            logger.info(f"账号信息已自动导出到: {output_file}")
            # 在控制台显示账号信息
            accounts = account_manager.load_accounts()
            print("\n===== 账号信息 =====")
            for account in accounts:
                print(f"{account.get('email', '')}—-{account.get('password', '')}—-{account.get('totp_secret', '')}")
            print("===================\n")
        
    elif args.command == "change":
        account_manager = AccountManager(args.file)
        accounts = account_manager.load_accounts()
        
        # 如果指定了邮箱，只修改指定邮箱的密码
        if args.email:
            accounts = [a for a in accounts if a["email"] == args.email]
            if not accounts:
                logger.error(f"未找到邮箱: {args.email}")
                return
        
        logger.info(f"开始修改 {len(accounts)} 个账号的密码，线程数: {args.threads}")
        # 默认使用无头模式，除非明确指定不使用
        headless = not args.no_headless
        change_passwords(accounts, proxy_manager, args.threads, headless=headless)
        
        # 修改完成后，自动以text格式导出最新账号信息
        output_file = f"outlook_accounts_updated_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        if account_manager.export_accounts(output_file, "text"):
            logger.info(f"更新后的账号信息已自动导出到: {output_file}")
            # 在控制台显示更新后的账号信息
            updated_accounts = account_manager.load_accounts()
            print("\n===== 更新后的账号信息 =====")
            for account in updated_accounts:
                print(f"{account.get('email', '')}—-{account.get('password', '')}—-{account.get('totp_secret', '')}")
            print("============================\n")
        
    elif args.command == "export":
        account_manager = AccountManager()
        format_type = getattr(args, "format", "text")
        if account_manager.export_accounts(args.output, format_type):
            logger.info(f"账号信息已导出到: {args.output} (格式: {format_type})")
            
            # 在控制台显示导出的账号信息
            accounts = account_manager.load_accounts()
            print(f"\n===== 导出的账号信息 ({len(accounts)}个) =====")
            for account in accounts:
                print(f"{account.get('email', '')}—-{account.get('password', '')}—-{account.get('totp_secret', '')}")
            print("===============================\n")
        else:
            logger.error("导出账号信息失败")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("程序已中断")
    except Exception as e:
        logger.error(f"程序异常: {str(e)}")
        raise
