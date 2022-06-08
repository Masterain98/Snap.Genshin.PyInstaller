import json
import os
import re
import subprocess
import sys
import tempfile
import time
import winreg
import zipfile

import requests as requests

from cert_data import cert_data_static

LOCAL_VERSION = "1.2"
has_req = False

def progressbar(url, path, fileName):
    if not os.path.exists(path):  # 看是否有该文件夹，没有则创建文件夹
        os.mkdir(path)
    start = time.time()  # 下载开始时间
    response = requests.get(url, stream=True)
    size = 0  # 初始化已下载大小
    chunk_size = 1024  # 每次下载的数据大小
    content_size = int(response.headers['content-length'])  # 下载文件总大小
    try:
        if response.status_code == 200:  # 判断是否响应成功
            print('开始下载，文件大小：{size:.2f} MB'.format(
                size=content_size / chunk_size / 1024))  # 开始下载，显示下载文件大小
            filepath = path + fileName
            with open(filepath, 'wb') as file:  # 显示进度条
                for data in response.iter_content(chunk_size=chunk_size):
                    file.write(data)
                    size += len(data)
                    print('\r' + '[下载进度]:%s%.2f%%' % (
                        '>' * int(size * 50 / content_size), float(size / content_size * 100)), end=' ')
        end = time.time()  # 下载结束时间
        print('下载完成\n用时: %.2f秒' % (end - start))  # 输出下载用时时间
    except Exception as e:
        print(e)


if __name__ == "__main__":
    # 安装证书
    try:
        print("安装临时证书中...")
        temp_cert = tempfile.NamedTemporaryFile(delete=False, encoding="utf-8", mode="w")
        temp_cert.write(cert_data_static)
        temp_cert.flush()
        os.environ['REQUESTS_CA_BUNDLE'] = temp_cert.name
    except Exception as e:
        print(e)

    # 设置变量
    output_stream = os.popen('echo %USERPROFILE%')
    user_home_path = str(output_stream.read())
    user_home_path = user_home_path.replace("\\", "/")
    user_home_path = user_home_path.replace("\n", "/")

    # 设置桌面路径
    process = subprocess.Popen(["powershell", '[Environment]::GetFolderPath("Desktop")'], stdout=subprocess.PIPE);
    user_desktop_path = result = process.communicate()[0].decode("utf-8")
    user_desktop_path = user_desktop_path.replace("\\", "/")
    user_desktop_path = user_desktop_path.replace("\n", "/")
    user_desktop_path = user_desktop_path.replace("\r", "")
    print("桌面目录： " + user_desktop_path)

    # 检查更新
    remote_version = requests.get('https://api.snapgenshin.com/installer/patch').text.replace('"', '')
    proc_arch = os.environ['PROCESSOR_ARCHITECTURE'].lower()

    print("="*30)
    print("Snap Genshin 一键安装器")
    print("本地安装器版本：" + LOCAL_VERSION)
    print("最新版：" + remote_version)
    if float(remote_version) > float(LOCAL_VERSION):
        print("检测到有新版本安装器")
        print("请在 https://resource.snapgenshin.com 下载最新版本")
    print("系统架构: " + proc_arch)
    print("="*30)
    print("请选择你需要执行的功能：")
    print("1. 自动安装所需系统环境")
    print("2. 自动安装系统环境和最新版 Snap Genshin 客户端")
    while not has_req:
        user_req = input("输入序号以执行：")
        if user_req == "1" or user_req == "2":
            has_req = True

    print("\n[检查 WebView2 Runtime 环境]")
    local_64bit = r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    user_64bit = r"Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    local_32bit = r"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    user_32bit = r"Software\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
    keyLocation = []

    # 判断系统架构
    try:
        if "86" in proc_arch:
            keyLocation = [local_32bit, user_32bit]
            print("建议升级到 x64 系统")
        elif "64" in proc_arch:
            keyLocation = [local_64bit, user_64bit]
        elif "arm" in proc_arch:
            print("不支持 arm 平台")
            sys.exit()
        else:
            print("未知平台，程序终止")
            sys.exit()

        findWebView = False
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, keyLocation[0])
            pv_value = winreg.QueryValueEx(key, "pv")[0]
            findWebView = True
        except FileNotFoundError:
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, keyLocation[1])
                pv_value = winreg.QueryValueEx(key, "pv")[0]
                findWebView = True
            except FileNotFoundError:
                print("未检测到 WebView2 Runtime 环境")
        if not findWebView:
            webView2_url = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
            progressbar(webView2_url, user_desktop_path, "webview2.exe")
            webview2_installer_path = user_desktop_path + "webview2.exe"
            install_result = str(os.system(webview2_installer_path + " /silent /install"))
            if install_result == "3010":
                print("安装成功")
            else:
                print("安装结果: " + str(install_result))
            try:
                os.remove(webview2_installer_path)
                print("删除安装文件成功")
            except:
                print("删除文件失败")
                print("Expected file: " + webview2_installer_path)
        else:
            print("WebView2 Runtime 本地版本：" + pv_value)
    except Exception as e:
        print("检测 WebView2 Runtime 环境时发生错误")
        input(e)

    try:
        print("\n[检查 .NET Desktop Runtime 环境]")
        output_stream = os.popen('dotnet --list-runtimes')
        output = str(output_stream.read())
        r = requests.get('https://api.snapgenshin.com/requirement/dotNet')
        requirement = json.loads(r.text)
        required_version = requirement["version"]
        dotNet_required_url = requirement["url"]
        m = re.search(required_version, output)
        try:
            dotNetLookupResult = m.group()
        except:
            dotNetLookupResult = None
        if dotNetLookupResult is not None:
            print(required_version + " 已安装")
        else:
            print(required_version + " 需要被安装")
            progressbar(dotNet_required_url, user_desktop_path, "dotnet.exe")
            installer_path = user_desktop_path + "dotnet.exe"
            install_result = str(os.system(installer_path + " /install /quiet /norestart"))
            if install_result == "3010":
                print("安装成功")
            try:
                os.remove(installer_path)
                print("删除安装文件成功")
            except:
                print("删除文件失败")
                print("Expected file: " + installer_path)
    except Exception as e:
        print("检测 dotNet Runtime 环境时发生错误")
        input(e)

    if user_req == "2":
        try:
            print("\n[开始下载 Snap Genshin]")
            r = requests.get('https://patch.snapgenshin.com/getPatch')
            patch_info = json.loads(r.text)
            print("当前版本：" + patch_info["tag_name"])

            progressbar(patch_info["browser_download_url"], user_desktop_path, "SnapGenshin-installer.zip")
            with zipfile.ZipFile(user_desktop_path + "/SnapGenshin-installer.zip", "r") as zip_ref:
                zip_ref.extractall(user_desktop_path + "/SnapGenshin/")
            print("Snap Genshin 已存放于桌面 SnapGenshin 文件夹")
            input("\n安装完成，按回车键退出本程序")
        except Exception as e:
            print("程序意外中断")
            input(e)
    else:
        input("\n安装完成，按回车键退出本程序")