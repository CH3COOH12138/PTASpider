import re
import json
import time
import pandas as pd
from selenium import webdriver

path = r'.\studata.xlsx'
urlpath = r'.\urls.txt'
cookiepath = r'.\pintia_cookies.txt'


def readexcelbyid(stuid, sheets):
    info = [0 for _ in range(5)]
    file = pd.read_excel(path, sheet_name=sheets)
    index = 0
    data = file.iloc[index, 1]
    while data != 0:
        if str(data) == stuid:
            for i in range(5):
                info[i] = file.iloc[index, i]
            break
        index += 1
        data = file.iloc[index, 1]
    return info


def readexcelbyname(name, sheets):
    info = [0 for _ in range(5)]
    file = pd.read_excel(path, sheet_name=sheets)
    index = 0
    data = file.iloc[index, 2]
    while data != '0':
        if data == name:
            for i in range(5):
                info[i] = file.iloc[index, i]
            break
        index += 1
        data = file.iloc[index, 2]
    return info


def getcookie():
    # 填写完整url
    url = r'https://pintia.cn'
    driver = webdriver.Chrome()
    driver.get(url)
    # 程序打开网页后50秒内 “手动登陆账户”
    time.sleep(50)
    with open('pintia_cookies.txt', 'w+') as f:
        # 将cookies保存为json格式
        f.write(json.dumps(driver.get_cookies()))
    driver.close()


def ptaspider(url, stuid):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--incognito')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    browser = webdriver.Chrome(options=chrome_options)
    browser.implicitly_wait(10)
    browser.get(url)
    with open(cookiepath, 'r') as f:
        cookies_list = json.load(f)
        for cookie in cookies_list:
            browser.add_cookie(cookie)
    browser.get(url)
    data = browser.find_element("class name", 'DataTable_1vh8W').text
    scoreRe = re.compile('([0-9]+) ' + str(stuid) + ' ([0-9]+)')
    score = scoreRe.findall(data)
    page = 1
    if len(score) == 0:
        maxpage = browser.find_elements("class name", 'pageItem_3P4fJ')[-2].text
        while len(score) == 0 and page < int(maxpage):
            browser.get(url + '?page=' + str(page))
            data = browser.find_element("class name", 'DataTable_1vh8W').text
            score = scoreRe.findall(data)
            page += 1
    browser.quit()
    return score


if __name__ == "__main__":

    str1 = input("请输入学号或姓名：\n")
    if str1[0:3] == '202':
        for i in range(3):
            info = readexcelbyid(str1, i)
            if info[0] != 0:
                break
    else:
        for i in range(3):
            info = readexcelbyname(str1, i)
            if info[0] != 0:
                break

    if info[0] != 0:
        print("班级：" + str(info[0]))
        print("学号：" + str(info[1]))
        print("姓名：" + info[2])
        print("性别：" + info[3])
        print("宿舍：" + info[4][0:3])
        print("床位：" + info[4][-1])
        urlRe = re.compile(r'(.+?):(.+)')
        names = []
        urls = []
        with open(urlpath) as f:
            str2 = f.readlines()
        for i in range(len(str2)):
            names.append(urlRe.findall(str2[i])[0][0])
            urls.append(urlRe.findall(str2[i])[0][1])
        for i in range(len(urls)):
            score = ptaspider(urls[i], info[1])
            if len(score) != 0:
                print('\n' + names[i] + '\n排名：' + score[0][0] + '\n分数：' + score[0][1])
            else:
                print('\n' + names[i] + '\n无此记录')
    else:
        print("查无此人")

    input()
