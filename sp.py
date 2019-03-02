import requests
import re
import random
from PIL import Image
from urllib.parse import urlencode
import pandas as pd
import numpy as np

def get_img(Session, id_numm):
    try:
        headers = {
            'Connection': 'keep - alive',
            'Host': 'cache.neea.edu.cn',
            'Referer': 'http://cet.neea.edu.cn/cet',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3486.0 Safari/537.36',
        }
        Session.headers = headers
        get_url = 'http://cache.neea.edu.cn/Imgs.do?'
        params = {
            'c': 'CET',
            'ik': id_numm,
            't': random.random()
        }
        response = Session.get(get_url, params=params)
        img_url = re.compile('"(.*?)"').findall(response.text)[0]
        img = requests.get(img_url, timeout=None)
        with open('C:/Users/bjw10/Desktop/kaust/para/img.png', 'wb') as f:
            f.write(img.content)
        Image.open('C:/Users/bjw10/Desktop/kaust/para/img.png').show()
    except Exception as e:
        print("Imgae_Error:", e.args)

def get_score(Session, id_num, name, level):
    capcha = input('Input verification code:')

    headers = {
        'Connection': 'keep - alive',
        'Host': 'cache.neea.edu.cn',
        'Origin': 'http://cet.neea.edu.cn',
        'Referer': 'http://cet.neea.edu.cn/cet',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3486.0 Safari/537.36',
    }

    query_url = "http://cache.neea.edu.cn/cet/query"

    test = {
        '1': 'CET4_182_DANGCI',
        '2': 'CET6_182_DANGCI',
    }

    data = {
        'data': test.get(level) + ',' + str(id_num) + ',' + name,
        'v': capcha
    }
    data = urlencode(data)
    response = Session.post(query_url, data=data, headers=headers)
    if 'error' in response.text:
        e = re.compile("'error':'(.*?)'|error:'(.*?)'").findall(response.text)[0]
        if e is not None:
            #print(e)
            if '验证码错误' in e[1]:
                print("Verification code error!")
                get_img(Session, id_num)
                get_score(Session, id_num, name, level)
            else:
                print(e[0])
    else:
        id_num = re.compile("z:'(.*?)'").findall(response.text)[0]
        name = re.compile("n:'(.*?)'").findall(response.text)[0]
        score = re.compile("s:(.*?),").findall(response.text)[0]
        listening = re.compile("l:(.*?),").findall(response.text)[0]
        reading = re.compile("r:(.*?),").findall(response.text)[0]
        writing = re.compile("w:(.*?),").findall(response.text)[0]
        rank = re.compile("kys:'(.*?)'").findall(response.text)[0]
        return [str(name),str(id_num),str(score),str(listening),str(reading),str(writing),str(rank)]
        

def cetSpider(inputName, outputName):
    xls = pd.read_excel(inputName)
    ret = pd.DataFrame(xls, columns = ['姓名', '准考证号', '总分', '听力','阅读','写作与翻译','口语等级'])
    for it in range(27):
        id_num =xls['准考证号'][it]
        name = xls['姓名'][it]
        level = str(id_num)[9]
        s = requests.Session()
        get_img(s, id_num)
        ret.loc[it] = get_score(s, id_num, name, level)
        ret.to_excel(outputName)

if __name__ == '__main__':
    cetSpider('demo/tst1.xlsx', 'demo/op.xlsx')