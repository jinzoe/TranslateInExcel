# coding=gbk
import urllib
import urllib.request
import json
import time
import hashlib
import openpyxl


class YouDaoFanyi:
    def __init__(self, appKey, appSecret):
        self.url = 'https://openapi.youdao.com/api/'
        self.headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.109 Safari/537.36",
        }
        self.appKey = '****************'  # Ӧ��id
        self.appSecret = '*************'  # Ӧ����Կ
        self.langFrom = 'auto'   # ����ǰ��������,autoΪ�Զ����
        self.langTo = 'auto'     # �������������,autoΪ�Զ����

    def getUrlEncodedData(self, queryText):
        '''
        ������url����
        :param queryText: �����������
        :return: ����url�����������
        '''
        salt = str(int(round(time.time() * 1000)))  # ��������� ,��ʵ�̶�ֵҲ����,����"2"
        sign_str = self.appKey + queryText + salt + self.appSecret
    
        sign = hashlib.md5(sign_str.encode("utf-8")).hexdigest()
        payload = {
            'q': queryText,
            'from': self.langFrom,
            'to': self.langTo,
            'appKey': self.appKey,
            'salt': salt,
            'sign': sign
        }

        # ע����get���󣬲�������
        data = urllib.parse.urlencode(payload)
        return data

    def parseHtml(self, html):
        '''
        ����ҳ�棬���������
        :param html: ���뷵�ص�ҳ������
        :return: None
        '''
        data = json.loads(html)
       # print ('-' * 10)
        translationResult = data['translation']
       
        if isinstance(translationResult, list):
            translationResult = translationResult[0]
        print (translationResult)
        if "basic" in data:
            youdaoResult = data['basic']
           # print ('�е��ʵ���')
            return youdaoResult 
        #print ('-' * 10)

    def translate(self, queryText):
        data = self.getUrlEncodedData(queryText)  # ��ȡurl�����������
        target_url = self.url + '?' + data    # ����Ŀ��url
       # print(target_url)
        request = urllib.request.Request(target_url, headers=self.headers)  # ��������
        response = urllib.request.urlopen(request)  # ��������
        return self.parseHtml(response.read().decode())    # ��������ʾ������


if __name__ == "__main__":

    path="C:/Users/*********************"

    fanyi = YouDaoFanyi(appKey, appSecret)
       
    wb = openpyxl.load_workbook(path)
    sheet = wb.get_sheet_by_name('Sheet1')
    
    i=0
    for row in sheet.rows:
        i+=1
       # print(row[0].value, "\t", end="")	   
        basicJson=fanyi.translate(row[0].value.strip())
        
        if basicJson:
            try:
                sheet.cell(column=2,row=i).value="["+basicJson['us-phonetic']+"]"
            except:
                try:
                    sheet.cell(column=2,row=i).value="["+basicJson['uk-phonetic']+"]"
                except:
                    print("error1")
            try:
                sheet.cell(column=3,row=i).value=basicJson['explains'][0]
            except:
                print("error2")
		       
       # print(basicJson['explains'][0])
       # print(basicJson['us-phonetic'])
    print()
    
    wb.save(path)