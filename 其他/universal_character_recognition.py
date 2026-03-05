from datetime import datetime
from wsgiref.handlers import format_date_time
from time import mktime
import hashlib
import base64
import hmac
from urllib.parse import urlencode
import json
import requests

'''
1、通用文字识别,图像数据base64编码后大小不得超过10M
2、appid、apiSecret、apiKey请到讯飞开放平台控制台获取并填写到此demo中
3、支持中英文,支持手写和印刷文字。
4、在倾斜文字上效果有提升，同时支持部分生僻字的识别
'''


APPId = "6c8f82e6"
APISecret = "NGQ2ZTg3NWY3NDkxMTMyYWJlYWQwNTJm"
APIKey = "73ed4e3a092cc935af22e2c420bd9cd8"
FILE_PATH = "TESTPIC.png" # 上传图片地址

with open("TESTPIC.png", "rb") as f:
    imageBytes = f.read()


class AssembleHeaderException(Exception):
    def __init__(self, msg):
        self.message = msg


class Url:
    def __init__(self, host, path, schema):
        self.host = host
        self.path = path
        self.schema = schema
        pass


# calculate sha256 and encode to base64
def sha256base64(data):
    sha256 = hashlib.sha256()
    sha256.update(data)
    digest = base64.b64encode(sha256.digest()).decode(encoding='utf-8')
    return digest


def parse_url(requset_url):
    stidx = requset_url.index("://")
    host = requset_url[stidx + 3:]
    schema = requset_url[:stidx + 3]
    edidx = host.index("/")
    if edidx <= 0:
        raise AssembleHeaderException("invalid request url:" + requset_url)
    path = host[edidx:]
    host = host[:edidx]
    u = Url(host, path, schema)
    return u


# build websocket auth request url
def assemble_ws_auth_url(requset_url, method="POST", api_key="", api_secret=""):
    u = parse_url(requset_url)
    host = u.host
    path = u.path
    now = datetime.now()
    date = format_date_time(mktime(now.timetuple()))
    # 注释掉调试输出
    # print(date)
    # date = "Thu, 12 Dec 2019 01:57:27 GMT"
    signature_origin = "host: {}\ndate: {}\n{} {} HTTP/1.1".format(host, date, method, path)
    # print(signature_origin)
    signature_sha = hmac.new(api_secret.encode('utf-8'), signature_origin.encode('utf-8'),
                             digestmod=hashlib.sha256).digest()
    signature_sha = base64.b64encode(signature_sha).decode(encoding='utf-8')
    authorization_origin = "api_key=\"%s\", algorithm=\"%s\", headers=\"%s\", signature=\"%s\"" % (
        api_key, "hmac-sha256", "host date request-line", signature_sha)
    authorization = base64.b64encode(authorization_origin.encode('utf-8')).decode(encoding='utf-8')
    # print(authorization_origin)
    values = {
        "host": host,
        "date": date,
        "authorization": authorization
    }

    return requset_url + "?" + urlencode(values)


url = 'https://api.xf-yun.com/v1/private/sf8e6aca1'

body = {
    "header": {
        "app_id": APPId,
        "status": 3
    },
    "parameter": {
        "sf8e6aca1": {
            "category": "ch_en_public_cloud",
            "result": {
                "encoding": "utf8",
                "compress": "raw",
                "format": "json"
            }
        }
    },
    "payload": {
        "sf8e6aca1_data_1": {
            "encoding": "jpg",
            "image": str(base64.b64encode(imageBytes), 'UTF-8'),
            "status": 3
        }
    }
}

request_url = assemble_ws_auth_url(url, "POST", APIKey, APISecret)

headers = {'content-type': "application/json", 'host': 'api.xf-yun.com', 'app_id': APPId}
# 注释掉调试输出，避免显示乱码
# print(request_url)
response = requests.post(request_url, data=json.dumps(body), headers=headers)
# print(response)
# print(response.content)

# 只处理响应，不打印原始内容
# print("resp=>" + response.content.decode())
tempResult = json.loads(response.content.decode())

# 解码返回的文本（可能是JSON格式，包含坐标信息）
decoded_text = base64.b64decode(tempResult['payload']['result']['text']).decode()

# 尝试解析为JSON，如果包含坐标信息则只提取文本
try:
    result_json = json.loads(decoded_text)
    # 提取所有文本内容，忽略坐标信息
    text_lines = []
    
    # 处理不同的JSON结构
    if 'pages' in result_json:
        # 包含pages的结构
        for page in result_json['pages']:
            if 'lines' in page:
                for line in page['lines']:
                    if 'content' in line:
                        text_lines.append(line['content'])
    elif 'lines' in result_json:
        # 直接包含lines的结构
        for line in result_json['lines']:
            if 'content' in line:
                text_lines.append(line['content'])
    elif isinstance(result_json, list):
        # 列表格式
        for item in result_json:
            if isinstance(item, dict) and 'content' in item:
                text_lines.append(item['content'])
    else:
        # 其他格式，尝试提取所有文本字段
        if 'content' in result_json:
            text_lines.append(result_json['content'])
        elif 'text' in result_json:
            text_lines.append(result_json['text'])
    
    # 合并所有文本行，智能处理换行
    import re
    
    if text_lines:
        # 智能合并文本：将同一句子的行合并，保留段落分隔
        merged_lines = []
        current_line = ""
        
        for i, line in enumerate(text_lines):
            line = line.strip()
            if not line:
                # 空行可能是段落分隔，如果当前行有内容则保存
                if current_line:
                    merged_lines.append(current_line.strip())
                    current_line = ""
                continue
            
            # 如果当前行为空，开始新行
            if not current_line:
                current_line = line
                continue
            
            # 判断是否应该合并到当前行
            should_merge = False
            
            # 情况1：当前行以句号、问号、感叹号结尾，下一行应该新起
            if current_line.rstrip().endswith(('.', '!', '?', '。', '！', '？')):
                should_merge = False
            # 情况2：当前行很短（少于5个字符），可能是单词的一部分，应该合并
            elif len(line) <= 5 and not line[0].isupper():
                should_merge = True
            # 情况3：当前行不以句号结尾，且下一行不以大写字母开头，可能是同一句
            elif not current_line.rstrip().endswith(('.', '!', '?', '。', '！', '？', ':', ';')) and not line[0].isupper():
                should_merge = True
            # 情况4：当前行以逗号、分号结尾，下一行应该合并（除非下一行以大写字母开头）
            elif current_line.rstrip().endswith((',', ';', '，', '；')) and not line[0].isupper():
                should_merge = True
            # 情况5：下一行以大写字母开头，可能是新句子，不合并
            elif line[0].isupper():
                should_merge = False
            else:
                # 默认情况：合并
                should_merge = True
            
            if should_merge:
                # 合并到当前行，添加空格
                current_line += " " + line
            else:
                # 保存当前行，开始新行
                merged_lines.append(current_line.strip())
                current_line = line
        
        # 添加最后一行
        if current_line:
            merged_lines.append(current_line.strip())
        
        # 合并为最终文本
        finalResult = '\n'.join(merged_lines).strip()
        
        # 清理多余的空白字符
        # 将多个连续空格替换为单个空格
        finalResult = re.sub(r' +', ' ', finalResult)
        # 将多个连续换行（超过2个）替换为2个换行（段落分隔）
        finalResult = re.sub(r'\n{3,}', '\n\n', finalResult)
        # 清理行首行尾的空白
        finalResult = '\n'.join([line.strip() for line in finalResult.split('\n')])
    else:
        # 如果无法解析，使用原始文本，但进行智能清理
        finalResult = decoded_text.strip()
        # 清理多余的空白，但保留基本结构
        finalResult = re.sub(r' +', ' ', finalResult)
        finalResult = re.sub(r'\n{3,}', '\n\n', finalResult)
        # 清理行首行尾的空白
        finalResult = '\n'.join([line.strip() for line in finalResult.split('\n')])
except (json.JSONDecodeError, KeyError, TypeError):
    # 如果不是JSON格式或解析失败，直接使用原始文本，但进行智能清理
    import re
    finalResult = decoded_text.strip()
    # 清理多余的空白，但保留基本结构
    finalResult = re.sub(r' +', ' ', finalResult)
    finalResult = re.sub(r'\n{3,}', '\n\n', finalResult)
    # 清理行首行尾的空白
    finalResult = '\n'.join([line.strip() for line in finalResult.split('\n')])

# 只输出最终的纯文本内容，不输出其他调试信息
print(finalResult)
