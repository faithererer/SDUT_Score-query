# _*_ coding : utf-8 _*_
# @Time : 2022/7/16 18:36
# @File : sdut_sc
# @Project : CODE_PY_single
import json

import requests
import json
import xlwt
url = ''
header ={

}
# 成绩查询界面的url
sc_url = ''

# 抓包找到提交全部学期成绩表单请求url
post_url = ''

page = 0

# from data 需要自己抓包
data = {
    "xnm": "",
    "xqm": "",
    "_search": "false",
    "nd": "1657969924833",
    "queryModel.showCount": "15",
    "queryModel.currentPage": "1",
    "queryModel.sortName": "",
    "queryModel.sortOrder": "asc",
    "time": "1",
}

# 发起一个会话
def creat_session(url, header):
    session = requests.session()
    session.get(url=url, headers=header)
    return session

# 会话
def session_next(url, header, session):
    session.get(url=url, headers=header)

# 提交表单
def session_post(url,header,data, session):
    response = session.post(url=url, headers=header, data=data)
    return response.text

# 创建xls表格
def creat_sheet(sheet_name):
    wb = xlwt.Workbook()
    sheet = wb.add_sheet(sheet_name)
    return [wb, sheet]

# 把字符串转变字典
def turn_dic(response_text):
    return json.loads(response_text)

# 从字典中获得成绩单页数
def get_total_page(dic):
    return dic.get('totalPage')

# 找到成绩主体json
def find_items(dic):
    return dic.get("items")

# 为表格写表头
def write_title(items):
    for index0, title in enumerate(items[0].keys()):
        sheet.write(0, index0, title)

# 写入表格
def write_sheet(items, index_data, sheet):
    for index_d, data in enumerate(items):
        # print(index_d)
        for index_obj, obj in enumerate(items[index_d].values()):
            sheet.write(index_data+index_d+1, index_obj, str(obj))
    return int(len(items))

# 保存表格
def save_xls(wb, xls_name):
    wb.save(xls_name)


if __name__ == '__main__':
    session = creat_session(url, header)
    session_next(sc_url, header, session)
    response_text = session_post(post_url, header, data, session)
    dic = turn_dic(response_text)
    total_page = get_total_page(dic)
    sheet_about = creat_sheet('成绩')
    sheet = sheet_about[1]
    wb = sheet_about[0]
    index_data = 0
    for page in range(1, total_page+1):
        data['time'] = str(page)
        data['queryModel.currentPage'] = str(page)
        response_text = session_post(post_url, header, data, session)
        dic = turn_dic(response_text)
        items = find_items(dic)  # 返回一个字典
        items = json.dumps(items)  # 字典转换为json对象
        items = turn_dic(items)  # 转换为字典
        index_data = write_sheet(items, index_data, sheet)
    write_title(items)
    save_xls(wb, '成绩.xls')
