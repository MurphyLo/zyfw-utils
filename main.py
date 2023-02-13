import re
import os
import sys
import time
import json
import execjs
import requests
import pandas as pd
from faker import Factory
from bs4 import BeautifulSoup


INFO = '[\033[32m+\033[0m] '
WARNING = '[\033[34m!\033[0m] '
ERROR = '[\033[31m-\033[0m] '

map_cols = {
    'kc': '课程名称',
    'xqmc': '校区',
    'xf': '学分',
    'zongxs': '总学时',
    'kclb': '课程类别',
    'cddw': '开课单位',
    'curent_skbjdm': '上课班号',
    'xkrssx': '限选人数',
    'xkrs': '选课/含免听',
    'qdrs': '可选人数',
    'qsz': '周次',
    'skfs': '授课方式',
    'rkjs': '任课教师',
    'sksj': '上课时间',
    'skdd': '上课地点',
    'syjx': '双语教学',
    'jpkc': '精品课程',
    'skbjdm': '课程代码-班号',
    'skfs_m': '授课方式代码',
    'xqdm': '校区代码',
}


def download_des():
    url = 'https://gitee.com/aksprince/bnu-utils/raw/master/des.js'
    try:
        res = requests.get(url, timeout=2)
    except requests.exceptions.ConnectTimeout:
        print(ERROR + 'Connection failed. Make sure you are connected to the Internet!')
        os.system('pause')
        sys.exit()
    with open('des.js', 'wb') as f:
        f.write(res.content)
    print(INFO + 'Download des.js successfully!')
    

def get_session(u, p, func_des) -> object:
    print('''
    ------------------------------
    教务课程表下载工具
    * 最后修改 2022.7.30
    
    - 请勿过高频、短时间内重复使用本工具；
    - 请勿在不完全了解程序原理的情况下
      修改代码并投入使用；
    以上行为可能导致的教务系统账号被锁等情况，
    本人无法提供支持与帮助。
    -----------------------------
    ''')
    session = requests.session()
    session.headers.update({'User-Agent': Factory.create().user_agent()})
    url = "http://zyfw.bnu.edu.cn/"
    try:
        res = session.get(url, timeout=2)
    except requests.exceptions.ConnectTimeout:
        print(ERROR + 'Connection failed. Make sure you are connected to campus WiFi or VPN!')
        os.system('pause')
        sys.exit()
    time.sleep(1.000)

    lt = re.search('name="lt" value="(.+)"', res.text).group(1)
    execution = re.search('name="execution" value="(.+)"', res.text).group(1)
    payload = {
        "rsa": func_des.call("strEnc", u + p + lt, "1", "2", "3"),
        "ul": len(u),
        "pl": len(p),
        "lt": lt,
        "execution": execution,
        "_eventId": "submit"
    }
    session.post(res.url, data=payload)
    if len(session.cookies.items()) >= 4:
        print(INFO + 'Login successfully!')
        session.get(url)
    else:
        print(ERROR + 'Login failed! Perhaps using the wrong password.')
        os.system('pause')
        sys.exit()
    time.sleep(2.000)
    return session


def get_account_info(session):
    info = {}
    session.headers.update({'Referer': 'http://zyfw.bnu.edu.cn/student/wsxk.tx.nopre.html?menucode=JW130406'})
    info.update(_get_time_range(session))
    info.update(_get_select_lesson_score(info['xn'], info['xqM'], info['xh'], session))
    info.update(_get_grade_speciaty(info['xh'], session))
    # session.headers.update({'Referer': 'http://zyfw.bnu.edu.cn/student/wsxk.kcbcx10319.html?menucode=JW130417'})
    # xnxq2 = _get_xnxq2(session)
    # if xnxq2['code'].split(',') != [info.get('xn'), info.get('xqM')]:
    #     print(WARNING + 'School year & semester inconsistency found!')
    #     info['xn'], info['xqM'] = xnxq2['code'].split(',')
    print(INFO + 'Successfully abtained grade, academic year, semester, id for tables.')
    time.sleep(5.000)
    return info


def _get_time_range(session) -> dict:
    url = "http://zyfw.bnu.edu.cn/jw/common/getWsxkTimeRange.action"
    payload = {'xktype': '2'}
    res = session.post(url, data=payload)
    data = json.loads(res.json()['result'])
    keys = ['qssj', 'jssj', 'nj', 'xh', 'xn', 'xqM']
    return {key: data.get(key) for key in keys}


def _get_select_lesson_score(xn, xqM, xh, session):
    url = "http://zyfw.bnu.edu.cn/jw/common/getSelectLessonScoreKcsInfo.action"
    payload = {
        'xn': xn,
        'xq_m': xqM,
        'xh': xh
    }
    res = session.get(url, data=payload)
    data = json.loads(res.json()['result'])
    keys = ["yxxf", "yxms", "zxf", "zms"]
    return {key: data.get(key) for key in keys}


def _get_grade_speciaty(xh, session):
    url = "http://zyfw.bnu.edu.cn/jw/common/getStuGradeSpeciatyInfo.action"
    payload = {'xh': xh}
    res = session.post(url, data=payload)
    data = json.loads(res.json()['result'])
    keys = ["zydm", "zymc", "pycc", "dwh"]
    return {key: data.get(key) for key in keys}


def _get_xnxq2(session):
    url = "http://zyfw.bnu.edu.cn/frame/droplist/getDropLists.action"
    payload = {
        'comboBoxName': 'Ms_KBBP_FBXQLLJXAP',
        'paramValue': '',
        'isYXB': '0',
        'isCDDW': '0',
        'isPYCC': '0',
        'isPYCCNJ': '0',
        'zyPyccFiled': '',
        'isKBLB': '0',
        'isXQ': '0',
        'isBJ': '0',
        'isSXJD': '0',
        'isDJKSLB': '0'
    }
    res = session.post(url, data=payload)
    data = json.loads(res.text)
    return data[0]


def read_file(path):
    try:
        with open(path, 'r', encoding='utf-8') as f:
            return f.read()
    except FileNotFoundError:
        print(ERROR + f'File {path} not found!')
        os.system('pause')
        sys.exit()


def get_course_table(info, session):
    print(INFO + f'Start downloading tables for school year {info["xn"]}, semester {info["xqM"]}.')
    index = 1
    tables = list()
    while True:
        url = "http://zyfw.bnu.edu.cn/taglib/DataTable.jsp"
        session.headers.update({'Referer': 'http://zyfw.bnu.edu.cn/taglib/DataTable.jsp?tableId=5327042'})
        payload = {
            "tableId": "5327042",
            "initQry": "0",
            "xktype": "2",
            "xh": info.get('xh'),
            "xn": info.get('xn') if not modify_xn else modify_xn,
            "xq": info.get('xqM') if not modify_xq else modify_xq,
            "nj": info.get('nj'),
            "pycc": info.get('pycc'),
            "dwh": info.get('zydm')[:2],
            "zydm": info.get('zydm'),
            "kclb1": "",
            "kclb2": "",
            "isbyk": "",
            "items": "",
            "btnFilter": "%C0%E0%B1%F0%B9%FD%C2%CB",
            "btnSubmit": "%CC%E1%BD%BB",
            "xnxq": info.get('xn') + "," + info.get('xqM'),
            "sel_pycc": info.get('pycc'),
            "sel_nj": info.get('nj'),
            "sel_yxb": info.get('zydm')[:2],
            "sel_zydm": info.get('zydm'),
            "kcmk": "",
            "sel_schoolarea": "",
            "sel_kclb1": "",
            "sel_kclb2": "",
            "sel_kcxz": "",
            "sel_kc": "",
            "sel_rkjs": "",
            "kkdw_range": "all",
            "sel_cddwdm": "",
            "menucode_current": "JW130417",
        }
        params = {'currPageCount': str(index)}
        res = session.post(url, data=payload, params=params)

        tables.append(_parse_page(res.text))
        index += 1

        sign = re.search("\('/taglib/DataTable.jsp',(\d+),(\d+)\)", res.text)
        print(INFO + f'Parsed page {sign.group(2)}/{sign.group(1)}...')
        if sign.group(1) == '0':
            print(ERROR + f'Table for school year {info["xn"]}, semester {info["xqM"]} currently unavailable!')
            os.system('pause')
            sys.exit()
        elif sign.group(1) == sign.group(2):
            print(INFO + f'Downloaded all tables!')
            break
        else:
            continue
    return pd.concat(tables)


def _parse_page(html):
    soup = BeautifulSoup(html, 'lxml')
    tbody = soup.body.div.div.table.tbody
    items = list()
    for tr in tbody:
        if not tr.name == 'tr':
            continue
        names, strings = list(), list()
        for td in tr:
            try:
                names.append(td['name'])
                strings.append(td.text)
            except KeyError:
                continue
        item = pd.Series(strings, index=names, name=tr.td.string).map(_strip_str)
        items.append(item)
    return pd.DataFrame(items)


def _strip_str(string):
    return string.strip() if type(string) == str else string


def save_table(df, file):
    t = time.strftime('%Y-%m-%d %H..%M..%S')
    with pd.ExcelWriter(file, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=t)
        worksheet = writer.sheets[t]
        col_width = [6, 30, 6, 6, 6, 32, 9, 9, 9, 12, 9, 9, 9, 9, 16, 16, 12, 12]
        for i in range(len(df.columns)+1):
            worksheet.set_column(i, i, col_width[i])
        workbook = writer.book
        fmt = workbook.add_format({
            'bg_color': '#e2efda',
            'font_color': '#000000',
        })
        rang = 'A2:' + chr(ord('A') + len(df.columns)) + str(len(df.index) + 1)
        worksheet.conditional_format(
            rang,
            {
                'type': 'formula',
                'criteria': '=mod(row(),2)=0',
                'format': fmt
            }
        )
    print(INFO + f'Saved tables to {file}!')


if __name__ == '__main__':
    '''
    将 student_number, password 变量分别修改为学号和数字京师密码后运行程序，即可自动获取最新课表的 excel 文件并存储到同一文件夹下;
    如需下载指定学年学期的课表，修改 modify_xn, modify_xq 变量，如：
    - 2022-2023学年，秋季学期
    modify_xn = '2022'
    modify_xq = '0'
    - 2021-2022学年，春季学期
    modify_xn = '2021'
    modify_xq = '1'

    * 注意：请确保修改后的变量为字符串类型
    '''
    student_number = input('请输入12位学号：')
    password = input('请输入数字京师密码：')
    modify_xn = None
    modify_xq = None


    # 使用账号密码登录
    download_des()
    func_des = execjs.compile(read_file("des.js"))
    s = get_session(student_number, password, func_des)
    info = get_account_info(s)

    # 获取自身专业最新、可用课表
    df = get_course_table(info, s)

    # 映射表头名称，无效数据清理
    df.index.name = '序号'
    df.columns = df.columns.map(map_cols)
    df = df.drop(['精品课程', '双语教学', '校区代码', '授课方式代码'], axis=1)

    # 优化数据内容（上课时间），以便于在 excel 中作数据筛选
    new_cols = ['上课时间(日)', '上课时间(节)']
    df_new_cols = df['上课时间'].str.replace(')', '', regex=False).str.split('(', expand=True)
    df_new_cols.columns = new_cols
    df = df.join(df_new_cols)
    df = df.drop('上课时间', axis=1)

    # 存储课表到同一文件夹内，以时间为 sheet 名称
    map_xq = {'0': '秋季学期', '1': '春季学期'}
    file = f"{info['zydm']}专业{info['nj']}级{info['xn']}-{str(int(info['xn'])+1)}学年{map_xq[info['xqM']]}课程安排明细.xlsx"
    save_table(df, file)
    os.system('pause')
    sys.exit()
