# TODO:伪装成浏览器，使用代理IP
import re
import json
import urllib.request
import pandas as pd
import numpy as np
import xlwings as xw
from bs4 import BeautifulSoup
import pandas.io.formats.excel

url = "http://deal.ggzy.gov.cn/ds/deal/dealList_find.jsp"


def build_params(time_begin, time_end, page_num, text):
    """
    构建请求参数
    :param time_begin: 起始时间
    :param time_end: 结束时间
    :param page_num: 当前页数
    :param text: 查询字段
    :return: json
    """
    req_params = {
        # 起始时间
        'TIMEBEGIN': time_begin,
        # 结束时间
        'TIMEEND': time_end,
        # 数据来源（1省平台，2央企招投标平台）
        'SOURCE_TYPE': 1,
        # 交易时间（01当天，02近三天，03近十天……）
        'DEAL_TIME': '03',
        # 业务类型（00不限，01工程建设，02政府采购……）
        'DEAL_CLASSIFY': '00',
        # 信息来源（若DEAL_CLASSIFY=00：0000不限，0001交易公告，0002成交公示；若DEAL_CLASSIFY=01：0100不限，
        # 0101招标/资审公告……；……）
        'DEAL_STAGE': '0000',
        # 交易省份
        'DEAL_PROVINCE': 0,
        # 交易城市
        'DEAL_CITY': 0,
        # 交易平台，默认
        'DEAL_PLATFORM': 0,
        # 来源平台（如果SOURCE_TYPE=2（央企招投标平台），0代表不限，1代表中国化工装备招投标平台……）
        'BID_PLATFORM': 0,
        # 行业（0不限，A01农业，A02林业……）
        'DEAL_TRADE': 0,
        'isShowAll': 1,
        # 分页,0不返回记录
        'PAGENUMBER': page_num,
        # 查询字段
        'FINDTXT': text
    }
    return req_params


def send(req_url, req_params):
    """
    发送请求
    :param req_url:请求地址
    :param req_params:请求参数
    :return:
    """
    try:
        req_params = bytes(urllib.parse.urlencode(req_params), encoding='utf8')
        res = urllib.request.urlopen(req_url, data=req_params, timeout=50)
        res_data = res.read()
    except:
        print('Request timeout：', req_url)
        return ''
    else:
        return res_data


def extract_text(content_text, key_word):
    """
    提取与关键词有关的文本内容
    :param content_text: 待提取的文本(str)
    :param key_word: 关键字(str)
    :return: 主要内容(str)
    """
    reg3 = [k + '(.*)\n' for k in key_word]
    pattern3 = re.compile(r"|".join(reg3))
    content3 = re.findall(pattern3, content_text)
    if content3:
        for i in range(len(content3[0])):
            if content3[i]:
                index = i
                key_info = key_word[index] + ' ' + content3[0][index] + ' '
                break
    else:
        key_info = ''
    if '{' in key_info or '}' in key_info:
        key_info = ''.join(re.findall('[\u4e00-\u9fa5]', key_info))
    return key_info


def deep_content(url, title):
    """
    通过跳转链接，爬取第二层内容
    :param url: 第一层的url地址
    :return: 第二层获取的主题内容
    """
    urls_jump = get_jump_urls(url)
    contents = send(urls_jump, {})
    info = ''
    focus_sub = ''
    if contents:
        # 提取网页正文内容
        soup = BeautifulSoup(contents, "html.parser")
        all_html = soup.find('div', id="mycontent")
        # 判断该链接中的内容是否和疫情相关
        pattern3 = re.compile(r'肺炎|疫情|冠状病毒', re.S)
        focus_sub = re.findall(pattern3, str(all_html))
        # 从正文中提出表格内容
        pattern1 = re.compile(r'<table(.*)</table>', re.S)
        table_sub = re.findall(pattern1, str(all_html))
        # 提取表格中的文字
        if table_sub:
            table_str = '<table' + table_sub[0]
            pattern2 = re.compile(r'<[^>]+>', re.S)
            table_text = pattern2.sub('', table_str)
            info += table_text.replace('\n', '')
            # 删除空白行
            all_text = re.sub(r'[\r\n\f]{2,}', '\n', all_html.text)
            table_text = re.sub(r'[\r\n\f]{2,}', '\n', table_text)
            # 得到表格外的文字部分 TODO:部分网页去重失败
            content_text = all_text.replace(table_text, ' ')
        else:
            content_text = soup.find('div', id="mycontent").text
        content_text = content_text.replace('。', '\n')
        content_text = content_text.replace('；', ' ')
        content_text = content_text.replace('、', '\n')

        if '更正' in title or '更换' in title or '变更' in title or '澄清' in title:
            key = ['主要内容：', '变更内容：', '更正事项和内容：', '截止时间变更为' '更正事项：', '更正/变更：', '更正：', '变更：',
                   '更正内容：', '更正（补充）事项及内容：', '更改：', '变更为：', '延长至', '更正为']
            info += extract_text(content_text, key)
        elif '中标' in title:
            key3 = ['项目名称', '工程名称', '产品（项目）名称', '项目标名', '标段：']
            key4 = ['中标人', '中标供应商', '中标单位', '成交供应商', '竞得人', '中标单位名称']
            key5 = ['中标价', '中标金额', '项目金额', '项目成交金额', '成交金额', '成交总金额', '项目预算金额']
            info += extract_text(content_text, key3)
            info += extract_text(content_text, key4)
            info += extract_text(content_text, key5)
        elif '招标' in title or '采购' in title or '磋商' in title:
            key3 = ['项目名称', '工程名称', '产品（项目）名称', '项目标名', '标段：']
            key4 = ['中标人', '中标供应商', '中标单位', '成交供应商', '竞得人', '中标单位名称']
            key5 = ['中标价', '中标金额', '项目金额', '项目成交金额', '成交金额', '成交总金额', '项目预算金额']
            key6 = ['采购人', '招标人']
            info += extract_text(content_text, key3)
            info += extract_text(content_text, key4)
            info += extract_text(content_text, key5)
            info += extract_text(content_text, key6)
        if len(info) < 25:
            info = ''
            all_html = re.sub(r'[\r\n\f]{2,}', '\n', all_html.text)
            all_text = ''.join(re.findall('[\u4e00-\u9fa5]', str(all_html)))
            info += all_text
    else:
        info += '无法访问跳转链接'
    # print(info)
    return info, focus_sub


def set_excel(writer, key, important_rows):
    """
    设置excel表格的格式
    :param writer: 工作簿
    :param key: 关键字
    :return: 无
    """
    # 文件位置：path，打开表格，然后保存，关闭，结束程序
    ws = writer.sheets[key]

    # 获取最后一列
    last_column = ws.range(1, 1).end('right').get_address(0, 0)[0]
    # 获取最后一行
    last_row = ws.range(1, 1).end('down').row
    # 生成表格的数据范围
    a_range = f'A1:{last_column}{last_row}'

    ws.autofit()

    # 设置字体：字号为10，第一行标题加粗
    ws.range(f'A1:{last_column}{last_row}').api.Font.Size = 10
    ws.range(f'A1:G1').api.Font.Bold = True

    # 设置列宽
    width_list = [12, 36, 12, 12, 12, 50, 60]
    for i in range(len(width_list)):
        ws.range(i+1, i+1).column_width = width_list[i]

    # 设置行高
    ws.range(a_range).row_height = 13.2  # 设置第1行

    # 设置水平方向靠左-4131，F列垂直方向自动换行    
    # VerticalAlignment -4108 垂直居中（默认）。 -4160 靠上，-4107 靠下， -4130 自动换行对齐。
    ws.range(f'G1:G{last_row}').api.VerticalAlignment = -4130
    ws.range(f'F1:F{last_row}').api.VerticalAlignment = -4130
    ws.range(f'B1:B{last_row}').api.VerticalAlignment = -4130

    # HorizontalAlignment -4108 水平居中。 -4131 靠左，-4152 靠右。
    ws.range(f'A1:{last_column}1').api.HorizontalAlignment = -4108  # 第一行标题居中
    ws.range(f'B2:B{last_row}').api.HorizontalAlignment = -4131  # 第2, 6, 7列的值靠左
    ws.range(f'F2:F{last_row}').api.HorizontalAlignment = -4131
    ws.range(f'G2:G{last_row}').api.HorizontalAlignment = -4131
    ws.range(f'A2:A{last_row}').api.HorizontalAlignment = -4108  # 第1，3，4, 5列的值水平居中
    ws.range(f'C2:C{last_row}').api.HorizontalAlignment = -4108
    ws.range(f'D2:D{last_row}').api.HorizontalAlignment = -4108
    ws.range(f'E2:D{last_row}').api.HorizontalAlignment = -4108

    # 设置边框
    ws.range(a_range).api.Borders(8).LineStyle = 1  # 上边框
    ws.range(a_range).api.Borders(8).Weight = 2  # 边框线宽
    ws.range(a_range).api.Borders(9).LineStyle = 1  # 下边框
    ws.range(a_range).api.Borders(9).Weight = 2
    ws.range(a_range).api.Borders(7).LineStyle = 1  # 左边框
    ws.range(a_range).api.Borders(7).Weight = 2
    ws.range(a_range).api.Borders(10).LineStyle = 1  # 右边框
    ws.range(a_range).api.Borders(10).Weight = 2
    ws.range(a_range).api.Borders(12).LineStyle = 1  # 内横边框
    ws.range(a_range).api.Borders(12).Weight = 2
    ws.range(a_range).api.Borders(11).LineStyle = 1  # 内纵边框
    ws.range(a_range).api.Borders(11).Weight = 2

    # 将重要的数据填充黄色
    # for index in important_rows:
    #     row = index+2
    #     ws.range(f'A{row}:G{row}').color = 255, 255, 0


def judge_company(province):
    """
    根据省份得到对应的区域公司
    :param province: 省份(str)
    :return: 区域公司(str)
    """
    company_dict = {'云网本部': ['北京', '安徽', '山西'], '天智公司': ['天津', '河北', '黑龙江'],
                    '江苏公司': ['山东', '江苏'], '河南办事处': ['河南'], '中之杰公司': ['辽宁', '吉林'],
                    '航联公司': ['内蒙古'], '四川公司': ['陕西', '四川'],
                    '甘肃公司': ['甘肃', '宁夏', '青海', '西藏', '新疆'],
                    '湖北办事处': ['湖北', '湖南'], '浙江公司': ['上海', '浙江'], '江西公司': ['江西'],
                    '广东公司': ['福建', '广东', '海南'], '重庆公司': ['重庆'], '贵州公司': ['贵州', '广西', '云南']}

    company = ''
    for comp, provinces in company_dict.items():
        if province in provinces:
            company = comp
            break
    return company


def export(data, writer, sheet_name, path):
    """
    输出数据到excel表格
    :param data: 获取的数据
    :param writer: 工作簿
    :param sheet_name: sheet标签名
    :param path:  excel所在的保存路径
    :return: 含有重要信息的字典{关键词1：[索引列表], 关键词2：[索引列表]}
    """
    data1 = pd.DataFrame(data[1:])
    important_indexs = []
    focus_sub = ''
    if len(data1) != 0:
        result = data1[['title', 'timeShow', 'stageShow', 'districtShow', 'url']]
        result.columns = ['标题', '公示时间', '信息类型', '省份', '链接地址']
        theme = []
        indexs = []
        urls = result['链接地址']
        for i in range(result.shape[0]):
            print("正在写入%d数据" % (i+1))
            title = result[['标题']].iloc[i, :].values[0]
            main_content, focus_sub = deep_content(urls[i], title)
            theme.append(main_content)
            # 如果此数据和“疫情”有关，则将所在的数据号加入列表中
            # if title == '山东省泰安市宁阳县第一人民医院区域PACS硬件平台采购及虚拟化平台升级项目变更公告':
            #     print('山东111111111111111111111111：', focus_sub)
            # if focus_sub or ('肺炎' in title) or ('冠状病毒' in title) or ('疫情' in title):
            #     indexs.append(i)
        result.insert(5, '主要内容', np.array(theme))
        result.drop_duplicates()
    else:
        result = pd.DataFrame(columns=['标题', '公示时间', '信息类型', '省份', '链接地址', '主要内容'])
    company_list = []
    for province in result['省份'].values.tolist():
        company_list.append(judge_company(province))
    result_final = pd.concat([pd.DataFrame(np.array(company_list).reshape(-1, 1), columns=['区域公司']), result], axis=1)
    result_final.to_excel(writer, sheet_name=sheet_name, index=False)  # 保存工作簿
    # if focus_sub:
    #     data2 = result_final.iloc[indexs, :]
    #     important_indexs = data2[data2['信息类型'].str.contains(r'采购/资审公告|采购合同|更正事项|开标记录|招标/资深公告|招标/资审文件澄清|交易公告')].index.tolist()
    print("'%s'相关数据已成功写入：%s！" % (sheet_name, path))
    print("------等待------")
    return important_indexs


def search(time_begin, time_end, text):
    results = [{}]
    response = send(url, build_params(time_begin, time_end, 0, text))
    if response:
        data = response.decode('utf-8')
        result = json.loads(data)
        page_size = result['ttlpage']
        print('查询到%s相关内容，一共%d页' % (text, page_size))
        # 根据页数循环请求
        for i in range(1, page_size + 1):
            # print("正在读取与'%s'相关的第%d页:%s" % (text, i, url))
            response = send(url, build_params(time_begin, time_end, i, text))
            if response:
                data = response.decode('utf-8')
                result = json.loads(data)
                # print('与%s相关的第%d页内容:%s' % (text, i, result))
                for row in result['data']:
                    results.append(row)
    return results


def get_jump_urls(raw_url):
    """
    获得某网页对应的的跳转链接地址
    :param raw_url: 原始url地址
    :return:对应的跳转url地址
    """
    # 获取raw_url网页的内容
    try:
        page = urllib.request.urlopen(raw_url, timeout=50)
        contents = page.read().decode('utf-8')
    except:
        print('Request timeout：', raw_url)
        return ''
    else:
        # 获取当前网页对应的跳转链接所在标签中的文本
        soup = BeautifulSoup(contents, "html.parser")
        text = str(soup.find('li', class_='li_hover'))
        # 利用正则，提取跳转链接地址，并返回结果
        pattern = re.compile(r'onclick="showDetail(.*?).shtml')
        urls_jump = re.findall(pattern, text)[0].split(',')[2].replace("'", "")
        urls_jump = 'http://www.ggzy.gov.cn/information' + urls_jump + '.shtml'
        return urls_jump


def main():
    # url = 'http://www.ggzy.gov.cn/information/html/a/370000/0204/202002/21/00376584e76a4b7a41a4b50cca1692ea6b4d.shtml'
    # deep_content(url, '平顶山学院智慧校园软件及实验室管理平台（智慧校园-')
    # 获取近十天平台数据
    key_words = ['疫情防控', '复工复产', '新冠肺炎', '医用口罩', '额温枪', '红外检测']
    path = "全国公共资源平台招投标信息汇总（截至2020030417）.xlsx"
    important_indexs_dict = {}
    # 生成excel,并写入数据
    pandas.io.formats.excel.header_style = None
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    for key in key_words:
        result = search('2020-02-23', '2020-03-03', key)
        important_indexs = export(result, writer, key, path)
        important_indexs_dict[key] = important_indexs
    writer.save()
    writer.close()

    # 设置单元格格式
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.open(path)
    for key in key_words:
        set_excel(wb, key, important_indexs_dict[key])
    wb.save()
    wb.close()
    app.quit()

    print('----完成爬虫任务----')


if __name__ == "__main__":
    main()
