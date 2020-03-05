import urllib.request
import urllib.request as url_request
import urllib.parse as url_parse
import chardet
import datetime
import pandas as pd
import numpy as np
import pandas.io.formats.excel
import pandas as pd
import xlwings as xw
import pandas.io.formats.excel
import time


def post(req_url, req_params):
    """
    post请求
    :param req_url:请求地址
    :param req_params:请求参数
    :return:
    """
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) "
                             "Chrome/47.0.2526.106 Safari/537.36"}
    params = bytes(url_parse.urlencode(req_params), encoding='utf8')
    req = url_request.Request(url=req_url, data=params, headers=headers)
    # open the url
    res = url_request.urlopen(req)
    res_data = res.read()
    encoding = chardet.detect(res_data)['encoding']
    return res_data.decode(encoding, 'ignore')


def get(req_url):
    """
    get参数请求
    :param req_url:请求地址
    :return:
    """
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) "
                             "Chrome/47.0.2526.106 Safari/537.36"}
    req = url_request.Request(req_url, headers=headers)
    # open the url
    res = url_request.urlopen(req)
    res_data = res.read()
    encoding = chardet.detect(res_data)['encoding']
    return res_data.decode(encoding, 'ignore')


class AbstractSpider(object):
    def __init__(self, url, name):  # 约定成俗这里应该使用r，它与self.r中的r同名
        self.url = url
        self.name = name
        self.path = name + "信息汇总（截至" + str(datetime.date.today()) + "）.xlsx"

    @staticmethod
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

    def write_excel(self, writer, data, sheet_name):
        """
        输出数据到excel表格
        :param data: 获取的数据{title,timeShow,stageShow,districtShow,url,content}
        :param writer: 工作簿
        :param sheet_name: sheet标签名
        :return: 含有重要信息的字典{关键词1：[索引列表], 关键词2：[索引列表]}
        """
        data1 = pd.DataFrame(data)
        important_index = []
        if len(data1) != 0:
            result = data1[['title', 'timeShow', 'stageShow', 'districtShow', 'url']]
            result.columns = ['标题', '公示时间', '信息类型', '省份', '链接地址']
            theme = []
            urls = result['链接地址']
            for i in range(result.shape[0]):
                title = result[['标题']].iloc[i, :].values[0]
                print("正在写入%d数据%s" % (i + 1, title))
                main_content = self.deep_content(urls[i], title)
                theme.append(main_content)
            result.insert(5, '主要内容', np.array(theme))
            result.drop_duplicates()
        else:
            result = pd.DataFrame(columns=['标题', '公示时间', '信息类型', '省份', '链接地址', '主要内容'])
        company_list = []
        for province in result['省份'].values.tolist():
            company_list.append(self.judge_company(province))
        result_final = pd.concat([pd.DataFrame(np.array(company_list).reshape(-1, 1), columns=['区域公司']), result],
                                 axis=1)
        result_final.to_excel(writer, sheet_name=sheet_name, index=False)  # 保存工作簿
        print("'%s'相关数据已成功写入：%s！" % (sheet_name, self.path))
        print("------等待------")
        return important_index

    @staticmethod
    def set_excel(writer, key):
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
            ws.range(i + 1, i + 1).column_width = width_list[i]

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

    @staticmethod
    def build_params(left_day, search_text, page_num):
        """
        构建查询参数，子类覆盖此方法
        :param page_num:
        :param left_day:
        :param search_text:
        :return:
        """
        pass

    @staticmethod
    def parse_html(content):
        """
        解析网页，子类覆盖此方法
        :param content:
        :return:
        """
        return []

    @staticmethod
    def url_encode(key):
        """
        url编码，子类覆盖此方法
        :return:
        """
        pass

    @staticmethod
    def deep_content(url, title):
        """
        爬主要内容，子类需覆盖此方法
        :param url:
        :param title:
        :return:
        """
        return ''

    def process(self, key_words, total_page):
        # 生成excel,并写入数据
        pandas.io.formats.excel.header_style = None
        writer = pd.ExcelWriter(self.path, engine='xlsxwriter')
        for key in key_words:
            # 打开网页
            print("读取网页%s, 参数:%s" % (self.url, key))
            results = []
            # 读取7天的,total_page页面
            for i in range(total_page):
                content = get(self.url + self.build_params(7, self.url_encode(key), i))
                # 解析网页
                result = self.parse_html(content)
                results.extend(result)

            # 把页面读完再写
            self.write_excel(writer, results, key)
            time.sleep(1)

        # 关闭excel
        writer.save()
        writer.close()
        # 设置单元格格式
        app = xw.App(visible=True, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.open(self.path)
        for key in key_words:
            self.set_excel(wb, key)
        wb.save()
        wb.close()
        # app.quit()
