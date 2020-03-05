# 中国招投标网
import spider_common as base
from lxml import etree
from lxml import html
# 主站
spider_url = "http://www.bidchance.com/freesearch.do"
spider_name = '中国招投标网'


class ChinaBid(base.AbstractSpider):
    @staticmethod
    def build_params(left_day, search_text, page_num):
        """
        请求参数
        :param page_num:
        :param left_day: 近期天数
        :param search_text: 搜索内容
        :return:
        """
        data = {
            'leftday': left_day,
            'filetype': '',
            'channel': '',
            'currentpage': str(page_num),
            'searchtype': 'zb',
            'queryword': search_text,
            'displayStyle': '',
            'pstate': '',
            'field': '',
            'leftday': '',
            'province': '',
            'bidfile': '',
            'project': '',
            'heshi': '',
            'recommand': '',
            'jing': '',
            'starttime': '',
            'endtime': '',
            'attachment': ''
        }
        # 拼接字符串
        query = ''
        for key in data:
            query += '&' + key + '=' + data.get(key)
        return '?' + query[1:]

    @staticmethod
    def parse_html(content):
        """
        解析网页内容
        :param content: 网页内容
        :return:
        """
        arr = []
        _html = etree.HTML(content)  # 将网页源码转换为 XPath 可以解析的格式
        """
        <tr class="datatr">
            <th width="5%" scope="col"><input type="hidden" id="channel50182601" value="gonggao"><input name="datalist" type="checkbox" value="50182601" id="datalist50182601"></th>
            <td align="center" id="channelname50182601">招标公告</td>
            <td id="title50182601"><a href="http://www.bidchance.com/info-gonggao-50182601.html" target="_blank"><span id="onlytitle50182601">安丘市垃圾焚烧发电项目工程监理招标公告</span></a><img src="http://www.bidchance.com/css/biaoshu.gif" width="33" height="15" title="招标文件下载">
            </td>
            <td align="center" id="prov50182601"><a href="http://shandong.bidchance.com" target="_blank">山东</a></td>
            <td align="center" id="pubdate50182601">2020-03-04</td>
        </tr>
        """
        el_tr = _html.xpath('//table[@class="searchaltab-table"]/tbody/tr[@class="datatr"]')
        for tr in el_tr:
            try:
                ele_td = tr.findall('td')
                result = {
                    # 名称
                    'title': ele_td[1].find("a/span").text,
                    # 链接
                    'url': ele_td[1].find("a").attrib.get('href'),
                    # 省份
                    'districtShow': ele_td[2].find("a").text,
                    # 信息类型
                    'stageShow': ele_td[0].text,
                    # 时间
                    'timeShow': ele_td[3].text
                }
                # 标题不为空的加进去
                if result.get("title") != "":
                    arr.append(result)
            except AttributeError as e:
                print("读取tr出错%s, %s:" % (e, html.tostring(tr)))
        return arr

    @staticmethod
    def url_encode(key):
        """
        url编码转换
        :param key:
        :return:
        """
        kv = {'疫情防控': '%D2%DF%C7%E9%B7%C0%BF%D8',
              '复工复产': '%B8%B4%B9%A4%B8%B4%B2%FA',
              '新冠肺炎': '%D0%C2%B9%DA%B7%CE%D1%D7',
              '医用口罩': '%D2%BD%D3%C3%BF%DA%D5%D6',
              '额温枪': '%B6%EE%CE%C2%C7%B9',
              '红外检测': '%BA%EC%CD%E2%BC%EC%B2%E2'}
        return kv.get(key)

    @staticmethod
    def deep_content(url, title):
        """
        爬主要内容，子类需覆盖此方法
        :param url:
        :param title:
        :return:
        """
        content = base.get(url)
        try:
            # 解析网页
            _html = etree.HTML(content)  # 将网页源码转换为 XPath 可以解析的格式
            div = _html.xpath('//dd[@id="infohtmlcon"]')
            return div[0].text.strip()
        except IndexError as e:
            print('主要内容解析失败%s, %s' % (e, html.tostring(_html)))
        except TypeError as e:
            print('主要内容解析失败%s, %s' % (e, html.tostring(_html)))


def main():
    key_words = ['疫情防控', '复工复产', '新冠肺炎', '医用口罩', '额温枪', '红外检测']
    spider = ChinaBid(spider_url, spider_name)
    # 默认取前3页
    spider.process(key_words, 3)


if __name__ == "__main__":
    main()
