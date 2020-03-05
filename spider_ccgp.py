# 中国政府网
import spider_common as base
from lxml import etree
from lxml import html

# 主站
url = "http://search.ccgp.gov.cn/bxsearch"


def build_params(time_type, kw):
    """
    请求参数
    :param timeType: 近期天数
    :param kw: 搜索内容
    :return:
    """
    data = {
        'searchtype': 1,
        'page_index': 1,
        'start_time': '',
        'end_time': '',
        'timeType': time_type,
        'searchparam': '',
        'searchchannel': 0,
        'dbselect': 'bid',
        'kw': kw,
        'bidSort': 0,
        'pinMu': 0,
        'bidType': 0,
        'buyerName': '',
        'projectId': '',
        'displayZone': '',
        'zoneId': '',
        'agentName': ''
    }
    # 拼接字符串
    query = ''
    for key in data:
        query += '&' + key + '=' + str(data.get(key))
    return '?' + query[1:]


def parse_data(content):
    """
    解析网页内容
    :param content: 网页内容
    :return:
    """
    arr = []
    _html = etree.HTML(content)  # 将网页源码转换为 XPath 可以解析的格式
    """
    <li>
        <a href="http://www.ccgp.gov.cn/cggg/dfgg/gzgg/202003/t20200304_13957968.htm" style="line-height:18px" target="_blank">
            德州市公安局<font color="red">互联网</font>门户网站建设项目（恢复）更正公告
        </a>
        <p>德州市公安局互联网门户网站建设项目（恢复）更正公告一、采购人：德州市公安局地址：德州市经济开发区长河大道198号联系方式：17605349909采购代理机构：山东信一项目管理有限公司地址：山东省烟台市莱山区县（区）</p>
        <span>2020.03.04 16:02:21
            | 采购人：德州市公安局
            | 代理机构：山东信一项目管理有限公司
            <br>
            <strong style="font-weight:bolder">
                    更正公告
            </strong>
            | <a href="javascript:void(0)">山东</a>
            | <strong style="font-weight:bolder"> </strong>
        </span>
    </li>
    """
    el_li = _html.xpath('//ul[@class="vT-srch-result-list-bid"]/li')
    for _li in el_li:
        result = {}
        # 名称
        result['name'] = _li.find("a").text.strip()
        # 链接
        result['link'] = _li.find("a").attrib.get('href')
        # 省份
        result['province'] = _li.find("span/a").text.strip()
        # 时间
        result['time'] = _li.find('span').text.strip()[0:10]
        # 内容
        result['content'] = _li.find("p").text.strip()
        arr.append(result)
    return arr


def main():
    # 互联网
    _url = url + build_params(2, '%E4%BA%92%E8%81%94%E7%BD%91')
    content = base.get(_url)
    arr = parse_data(content)
    print(arr)


if __name__ == "__main__":
    main()
