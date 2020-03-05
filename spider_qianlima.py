# 中国政府网
import spider_common as base
from lxml import etree
from lxml import html

# 主站
url = "http://zb.yfb.qianlima.com/yfbsemsite/mesinfo/zbpglist"


def build_params(searchword):
    """
    请求参数
    :param timeType: 近期天数
    :param kw: 搜索内容
    :return:
    """
    data = {
        'pageNo': 1,
        'pageSize': 15,
        'searchword': '互联网',
        'searchword2': '',
        'hotword': '',
        'provinceId': '2703',
        'provinceName': '',
        'areaId': '2703',
        'areaName': '',
        'infoType': '1',
        'infoTypeName': '',
        'noticeTypes': 0,
        'noticeTypesName:': '',
        'secondInfoType': '',
        'secondInfoTypeName': '',
        'timeType': 1,
        'timeTypeName': '近一周',
        'searchType': '2',
        'clearAll': 'false',
        'e_keywordid': '116479021597',
        'e_creative': '28381277451',
        'flag': 1,
        'source': 'baidu',
        'firstTime': 1,
    }
    return data


def parse_data(content):
    """
    解析网页内容
    :param content: 网页内容
    :return:
    """
    arr = []
    _html = etree.HTML(content)  # 将网页源码转换为 XPath 可以解析的格式
    """
    <tr>
        <td style="text-align: center">2020-03-04</td>
        <td style="text-align: center">北京</td>
        <td style="text-align: center">招标</td>
        <td style="text-align: left">
            <a href="javascript:;" onclick="popUpQRcodeImg('172046910','172038256')">
                    2020-2022年北京联通网络交付运营中心公众客户支撑中心、<font color="red">互联网</font>中心、智能云中心维保服务项目（华为设备）招标公告
            </a>
        </td>
    </tr>
    """
    el_tr = _html.xpath('//table[@id="contentTable"]/tbody/tr')
    for tr in el_tr:
        ele_td = tr.findall('td')
        result = {}
        # 名称
        result['name'] = ele_td[3].find('a').text.strip()
        # 链接
        result['link'] = ele_td[3].find('a').attrib.get('href')
        # 省份
        result['province'] = ele_td[1].text.strip()
        # 时间
        result['time'] = ele_td[0].text.strip()
        arr.append(result)
    return arr


def main():
    # 互联网
    content = base.post(url, build_params('互联网'))
    arr = parse_data(content)
    print(arr)


if __name__ == "__main__":
    main()
