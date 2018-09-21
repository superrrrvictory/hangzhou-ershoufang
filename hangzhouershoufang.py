import requests
import json
import time
import xlsxwriter

headers = {
    'Cookie': 'Hm_lvt_70e93e4ca4be30a221d21f76bb9dbdfa=1537470723; ROUTEID=.lb6; JSESSIONID=7E499CFC1310B2AD66937C6D8DB3DD5B.lb6; Hm_lpvt_70e93e4ca4be30a221d21f76bb9dbdfa=1537472188',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/69.0.3497.100 Safari/537.36'}
base_url = 'http://jjhygl.hzfc.gov.cn/webty/WxAction_getGpxxSelectList.jspx?page='


def parse_one_page(url):
    r = requests.get(url, headers=headers)
    hjson = r.json()
    lists = hjson.get('list')
    for i in lists:
        item = {}
        item[u'房屋编号'] = i['fwtybh']
        item[u'房屋面积'] = i['jzmj']
        item[u'委托价格'] = i['wtcsjg']
        item[u'小区名称'] = i['xqmc']
        item[u'地区'] = i['cqmc']
        item[u'挂牌机构'] = i['mdmc']
        item[u'首次挂牌时间'] = i['scgpshsj']
        yield item
    time.sleep(3)
    # print(item)
    # print("=========")
    # print("done")


# #
def save_to_my_computer(result1):
    wb = xlsxwriter.Workbook('./result1.xlsx')
    worksheet = wb.add_worksheet()
    bold_format = wb.add_format({'bold': True})
    money_format = wb.add_format({'num_format': '$#,##0'})
    date_format = wb.add_format({'num_format': 'mmmm d yyyy'})
    worksheet.set_column("A:G", 15)
    worksheet.write('A1', '房屋编号', bold_format)
    worksheet.write('B1', '房屋面积', bold_format)
    worksheet.write('C1', '委托价格', bold_format)
    worksheet.write('D1', '小区名称', bold_format)
    worksheet.write('E1', '地区', bold_format)
    worksheet.write('F1', '挂牌机构', bold_format)
    worksheet.write('G1', '首次挂牌时间', bold_format)
    row = 1
    col = 0
    # 使用write_string方法，指定数据格式写入数据
    for item in result1:
        worksheet.write_string(row, col, str(item['房屋编号']))
        worksheet.write_string(row, col + 1, str(item['房屋面积']))
        worksheet.write_string(row, col + 2, str(item['委托价格']))
        worksheet.write_string(row, col + 3, str(item['小区名称']))
        worksheet.write_string(row, col + 4, str(item['地区']))
        worksheet.write_string(row, col + 5, str(item['挂牌机构']))
        worksheet.write_string(row, col + 6, str(item['首次挂牌时间']))
        row += 1
    wb.close()


def main():
    urls = []
    result1 = []
    print("请输入想要查询的起始页：")
    m = int(input())
    print("请输入想要查询的末页：")
    n = int(input())
    for i in range(m, n):
        urls.append('http://jjhygl.hzfc.gov.cn/webty/WxAction_getGpxxSelectList.jspx?page=' + str(i))
    for k, url in enumerate(urls):
        print("正在加载第" + str(k + 1) + "页")
        results = parse_one_page(url)
        for result in results:
            print(result)
            result1.append(result)
        print(result1)
        save_to_my_computer(result1)
        print("=========")
    # save_to_my_computer(result1)


if __name__ == "__main__":
    main()
