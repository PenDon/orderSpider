# _*_coding:utf-8_*_
import json
import re
import time
import aiohttp
import asyncio
import requests
from lxml import etree
import xlsxwriter
import os
from io import BytesIO
from urllib.request import urlopen
import traceback

try:
    with open("./config.json", 'r') as f:
        json_data = json.loads(f.read())
        f.close()
except Exception as fileException:
    print(fileException, "读取config.json文件失败，请检查!")
    time.sleep(5)
    exit(0)
dxm_cookie = json_data['dmx_cookie']
selection_1 = 999
selection_2 = 999
site_dict = json_data['site']
while selection_2 not in range(len(site_dict.keys()) + 1):
    print('  请选择：')
    print('\t0 : 导出所有商铺')
    i = 1
    for key in site_dict.keys():
        print('\t' + str(i) + ' : ' + key)
        i = i + 1
    try:
        selection_2 = eval(input("\t请选择商铺名（根据上表输入一个整数）："))
    except Exception as e:
        print(e, '\n\t请输入一个有效的整数')
        continue
d = {}
site_list = []
if selection_2:
    site_name = list(site_dict.keys())[selection_2 - 1]
    print('\t选择商铺：', site_name)
    d[site_name] = site_dict[site_name]
    site_list.append(d)
else:
    site_list = [site_dict]
    print('\t选择更新所有商铺')
# print(site_list)
shopify_cookie_dict = {}
shopify_cookie = json_data['shopify_cookie']
for name, value in shopify_cookie.items():
    newDict = {}
    for cookieobj in value.split(";"):
        cookie_name = cookieobj.split("=")[0]
        cookie_name = re.sub('^ ', '', cookie_name)
        cookie_value = cookieobj.split("=")[1]
        cookie_value = re.sub('^ ', '', cookie_value)
        newDict[cookie_name] = cookie_value
    shopify_cookie_dict[name] = newDict
limitation = json_data['limitation']  # config中添加字段limitation限制爬取条数



# print(shopify_cookie_dict, dxm_cookie_dict, sep='\n')


def get_id_list(url, coreurl, order_number_filename, limitation, site):
    headers = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
        'accept-language': 'zh-CN,zh;q=0.9',
        'cache-control': 'no-cache',
        'pragma': 'no-cache',
        'upgrade-insecure-requests': '1',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.157 Safari/537.36',
        'accept-encoding': 'gzip, deflate, br',
    }

    s = requests.session()
    s.headers.update(headers)
    # print(shopify_cookie_dict[site])
    try:
        for name, value in shopify_cookie_dict[site].items():
            s.cookies.set(name, value)
        page = s.get(url)
    except Exception as ide:
        print(ide, f"无法访问{site}!请检查")
        return []
    tree = etree.HTML(page.text)
    r = tree.xpath('//script[@data-serialized-id="csrf-token"]/text()')
    if not r:
        print("响应没有内容，请检查。可能的原因：shopify cookie过期,请重新获取")
        time.sleep(10)
        return []
    token = r[0].replace('"', '')
    print('token:', token)

    s.headers.update({
        'content-type': 'application/json',
        'x-csrf-token': token,
        'accept': 'application/json',
    })
    # coreurl = 'https://loveintheboxstore.myshopify.com/admin/internal/web/graphql/core'
    data = {
        "operationName": "OrderIndex",
        "variables": {
            "ordersFirst": 50,
            "ordersLast": None,
            "query": "fulfillment_status:'unshipped'",
            "before": None,
            "after": None,
            "sortKey": "PROCESSED_AT",
            "reverse": True,
            "savedSearchId": None,
            "skipCustomer": True,
            "skipAppsAttributedToOrders": True
        },
        "query": "query OrderIndex($ordersFirst: Int, $ordersLast: Int, $before: String, $after: String, $sortKey: OrderSortKeys, $reverse: Boolean, $query: String, $savedSearchId: ID, $locationId: ID, $skipCustomer: Boolean!, $skipAppsAttributedToOrders: Boolean!) {\n  location(id: $locationId) {\n    name\n    __typename\n  }\n  staffMember {\n    id\n    isShopOwner\n    __typename\n  }\n  shop {\n    id\n    ordersDelayedSince\n    appLinks(type: ORDERS, location: INDEX) {\n      id\n      text\n      url\n      icon {\n        transformedSrc(maxWidth: 80)\n        __typename\n      }\n      __typename\n    }\n    appActions: appLinks(type: ORDERS, location: ACTION) {\n      id\n      text\n      url\n      icon {\n        transformedSrc(maxWidth: 80)\n        __typename\n      }\n      __typename\n    }\n    plan {\n      trial\n      __typename\n    }\n    showInstallMobileAppBanner\n    features {\n      fraudProtectEligibility\n      eligibleForBulkLabelPurchase\n      __typename\n    }\n    currencyCode\n    __typename\n  }\n  shopHasOrders: orders(first: 1, reverse: true) {\n    edges {\n      node {\n        id\n        __typename\n      }\n      __typename\n    }\n    __typename\n  }\n  orders(first: $ordersFirst, after: $after, last: $ordersLast, before: $before, sortKey: $sortKey, reverse: $reverse, query: $query, savedSearchId: $savedSearchId) {\n    edges {\n      cursor\n      node {\n        ...OrderListItem\n        customer @skip(if: $skipCustomer) {\n          id\n          firstName\n          lastName\n          email\n          __typename\n        }\n        __typename\n      }\n      __typename\n    }\n    pageInfo {\n      hasNextPage\n      hasPreviousPage\n      __typename\n    }\n    __typename\n  }\n  appsAttributedToOrders @skip(if: $skipAppsAttributedToOrders) {\n    id\n    title\n    __typename\n  }\n}\n\nfragment OrderListItem on Order {\n  id\n  name\n  note\n  closed\n  cancelledAt\n  processedAt\n  hasPurchasedShippingLabels\n  hasTimelineComment\n  displayFinancialStatus\n  displayFulfillmentStatus\n  multiCurrencyStatus {\n    message\n    __typename\n  }\n  totalPriceSet {\n    shopMoney {\n      amount\n      currencyCode\n      __typename\n    }\n    presentmentMoney {\n      amount\n      currencyCode\n      __typename\n    }\n    __typename\n  }\n  expiringAuthorization\n  riskRecommendation {\n    recommendation\n    __typename\n  }\n  fraudProtection {\n    level\n    __typename\n  }\n  disputes {\n    status\n    protectedByFraudProtect\n    initiatedAs\n    id\n    __typename\n  }\n  __typename\n}\n"
    }
    ids = []
    i = 1

    order_old_list = []

    if os.path.exists(order_number_filename):
        with open(order_number_filename, 'r') as of:
            order_old_list = of.read().split(',')
        # print(order_old_list)
        # print(type(order_old_list))
        # exit()
    flag = True
    while flag:
        while True:
            try:
                res = s.post(coreurl, json=data, timeout=10)
                break
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                print(e)
                continue
        res = res.json()
        edges = res['data']['orders']['edges']
        pat = re.compile(r'([0-9]+)')
        print(f'读取shopify商铺{site}第{i}页')
        time.sleep(2)
        for item in edges:
            node = item['node']
            id = node['id']
            id = pat.search(id).group(1)
            if order_old_list:
                if id not in order_old_list:
                    print(node['name'] + '加入待处理队列')
                    ids.append(id)
                else:
                    print(f"{node['name']}已处理，忽略!")
                    continue
            else:
                print(node['name'] + '加入待处理队列')
                ids.append(id)
            if len(ids) >= limitation:
                flag = False
                print('待处理订单条数达到爬取条数上限，停止')
                break
        if i * 50 + 1 > limitation:
            print('页数达到爬取条数上限，停止')
            break
        cursor = edges[-1]['cursor']
        data['variables']['after'] = cursor
        if len(edges) < 50:
            break
        time.sleep(2)
        i += 1
    return ids


async def get_order_detail(order_id, session, site, shop_id):
    order_url = f'https://{site}.myshopify.com/admin/orders/{order_id}'
    while True:
        intraid = ''
        try:
            async with session.get(order_url) as response:
                if response.status == 200:
                    response = await response.text()
                    results = []
                    tree = etree.HTML(response)
                    intraid = tree.xpath('//h1/text()')
                    intraid = intraid[0]
                    orderId = re.findall('\d+', intraid)
                    status = tree.xpath('//h2/text()')
                    r = tree.xpath(
                        '//section[contains(@id, "unfulfilled-card-")]//div[@class="unfulfilled-card__line_item__secondary-details"]')
                    # r = tree.xpath('//section[@id="unfulfilled-card-0"]//a[contains(@href,"admin/products/")]')
                    # $x('//*[@class="unfulfilled-card__line_item__secondary-details"]')
                    for p in r:

                        result = {'订单号': f'{str(shop_id)}-' + orderId[0], 'id': order_id}
                        names_num = ''
                        # title = title_ele.text
                        title = p.xpath('./parent::div/p')[0]  # p标签
                        title_ele = title  # p标签
                        img = title_ele.xpath('./../../../../../../div[1]/div/img/@src')
                        # ('//section[@id="unfulfilled-card-0"]//a[contains(@href,"admin/products/")]/../../../../../../../div[1]/span/text()')
                        if img:
                            result['img'] = 'https:' + img[0]
                        else:
                            continue
                        title = re.sub(r'\n', '', title.xpath('string(.)'))
                        result['title'] = title.strip()
                        eles = title_ele.xpath('./following-sibling::div//p/text()')
                        result['material'] = eles[0]
                        if len(eles) > 1:
                            sku_split = eles[1].split(':')
                            if len(sku_split) > 1:
                                result['SKU'] = eles[1].split(':')[1]
                            else:
                                result['SKU'] = ''
                        else:
                            result['SKU'] = ''
                        # result['material&sku'] = ','.join(eles)
                        num = title_ele.xpath('./../../../../../../div[1]/span/text()')
                        result['num'] = num[0]
                        eles = title_ele.xpath('../../../following-sibling::div//ul/li')
                        names = []
                        for li in eles:
                            string = li.xpath('string(.)')
                            string = re.sub(r'\n', '', string)
                            if string.lower().find('gift') != -1:
                                continue
                            if "_" in string:
                                continue
                            if string.lower().find('number') != -1:
                                names_num = string.split(':')[1]
                                continue
                            if string.lower().find('size') != -1:
                                result['material'] += ' / ' + ''.join(string.strip().split())
                                continue
                            # string = string.split(':')
                            string = string.split(':')
                            name = string[1].strip()
                            if not name:
                                continue
                            names.append(string[1].strip())
                        result['names'] = ','.join(names)
                        result['names_num'] = names_num.strip()
                        results.append(result)
                    # print("res:", results)
                    return results
                else:
                    return -1
        except (aiohttp.client_exceptions.ServerDisconnectedError, aiohttp.client_exceptions.ClientConnectorError) as e:
            print(e)
            time.sleep(1)
            continue
        except Exception as e:
            print(intraid, e)
            with open('error_excel.txt', 'a+') as fe:
                fe.write(f"页面{order_url}出现错误,订单号为{intraid}!{traceback.print_exc()} \n")
                fe.close()
            return []


async def run(ids, site, shop_id, result_list):
    max_tasks_num = 20
    headers = {
        'accept': 'text/html, application/xhtml+xml, application/xml',
        'accept-language': 'zh-CN,zh;q=0.9',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.157 Safari/537.36',
        'accept-encoding': 'gzip, deflate, br',
        'x-shopify-web': '1',
    }

    async with aiohttp.ClientSession(headers=headers, cookies=shopify_cookie_dict[site]) as session:
        times = (len(ids) - 1) // max_tasks_num
        for i in range(0, times + 1):
            start = i * max_tasks_num
            end = max_tasks_num * (i + 1)
            ids_loop = ids[start: end]
            print(ids_loop)
            tasks = [asyncio.ensure_future(get_order_detail(id, session, site, shop_id)) for id in ids_loop]
            finished, pending = await asyncio.wait(tasks, timeout=90)
            print(f'第{i + 1}批次完成，成功{len(finished)},失败{len(pending)}')
            for i in finished:
                # print("return: ", i.result())
                result_list.extend(i.result())
            time.sleep(1)

    # pdf = df.DataFrame(result_list)
    # pdf.to_excel('./out.xlsx',columns=['title','内部id','订单id','sku','num','names','状态'],index=False)


def main():
    """
    主函数
    :return:
    """
    for select_site_dict in site_list:
        for site, shop_id in select_site_dict.items():
            result_list = []
            w = {}
            url = "https://{site}.myshopify.com/admin/orders".format(site=site)
            coreurl = 'https://{site}.myshopify.com/admin/internal/web/graphql/core'.format(site=site)
            order_number_filename = site + "_excel.txt"
            print(f"爬取商铺{site}")
            ids = get_id_list(url, coreurl, order_number_filename, limitation, site)
            if not ids:
                print(f'商铺{site}没有爬取到内容!')
                time.sleep(10)
                continue

            loop = asyncio.get_event_loop()
            try:
                loop.run_until_complete(run(ids, site, shop_id, result_list))
            except Exception as ex:
                print("请求订单详细页面出现错误!", ex)
                time.sleep(5)
                exit()
            completed_ids = []
            for result in result_list:
                completed_ids.append(result['id'])
            content_ids = ''
            with open(order_number_filename, 'a+') as f1:
                if f1.read():
                    for completed_id in completed_ids:
                        content_ids += ',' + completed_id
                    f1.write(content_ids)
                else:
                    f1.write(','.join(completed_ids))
                f1.close()
            # print(result_list)
            print("正在写入Excel...")
            try:
                if not os.path.exists('excel'):
                    os.mkdir('excel')
                excel_name = 'excel/' + site + '.xlsx'
                workbook = xlsxwriter.Workbook(excel_name)
                worksheet = workbook.add_worksheet(site)
                title_format = workbook.add_format()
                title_format.set_bold()
                title_format.set_font_size(15)

                cell_format = workbook.add_format()
                cell_format.set_align('vcenter')
                cell_format.set_font_name('微软雅黑')

                text_format = workbook.add_format()
                text_format.set_text_wrap()
                text_format.set_font_name('微软雅黑')
                columns = {"A": "序号", "B": "订单号", "C": "Material", "D": "SKU", "E": "产品名称", "F": "产品图片", "G": "刻字信息",
                           "H": "名字数量", "I": "产品数量", "J": "备注"}
                columns_size = {"0": 5, "1": 15, "2": 40, "3": 20, "4": 50, "5": 20, "6": 30,
                                "7": 15, "8": 15, "9": 20}
                for key, value in columns_size.items():
                    worksheet.set_column(int(key), int(key), value)
                for key, value in columns.items():
                    worksheet.write(key + '1', value, title_format)
                row = 2
                for result in result_list:
                    num = re.findall(r'([0-9]+)', result['names_num'])
                    if not num:
                        with open('error_excel.txt','a+') as fe:
                            fe.write(result['订单号'] + '数量出错\n')
                            fe.close()
                            num = 0
                    worksheet.write(f'A{row}', row - 1, cell_format)
                    worksheet.write(f'B{row}', result['订单号'], cell_format)
                    worksheet.write(f'C{row}', result['material'], cell_format)
                    worksheet.write(f'D{row}', result['SKU'], cell_format)
                    worksheet.write(f'E{row}', result['title'], cell_format)
                    image_data = BytesIO(urlopen(result['img']).read())
                    worksheet.insert_image(row - 1, 5, result['img'], {'image_data': image_data})
                    num = int(num[0])
                    if num != 0 and num != len(result['names'].split(',')):
                        color_format = workbook.add_format()
                        color_format.set_text_wrap()
                        color_format.set_font_name('微软雅黑')
                        color_format.set_bg_color('red')
                        worksheet.write(f'G{row}', result['names'], color_format)
                    else:
                        worksheet.write(f'G{row}', result['names'], text_format)
                    worksheet.write(f'H{row}', result['names_num'], cell_format)
                    worksheet.write(f'I{row}', result['num'], cell_format)
                    worksheet.set_row(row - 1, 100)
                    row += 1

                print("Done!")
                workbook.close()
            except Exception as ex:
                print("Error:")
                print(ex)
            time.sleep(5)


main()
