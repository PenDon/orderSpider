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
        jsonData = json.loads(f.read())
        f.close()
    dxmCookie = jsonData['dmx_cookie']
    siteDict = jsonData['site']
    selectionShop = 999
    while selectionShop not in range(len(siteDict.keys()) + 1):
        print(' Please choose')
        print('\t0 : Export all the shop')
        i = 1
        for key in siteDict.keys():
            print('\t' + str(i) + ' : ' + key)
            i = i + 1
        try:
            selectionShop = eval(input("\tPlease choose a shop with integer"))
        except Exception as e:
            print(e, '\n\tPlease enter a valid integer')
            continue
    d = {}
    siteList = []
    if selectionShop:
        siteName = list(siteDict.keys())[selectionShop - 1]
        print('\t选择商铺：', siteName)
        d[siteName] = siteDict[siteName]
        siteList.append(d)
    else:
        siteList = [siteDict]
        print('\t选择更新所有商铺')
    # print(site_list)
    shopifyCookieDict = {}
    shopifyCookie = jsonData['shopify_cookie']
    for name, value in shopifyCookie.items():
        newDict = {}
        for cookieObj in value.split(";"):
            cookieName = cookieObj.split("=")[0]
            cookieName = re.sub('^ ', '', cookieName)
            cookieValue = cookieObj.split("=")[1]
            cookieValue = re.sub('^ ', '', cookieValue)
            newDict[cookieName] = cookieValue
        shopifyCookieDict[name] = newDict
    limitation = jsonData['limitation']  # set the limitation
except Exception as fileException:
    print(fileException, "Read file [ config.json ] failed !")
    time.sleep(5)
    exit(0)


# print(shopify_cookie_dict, dxm_cookie_dict, sep='\n')


def get_id_list(url, coreurl, filename, limitation, site):
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
        for name, value in shopifyCookieDict[site].items():
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

    orderOldList = []

    if os.path.exists(filename):
        with open(filename, 'r') as of:
            orderOldList = of.read().split(',')
        # print(orderOldList)
        # print(type(orderOldList))
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
            if orderOldList:
                if id not in orderOldList:
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


async def getOrderDetail(order_id, session, site, shop_id):
    orderUrl = f'https://{site}.myshopify.com/admin/orders/{order_id}'
    while True:
        intraid = ''
        try:
            async with session.get(orderUrl) as response:
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
                        namesNum = ''
                        # title = titleEle.text
                        title = p.xpath('./parent::div/p')[0]  # p标签
                        titleEle = title  # p标签
                        img = titleEle.xpath('./../../../../../../div[1]/div/img/@src')
                        # ('//section[@id="unfulfilled-card-0"]//a[contains(@href,"admin/products/")]/../../../../../../../div[1]/span/text()')
                        if img:
                            result['img'] = 'https:' + img[0]
                        else:
                            continue
                        title = re.sub(r'\n', '', title.xpath('string(.)'))
                        result['title'] = title.strip()
                        eles = titleEle.xpath('./following-sibling::div//p/text()')
                        result['material'] = eles[0]
                        if len(eles) > 1:
                            skuSplit = eles[1].split(':')
                            if len(skuSplit) > 1:
                                result['SKU'] = eles[1].split(':')[1]
                            else:
                                result['SKU'] = ''
                        else:
                            result['SKU'] = ''
                        # result['material&sku'] = ','.join(eles)
                        num = titleEle.xpath('./../../../../../../div[1]/span/text()')
                        result['num'] = num[0]
                        eles = titleEle.xpath('../../../following-sibling::div//ul/li')
                        names = []
                        for li in eles:
                            string = li.xpath('string(.)')
                            string = re.sub(r'\n', '', string)
                            if string.lower().find('gift') != -1:
                                continue
                            if "_" in string:
                                continue
                            if string.lower().find('number') != -1:
                                namesNum = string.split(':')[1]
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
                        result['namesNum'] = namesNum.strip()
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
                fe.write(f"页面{orderUrl}出现错误,订单号为{intraid}!{traceback.print_exc()} \n")
                fe.close()
            return []


async def run(ids, site, shopId, resultList):
    maxTasks = 20
    headers = {
        'accept': 'text/html, application/xhtml+xml, application/xml',
        'accept-language': 'zh-CN,zh;q=0.9',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.157 Safari/537.36',
        'accept-encoding': 'gzip, deflate, br',
        'x-shopify-web': '1',
    }

    async with aiohttp.ClientSession(headers=headers, cookies=shopifyCookieDict[site]) as session:
        times = (len(ids) - 1) // maxTasks
        for i in range(0, times + 1):
            start = i * maxTasks
            end = maxTasks * (i + 1)
            idsLoop = ids[start: end]
            print(idsLoop)
            tasks = [asyncio.ensure_future(getOrderDetail(id, session, site, shopId)) for id in idsLoop]
            finished, pending = await asyncio.wait(tasks, timeout=90)
            print(f'第{i + 1}批次完成，成功{len(finished)},失败{len(pending)}')
            for i in finished:
                # print("return: ", i.result())
                resultList.extend(i.result())
            time.sleep(1)


def main():
    """
    主函数
    :return:
    """
    for selectSiteDict in siteList:
        for site, shopId in selectSiteDict.items():
            resultList = []
            w = {}
            url = "https://{site}.myshopify.com/admin/orders".format(site=site)
            coreUrl = 'https://{site}.myshopify.com/admin/internal/web/graphql/core'.format(site=site)
            filename = site + "_excel.txt"
            print(f"爬取商铺{site}")
            ids = get_id_list(url, coreUrl, filename, limitation, site)
            if not ids:
                print(f'商铺{site}没有爬取到内容!')
                time.sleep(10)
                continue

            loop = asyncio.get_event_loop()
            try:
                loop.run_until_complete(run(ids, site, shopId, resultList))
            except Exception as ex:
                print("请求订单详细页面出现错误!", ex)
                time.sleep(5)
                exit()
            completedIds = []
            for result in resultList:
                completedIds.append(result['id'])
            contentIds = ''
            with open(filename, 'a+') as f1:
                if f1.read():
                    for completedId in completedIds:
                        contentIds += ',' + completedId
                    f1.write(contentIds)
                else:
                    f1.write(','.join(completedIds))
                f1.close()
            # print(resultList)
            print("正在写入Excel...")
            try:
                if not os.path.exists('excel'):
                    os.mkdir('excel')
                excelName = 'excel/' + site + '.xlsx'
                workbook = xlsxwriter.Workbook(excelName)
                worksheet = workbook.add_worksheet(site)
                titleFormat = workbook.add_format()
                titleFormat.set_bold()
                titleFormat.set_font_size(15)

                cellFormat = workbook.add_format()
                cellFormat.set_align('vcenter')
                cellFormat.set_font_name('微软雅黑')

                textFormat = workbook.add_format()
                textFormat.set_text_wrap()
                textFormat.set_font_name('微软雅黑')
                columns = {"A": "序号", "B": "订单号", "C": "Material", "D": "SKU", "E": "产品名称", "F": "产品图片", "G": "刻字信息",
                           "H": "名字数量", "I": "产品数量", "J": "备注"}
                columnsSize = {"0": 5, "1": 15, "2": 40, "3": 20, "4": 50, "5": 20, "6": 30,
                                "7": 15, "8": 15, "9": 20}
                for key, value in columnsSize.items():
                    worksheet.set_column(int(key), int(key), value)
                for key, value in columns.items():
                    worksheet.write(key + '1', value, titleFormat)
                row = 2
                for result in resultList:
                    num = re.findall(r'([0-9]+)', result['names_num'])
                    if not num:
                        with open('error_excel.txt', 'a+') as fe:
                            fe.write(result['订单号'] + '数量出错\n')
                            fe.close()
                            num = 0
                    worksheet.write(f'A{row}', row - 1, cellFormat)
                    worksheet.write(f'B{row}', result['订单号'], cellFormat)
                    worksheet.write(f'C{row}', result['material'], cellFormat)
                    worksheet.write(f'D{row}', result['SKU'], cellFormat)
                    worksheet.write(f'E{row}', result['title'], cellFormat)
                    image_data = BytesIO(urlopen(result['img']).read())
                    worksheet.insert_image(row - 1, 5, result['img'], {'image_data': image_data})
                    num = int(num[0])
                    if num != 0 and num != len(result['names'].split(',')):
                        colorFormat = workbook.add_format()
                        colorFormat.set_text_wrap()
                        colorFormat.set_font_name('微软雅黑')
                        colorFormat.set_bg_color('red')
                        worksheet.write(f'G{row}', result['names'], colorFormat)
                    else:
                        worksheet.write(f'G{row}', result['names'], textFormat)
                    worksheet.write(f'H{row}', result['names_num'], cellFormat)
                    worksheet.write(f'I{row}', result['num'], cellFormat)
                    worksheet.set_row(row - 1, 100)
                    row += 1

                print("Done!")
                workbook.close()
            except Exception as ex:
                print("Error:")
                print(ex)
            time.sleep(5)


main()
