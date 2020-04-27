# _*_coding:utf-8_*_
import json
import re
import time
import aiohttp
import asyncio
import requests
from lxml import etree
from bs4 import BeautifulSoup

with open("./config.json", 'r') as f:
    json_data = json.loads(f.read())
dxm_cookie = json_data['dmx_cookie']
selection_1 = 999
selection_2 = 999
while selection_1 not in range(len(dxm_cookie.keys()) + 1):
    print('  Please select: ')
    print('\t0 : update all the dxm account')
    i = 1
    for key in dxm_cookie.keys():
        print('\t' + str(i) + ' : ' + key)
        i = i + 1
    try:
        selection_1 = eval(input("\tPlease select a dxm account to update (Enter a number according to the table above):"))
    except Exception as e:
        print(e, '\n\t请输入一个有效的整数')
        continue
if selection_1:
    dxm_name = list(dxm_cookie.keys())[selection_1 - 1]
    print('\t选择帐号：', dxm_name)
    dxm_cookie = [dxm_cookie[dxm_name]]
    dxm_name = [dxm_name]
else:
    dxm_name = list(dxm_cookie.keys())
    dxm_cookie = list(dxm_cookie.values())
    print('\t选择更新所有店小秘帐号')
print(dxm_name)
time.sleep(2)
dxm_cookie_dict_list = []
for i in range(len(dxm_cookie)):
    dxm_cookie_dict = {}
    for cookieobj in dxm_cookie[i].split(";"):
        name = cookieobj.split("=")[0]
        value = cookieobj.split("=")[1]
        dxm_cookie_dict[name] = value
    dxm_cookie_dict_list.append(dxm_cookie_dict)
site_dict = json_data['site']
while selection_2 not in range(len(site_dict.keys()) + 1):
    print('  请选择：')
    print('\t0 : 更新所有商铺')
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

result_list = []


# print(shopify_cookie_dict, dxm_cookie_dict, sep='\n')


def get_id_list(url, coreurl, site, dxm_list):
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
    for name, value in shopify_cookie_dict[site].items():
        s.cookies.set(name, value)

    page = s.get(url)
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
    ids = []
    for data in dxm_list:
        order_number = data['order_number']
        data = {
            "operationName": "OrderIndex",
            "variables": {
                "ordersFirst": 50,
                "ordersLast": None,
                "before": None,
                "after": None,
                "query": order_number,
                "sortKey": "PROCESSED_AT",
                "reverse": True,
                "savedSearchId": None,
                "skipCustomer": True,
                "skipAppsAttributedToOrders": True
            },
            "query": "query OrderIndex($ordersFirst: Int, $ordersLast: Int, $before: String, $after: String, $sortKey: OrderSortKeys, $reverse: Boolean, $query: String, $savedSearchId: ID, $locationId: ID, $skipCustomer: Boolean!, $skipAppsAttributedToOrders: Boolean!) {\n  location(id: $locationId) {\n    name\n    __typename\n  }\n  staffMember {\n    id\n    isShopOwner\n    __typename\n  }\n  shop {\n    id\n    ordersDelayedSince\n    appLinks(type: ORDERS, location: INDEX) {\n      id\n      text\n      url\n      icon {\n        transformedSrc(maxWidth: 80)\n        __typename\n      }\n      __typename\n    }\n    appActions: appLinks(type: ORDERS, location: ACTION) {\n      id\n      text\n      url\n      icon {\n        transformedSrc(maxWidth: 80)\n        __typename\n      }\n      __typename\n    }\n    plan {\n      trial\n      __typename\n    }\n    showInstallMobileAppBanner\n    features {\n      fraudProtectEligibility\n      eligibleForBulkLabelPurchase\n      __typename\n    }\n    currencyCode\n    __typename\n  }\n  shopHasOrders: orders(first: 1, reverse: true) {\n    edges {\n      node {\n        id\n        __typename\n      }\n      __typename\n    }\n    __typename\n  }\n  orders(first: $ordersFirst, after: $after, last: $ordersLast, before: $before, sortKey: $sortKey, reverse: $reverse, query: $query, savedSearchId: $savedSearchId) {\n    edges {\n      cursor\n      node {\n        ...OrderListItem\n        customer @skip(if: $skipCustomer) {\n          id\n          firstName\n          lastName\n          email\n          __typename\n        }\n        __typename\n      }\n      __typename\n    }\n    pageInfo {\n      hasNextPage\n      hasPreviousPage\n      __typename\n    }\n    __typename\n  }\n  appsAttributedToOrders @skip(if: $skipAppsAttributedToOrders) {\n    id\n    title\n    __typename\n  }\n}\n\nfragment OrderListItem on Order {\n  id\n  name\n  note\n  closed\n  cancelledAt\n  processedAt\n  hasPurchasedShippingLabels\n  hasTimelineComment\n  displayFinancialStatus\n  displayFulfillmentStatus\n  multiCurrencyStatus {\n    message\n    __typename\n  }\n  totalPriceSet {\n    shopMoney {\n      amount\n      currencyCode\n      __typename\n    }\n    presentmentMoney {\n      amount\n      currencyCode\n      __typename\n    }\n    __typename\n  }\n  expiringAuthorization\n  riskRecommendation {\n    recommendation\n    __typename\n  }\n  fraudProtection {\n    level\n    __typename\n  }\n  disputes {\n    status\n    protectedByFraudProtect\n    initiatedAs\n    id\n    __typename\n  }\n  __typename\n}\n"
        }
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
        print(f'查询shopify商铺{site} orderId 为 {order_number}')
        if len(edges) :
            item = edges[0]
            node = item['node']
            id = node['id']
            id = pat.search(id).group(1)
            ids.append(id)
        else :
            print(f"shopify商铺{site}中未查询到 orderId : {order_number}")
        # time.sleep(1)
    return ids


async def get_order_detail(order_id, session, site):
    order_url = f'https://{site}.myshopify.com/admin/orders/{order_id}'
    while True:
        try:
            async with session.get(order_url) as response:
                if response.status == 200:
                    response = await response.text()
                    results = []
                    tree = etree.HTML(response)
                    intraid = tree.xpath('//h1/text()')
                    intraid = intraid[0]
                    status = tree.xpath('//h2/text()')
                    status = status[0]
                    r = tree.xpath('//section[contains(@id, "unfulfilled-card-")]//div[@class="unfulfilled-card__line_item__secondary-details"]')
                    # r = tree.xpath('//section[@id="unfulfilled-card-0"]//a[contains(@href,"admin/products/")]')
                    # $x('//*[@class="unfulfilled-card__line_item__secondary-details"]')
                    for p in r:
                        result = {'订单id': order_id, '内部id': intraid, '状态': status}
                        # title = title_ele.text
                        title = p.xpath('./parent::div/p')[0]   # p标签
                        title_ele = title  # p标签

                        title = re.sub(r'\n', '', title.xpath('string(.)'))
                        result['title'] = title

                        eles = title_ele.xpath('./following-sibling::div/span/p/text()')
                        result['sku'] = ','.join(eles)
                        num = title_ele.xpath('./../../../../../../div[1]/span/text()')
                        # ('//section[@id="unfulfilled-card-0"]//a[contains(@href,"admin/products/")]/../../../../../../../div[1]/span/text()')
                        result['num'] = num[0]
                        eles = title_ele.xpath('../../../following-sibling::div//ul/li')
                        names = []
                        for li in eles:
                            string = li.xpath('string(.)')
                            string = re.sub(r'\n', '', string)
                            if "_" in string:
                                continue
                            names.append(string)
                        result['names'] = '、'.join(names)
                        results.append(result)
                    # print("res:", results)
                    # exit(0)
                    return results
                else:
                    return -1
        except (aiohttp.client_exceptions.ServerDisconnectedError, aiohttp.client_exceptions.ClientConnectorError) as e:
            print(e)
            time.sleep(1)
            continue
        except Exception as e:
            print(e)
            continue


async def run(ids, site):
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
            tasks = [asyncio.ensure_future(get_order_detail(id, session, site)) for id in ids_loop]
            finished, pending = await asyncio.wait(tasks, timeout=90)
            print(f'第{i + 1}批次完成，成功{len(finished)},失败{len(pending)}')
            for i in finished:
                # print("return: ", i.result())
                result_list.extend(i.result())
            time.sleep(2)

    # pdf = df.DataFrame(result_list)
    # pdf.to_excel('./out.xlsx',columns=['title','内部id','订单id','sku','num','names','状态'],index=False)


# 店小蜜

def get_list(shopId, prefix ,cookie_dict, limit):
    url = "https://www.dianxiaomi.com/package/list.htm?pageNo={pageNo}&pageSize=30&shopId={shopId}&state=approved"
    headers = {
        'accept': 'text/html, */*; q=0.01',
        'accept-language': 'zh-CN,zh;q=0.9',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.157 Safari/537.36',
        'accept-encoding': 'gzip, deflate, br',
        'Host': 'www.dianxiaomi.com',
    }
    # print(cookie_dict)
    s = requests.session()
    s.headers.update(headers)
    data = []
    for name, value in cookie_dict.items():
        s.cookies.set(name, value)
    i = 1
    last_order = None
    while True:
        while True:
            try:
                page = s.get(url.format(pageNo=i, shopId=shopId))
                break
            except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                print(e)
                continue

        soup = BeautifulSoup(page.text, 'lxml')
        print(f'商铺号{shopId}第{i}页')
        first_order = soup.select_one(".tableOrderId a")

        if first_order:
            first_order = first_order.get_text()
        else:
            break

        if last_order is None:
            last_order = first_order
        else:
            if last_order == first_order:
                break
            else:
                last_order = first_order

        OrderId = soup.select(".tableOrderId a")

        for order in OrderId:
            packageIds = re.findall('\d+', str(order['onclick']))
            order_number = order.get_text()
            order_number = str(order_number).split("-")

            data.append({'packageId': packageIds[0], 'order_number': prefix + order_number[1]})
        # print(f'dxm_data:\n{data}')
        i += 1
        if len(data) > limit:
            break
        time.sleep(2)
    return data


def write_remark(w, order_number_filename, cookie_dict):
    """
    写入备注
    :param w:
    :param order_number_filename:
    :param cookie_dict:
    :return:
    """
    print("写入备注")
    error_list = []
    success_dict = {}
    url = 'https://www.dianxiaomi.com/dxmPackageComment/add.json'
    has_remark_url = "https://www.dianxiaomi.com/dxmPackageComment/getByPackId.json"
    headers = {
        'accept': 'application/json, text/javascript, */*; q=0.01',
        'accept-language': 'zh-CN,zh;q=0.9',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.157 Safari/537.36',
        'accept-encoding': 'gzip, deflate, br',
        'Host': 'www.dianxiaomi.com',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
    }
    s = requests.session()
    s.headers.update(headers)
    for name, value in cookie_dict.items():
        s.cookies.set(name, value)
    order_new_id = []
    for k, v in w.items():
        print(f"内部ID为{v['intern_id']}正在监测")
        is_success = True
        has_remark = s.post(has_remark_url, {"packageId": k})
        if has_remark.json()['dxmPackageComment'] is not None:
            print(f"内部ID为{v['intern_id']}已经有备注，跳过")
            continue
        content = ""
        for i in v['remake'].split("、"):
            try:
                i = i.replace("::", ":")
                # i = i.replace(" ", "")
                remake = i.split(":")
                if '数量' in i:
                    content += i + "\n"
                else:
                    if remake[0] != '_boldOptionLocalStorageId':
                        if remake[1] != "":
                            content += remake[1] + "\n"
            except Exception as e1:
                print(f"内部ID为{v['intern_id']}出错{e1}")
                is_success = False
                error_list.append(v['intern_id'])
                break
        if is_success:
            if success_dict.get(v['intern_id'], None) is None:
                success_dict[v['intern_id']] = content
        else:
            continue

        data = {
            "packageId": k,
            "commentType": "sys_picking",
            "content": content,
            "color": "F00",
        }
        print(f"内部ID为{v['intern_id']}正在提交数据\n{content}")
        res = s.post(url, data=data)
        if res.status_code == 200:
            print(res.json())
            print(f'packageId为{k}完成')
            order_new_id.append(v['order_id'])
            time.sleep(2)

    with open(order_number_filename, 'a') as f1:
        f1.write(",".join(order_new_id))
        f1.close()

    with open("error.txt", 'a+') as fe:
        fe.write(",".join(error_list) + "\n")
        fe.close()

    with open("success.txt", 'a+', encoding='utf-8') as fs:
        for key, value in success_dict.items():
            fs.write("内部id：" + key + ", 写入内容：" + value + "\n")
        fs.close()


def match_goods_title(title):
    """
    匹配需要写入备注的商品名
    :param title:
    :return:
    """
    with open('config.json', 'r+') as f2:
        goods_names = json.loads(f2.read())['goods_names']
        for goods_name in goods_names:
            if re.search(goods_name, title):
                return True
            else:
                continue
        f2.close()
        return False


def main():
    """
    主函数
    :return:
    """
    for select_site_dict in site_list:
        for site, shop in select_site_dict.items():
            shop_id = shop['shop_id']
            shop_prefix = shop['prefix']

            w = {}
            url = "https://{site}.myshopify.com/admin/orders".format(site=site)
            coreurl = 'https://{site}.myshopify.com/admin/internal/web/graphql/core'.format(site=site)
            order_number_filename = site + ".txt"

            j = 0
            for cookie_dict in dxm_cookie_dict_list:
                print(f"读取店小秘帐号{dxm_name[j]},")
                dxm_list = get_list(shop_id, shop_prefix, cookie_dict, limitation)
                # print("result_list:", result_list)
                if not dxm_list:
                    print("店小蜜没有查询到订单。可能是cookie有错误，请检查")
                    time.sleep(3)
                    continue
                print(f' > 店小秘帐号{dxm_name[j]}一共{len(dxm_list)}\n')
                print(f" > 查询shopify商铺{site}")
                ids = get_id_list(url, coreurl, site, dxm_list)
                if not ids:
                    print(f'商铺{site}没有爬取到内容!')
                    time.sleep(10)
                    continue
                loop = asyncio.get_event_loop()
                try:
                    loop.run_until_complete(run(ids, site))
                except Exception as ex:
                    print("请求订单详细页面出现错误!", ex)
                    time.sleep(5)
                    exit()

                # print(result_list)
                # print(dxm_list)
                for result in result_list:
                    packageId = result['packageId']
                    if match_goods_title(result['title']):
                        if w.get(packageId, None) is None:
                            w[packageId] = {'packageId': packageId,
                                            'remake': "数量:" + str(result['num']) + "、" + result['names'],
                                            'order_id': result['订单id'],
                                            'intern_id': result['内部id'],
                                            'title': result['title']
                                }
                        else:
                            w[packageId]['remake'] += "、" + "数量:" + str(result['num']) + "、" + result['names']
                        w[packageId]['remake'] = re.sub(r'([、:])(\s+)', lambda x: x.group(1),
                                                        w[packageId]['remake'])
                        w[packageId]['remake'] = re.sub(r'(\s+)(、)', '、', w[packageId]['remake'])
                        w[packageId]['remake'] = re.sub(r'(\s+)$', '', w[packageId]['remake'])
                print(w)
                if w:
                    write_remark(w, order_number_filename, cookie_dict)
                else:
                    print(f"店小蜜账户{[j]}和shopify商铺{site}没有匹配数据，请检查")
                    time.sleep(10)
                j += 1


            time.sleep(5)


main()
