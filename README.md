# orderSpider
These files created for personal spider and git learning

## Introduction
- writeRemark.py: Get the orderList of shopify and write remark into dxm.
- toExcel.py: Get the orderList of shopify and write into excel.

## Require
- aiohttp - 3.6.1
- asyncio - 3.4.3
- bs4 - 0.0.1
- lxml - 4.4.1
- requests - 2.22.0
- xlrd - 1.2.0
- XlsxWriter - 1.2.8
- XlsxWriter - 1.2.8

### Editor your config for spider
- The key `dxm_cookies` contains key-values of your dxm account name and cookie.
- The key `shopify_cookies` contains key-values of your shop name and cookie.
- The key `site` contains key-values of your shop name and its shopId in dxm.
- The key `goods_name` contains the goods that need to write remark.
- The key `limitation` set to limit the num of order got.
