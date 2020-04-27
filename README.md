# orderSpider
These files created for personal spider and git learning

## Introduction
- writeRemark.py: Get the orderList of shopify and write remark into dxm.
- toExcel.py: Get the orderList of shopify and write into excel.

### Editor your config for spider
- The key `dxm_cookies` contains key-values of your dxm account name and cookie.
- The key `shopify_cookies` contains key-values of your shop name and cookie.
- The key `site` contains key-values of your shop name and its shopId in dxm.
- The key `goods_name` contains the goods that need to write remark.
- The key `limitation` set to limit the num of order got.

======

# 包裹接口
=========

## 获取包裹列表
GET /api/wuliu/package/index?access_token=:accessToken

### 查询参数
| 参数 | 说明 |
|---|---|
| package_number | 包裹号 |
| order_number | 订单号 |
| waybill_number | 运单号 |
| shop_name | 所属店铺 |
| delivery_begin_datetime | 发货时间开始筛选 |
| delivery_end_datetime | 发货时间结束筛选 |
| line_id | 所属线路 |
| company_id | 所属物流商 |
| status | 包裹状态, 0=>待处理 ,1=>已接单 ,2=>运输中 ,3=>到达代取 ,4=>成功签收 ,5=>查询不到 ,6=>运输过久 ,7=>可能异常 ,8=>投递失败, 100 => 已发货包裹  |
| expand | 扩展参数，可用的有country => 包裹目的地国家,line => 包裹运输线路/物流方式,routes => 包裹路由节点详情, line.company => 线路所属公司 |

## 编辑包裹
PUT/PATCH /api/wuliu/package/update?access_token=:accessToken&id=:id

### 参数
| 参数 | 类型 | 必填 | 默认值 | 说明 |
|---|:---:|:---:|:---:|---|
| status | string | 否 | 无 | 包裹状态 |

## 包裹详情
GET /api/wuliu/package/view?access_token=:accessToken

expand 扩展参数与 index 接口相同

## 包裹发货
POST /api/wuliu/package/delivery?access_token=:accessToken

### 参数 
| 参数 | 类型 | 必填 | 默认值 | 说明 |
|---|:---:|:---:|:---:|---|
| package_number | string | 是 | 无 | 包裹号 |
| weight | int | 是 | 无 | 重量 |

## 包裹批量发货
POST /api/wuliu/package/batch-delivery?access_token=:accessToken

### 参数 
| 参数 | 类型 | 必填 | 默认值 | 说明 |
|---|:---:|:---:|:---:|---|
| packages | array | 是 | 无 | 批量发货的包裹数组 |

提交示例:
- packages[0][package_number]:'XMEB154684'
- packages[0][weight]:'25'
- packages[1][package_number]:'XMEB165484'
- packages[1][weight]:'30'

## 编辑节点
PUT/PATCH /api/wuliu/package-route/process?access_token=:accessToken&id=:id

注意:id 应为 route 的 id
### 参数 
| 参数 | 类型 | 必填 | 默认值 | 说明 |
|---|:---:|:---:|:---:|---|
| process_status | int | 是 | 0=>无需处理,1=>待处理,2=>忽略,3=>已处理 | 包裹处理状态 |
| plan_datetime | string | 否 | 无 | 预测时间,格式:Y-m-d H:i:s |
| remark | string | 否 | 无 | 备注 |

## 导出包裹数据为excel
GET /api/wuliu/package/to-excel?access_token=:accessToken

### 查询参数
与 index 接口查询参数相同

## 包裹状态数据统计
GET /api/wuliu/package/status-options?access_token=:accessToken

## 批量修改路由节点预测时间
POST /api/wuliu/package-route/batch-change-plan-datetime?access_token=:accessToken&packageId=:packageId

### 提交参数
| 参数 | 类型 | 必填 | 默认值 | 说明 |
|---|:---:|:---:|:---:|---|
| routes | array | 是 | 无 | 数组格式,提交修改的路由 id 以及修改天数，见示例 |

注意:此处提交的为路由的id，与packageId注意区分<br>
提交示例:
- routes[0][id]:1
- routes[0][days]:2
- routes[1][id]:2
- routes[0][days]:1

## 发货统计接口
GET /api/wuliu/package/delivery-statistics?access_token=:accessToken

### 查询参数
| 参数 | 说明 |
|---|---|
| beginDate | 开始日期筛选,Y-m-d格式 |
| endDate | 结束日期筛选,Y-m-d格式 |
