# MySpider

<p align="center">
<b>E-commerce Product Data Collection & Automation Tool</b>
</p>


# 中文说明 🇨🇳


## 项目简介

MySpider 是一个面向电商平台的数据采集与处理工具，支持：

- 淘宝 (Taobao)
- 拼多多 (Pinduoduo)
- 京东 (JD.com)
- 1688


项目支持采集商品标题、价格、销量、评论数量、店铺信息、商品链接、商品图片等数据，并将结果导出为 Excel 文件。


本项目基于 Python 开发，融合：

- Selenium
- PyQuery
- Pillow
- Flask
- JavaScript Injection

实现：

- 浏览器自动化控制
- 动态页面数据解析
- 网络请求监听
- JSON 数据处理
- Excel 数据生成

> ⚠️ 本仓库代码已经经过脱敏处理，仅用于技术展示和学习参考。
>
> 当前代码无法直接运行，部分授权、账号验证以及商业模块已隐藏。

---

# 功能特性


## 支持四大电商平台

支持：

- 淘宝
- 拼多多
- 京东
- 1688


---

## 商品信息采集

支持获取：

- 商品标题
- 商品价格
- 销量
- 评论数量
- 店铺名称
- 商品链接
- 包邮信息
- 商品图片


支持商品数据自动解析，并转换为结构化数据。


---

## 浏览器网络请求监听与数据解析

项目支持基于 JavaScript Injection 的浏览器数据监听方案。


主要功能：

- 注入 JavaScript Hook 扩展浏览器运行环境。
- 监听页面加载过程中产生的 Request / Response 数据。
- 解析接口返回的 JSON 数据。
- 提取商品列表以及相关字段。
- 降低传统 DOM 页面解析对页面结构的依赖。


适用于：

- 异步加载页面
- 动态渲染页面
- 前端数据驱动页面


---

## 浏览器自动化控制

基于 Selenium 实现：

- 自动打开网页
- 页面滚动
- 点击操作
- 动态加载触发


使用：

- `undetected_chromedriver`
- `selenium-wire`

增强浏览器自动化能力。


---

## 平台状态检测

支持检测：

- 滑块验证
- 页面访问异常
- 登录状态异常
- 采集中断


当检测到异常情况时，会自动停止任务并提示用户。


---

## 图片处理与 Excel 导出


支持：

- 下载商品主图
- 图片插入 Excel
- 商品数据格式化
- 自动生成数据报表


---

## 增量 Hash 数据处理


使用 Hash 校验机制：

- 判断商品是否重复
- 避免重复写入
- 支持增量采集
- 支持断点继续


提高长时间运行稳定性。


---

## 懒加载支持


支持动态网页：

- 自动滚动页面
- 触发懒加载
- 获取完整商品列表


---

# 项目架构


```
WPF Client
    |
    |
Flask API Server
    |
    |
Spider Engine
    |
    |
Selenium Browser
    |
    |
JavaScript Injection
    |
    |
Request Monitoring
    |
    |
JSON Data Parsing
    |
    |
Excel Export
```


---

# 文件结构


```
MySpider

├── Spider/
│
│   ├── Tao_collection_tool.py
│   ├── PDD_collection_tool.py
│   ├── JD_collection_tool.py
│   └── Ali1688_collection_tool.py
│
├── app.py
│
├── path_utils.py
│
├── LicenceRecode.py
│
└── README.md
```


---

# 文件说明


## Spider/

主要包含：

- 平台数据采集逻辑
- 数据解析
- 图片处理
- Excel 数据处理


---

## app.py

Flask 后端服务：

- API 接口
- 爬虫任务管理
- 启动、暂停、停止控制


---

## path_utils.py

资源路径管理：

- 开发环境路径处理
- 打包环境路径兼容


---

## LicenceRecode.py

许可证管理模块：

- 软件授权验证
- 环境绑定
- 激活管理


---

# 安装与使用


## Clone 项目


```bash
git clone https://github.com/JZIIIIII/PDD-Collection-Tool.git

cd PDD-Collection-Tool
```


---

# 发布版本


当前版本：

```
v1.0.0 Trial Edition
```


体验版包含：

- 四个平台基础采集功能
- Excel 导出
- 图片处理


限制：

- 单次最多采集 10 条商品数据


由于目标电商平台页面结构可能变化，实际采集结果可能存在差异。


---

# 前端参考


本项目客户端 UI 参考：

UIKitTutorials

https://github.com/Jeyderht/UIKitTutorials


---

# 免责声明


本项目仅用于：

- 学习研究
- 自动化技术实践
- Python / Selenium 技术交流


使用者应遵守目标平台服务协议以及相关法律法规。



<br>



# English Description 🇬🇧


## Project Introduction


MySpider is an e-commerce data collection and processing tool supporting:

- Taobao
- Pinduoduo
- JD.com
- 1688


It collects product information including:

- Product title
- Price
- Sales volume
- Review count
- Store information
- Product URL
- Product images


Collected data can be exported into Excel files.


The project is developed with Python and integrates:

- Selenium
- PyQuery
- Pillow
- Flask
- JavaScript Injection


Providing:

- Browser automation
- Dynamic data parsing
- Network request monitoring
- JSON processing
- Excel generation


> ⚠️ The source code has been desensitized and is provided for demonstration and educational purposes only.
>
> The current version cannot run directly. Some license, authentication, and commercial modules have been removed.


---

# Features


## Multi-platform Support


Supported platforms:

- Taobao
- Pinduoduo
- JD.com
- 1688


---

## Product Information Collection


Collects:

- Product titles
- Prices
- Sales volume
- Review counts
- Store names
- Product URLs
- Shipping information
- Product images


---

## Browser Network Monitoring and Data Parsing


The project supports JavaScript Injection based browser monitoring.


Features:

- Injects JavaScript hooks into browser runtime.
- Monitors Request / Response data.
- Parses JSON responses.
- Extracts structured product information.
- Reduces dependency on traditional DOM parsing.


Suitable for:

- Asynchronous loading pages
- Dynamic rendering pages
- Modern frontend applications


---

## Browser Automation


Implemented with Selenium:


Features:

- Automated browsing
- Page scrolling
- Click simulation
- Dynamic content loading


Technologies:

- `undetected_chromedriver`
- `selenium-wire`


---

## Platform Status Detection


Detects:

- Slider verification
- Access restrictions
- Login interruptions
- Collection failures


Automatically stops tasks when abnormal conditions occur.


---

## Image Processing and Excel Export


Supports:

- Product image downloading
- Image embedding into Excel
- Data formatting
- Report generation


---

## Incremental Hash Processing


Uses hash-based verification:


- Duplicate detection
- Incremental collection
- Resume collection
- Stable long-running operation


---

## Lazy Loading Support


Supports:

- Automatic scrolling
- Triggering lazy loading
- Complete product acquisition


---

# Project Architecture


```
WPF Client

    |

Flask API Server

    |

Spider Engine

    |

Selenium Browser

    |

JavaScript Injection

    |

Request Monitoring

    |

JSON Parsing

    |

Excel Export
```


---

# Installation


```bash
git clone https://github.com/JZIIIIII/PDD-Collection-Tool.git

cd PDD-Collection-Tool
```


---

# Release Version


Current release:

```
v1.0.0 Trial Edition
```


Trial version includes:

- Basic four-platform collection
- Excel export
- Image processing


Limitation:

- Maximum 10 products per task


Results may vary due to changes in target e-commerce platforms.


---

# Frontend Reference


The frontend interface is based on:

UIKitTutorials

https://github.com/Jeyderht/UIKitTutorials


---

# Disclaimer


This project is intended for:

- Educational purposes
- Automation technology research
- Python / Selenium practice


Users should comply with the terms of service and applicable laws of target platforms.