# PDD Collection Tool

## 项目简介

PDD Collection Tool 是一个用于爬取拼多多/淘宝商品信息的工具。它支持爬取商品的标题、价格、成交量、评论数、店铺链接、图片等信息，并将结果以 Excel 格式保存。本工具使用 Python、Selenium、pyquery 等技术实现，具备防反爬策略，并支持懒加载处理。

## 功能特性

- **商品信息爬取**：支持爬取拼多多商品的标题、价格、销量、评论数等信息。
- **图片下载**：可以下载商品图片并保存到 Excel 中。
- **多线程支持**：加快爬取速度，提升效率。
- **防反爬策略**：使用 `undetected_chromedriver` 和 `selenium-wire` 解决反爬问题。
- **懒加载处理**：自动滚动页面加载更多商品。
- **Excel 导出**：将抓取到的数据保存为 Excel 格式，方便查看和分析。

## 安装与使用

1. 克隆本项目：

   ```bash
   git clone https://github.com/JZIIIIII/PDD-Collection-Tool.git
   cd PDD-Collection-Tool
