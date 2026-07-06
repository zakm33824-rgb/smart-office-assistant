# 轻量化智能办公助手

一个基于 Streamlit 的办公小程序，支持：

- 智能 PPT 生成与导出
- 页面级智能组装 PPT
- 表格文本生成
- CSV / XLSX 数据处理与导出
- PPT 模板资源库、页级模板库、组件级素材库

## 本地运行

```bash
pip install -r requirements.txt
python scripts/bootstrap_template_library.py
streamlit run app.py
```

## 模板资源库

资源库相关代码位于 `ppt_template_library/`，包含：

- 来源端点清单
- 页级模板库
- 组件级素材库
- 模板元数据结构
- SQLite 存储层
- 预览图生成器
- 下载与解析骨架
- Streamlit 页面入口

先运行：

```bash
python scripts/bootstrap_template_library.py
```

它会创建目录结构、初始化数据库、写入首批来源目录清单、页级模板与组件素材。

## 部署到 GitHub + Streamlit Community Cloud

1. 新建一个 GitHub 仓库。
2. 把本项目的 `app.py`、`requirements.txt`、`.streamlit/config.toml` 推送到仓库根目录。
3. 打开 [Streamlit Community Cloud](https://share.streamlit.io/)，用 GitHub 登录。
4. 选择仓库，入口文件填写 `app.py`。
5. 点击部署，等待生成公网链接。

## 说明

- 只要仓库是公开的，任何人都可以通过部署后的链接访问。
- 如果后续你想换成自己的域名，也可以把同一个项目迁移到自有服务器或其他平台。
- 模板库后续可以继续接真实下载、去重、分类、评分和跨模板组装流程。
