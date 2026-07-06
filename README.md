# 轻量化智能办公助手

一个基于 Streamlit 的单文件办公小程序，支持：

- 智能 PPT 生成与导出
- 表格文本生成
- CSV / XLSX 数据处理与导出

## 本地运行

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署到 GitHub + Streamlit Community Cloud

1. 新建一个 GitHub 仓库。
2. 把本项目的 `app.py`、`requirements.txt`、`.streamlit/config.toml` 推送到仓库根目录。
3. 打开 [Streamlit Community Cloud](https://share.streamlit.io/)，用 GitHub 登录。
4. 选择仓库，入口文件填写 `app.py`。
5. 点击部署，等待生成公网链接。

## 说明

- 只要仓库是公开的，任何人都可以通过部署后的链接访问。
- 如果后续你想换成自己的域名，也可以把同一个项目迁移到自有服务器或其他平台。
