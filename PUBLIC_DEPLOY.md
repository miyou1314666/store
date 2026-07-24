# 车架散件储备清单工具公网部署说明

这个工具需要后端处理 Excel，所以不能只部署成普通静态网页。建议部署到支持 Python Web Service 的平台，例如 Render、Railway、PythonAnywhere 或工厂内网服务器。

## 推荐方案：Render 公网演示版

1. 新建一个 GitHub 仓库。
2. 把本文件夹 `frame-parts-tool` 里的全部文件上传到仓库根目录。
3. 打开 Render，选择 New Web Service。
4. 连接 GitHub 仓库。
5. 配置如下：
   - Runtime: Python
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `python app.py`
6. 部署完成后，Render 会给出一个 `https://xxx.onrender.com` 公网网址。
7. 别人打开这个网址即可上传 SAP/SRM 表格并生成储备清单。

## 注意事项

- 不建议在公网演示版上传工厂真实敏感数据，除非得到允许。
- 免费平台可能会休眠，第一次打开会慢一些。
- 如果要正式落地，建议部署到工厂内网服务器，并接入 SAP/SRM 接口或数据中台。
