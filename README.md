# 信用债交易记录转换 - Cloudflare Pages 静态版

这是一个纯前端版本，适合部署到 Cloudflare Pages。解析、Word 读取、Excel 导出都在浏览器本地完成，没有 Python 后端，也不会因为后端空闲而休眠。

## 文件

- `index.html`：页面入口
- `styles.css`：页面样式
- `app.js`：解析、表格编辑、Excel 导出逻辑

## Cloudflare Pages 直接上传

1. 进入 Cloudflare Dashboard。
2. 打开 `Workers & Pages`。
3. 选择 `Pages`。
4. 创建项目，选择 `Direct Upload`。
5. 上传整个 `cloudflare_trade_parser` 文件夹里的文件。
6. 发布后即可得到一个 `*.pages.dev` 地址。

## Cloudflare Pages 连接 GitHub

如果你把这几个文件放进 GitHub 仓库：

- Framework preset：`None`
- Build command：留空
- Build output directory：如果这些文件在仓库根目录，填 `/`；如果在子目录，填对应目录名

## 依赖

页面通过 CDN 加载：

- SheetJS：生成 Excel
- Mammoth.js：读取 `.docx`
- Lucide：按钮图标

如果办公网络无法访问 CDN，可以把这些库文件下载到本地 `vendor/` 目录后改成本地引用。
