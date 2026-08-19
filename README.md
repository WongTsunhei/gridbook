# GridBook

GridBook 是一个开源、基于浏览器的多人实时协作电子表格。它提供接近 Excel / WPS 的编辑体验，并支持虚拟化网格、多工作表、单元格格式、合并单元格、公式文本、Excel / WPS 粘贴与增量实时同步。

GridBook is an open-source, browser-based collaborative spreadsheet for real-time multi-user editing. It provides an Excel / WPS-style experience with virtualized grids, multi-sheet workbooks, cell formatting, merged cells, formula text, spreadsheet clipboard paste, and incremental synchronization.

## 特性

- 稀疏的无限网格：空白单元格不创建存储记录，列标按 A、B、…、AA… 动态生成。
- 虚拟化渲染、冻结表头与首列、行列尺寸调整、合并单元格和填充柄。
- 矩形多选、行列选择、键盘导航、撤销重做、右键菜单和 Excel / WPS HTML、TSV 剪贴板粘贴。
- 字体、字号、颜色、填充、框线、对齐、换行和简化单元格数据格式。
- 多工作表、函数输入候选和公式文本存储；不内置公式计算服务。
- 通过 HTTP + SSE 的增量协作同步，避免为单个变更广播完整工作簿。

## 本地运行

先启动独立后端：

```powershell
$env:GRIDBOOK_PORT = "8767"
python backend/server.py
```

在另一个终端安装前端依赖并启动开发服务器：

```powershell
npm install
npm run dev
```

打开 <http://127.0.0.1:5274/>。前端通过 `/api` 代理访问本机的 GridBook 后端。

## 构建与测试

```powershell
npm run build
python -m unittest backend.test_server
```

## 部署

`deploy/gridbook.service` 是独立的 systemd 服务示例，默认监听 `127.0.0.1:8767`，运行数据放在 `/var/lib/gridbook/`。构建后的 `dist/` 可部署为 `/opt/gridbook/web/`，并使用 `deploy/nginx-gridbook.conf` 作为独立站点或子域名的 Nginx 配置参考。

运行时数据位于 `backend/data/`（开发环境）或服务配置中的独立数据目录，已被 Git 忽略，不应提交到开源仓库。
