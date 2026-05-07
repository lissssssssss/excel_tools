# 铝合金代发 Excel 月度合并脚本

## 功能

按文件名前缀（如 `3.`、`4.`）把 `/Users/lishengsheng/Documents/铝合金代发`
里的 `.xls` 文件按月合并成一个 `.xlsx`：

- 同一月内按文件名（日期）顺序拼接行
- 多 sheet 文件每个 sheet 都会被处理
- **按"快递单号"全局去重**：同一月份汇总文件里，每个非空单号只保留**最早出现的**那一条

输出在 `<源目录>/_合并/`：

```
_合并/
├── 2月汇总.xlsx
├── 3月汇总.xlsx
├── 4月汇总.xlsx
├── 5月汇总.xlsx
└── merge.log
```

## 保留的内容

| 项目 | 说明 |
|---|---|
| 行 | 按文件 (日) → sheet → 行号顺序拼接, "快递单号" 重复时丢后一条 |
| 图片 | JPEG / PNG / DIB / TIFF（PNG 保留为 PNG，含透明）；矢量图（emf/wmf/pict）暂跳过并记日志 |
| 字体 | 字体名 / 字号 / 加粗 / 斜体 / 下划线 / 颜色 |
| 对齐 | 水平 / 垂直 / **自动换行** / 缩进 / 旋转 |
| 数字格式 | 完整保留（日期、百分比、文本等） |
| 边框 | 上下左右四边的线型和颜色 |
| 行高 | 每行复用源文件行高 |
| 列宽 | 用第一个文件的列宽作为规范 |
| 合并单元格 | 按 "源行号 → 输出行号" 翻译重新合并 |

每行末尾追加两列：

| ... 原 12 列 ... | 源文件 | 源文件行号 |

- `源文件`：单 sheet 文件就是文件名；多 sheet 文件会写成 `文件名.xls / sheet名`
- `源文件行号`：Excel 1-based 行号（含表头）

任何报错或可疑日志都会带上这两个值，方便回查源数据。

## 快递单号去重

- 比较前先做**规范化**：去掉所有空格 + 大小写统一（即 `"sf 123"`、`"SF123"`、`" sf123 "` 视作同一条）
- **空单号** 不参与去重（直接保留）
- 跨**文件 + 跨 sheet** 全局去重，但每个月独立（2 月 / 3 月互不影响）
- 早出现的留下；后面再次出现会写到 `INFO` 日志：`快递单号=xxx 已存在, 跳过`
- 写入完成时汇总会打印：`数据行=N, 去重跳过=K, 唯一快递单号=U`

## 安装

```bash
cd /Users/lishengsheng/tmp
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 运行

```bash
cd /Users/lishengsheng/tmp
source .venv/bin/activate

# 默认: 源 = /Users/lishengsheng/Documents/铝合金代发, 输出 = 源/_合并/
python merge_xls.py

# 或自定义路径
python merge_xls.py "/path/to/source" "/path/to/output"
```

## 性能 / 体量参考

约 90 个 `.xls` 文件、3500+ 行、3500+ 张图片，跑完约 **30 秒**：

| 月份 | 源文件 | sheet | 数据行 (去重前) | 写入行 | 去重跳过 | 嵌入图片 | 输出大小 |
|---|---|---|---|---|---|---|---|
| 2月 | 7 | 7 | 135 | 134 | 1 | 134 | 1.3 MB |
| 3月 | 54 | 67 | 1598 | 1504 | 94 | 1487 | 14 MB |
| 4月 | 29 | 29 | 1686 | 1684 | 2 | 1667 | 15 MB |
| 5月 | 4 | 4 | 251 | 251 | 0 | 248 | 2.2 MB |

图片在写入前会被 PIL 缩到最大 220×220 像素（保持比例，PNG 仍是 PNG），
单元格里显示尺寸 100×100 像素。

## 日志说明

- **`INFO`**：每个文件的解析进度、行数、图片数。每行结尾附带:
  ```
  (BLIP仓库=N, 形状=M, 已定位行数=K)
  ```
  - `BLIP仓库`：源文件里所有图片字节数（含历史里被删除但 BLIP 没清掉的）
  - `形状`：源文件里实际放在格子里的图形对象数量
  - `已定位行数`：脚本成功把图片对应到一个数据行的数量
- **`INFO`** 关注点：
  - `快递单号=xxx 已存在, 跳过` —— 全局去重命中
- **`WARNING`** 关注点（每条都会告诉你**哪个文件、哪个 sheet、哪一行**）：
  - 表头与第一个 sheet 不一致
  - 表头里没找到 "图" 字 / "快递单号" 列
  - 某行有多张图片（顺移到下一行 / 后 3 行无空位则丢弃）
  - 矢量图片（emf/wmf）跳过
  - 图片锚点超出数据范围
  - 行只有图片没有任何文本
  - 隐藏 sheet 跳过
  - 文件名无法识别月份
  - 合并单元格因为源行被跳过而无法平移
- **`ERROR`**：单个文件读不开或写入失败时记录，跳过该文件继续处理其他文件。

## 实现要点

- `.xls` 是 Excel 97-2003 二进制（OLE 复合文档 + BIFF）。`xlrd` 能读单元格但
  读不到图片，所以脚本直接解析 `Workbook` 流：
  - 用 `BOF (0x0809) / EOF (0x000A)` 划分 substream，区分 workbook globals 和
    每个 sheet 各自的 substream
  - `MsoDrawingGroup (0xEB)` 在 globals 里 → `OfficeArtBStoreContainer` →
    `OfficeArtFBSE` → `OfficeArtBlip (0xF01D=jpeg, 0xF01E=png, ...)` 抽出图片字节流
  - 每个 sheet 的 `MsoDrawing (0xEC)` → `OfficeArtSpContainer` →
    `ClientAnchor + FOPT.Pib` 拿到 (行, 列) 锚点 + 图片仓库索引（1-based）
  - 多 sheet 文件按 sheet 索引分别归图片，避免不同 sheet 的图片混到一起
- 用 xlrd `formatting_info=True` 拿格式（XF / Font / Alignment / Border /
  ColInfo / RowInfo / merged_cells），转换成 openpyxl 的 `Font` / `Alignment` /
  `Border` 应用到对应单元格。
- `formatting_info=True` 时 `sh.nrows / sh.ncols` 会包含很多"空但有格式"的
  幽灵行/列。脚本通过表头里最后一个非空列定位真实列数；通过逆向扫描定位
  最后一个含数据的行。
- 图片解码后，PIL 缩略到 220×220，PNG 保持 PNG（含透明度），其余 JPEG。
- 输出格式必须是 `.xlsx` —— 纯 Python 没法把图片写回 `.xls`（`xlwt` 不支持）。
- **不需要安装 LibreOffice / Office**，纯 Python。
