# ERP凭证字段提取工具

批量从ERP导出的Excel文件中提取指定字段，合并导出到汇总文件。

## 功能

- 拖拽或选择多个 `.xlsx` / `.xls` 文件
- 自动扫描所有文件的表头，列出交集字段
- 自定义勾选需要提取的字段（默认6个核心字段）
- 勾选偏好自动记忆，下次启动自动恢复
- 进度条 + 完成弹窗提示
- 导出按源文件表头顺序排列的汇总 Excel

## 默认提取字段

- 凭证抬头文本
- 文本
- 凭证货币价值
- 借/贷标识
- 总账科目
- 总账科目：长文本

## 使用方式

### macOS
双击 `ERP凭证字段提取工具.app` 运行。

### Windows
下载 [Releases](https://github.com/dean611953-beep/erp-extract-tool/releases) 中的 `ERP凭证字段提取工具_Windows.zip`，解压后双击 `ERP凭证字段提取工具.exe`。

### 源码运行
```bash
pip install -r requirements.txt
python erp_extract_tool.py
```
