# 后端自测 — G 盘测试资源准备

> 配合 `测试用例_硬盘扫描专项.xlsx` 使用。  
> 测试盘：**G: Test5GB**（约 5GB HDD 分区）。**仅操作 G:，勿动 C/D/E/F。**

---

## 1. 仓库内相关文件

| 路径 | 用途 |
| --- | --- |
| `scripts/stage_user_preview_resources.ps1` | 清空 G: 根目录并生成根目录样本（情况 1～5 主流程） |
| `scripts/stage_g_folder_samples.ps1` | 生成嵌套文件夹样本（情况 6 路径抽测，不清空整盘） |
| `scripts/clear_gh_delete.ps1` | 普通清理 G:/H: 根目录内容（跳过 System Volume Information） |
| `scripts/clear_gh_secure.ps1` | 深度清理（含空闲区覆写，偶尔使用） |
| `scripts/user_resources/` | Office / 图片模板目录（见 §3） |

---

## 2. 快速开始

在**项目根目录**打开 PowerShell（建议管理员）：

```powershell
# 1) 普通清理 G:（可选，脚本 stage 也会先清 G: 根目录）
.\scripts\clear_gh_delete.ps1

# 2) 生成 G:\ 根目录样本（情况 1～5）
.\scripts\stage_user_preview_resources.ps1

# 3) 生成嵌套文件夹样本（情况 6，4 条路径抽测）
.\scripts\stage_g_folder_samples.ps1
```

清空 **G: 盘对应回收站** 后再执行删除类用例。

---

## 3. `user_resources` 模板说明

脚本会从 `scripts/user_resources/` **递归**查找模板文件，按扩展名复制到 G:。

### 3.1 当前仓库已有模板

| 扩展名 | 目录 | 说明 |
| --- | --- | --- |
| `.pptx` | `user_resources/ppt/` | 2 个 pptx |
| `.xlsx` | `user_resources/xlsx/` | 3 个 xlsx |
| `.jpg` | `user_resources/图片/` | 多张 jpg |

### 3.2 必须自备模板（无模板则跳过该格式）

以下格式**没有真实模板不会生成**（避免无效占位文件）：

- `doc` / `docx` / `xls` / `ppt`（非 pptx）

### 3.3 无模板时自动生成的格式

- `pdf` / `txt` / `rtf` / `odt` / `pages` / `numbers` / `key` → 简单占位内容

### 3.4 添加模板

将真实可预览文件放入 `scripts/user_resources/` 任意子目录，扩展名正确即可，例如：

```
scripts/user_resources/
  ppt/
  xlsx/
  docx/          ← 可自行新增
  图片/
```

---

## 4. 根目录样本（`stage_user_preview_resources.ps1`）

| 项目 | 说明 |
| --- | --- |
| 目标盘 | **G:\** 根目录（单层，无子文件夹） |
| 执行前 | **清空 G: 根目录**（保留 System Volume Information） |
| 每种格式 | 3 个文件 |
| 命名规则 | `<名称>_TIME_<yyyyMMdd_HHmmss>_PATH_G_<序号>.<ext>` |
| 示例 | `AUTO_PPTX_01_TIME_20260525_143000_PATH_G_001.pptx` |

**建议**：预览/完整性用例优先选 `pptx`、`xlsx`、`jpg` 等已有真实模板的文件。

---

## 5. 嵌套文件夹样本（`stage_g_folder_samples.ps1`）

| 用例ID | 层级 | 路径 | 样本文件名 |
| --- | --- | --- | --- |
| DISK_FOLDER_L01 | 1 层 | `G:\nest_L1\` | `folder_test_L01_sample.pptx` |
| DISK_FOLDER_L03 | 3 层 | `G:\nest_L1\nest_L2\nest_L3\` | `folder_test_L03_sample.pptx` |
| DISK_FOLDER_L05 | 5 层 | `G:\nest_L1\…\nest_L5\` | `folder_test_L05_sample.pptx` |
| DISK_FOLDER_L10 | 10 层 | `G:\nest_L1\…\nest_L10\` | `folder_test_L10_sample.pptx` |

操作：**Delete → 回收站保留** → 扫描 G: → 核对原路径是否完整。

> 此脚本**不会**清空 G: 整盘，可与根目录样本共存；若需干净环境请先 `clear_gh_delete.ps1`。

---

## 6. 隐藏文件（情况 ③④⑤）

```powershell
# 设为隐藏（路径按实际样本修改）
(Get-Item -LiteralPath "G:\AUTO_PPTX_01_TIME_....pptx" -Force).Attributes = "Hidden"

# 取消隐藏（如需）
(Get-Item -LiteralPath "G:\样本路径" -Force).Attributes = "Normal"
```

情况 ③：**设 Hidden 后不删除**，扫描 G:，按产品定义记录是否出现在结果中。

---

## 7. 执行注意

1. **只扫描分区 G:**，不要选整盘（避免扫到同盘 E/F/H）。
2. 恢复软件建议**管理员身份**运行。
3. 每条用例尽量**单独准备样本**，避免 G: 上残留干扰「检出」。
4. 导出恢复文件建议目录：`D:\恢复导出\`（如有）。

---

## 8. 与 xlsx 的对应关系

| xlsx Sheet | 内容 |
| --- | --- |
| `5情况x4状态` | 5 种情况 × 4 状态模型总览 |
| `G盘-明细` | 完整用例步骤与预期 |
| `G盘-资源准备` | 本文摘要（脚本与模板索引） |
