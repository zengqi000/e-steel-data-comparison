---
name: e-steel-data-comparison
description: 对比 e钢平台销售合同与销售实提的重量数据。当用户需要提供合同号来核对两个接口返回的重量数据时使用此技能。支持单个或多个合同号查询，生成详细的 Excel 对比报告。
---

# E钢数据对比技能

## 概述

此技能用于对比 e钢平台（itgmetals.com）的销售合同接口与销售实提接口返回的重量数据，识别差异并生成详细的 Excel 报告。

## 使用场景

当用户提出以下类型的问题时，应使用此技能：
- "帮我对比合同号 XXX 的数据"
- "核对合同 XXX, YYY, ZZZ 的重量"
- "检查这些合同的数据是否一致：[合同号列表]"
- "生成合同对比报告"

## 工作流程

### 1. 获取合同号

从用户对话中提取合同号。用户可以：
- 提供单个合同号
- 提供多个合同号（用逗号、空格或换行分隔）
- 不提供合同号（表示查询全部合同）

### 2. 执行对比脚本

使用以下命令调用 Python 脚本：

```bash
/Library/Frameworks/Python.framework/Versions/3.11/bin/python3 ~/.workbuddy/skills/e-steel-data-comparison/scripts/compare_contracts.py "合同号1" "合同号2" ...
```

**注意：** 脚本需要使用系统 Python（3.11），因为依赖包（requests, openpyxl）已在该环境中安装。

### 3. 解析输出

脚本输出包含两部分：

**A. 文字摘要**（在 `__WORKBUDDY_FILE_START__` 之前）
- 包含核对结果的文本摘要
- 显示合同表合同数、实提表合同数、差异条数等
- 如有差异，会列出差异明细

**B. Excel 文件**（在 `__WORKBUDDY_FILE_START__` 和 `__WORKBUDDY_FILE_END__` 之间）
- 使用特殊的标记格式输出
- 包含 base64 编码的 Excel 文件数据
- WorkBuddy 会自动识别并提供下载

### 4. 返回结果给用户

- 向用户展示文字摘要
- 如果有 Excel 文件，WorkBuddy 会自动处理下载标记
- 告知用户核对结论（是否有差异）

## 示例对话

**用户：** "帮我对比合同号 HT2024001 和 HT2024002"

**执行步骤：**
1. 提取合同号：HT2024001, HT2024002
2. 执行命令：`python ~/.workbuddy/skills/e-steel-data-comparison/scripts/compare_contracts.py "HT2024001" "HT2024002"`
3. 解析输出，提取文字摘要和 Excel 文件
4. 返回核对结果给用户

## 脚本说明

脚本文件：`scripts/compare_contracts.py`

**功能：**
- 调用两个 API 接口获取数据
- 对比销售合同表的 `sumRealPickWeight` 字段和销售实提表的 `weight` 字段
- 识别重量差异（容忍度：0.001 吨）
- 生成包含多个工作表的 Excel 报告

**输出文件包含的工作表：**
- 📊 汇总概览：核对时间、查询合同、统计信息
- ❌ 重量差异明细：两表均有但重量不一致的合同
- ⚠️ 仅合同表有：只在一个表中存在的合同
- ⚠️ 仅实提表有：只在一个表中存在的合同
- 📋 合同表数据：完整的合同表数据
- 📋 实提表汇总：完整的实提表数据

## 注意事项

1. **TOKEN 配置**：脚本中已包含 API 认证 TOKEN。如接口返回 TOKEN 过期，WorkBuddy 会自动提示用户输入新 TOKEN 并替换脚本中的值后重试，无需手动修改
2. **网络连接**：执行脚本需要能访问 `manage.itgmetals.com` 接口
3. **Python 环境**：脚本使用系统 Python（/Library/Frameworks/Python.framework/Versions/3.11/bin/python3），依赖包 requests 和 openpyxl 需要已安装
   
   如系统 Python 缺少依赖，请执行：
   ```bash
   /Library/Frameworks/Python.framework/Versions/3.11/bin/python3 -m pip install requests openpyxl
   ```

## TOKEN 过期处理流程

当脚本检测到 TOKEN 过期时，会以退出码 99 退出，并在 stderr 输出 `__TOKEN_EXPIRED__`。

**WorkBuddy 的处理步骤：**

1. 执行脚本后，检查退出码是否为 99，或 stderr 是否包含 `__TOKEN_EXPIRED__`
2. 如果是 TOKEN 过期，使用 AskUserQuestion 工具提示用户输入新的 TOKEN：
   - 提示语："🔑 e钢平台 TOKEN 已过期，请输入新的 TOKEN（以 Bearer 开头）："
3. 用户输入新 TOKEN 后，使用 Edit 工具替换脚本中的 TOKEN 变量值
   - 脚本路径：`~/.workbuddy/skills/e-steel-data-comparison/scripts/compare_contracts.py`
   - 替换 `TOKEN = '...'` 整行
4. 重新执行脚本
5. 如果仍然报 TOKEN 过期，再次提示用户（最多重试 3 次）

**TOKEN 替换示例：**
```python
# 旧值
TOKEN = 'Bearer eyJhbGci...旧token...'
# 新值
TOKEN = 'Bearer eyJhbGci...新token...'
```

## 错误处理

如脚本执行失败：
- **TOKEN 过期**：按照上述"TOKEN 过期处理流程"操作
- 检查网络连接
- 检查合同号格式是否正确
- 查看错误信息，向用户报告具体错误原因
