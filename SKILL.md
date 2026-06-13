---
name: e-steel-data-comparison
description: 对比 e钢平台销售合同业务检查数据与3个相关业务表（销售提单、销售实提、销售退货）的重量数据。支持单个或多个合同号查询，直接输出文字核对结果。
---

# E钢数据对比技能

## 概述

此技能用于对比 e钢平台（itgmetals.com）的销售合同业务检查接口与3个相关业务表接口返回的数据，识别差异并直接输出文字核对结果。

## 对比逻辑

### 新版本（当前）

对比销售合同业务检查接口与3个相关业务表：

1. **销售合同业务检查接口** (`/sell-contract/sell-contract-check-scroll`)
   - 提单重量：`newPickupWeight`
   - 提单金额：`newPickupSumAmount`
   - 实提重量：`newRealPickWeight`
   - 销售退货重量：`newSellReturnWeight`

2. **销售提单业务检查汇总表** (`/sell-pickup/scroll-check`)
   - 对比：`sumWeight` 之和 vs `newPickupWeight` 和 `newPickupSumAmount`

3. **销售实提业务检查汇总表** (`/sell-real-pickup/scroll-check`)
   - 对比：`sumWeight` 之和 vs `newRealPickWeight`

4. **销售退货报表** (`/sell-return-product/page`)
   - 对比：`weight` 之和 vs `newSellReturnWeight`

## 使用场景

当用户提出以下类型的问题时，应使用此技能：
- "帮我核对合同号 XXX 的数据"
- "核对合同 XXX, YYY, ZZZ 的重量"
- "检查这些合同的数据是否一致：[合同号列表]"
- "生成合同对比报告"
- "核对 20260610-20260612 这段时间的合同数据"
- "帮我核对6月10号到12号的合同重量"
- "对比合同的提单重量、实提重量、退货重量"

## 工作流程

### 1. 获取合同号

从用户对话中提取合同号或时间段。支持三种模式：

**模式 A：指定合同号**
- 用户提供单个合同号
- 用户提供多个合同号（用逗号、空格或换行分隔）
- 不提供合同号（表示查询全部合同）

**模式 B：按时间段查询**
- 用户提供时间段，格式：`20260610-20260612`（起始日期-结束日期）
- 也可以只提供单个日期：`20260612`（表示只查当天）
- WorkBuddy 需要将用户自然语言转换为 `--date-range` 参数

### 2. 执行对比脚本

**模式 A：指定合同号**
```bash
/Library/Frameworks/Python.framework/Versions/3.11/bin/python3 ~/.workbuddy/skills/e-steel-data-comparison/scripts/compare_contracts.py "合同号1" "合同号2" ...
```

**模式 B：按时间段查询**
```bash
/Library/Frameworks/Python.framework/Versions/3.11/bin/python3 ~/.workbuddy/skills/e-steel-data-comparison/scripts/compare_contracts.py --date-range 20260610-20260612
```

**注意：** 脚本需要使用系统 Python（3.11），因为依赖包（requests, openpyxl）已在该环境中安装。

### 3. 解析输出

脚本直接输出文字格式的核对结果：

- 每个合同逐条显示提单重量、实提重量、销售退货重量的对比数据
- ✅ 表示数据一致，❌ 表示有差异
- 末尾显示汇总信息：合同总数、差异条数

### 4. 返回结果给用户

- 将脚本输出结果直接展示给用户
- 告知用户核对结论（是否有差异）

## 示例对话

**用户：** "帮我对比合同号 HT2024001 和 HT2024002"

**执行步骤：**
1. 提取合同号：HT2024001, HT2024002
2. 执行命令：`python ~/.workbuddy/skills/e-steel-data-comparison/scripts/compare_contracts.py "HT2024001" "HT2024002"`
3. 解析输出，提取文字摘要和 Excel 文件
4. 返回核对结果给用户

**用户：** "核对6月10号到12号的合同重量"

**执行步骤：**
1. 识别时间段，转换为格式：`20260610-20260612`
2. 执行命令：`python ~/.workbuddy/skills/e-steel-data-comparison/scripts/compare_contracts.py --date-range 20260610-20260612`
3. 脚本会先调用销售合同列表接口获取该时间段所有合同名称
4. 再用合同名称调用对比接口进行核对
5. 解析输出，返回核对结果给用户

## 脚本说明

脚本文件：`scripts/compare_contracts.py`

**功能：**
- 支持三种查询模式：指定合同号、按时间段查询、查询全部
- 按时间段查询时，先调用销售合同列表接口（`/sell-contract/scroll`）获取该时间段所有合同
- 调用4个 API 接口获取数据：
  1. 销售合同业务检查接口（`/sell-contract/sell-contract-check-scroll`）
  2. 销售提单业务检查汇总表（`/sell-pickup/scroll-check`）
  3. 销售实提业务检查汇总表（`/sell-real-pickup/scroll-check`）
  4. 销售退货报表（`/sell-return-product/page`）
- 对比4个关键字段：
  - 提单重量（newPickupWeight）
  - 提单金额（newPickupSumAmount）
  - 实提重量（newRealPickWeight）
  - 销售退货重量（newSellReturnWeight）
- 识别重量差异（容忍度：0.001 吨）
- 直接输出文字格式的核对结果，不生成 Excel 文件

## 注意事项

1. **TOKEN 配置**：脚本中已包含 API 认证 TOKEN。如接口返回 TOKEN 过期，WorkBuddy 会自动提示用户输入新 TOKEN 并替换脚本中的值后重试，无需手动修改
2. **网络连接**：执行脚本需要能访问 `manage.itgmetals.com` 接口
3. **Python 环境**：脚本使用系统 Python（/Library/Frameworks/Python.framework/Versions/3.11/bin/python3），依赖包 requests 需要已安装
   
   如系统 Python 缺少依赖，请执行：
   ```bash
   /Library/Frameworks/Python.framework/Versions/3.11/bin/python3 -m pip install requests
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
