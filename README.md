# e-steel-data-comparison

e钢平台销售合同与销售实提重量核对技能 - WorkBuddy Skill

## 功能

对比 e钢平台（itgmetals.com）的销售合同接口与销售实提接口返回的重量数据，识别差异并生成详细的 Excel 报告。

## 安装

在 WorkBuddy 中输入：
```
安装技能 https://github.com/zengqi000/e-steel-data-comparison
```

## 使用

安装后，在对话中说：
- "帮我核对合同 XXX 的数据"
- "核对合同 XXX, YYY, ZZZ 的重量"

## 配置

首次使用需要设置 TOKEN：
1. 脚本执行时如果提示 TOKEN 过期，输入新的 e钢平台 TOKEN 即可
2. TOKEN 格式：`Bearer eyJhbGci...`

## 依赖

- Python 3.11+
- requests
- openpyxl

安装依赖：
```bash
pip install requests openpyxl
```

## 报告内容

Excel 报告包含以下工作表：
- 📊 汇总概览
- ❌ 重量差异明细（如有差异）
- ⚠️ 仅合同表有 / 仅实提表有
- 📋 合同表数据 / 实提表汇总
