# 六合彩金额统计系统

## 功能特性

- **智能文本解析** — 支持 `1.2.3各5`、`马.羊各10元`、`鼠各20`、`1.2.3各5万` 等多种格式
- **实时统计** — 输入即时计算，金额自动汇总，按号码/生肖/波色分类展示
- **Excel 同步** — 数据实时写入 Excel 文件（openpyxl，无需安装 Office）
- **界面配置** — 双击生肖/波色单元格直接编辑，自动同步到 config.json 和 Excel
- **备份回滚** — 每次操作自动备份，支持随时恢复到任意历史版本
- **每日重置** — 每天 00:00 自动重置数据，重置前自动备份
- **时间命名** — Excel 文件名格式：`YYYYMMDD_六合彩金额统计表.xlsx`

## 快速开始

### 方式一：一键启动（推荐）
双击 `setup.bat`，自动安装依赖并启动程序。

### 方式二：手动启动
```bash
pip install openpyxl
python main.py
```

### 方式三：打包成 exe
```bash
pip install pyinstaller openpyxl
pyinstaller --noconsole --onefile main.py
```

## 项目结构

```
sixhecai/
├── main.py          # 主程序
├── config.json      # 配置文件（生肖映射 + 界面设置）
├── setup.bat        # 一键安装启动脚本
├── requirements.txt # Python 依赖
├── excel/           # Excel 输出目录
│   └── YYYYMMDD_六合彩金额统计表.xlsx
└── backup/          # 备份目录
    └── YYYYMMDD_HHMMSS_标签/
```

## 输入格式

| 格式 | 示例 | 说明 |
|------|------|------|
| 数字多点 | `1.2.3各5` | 1、2、3号各5元 |
| 生肖多点 | `马.羊各10` | 马和羊各10元 |
| 单一生肖 | `鼠各20` | 鼠20元 |
| 带单位 | `1.2.3各5万` | 支持万/亿 |
| 混合 | `1.2.3各5元` | 自动去除"元"字 |
