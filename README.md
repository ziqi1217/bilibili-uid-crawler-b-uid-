# B站UID批量爬虫

批量扫描 B站（bilibili）用户 UID，筛选出等级0、关注0、粉丝0、昵称不含中文的账号。

## 功能特性

- **多位数扫描**：支持 7/8/9/10 位 UID，通过前缀/后缀组合指定范围
- **多线程并发**：默认8线程，可自定义（1~30），速率可达 3~10 uid/s
- **Cookie加速**：填入 SESSDATA 后绕过未登录风控，速度提升 5~10 倍
- **断点续传**：中断后重新运行可跳过已扫描的 UID
- **Ctrl+C 保护**：随时中断，自动保存已找到的结果
- **结果导出**：自动保存 CSV + JSON，另有 Excel 导出工具（仅 UID + 昵称）

## 筛选条件

同时满足以下 4 项才会被记录：

| 条件 | 要求 |
|------|------|
| 等级 | = 0 |
| 关注数 | = 0 |
| 粉丝数 | = 0 |
| 昵称 | 不含中文字符 |

## 快速开始

### 1. 安装依赖

```bash
pip install requests openpyxl
```

### 2. 运行爬虫

```bash
python crawler.py
```

按提示依次输入：

```
UID位数（7/8/9/10，回车=7）: 7
UID前缀（可不填）: 100
UID后缀（可不填）:
B站Cookie SESSDATA值（不填则慢）: xxx
线程数（回车=默认8）: 15
```

### 3. 导出 Excel

爬虫跑完后运行：

```bash
python export_excel.py
```

生成 `bilibili_results/B站账号清单.xlsx`，只含 UID 和昵称两列。

## UID 范围说明

| 输入 | 扫描范围 | 数量 |
|------|----------|------|
| 前缀 `100`，7位 | 1000000 ~ 1009999 | 10,000 |
| 前缀 `100`，8位 | 10000000 ~ 10099999 | 100,000 |
| 后缀 `888`，7位 | xxx0888 ~ xxx9888 | 9,000 |
| 前缀`12` + 后缀`99`，7位 | 1200099 ~ 1299999 | 10,000 |

位数越多 → 前缀/后缀越长 → 范围越小 → 耗时越短。

## Cookie 获取方法

1. 打开 [bilibili.com](https://www.bilibili.com) 并登录
2. 按 F12 打开开发者工具
3. 切换到 **Application**（应用）标签
4. 左侧展开 **Cookies** → 点击 `https://www.bilibili.com`
5. 找到 **SESSDATA** 那一行，复制它的值
6. 粘贴到爬虫的 Cookie 输入框

> SESSDATA 是登录凭证，请勿泄露给他人。

## 文件结构

```
├── crawler.py          # 爬虫主程序（双击运行）
├── export_excel.py     # Excel 导出工具（双击运行）
├── requirements.txt    # 依赖列表
├── .gitignore          # Git 忽略规则
├── LICENSE             # MIT 开源协议
└── README.md           # 本文件
```

运行后自动生成：

```
bilibili_results/
├── result_xxxxx.csv       # 每次运行的 CSV 结果
├── result_xxxxx.json      # 每次运行的 JSON 结果
├── _progress.json         # 断点续传进度文件
└── B站账号清单.xlsx        # Excel 导出（需手动运行导出工具）
```

## 注意事项

- 无 Cookie 时速度约 0.3~1 uid/s，有 Cookie 后可达 3~10 uid/s
- 线程数建议 8~15，超过 20 可能触发 B站风控（code=-352/-799）
- 限速后程序会自动等待重试，无需手动干预
- 扫描超过 50 万 UID 时会弹出确认提示
- 本工具仅供学习研究，请勿用于商业用途或大规模滥用

## License

MIT License - 详见 [LICENSE](./LICENSE)
