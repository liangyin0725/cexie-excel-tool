# 测斜 Excel 修正工具

一个 Windows 桌面小程序，用来批量读取测斜 `.xlsx` 原始数据，按点号和列名设置修正规则，并生成新的修正后 Excel。

## 使用方式

1. 运行 `测斜Excel修正工具.exe`。
2. 在 Step 1 选择包含原始 Excel 的文件夹。
3. 在 Step 2 选择点号，为每一列设置操作和值。
   如果所有 Excel 使用同一套修正规则，先在任意一个点号填好规则，再点击“应用到全部点号”。
4. 可选：点击“保存参数”把当前规则保存为 `.json`，下次用“加载参数”复用。
5. 在 Step 3 点击“生成修正后 Excel”。

输出文件会放在原文件夹下的 `corrected` 子文件夹中，原始文件不会被修改。

## 修正规则

- `不修改`：该列保持原样。
- `加 +`：数值列统一加一个数。
- `减 -`：数值列统一减一个数。
- `乘 ×`：数值列统一乘一个数。
- `除 ÷`：数值列统一除以一个非零数。
- `公式`：用公式计算新值，`x` 表示本列原单元格数值，`[列名]` 表示同一行其他列的原始数值，`rand()` 表示 0 到 1 之间的均匀随机数。例如 `x + [A0] * 0.1 + rand() * 0.5`。
- `替换`：该列统一替换为填写的内容。

空单元格会保持空白。数值运算遇到非数字内容时会保持原值，并在生成结束后提示。

生成文件时，`A轴管口起算`、`A轴管底起算`、`B轴管口起算`、`B轴管底起算` 会自动重算：

- A轴单行增量：`(A180 - A0) / 2`，先四舍五入到 2 位
- B轴单行增量：`(B180 - B0) / 2`，先四舍五入到 2 位
- 管口起算：从上到下累加单行增量
- 管底起算：从当前行到最后一行反向累计，并取负值

这些列基于修正后的 `A0`、`A180`、`B0`、`B180` 计算。如果对四个起算列本身也设置了修正规则，最终会被自动重算结果覆盖。

## 开发运行

```powershell
$py='C:\Users\liangyin725\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
& $py run_app.py
```

## 打包

先安装依赖：

```powershell
$py='C:\Users\liangyin725\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
& $py -m pip install -r requirements.txt
```

然后打包：

```powershell
.\build_exe.ps1
```

生成文件：

```text
dist\测斜Excel修正工具.exe
```
