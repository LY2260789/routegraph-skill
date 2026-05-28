# RouteGraph 文本与符号规范

## 核心原则

RouteGraph 默认采用 **可编辑优先** 原则。也就是说，生成的 PPT 应方便用户二次修改，而不是强行一次性复刻所有上下标和数学格式。

## 1. 下标默认简化

论文图里常见下标，默认识别为普通数字。

| 原图 | RouteGraph 默认输出 | 原因 |
|---|---|---|
| PM₂.₅ | PM2.5 | PPT 中手动下标很方便 |
| NO₂ | NO2 | 避免 Unicode 下标字体不一致 |
| SO₂ | SO2 | 避免编码和字体问题 |
| O₃ | O3 | 保持可编辑 |
| CO₂ | CO2 | 保持可编辑 |

默认不要在 PowerShell 字符串里写：

```powershell
"PM$_2.5"
```

PowerPoint 不识别 LaTeX，会显示错误。

## 2. Delta 符号

`Delta` 应尽量显示为 `Δ`，但不要直接依赖脚本文本编码。

推荐写法：

```powershell
$Delta = [char]0x0394
Add-TextBox $slide "$Delta PM2.5 air quality" 340 213 190 22 13 $true $black
```

或：

```powershell
$text = $Delta + " PM2.5 air quality"
```

## 3. 常见希腊字母转义

| 符号 | PowerShell 写法 |
|---|---|
| Δ | `[char]0x0394` |
| α | `[char]0x03B1` |
| β | `[char]0x03B2` |
| γ | `[char]0x03B3` |
| μ | `[char]0x03BC` |
| σ | `[char]0x03C3` |
| ρ | `[char]0x03C1` |

## 4. 单位规则

第一版 RouteGraph 默认以可编辑为主。

| 原图 | 默认输出 |
|---|---|
| μg m⁻³ | ug m-3 或 μg m-3 |
| m s⁻¹ | m s-1 |
| W m⁻² | W m-2 |

如用户要求论文级格式，再考虑 Unicode 上标或 PPT 二次手动精修。

## 5. 手动换行

需要换行时，主动使用 PowerShell 换行符：

```powershell
"Counterfactual`nemission strength"
```

不要依赖 PowerPoint 自动换行。

## 6. 自动换行

默认关闭：

```powershell
$shape.TextFrame.WordWrap = 0
```

如果文字超出或被裁切，优先增大框宽，而不是让 PPT 自动拆单词。

## 7. 可接受差异

二次审图时，下列情况属于可接受简化，不需要自动修正：

- `PM₂.₅` → `PM2.5`
- `NO₂` → `NO2`
- `SO₂` → `SO2`
- `O₃` → `O3`

下列情况应修正：

- `Delta` 应改为 `Δ`
- `PM$_2.5` 显示为错误字符
- `XGBoost` 被拆成两行
- 英文单词被无意义拆开
