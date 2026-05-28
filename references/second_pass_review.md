# RouteGraph 二次审图规范

## 目标

二次审图用于比较：

```text
原始论文图 / 用户参考图
vs.
PowerShell COM 生成图截图
vs.
当前 PowerShell 代码
```

并给出修正方案。

## 审图优先级

### 1. 箭头拓扑

优先检查：

- From 是否正确；
- To 是否正确；
- 方向是否正确；
- 样式是否正确；
- 是否缺少关键箭头；
- 是否多了误导性箭头。

### 2. 箭头路由

检查：

- 是否斜穿框；
- 是否压住文字；
- 是否和其他箭头混淆；
- 是否该使用分段折线；
- 箭头头部是否在正确端点。

### 3. Panel 结构

检查：

- panel 顺序是否正确；
- panel 标题是否对应；
- A/B/C/D 位置是否接近原图；
- 各元素是否归属正确 panel。

### 4. 文本与符号

检查：

- 下标是否按可编辑规则简化；
- `Delta` 是否应显示为 `Δ`；
- 是否有乱码；
- 是否有单词被拆行；
- 字体大小是否稳定。

### 5. 布局和美观

检查：

- 对齐；
- 间距；
- 背景块大小；
- 颜色饱和度；
- 右侧或底部是否被裁切。

## 审图报告模板

每次用户发生成图截图后，先输出简短审图报告：

```text
Second-pass Review

1. Arrow topology
   - A1: A.XGBoost → B.Counterfactual conc., correct / incorrect
   - B1: B.Counterfactual conc. → C.Counterfactual PM2.5, correct / incorrect

2. Arrow routing
   - B→C should use segmented dashed route instead of diagonal line.

3. Text
   - PM2.5 simplification is acceptable.
   - Delta should be generated using [char]0x0394.

4. Layout
   - D panel is too far right and should be shifted left by 20 px.

5. Code action
   - Replace B→C arrow block.
   - Shift D background and boxes left.
```

## 输出修正代码

如果问题集中在某个 panel，优先输出局部替换代码。

如果问题较多，输出完整修订版脚本。
