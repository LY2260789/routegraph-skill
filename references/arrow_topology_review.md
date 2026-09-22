# RouteGraph 箭头拓扑审查规范

## 核心原则

技术路线图的核心不是框，而是箭头。箭头决定科学流程、方法逻辑和因果/数据流关系。

RouteGraph 的二次审图中，箭头优先级最高。

## 1. 生成代码前必须输出箭头拓扑表

建议格式：

| Arrow ID | From | To | Style | Meaning | Priority |
|---|---|---|---|---|---|
| A1 | A.XGBoost | B.Counterfactual conc. | purple dashed | BAU prediction | high |
| B1 | B.Counterfactual conc. | C.Counterfactual PM2.5 | purple dashed | counterfactual transfer | high |
| C1 | C.XGBoost | D.Counterfactual emission strength | purple dashed | fixed-met prediction | high |
| C2 | C.XGBoost | D.Emission strength | purple dashed | fixed-met prediction | high |

## 2. 箭头错误分级

### 高优先级错误

必须修正：

- 箭头源对象错误；
- 箭头目标对象错误；
- 箭头方向相反；
- 关键箭头缺失；
- 多余箭头改变了方法逻辑；
- 箭头连接到错误 panel。

### 中高优先级错误

通常需要修正：

- 紫色虚线被画成蓝色实线；
- 蓝色实线被画成紫色虚线；
- 一个逻辑虚线箭头被拆成多个独立短线对象，导致虚线节奏在拐点处断裂；
- 箭头像由多个刚性线段硬拼，缺少连续 PPT 连接符/自由曲线的整体感；
- 箭头头部位置错误；
- 箭头穿过关键文字；
- 箭头被方框遮挡导致无法理解。

### 低优先级问题

可后期微调：

- 箭头略长；
- 端点扎进框内一点点；
- 虚线间距不完全一致；
- 线条位置不够美观但逻辑正确。

## 3. 颜色与语义

| 样式 | 含义 |
|---|---|
| 蓝色实线 | 模型训练、输入输出、内部结构关系 |
| 紫色虚线 | 反事实预测、情景传递、跨阶段数据流 |
| 黑色实线 | 注释输入、模拟输入、局部说明 |
| 灰色虚线 | 辅助关系、弱关系、可选路径 |

## 4. 方向判断

PowerShell 里：

```powershell
Add-LineArrow $slide $x1 $y1 $x2 $y2 ...
```

代表方向：

```text
(x1, y1) → (x2, y2)
```

因为箭头头部在终点：

```powershell
$line.Line.EndArrowheadStyle = 3
```

二次审图必须检查起点和终点是否符合原图。

## 5. 路由规范

复杂箭头优先使用一个连续 shape：

```text
水平线段 + 垂直线段 + 水平箭头
```

或：

```text
垂直线段 + 水平线段 + 最后一段箭头
```

在 PowerShell 中应优先用 `Add-PolylineArrow` 一次生成这些折点，而不是分别调用多个 `Add-LineSegment` 和 `Add-LineArrow`。这样虚线会沿同一个 shape 连续排布。

避免斜线穿过图形。

## 6. 模型框遮挡路线

如果原图中路线像是从模型框后面穿过，可使用：

```text
先画路线，再画模型框。
```

这样模型框会覆盖中间线段，视觉上更接近论文图。
