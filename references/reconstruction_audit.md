# Scientific Figure Reconstruction：Phase 0 审计

审计日期：2026-10-26。实施目标为本仓库的增量升级，不修改安装目录或另一个工作副本。

## 实际调用关系

`SKILL.md → Agent 读图与拓扑规划 → assets/templates/*.ps1 → PowerPoint COM → 原生 PPTX / Slide.Export PNG`。
现有模板是可复制脚本，不是库。`horizontal_pipeline_template.ps1` 只有注释，`four_panel_route_template.ps1` 有绘图 helper 和空白布局占位；`pm25_counterfactual_framework.ps1` 是完整硬编码例图；`continuous_arrow_regression.ps1` 是可运行回归。

工作区 `.routegraph-build/recreate_charging_collection.ps1` 实际生成仪器形状。`Poly/Ellipse/Grad/Segment/Cylinder/Shadow` 均直接调用 COM，Cylinder 内部已做原生分组。Collection 的多边形、槽纹、电极尚是单图坐标。该脚本未进入 Skill，也未被模板调用。

`.routegraph-build/finalize_charging_collection.mjs` 调用演示文稿 Skill 的 `finalizePresentation`，检查 PPTX 包、布局、字体并尝试 Artifact Tool 导入。它没有生成本案例的原生形状。原校验 JSON 的通过不证明科学准确性或实际视觉正确。

演示文稿 Skill 的通用默认创作工具是 Artifact Session / Artifact Tool；它也明确允许用户要求的可编辑图使用原生对象。本任务明确要求保留 COM 与旧接口，因此 Routegraph 负责科研语义、几何和 COM 绘制，通用演示文稿工具只作为可选最终文件检查。无需复制另一套幻灯片引擎，不新增该工具的强制依赖。

## 复用与风险

- 复用现有 Freeform / Ellipse / TwoColorGradient / Shapes.Range.Group / Slide.Export 的实现模式和 single 类型转换。
- 旧箭头 helper 与模板接口不改。历史 case notes 中分段箭头建议已过时，以 SKILL 的连续箭头规则为准。
- 旧模板会写桌面、显示 PowerPoint、甚至 Quit 整个应用；新入口只关闭自己创建的演示文稿，不退出用户 PowerPoint。测试旧模板使用临时隔离副本，只替换 Quit 行。
- 无 Python 库、统一几何模型、自动图像识别、机器视觉评分或单元测试框架。不把 Agent 读图描述为已经实现的 OCR。
- 工作区有 `routegraph-skill-repo`（独立 Git 仓库，基线 3b1d585）、`routegraph-skill` 和用户安装目录。三者入口 SHA256 相同。本次以 repo 为唯一维护源，其他副本不自动同步。
- 可访问旧仪器 PPT、预览、生成脚本和连续箭头案例。论文原图、Photonion 照片、历史简单路线原图未在本工作区找到。旧输出只能作为回归基准。

## 最小架构与归属

`scripts/reconstruction/scene.py`：与 COM 解耦的 Instrument Scene 校验、投影、参数化几何编译、局部补丁。仅 Python 标准库。

`scripts/render_instrument.ps1`：消费中性的绘图计划，将 ellipse/polygon/line/text 与 native_reference 转成 PPT 原生对象。图层由场景 z_order 控制，组件拥有语义分组。

`scripts/reconstruction/verify.py`：检查原生对象、标签、组件边界、关系顺序、来源校验和参考区域漂移；输出分维度报告，不给伪综合相似度。

`assets/style_presets/scientific_metallic.json`：材质参数独立于几何；Flat 和 Anime 只设计接口，本轮不拓展完整风格。

`assets/examples/instrument_reconstruction/`：旧产物的只读副本与语义化迁移 Scene。复杂细节通过显式 native_reference 适配器保留，不能把该适配器宣称为通用几何或视觉识别能力。

`tests/`：模型约束、几何尺寸/旋转/孔径、补丁不可变性、PPT 编辑性和旧箭头回归。

## 两类图模型

共享 schema_version、source_reference、confidence、label、style、坐标单位 pt。Instrument Scene 使用 components、assemblies、projection、装配/遮挡关系；Diagram Graph 使用 nodes、ports、edges、direction、route、panel。流程边不是装配边，禁止强行共用同一含义字段。

Phase 1 只实现 Instrument Scene。Diagram Graph、OCR、自动布局、多图实体对齐在后续阶段设计实施；旧技术路线直接走旧模板。

## 验证与后续顺序

先验证科学结构，再验证可编辑性，之后才比较视觉。结构化规则检查不替代人读论文原图。局部修正通过有来源理由的 JSON 补丁生成新 Scene，再编译和渲染，永不覆盖输入。

下一步依次为：取得论文原图并人工核对当前迁移；Photonion 可见外观与不确定性标注；Diagram Graph 和节点/箭头可靠复刻；多图整合；Scientific Flat 的几何不变性回归；最后才考虑 Anime Tech。
