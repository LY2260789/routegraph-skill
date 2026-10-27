# Phase 0 / Phase 1 交付与验收

日期：2026-10-26。维护源为 `routegraph-skill-repo`。本次不修改已安装 Skill 目录，不覆盖用户原始 PPT、预览或输入素材。安装包可用于之后部署。

## 架构诊断与实施结论

原生对象实际由 PowerShell + PowerPoint COM 创建。演示文稿 Skill 在原仪器案例中承担文件校验和导入检查，没有承担几何绘图。具体审计和旧脚本调用关系见 `reconstruction_audit.md`。

新增轻量链路：Instrument Scene → Python 标准库编译器 → 中性绘图计划 → COM 渲染 → PPTX/PNG/回执 → 原生结构验证与预览审阅 → 局部补丁。技术路线模板保留原路径和原接口。

Instrument Scene 与未来 Diagram Graph 分开。仪器关系承载装配/空间，流程图边承载有向流。两者共享来源、置信度、标注和样式基础字段。多图整合使用组合清单设计，本轮没有实现跨图识别或自动合并。

## 已实现范围

- Cylinder、HollowCylinder、Cuboid 的尺寸、位置、旋转、材质参数化。
- 统一 oblique 投影参数；分面命名，显式绘图顺序，遮挡关系检查。
- Scientific Metallic 的 steel、gray_polymer、blue_polymer 材质。保持旧案例中的打印聚合物语义。
- NativeReference 兼容层：按名称复制已有 PPT 的原生细节，校验文件哈希和对象类型，保留编辑性。
- 语义组件、引用与置信度校验、标签和连续引线、原生分组。
- 输出 PNG、原生对象统计、标签溢出、画布越界、位置比例、遮挡顺序检查。
- 带来源哈希和修改理由的局部 JSON 补丁；修正后重新编译和渲染，不覆盖输入。

未实现：照片自动分割、OCR、真实内部结构推断、物理 CAD、自动视觉评分、一般化槽盒、电极/复杂板件的全部参数化、全场景自动风格切换。

## 修改与新增文件

原有文件仅改：

- `SKILL.md`：增加仪器模式路由、能力边界和本地预览检查。
- `README.md`：新增架构、运行与回归命令。
- `.gitignore`：允许跟踪作为测试依据的只读旧 PPT 副本。

新增：

- `scripts/reconstruction/scene.py`：模型校验、几何编译、局部补丁。
- `scripts/reconstruction/verify.py`：原生 PPTX 与渲染回执验证。
- `scripts/render_instrument.ps1`：COM 绘制、语义分组、导出及文本边界测量。
- `assets/style_presets/scientific_metallic.json`：材质配置及后续风格接口。
- `assets/examples/instrument_reconstruction/legacy_reference.pptx`：旧产物原样副本。
- `assets/examples/instrument_reconstruction/charging_collection.scene.json`：迁移场景。
- `assets/examples/instrument_reconstruction/label_fit.patch.json`：标签框局部修正记录。
- `assets/examples/instrument_reconstruction/charging_collection.reviewed.scene.json`：修正后的可运行场景。
- `tests/test_reconstruction.py`：21 项单元测试。
- `tests/make_gallery.py`：独立解析几何案例。
- `tests/check_legacy_arrows.py`：旧路线拓扑断言。
- `tests/test_native_editing.ps1`：实际拆组/编辑/保存重开检查。
- `tests/run_regressions.ps1`：完整回归入口。
- `references/reconstruction_audit.md`：Phase 0 审计与增量架构。
- `references/instrument_reconstruction.md`：字段约定、调用关系、限制与后续设计。
- 本验收报告。

旧模板、旧样式、旧生成脚本、另一个工作副本及用户安装目录未修改。工作区 `.routegraph-build/phase1` 保留开发过程与校验回执，不属于 Skill 的运行依赖。

## 测试与可编辑性

最终一键运行目录为工作区 `.routegraph-build/phase1/regression-complete`，退出码 0。21/21 单元测试、仪器场景校验、独立基础几何校验、旧箭头拓扑检查和原生编辑保存重开检查全部通过。仪器与基础几何规则报告均无 error 或 warning。Skill 入口检查 `quick_validate.py` 通过；该检查使用已有 Anaconda 的 PyYAML，无新增安装依赖。新重建模块本身只使用 Python 标准库。

运行命令见 README。21 项单元测试覆盖非法孔径、非有限数、未知几何、缺失来源/关系/标签锚点、遮挡顺序、原生引用缺失/哈希不匹配、禁止覆盖、补丁不变性/过期补丁、材质替换保持几何等。

COM 集成测试实际拆开空心圆柱的 5 个原生表面，修改外壁尺寸与颜色、修改标签文字、重新分组、保存并重开检查。测试使用副本，源 PPT 保持不变。

| 案例 | 输入可用性与结果边界 | 输出原生对象 |
|---|---|---|
| 旧简单路线/连续箭头回归 | 使用仓库现有可运行模板。原来的历史简单路线参考图片未找到，不宣称已做该图片 OCR 复刻回归 | 一个逻辑箭头对应一个对象；检查直线、虚线折线、多折点路径和真实分叉 |
| 静电采样装置爆炸图 | 旧 PPT、旧预览、旧生成脚本可用；论文原图不可用。本次是对旧产物的结构化迁移与回归 | 19 个语义组件：10 个参数化组件、9 个原生引用组件；15 个 PPT 分组；88 个 p:sp、38 个连接/线对象；12 个文本对象包含在 p:sp 中 |
| Photonion 照片 | 工作区未找到实际照片，未执行，未生成假案例 | 不适用 |
| 独立基础几何图 | 解析几何合成测试，用于验证通用模块；不计作 Photonion 替代 | 3 个分组，15 个原生形状，其中 4 个文本对象 |

仪器图的 126 个底层形状/线对象与 15 个分组为不同计数层级；文本是形状的子集，不重复相加。图片、SVG、OLE、graphic frame、嵌入媒体均为 0。所有 12 个原标签/标题和 9 条标注箭头保留。

## 与旧图相比的改进与剩余问题

改进：不再为基础组件逐次硬编码绘图函数；空心管分成外壁、端面、内壁、暗面等可命名表面；金属环与聚合物主体有独立材质；盒体由尺寸和基向量生成；细节有语义分组；组件来源与不确定性可追踪；产生机器可读校验报告和局部修正记录。

实际修复：关闭 PowerPoint 文本框继承的 AutoSize，固定文本框尺寸；增加透明标签框余量；修正旋转椭圆包围盒计算；阴影保留为独立原生对象，避免多椭圆分组包围盒近似造成误判。

剩余问题：旧图中的电极、槽纹、复杂收集板与阴影仍通过原生引用保留，未自动参数化或重着色。充电与收集区沿用各自的局部投影约定，尚非全局求解的同一 3D 相机。空心管暗面为科研示意式绘法，不能据此断言内部构造。不能计算相对于缺失论文原图的准确率。

验证以语义/规则和组件区域比较为主，未使用全图像素相似度。科学事实最终核验仍需论文原图及相关说明。

## 兼容性与后续顺序

旧模板内容与 helper 签名保持不变。回归中仅在临时副本去除 `Quit`，避免关闭用户 PowerPoint。旧模板独立运行正常；与新渲染器同一 PowerShell 进程连续运行时出现 Boolean/Int32 COM 绑定错误，回归入口将旧模板放到独立 shell 进程，保持其原绘图行为。模板本身还保留浅底白字等原有视觉弱点，本次未作为仪器升级顺带修改。

后续先取得论文原图进行部件/标签/遮挡的科学核验，再用实际 Photonion 照片只重建可见外观。随后完成 Diagram Graph 的文本节点/有向边可靠复刻，再实施多图组合。Scientific Flat 在原生细节逐步参数化后做几何不变性回归；Anime Tech 最后考虑。

Phase 1 达成的是可复用的最小重建链路与旧案例迁移，不等于全部仪器照片识别或全部高级几何已完成。

## 交付位置

工作区 `output/phase1/`：

- `charging_collection.pptx` / `.png`：迁移后的采样装置图。
- `primitives.pptx` / `.png`：独立基础几何验证图。
- `legacy_arrows.pptx` / `.png`：旧路线箭头回归图。
- `validation-summary.json`：最终机器可读汇总。
- `routegraph-phase1.zip`：完整 Skill 源码、说明、测试与必要的旧原生参考副本。排除 Git 元数据和缓存。

额外的测试修改副本 `native_edit_roundtrip.pptx` 仅留在构建目录，避免将故意改成红色的编辑性测试图误认为最终科研图。
