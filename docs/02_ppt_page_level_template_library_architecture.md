# PPT 页面级模板库技术架构方案

## 1. 目标

把完整 PPT 拆成可独立使用的资源单元，形成四层资产体系：

1. PPT 页面级模板库
2. PPT 组件级素材库
3. 智能页面检索系统
4. 智能组装系统

最终让程序不是“找一个最像的完整模板”，而是“按需求拼出一套最合适的页面组合”。

## 2. 推荐架构分层

### 2.1 资源发现层

负责收集合法、可复用的模板来源：

- GitHub / GitLab 开源仓库
- Creative Commons / Public Domain 资源
- 免费模板站点
- HTML / Reveal.js / Marp / Slidev / Beamer 模板

### 2.2 解析与验证层

负责把原始资源变成结构化对象：

- 文件类型识别
- 损坏检测
- 许可证识别
- 是否允许商用
- 页面数、比例、语言、作者、来源站点提取

### 2.3 页面拆解层

负责把一个完整 PPT 拆成单页：

- slide 级解包
- 页面类型识别
- 文本 / 图片 / 图表 / 表格 / 流程 / 时间轴识别
- 页面截图生成
- 页面独立 metadata 建档

### 2.4 组件提取层

从高质量页面中抽出可复用组件：

- 标题区
- 副标题区
- KPI 卡片
- 图表区
- 表格区
- 时间轴区
- 流程区
- 对比区
- 组织架构区
- 页面底栏 / 页眉 / 页码

### 2.5 搜索与推荐层

采用“标签 + 向量 + 规则”三路并行：

- 标签检索：行业、场景、风格、颜色、页型、内容结构
- 语义检索：标题、正文、描述、用途 embedding
- 视觉检索：页面截图相似度、结构相似度

### 2.6 组装与设计层

负责把选中的页面重新变成统一风格的 PPT：

- 统一主题色
- 统一字体与字号
- 统一页边距和留白
- 统一图标风格
- 统一图表配色
- 统一页面节奏

## 3. 建议的数据模型

### 3.1 source_catalog

记录资源来源站点或仓库。

关键字段：

- source_id
- source_url
- source_website
- license_hint
- commercial_use_hint
- modification_allowed_hint
- download_method
- programmatic_access
- quality_hint
- tags
- score

### 3.2 template_catalog

记录完整模板或模板集合。

关键字段：

- template_id
- template_name
- source_url
- license
- commercial_use
- modification_allowed
- file_format
- slide_count
- aspect_ratio
- language
- quality_score
- duplicate_hash
- preview_image
- local_path

### 3.3 slide_catalog

记录单页页面资源。

关键字段：

- slide_id
- source_template_id
- slide_number
- slide_type
- slide_subtype
- industry
- scenario
- style
- layout_type
- primary_color
- secondary_color
- aspect_ratio
- text_density
- image_density
- chart_count
- table_count
- has_timeline
- has_process
- has_map
- has_people
- has_infographic
- overall_quality_score
- preview_image
- thumbnail_path
- slide_file_path
- embedding_vector

### 3.4 component_catalog

记录页面内部高质量组件。

关键字段：

- component_id
- component_type
- component_subtype
- source_slide_id
- bounding_box
- style_token
- color_token
- preview_image
- metadata_json

## 4. 目录结构建议

```text
ppt_intelligence_system/
  template_library/
  slide_library/
  component_library/
  metadata/
  embeddings/
  preview/
  database/
  search_engine/
  recommendation_engine/
  layout_engine/
  content_adapter/
  chart_engine/
  excel_analyzer/
  ppt_assembler/
  quality_scoring/
  duplicate_detection/
  user_feedback/
  logs/
```

## 5. 搜索与评分策略

先过滤，再排序，再组合。

### 5.1 过滤

- license 是否允许
- 商用是否允许
- 比例是否匹配
- 页型是否匹配
- 是否存在明显损坏

### 5.2 排序

建议综合分数：

- 主题匹配
- 行业匹配
- 场景匹配
- 页型匹配
- 内容结构匹配
- 视觉风格匹配
- 颜色匹配
- 页面质量
- 可编辑程度
- 许可证友好度

### 5.3 页面级组合规则

不能只取“最像”的页面。应确保：

- 内容数量匹配
- 图表密度匹配
- 文字密度匹配
- 前后页节奏一致
- 封面 / 目录 / 内容 / 总结页完整

## 6. 技术边界

### 本地即可完成

- 文件解析
- 哈希去重
- 页面切片
- 预览图生成
- 标签入库
- 规则排序
- Excel 图表渲染

### 建议 LLM 参与

- 用户意图归纳
- 页面用途推断
- 内容改写与摘要
- 页面组合建议
- 生成章节结构

### 建议视觉模型参与

- 页面类型识别
- OCR
- 视觉质量评分
- 感知去重
- 页面风格分类

## 7. 结论

这个架构的关键不是“更多模板”，而是“更小粒度、更强索引、更强组合”。只要页面与组件都能被稳定识别和检索，后续生成就能从“套模板”升级成“智能组装”。
