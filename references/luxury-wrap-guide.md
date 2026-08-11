# 豪华版封面/目录/结束页设计规范

> 当用户要求「封面 / 目录 / 结束页」升级为 **豪华版式、多装饰、多层次、信息丰富** 时，按本规范生成。
> 配合 `scripts/build_wrap_pages.py` 一键套用。

## 适用场景

模板（如 architecture-deck / premium-corp）的 `page_roles.cover / agenda / ending` 为空、没有封面目录结束页时；
或用户明确要求「增加封面、目录、结束页」时。

## 设计原则

1. **强对比标题**：主标题 40-46pt 加粗，白字 / 深色底，英文衬底 12-13pt。
2. **渐变背景**：主色→深色垂直渐变（`fill.gradient()`），不用纯色平铺。
3. **几何装饰**：2-3 个圆形（半/不透明）+ 金色圆点 + 顶部金色横线 + 左侧装饰竖线。
4. **图标系统**：用 Unicode 符号（▣ ◉ ◈ ✉ ✆ ⌂ ◆ ▸）代替纯文字标签。
5. **层次化排版**：标签 → 主标题 → 副标题 → 分隔线 → 信息卡 → 装饰。
6. **信息丰富**：封面右侧加 `HIGHLIGHTS` 亮点清单（≤4 项）；结束页加「下一步计划」板块 + 联系方式。
7. **留白控制**：每元素距画布边 ≥ 1.2cm，双列布局均衡。

## 主题色（--accent）

| 名称 | primary | dark | mid | gold | 适用模板 |
|---|---|---|---|---|---|
| deepblue | #1F3A93 | #0E1B42 | #2A4FB8 | #C99D52 | architecture-deck |
| red | #A52524 | #5E1212 | #C93A37 | #E8C57E | premium-corp / mckinsey-style |
| gold | #8A6D1F | #3E320E | #B89535 | #E8D99E | 政务 / 高端商务 |
| green | #1E5B4C | #0F2E26 | #2E7D69 | #C9DD9A | 环保 / 科技 |

## 三页版式

### 封面
- 左上：金色标签 + 装饰竖线
- 主标题（两行，42pt）+ 副标题（26pt）+ 英文衬底
- 金色分隔线
- 3 个信息卡（部门 / 版本 / 日期）
- 右侧 HIGHLIGHTS 亮点（4 项）
- 右上 LOGO 占位 + 底部项目条

### 目录
- 顶部主色条 + 金色细线
- 左侧 CONTENTS 标签 + 大标题 + "N SECTIONS"
- 章节条目双列：序号方块（主色）+ 标题（16pt bold）+ 副标（10pt）+ 装饰线 + 金色圆点
- 底部项目信息

### 结束页
- 渐变背景 + 四角金色装饰角 + 圆点
- 左下 "END" 水印（用接近背景的深色，不抢主视觉）
- 中央「感谢聆听」56pt + "THANKS FOR LISTENING"
- 「下一步计划 NEXT STEPS」板块（深色底 + 金色标题）
- 4 列联系方式 + 底部项目条

## 实现方式（关键）

- **deepcopy 方案**：`build_pptx.py` 输出带残留 parts 的文件（如 architecture-deck 模板含
  slide1-slide37），直接 `add_slide` 会触发 part 命名冲突导致 `soffice` 无法加载。
  正确做法：
  ```python
  new_prs = Presentation()          # 全新空 PPT
  new_prs.slide_width = Cm(33.87)   # 16:9
  for src in src_prs.slides:        # deepcopy 内容页
      ns = new_prs.slides.add_slide(blank)
      for shape in src.shapes:
          ns.shapes._spTree.append(copy.deepcopy(shape._element))
  # 然后 add_slide 加 3 张豪华页，最后重排 sldIdLst
  ```
- **重排 sldIdLst**：封面, 目录, 内容..., 结束页。

## 脚本用法

**最小一行调用（目录自动提取、标题取文件名）：**

```bash
python3 scripts/build_wrap_pages.py 输入.pptx
# → 生成 输入_wrapped.pptx（封面 + 目录 + 原内容 + 结束页）
```

**完整自定义调用：**

```bash
python3 scripts/build_wrap_pages.py out.pptx \
  --title "核电研发与工程软件产品线" \
  --subtitle "组织运作方案（初稿）" \
  --tagline "NUCLEAR POWER · SOFTWARE LINE" \
  --dept "核电软件产品线" --version "V1.0" --date "2026 年" \
  --toc '{"01":"方案总则","02":"组织架构与整体定位","03":"各单元核心岗位职责","04":"核心运作流程","05":"管理制度与规范","06":"阶段运作目标","07":"保障措施","08":"附则"}' \
  --toc-sub '{"01":"编制目的 · 运作定位 · 适用范围","02":"1产线/3赛道/2科室","03":"三大赛道 + 两大科室","04":"需求闭环 · 协同联动","05":"进度 · 质量 · 考核","06":"短期 · 中期 · 长期","07":"人员 · 资源 · 技术","08":"修订说明"}' \
  --highlights "效率提升 60%+;重复事务自动处理;知识资产沉淀;全员 AI 助手" \
  --accent deepblue \
  --out 输出.pptx
```

- `--toc` 缺省时自动从每张内容页提取标题（取字号最大的文本 run）。
- `--title` 缺省取文件名；`--out` 缺省为 `<文件名>_wrapped.pptx`。
- `--accent` 见上方主题色表。
