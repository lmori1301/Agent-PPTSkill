# Agent-PPTSkill 安装说明

> **Agent-PPTSkill**（原名 GordenPPTSkill）是一款面向 AI 助手的 PPT 生成技能：
> 内置 21 套中文 PPT 模板，从模板生成真实 `.pptx`，只替换文字、不破坏原排版。
> 仅供个人学习与研究，**严禁商业用途**。

---

## 一、安装步骤

### 1. 放置技能目录

将本压缩包解压后得到的 `Agent-PPTSkill` 文件夹，放到 AI 助手的技能目录：

```bash
# macOS / Linux
mkdir -p ~/.workbuddy/skills
cp -r Agent-PPTSkill ~/.workbuddy/skills/

# Windows（PowerShell）
# 将 Agent-PPTSkill 文件夹复制到 %USERPROFILE%\.workbuddy\skills\
```

解压后应确认以下关键文件存在：

```
Agent-PPTSkill/
├── SKILL.md          ← 技能入口（AI 通过它识别）
├── VERSION           ← 版本号（当前 v1.0.20）
├── scripts/          ← 构建/渲染脚本
├── templates/        ← 21 套内置模板
└── references/       ← 工作流与编辑规则
```

### 2. 安装 Python 依赖

```bash
pip3 install python-pptx
```

### 3. 可选：渲染预览图所需工具（仅预览/自检需要，生成 PPT 不需要）

```bash
# macOS（推荐 Homebrew）
brew install --cask libreoffice poppler

# Debian/Ubuntu
sudo apt install libreoffice poppler-utils

# Windows
# 安装 LibreOffice：https://www.libreoffice.org/download
# 安装 poppler：https://github.com/oschwartz10612/poppler-windows/releases
```

---

## 二、验证安装

1. 打开你的 AI 助手（如 WorkBuddy / Cursor）
2. 发送：`用 PPT Skill 做一份季度总结 PPT`
3. 若能列出模板选择，说明安装成功

---

## 三、使用方式（三种模式）

| 模式 | 适用场景 | 说明 |
|---|---|---|
| **模式 A** | 用内置模板 | 默认方式。AI 读取 `templates/INDEX.md`，按场景/颜色匹配模板，展示预览图让你选 |
| **模式 B** | 用自己的 .pptx 做模板 | 上传你自己的 PPT，AI 按它的排版替换文字 |
| **模式 C** | 完全原创 | 不用任何模板，AI 直接生成简洁版式 |

### 常用指令示例

```
做一个复杂、豪华的 PPT，介绍 XXX 项目          # 模式 A，AI 会先让你选模板
按这个模板做一份新的：<上传 .pptx>              # 模式 B
做一份原创简洁风格的 PPT，主题是 YYY            # 模式 C
```

---

## 四、常见问题

**Q：生成 PPT 时报 `ModuleNotFoundError: No module named 'pptx'`**
A：未安装 python-pptx，执行 `pip3 install python-pptx`。

**Q：输出提示"N 处文字偏长"**
A：正常提示，不影响保存。这是模板容量估算，1-2 字超出可接受。

**Q：渲染预览图失败（render_slides.py）**
A：缺少 LibreOffice 或 poppler，按第一部分第 3 步安装。

**Q：如何更新到最新版？**
A：技能目录里运行：
```bash
python3 scripts/apply_update.py
```
会自动从 GitHub 拉取增量更新（更新源：`git+https://github.com/lmori1301/Agent-PPTSkill.git#main`）。

**Q：占位文字没被替换？**
A：向 AI 明确要求"替换所有占位文字"，或检查模板容量是否限制。完整 PPT 不应残留任何占位词。

---

## 五、非商业声明

本技能与内置 PPT 模板**仅供个人学习和研究**，**严禁用于任何商业用途**
（含商业演示、销售、培训分发、客户提案、企业内部以营利为目的的使用等）。
模板素材来自第三方设计师作品，二次商用需获取原作者授权。

---

*安装有问题可联系维护者排查。*
