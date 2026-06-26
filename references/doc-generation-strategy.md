# 大文档生成策略

> 当 Step 5 构建的投标响应 JSON 超过单次工具调用传输限制时，采用分步构建 + 管道串联策略。

## 何时需要此策略

单次 Bash heredoc 实用上限约 50-80KB（取决于终端/平台）。以下信号表示需要分步构建：

- 响应总字数 > 18,000 中文字符（含 JSON 结构后 ≈ 60KB+）
- 章节数 > 5 章且每章 > 2,000 字
- Bash 工具调用报 `SyntaxError` 或静默无输出

## 分步构建模式

### A. 管道模式（3 步标准模板）

将完整 JSON 拆分为 3 个 Python 脚本，每步：stdin 读 JSON → 追加章节 → stdout 输出 JSON。最后一级管道到 `generate_docx.py`。

**文件分工**：

```
/tmp/build_step1.py    # 初始构建（含第1-2章）
/tmp/build_step2.py    # 追加第3-5章
/tmp/build_step3.py    # 追加第6-8章（末尾 print json）
```

**执行**：

```bash
python3 /tmp/build_step1.py | python3 /tmp/build_step2.py | python3 /tmp/build_step3.py | \
  python3 generate_docx.py output.docx
```

### B. 文件模式（推荐，更稳健）

将管道拆为两步：构建 JSON → 写文件，读文件 → 生成 docx。中间 JSON 可检查验证。

```bash
# 步骤1：构建 JSON 并写入文件
python3 /tmp/build_step1.py | python3 /tmp/build_step2.py | python3 /tmp/build_step3.py > /tmp/final.json

# 步骤2：从文件生成 docx
python3 generate_docx.py output.docx /tmp/final.json
```

## 每步脚本模板

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Step N: Read JSON from stdin, add chapters X-Y, output JSON"""
import json, sys

data = json.load(sys.stdin)

def ch(h, c): data["sections"].append({"heading": h, "level": 2, "content": c})
def p(t): return {"type": "p", "text": t}
def s(h, items): return {"heading": h, "level": 3, "content": items}

# ---- 追加章节 ----
ch("第X章标题", [
    p("开篇段落，按需引用招标条款..."),
    s("子章节标题", [
        p("段落内容..."),
    ]),
])

# ⚠️ 此行必须存在，否则 stdout 为空
print(json.dumps(data, ensure_ascii=False, indent=2))
```

## JSON 内容元素速查

| 元素类型 | JSON 结构 | 用途 |
|---------|----------|------|
| 段落 | `{"type": "p", "text": "..."}` | 正文叙述 |
| 子章节 | `{"heading": "...", "level": N, "content": [...]}` | 嵌套标题+内容 |
| 表格 | `{"type": "table", "headers": [...], "rows": [[...], ...]}` | 结构化数据 |

## 常见陷阱及对策

| 陷阱 | 症状 | 对策 |
|------|------|------|
| 中文弯引号被转换 | Python SyntaxError | 用 Write 工具写 .py 文件，不用 `cat > file << 'EOF'` |
| 管道 stderr 污染 | JSON 解析失败 "Expecting value" | 管道中禁用 `2>&1`，让 stderr 自然输出到终端 |
| 脚本缺 print 语句 | stdout 为空，"stdin 无数据" | 每步模板末尾固定加 `print(json.dumps(...))` |
| 编码声明缺失 | "Non-UTF-8 code... no encoding declared" | 每个 .py 首行加 `# -*- coding: utf-8 -*-` |
| 管道链条过长 | 中间步骤缓冲区溢出 | 超过 4 步时改用文件模式（写中间 JSON） |

## 可复用的辅助函数

所有分步脚本中保持一致的辅助函数定义，避免不同步骤间的函数签名不一致：

```python
def ch(h, c):   # 添加 H2 章节
    data["sections"].append({"heading": h, "level": 2, "content": c})

def p(t):       # 添加正文段落
    return {"type": "p", "text": t}

def s(h, items): # 添加 H3 子章节
    return {"heading": h, "level": 3, "content": items}

def tbl(h, r):   # 添加表格（可选）
    return {"type": "table", "headers": h, "rows": r}
```

## 步骤数估算

| 总字数（约） | JSON 体积（约） | 建议步数 | 每步章节数 |
|------------|---------------|---------|----------|
| < 8,000 字 | < 30 KB | 1 步 | 全部 |
| 8,000-16,000 字 | 30-60 KB | 2 步 | 各半 |
| 16,000-24,000 字 | 60-90 KB | 3 步 | 约 3 章/步 |
| > 24,000 字 | > 90 KB | 3-4 步 | 约 2-3 章/步 |

## 递归大纲：大章节的二次拆分

当某个章节的字数目标 > 3,000 字时，单一 `p()` 段落堆砌会导致结构扁平、缺乏层次感。此时应将该章节作为一个**微型项目**，递归应用"生成大纲 → 确认大纲 → 逐节编写"的完整流程。

### 触发条件

- 单章目标字数 > 3,000 字（如 [核心·标星] 类的 5,000+ 字章节）
- 章节内容包含多个独立主题（如"重难点"章包含 6 个独立难点）

### 递归流程

```
Step 4 编写某章时：
  1. 展示该章的子大纲（H3 标题列表 + 各节字数分配 + 覆盖条款）
  2. 等待用户确认子大纲
  3. 逐节编写（每节写完后用户确认，与 Step 4 主流程一致）
  4. 全节完成后汇总为该章的完整 content 数组
```

### 示例

第 3 章「项目重难点及应对策略」目标 5,500 字，在 Step 4 编写时的子大纲确认：

```markdown
## 📋 第3章子大纲

**重难点一：人员质量的系统性前置管控** [核心·标星] ~1000字
  覆盖：试用期考核🔴、人员质量、四层过滤选拔体系
**重难点二：多层考核体系下的持续达标管理** [核心·标星] ~1000字
  覆盖：供应商考核🔴、人员三级考核🔴、双轨达标
...

⚠️ 请确认子大纲，我将逐节编写。
```

### 递归大纲的 JSON 对应

子大纲的每个 H3 节对应一个 `s()` 调用，内含多个 `p()` 段落：

```python
s("重难点一：人员质量的系统性前置管控", [
    p("招标文件明确规定..."),
    p("我方建立了"四层过滤"人员选拔体系..."),
]),
```

### 与主流程的衔接

- Step 3 生成的大纲中，对于字数 > 3,000 的章节，标注「建议递归拆分」
- Step 4 编写到该章时，先展示子大纲，待用户确认后再逐节生成
- 子大纲确认的交互模式与 Step 3 完全相同，只是范围限定在单章内

- Step 5 中引用本策略：「如 JSON 超过单次传输限制，参考 `references/doc-generation-strategy.md` 的分步构建模式」
- 生成脚本 `generate_docx.py` 支持的文件模式是本策略的配套改进
