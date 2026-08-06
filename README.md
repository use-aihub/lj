# reviewing-reports：专业报告原生批注审核技能

本项目提供一个面向 AI 代理的专业报告审核技能，用于审查 DOCX/XLSX 报告中的专业问题、计算错误、引用错误和跨文件矛盾，并交付带 Word/Excel 原生批注的审核副本。

当前内置输变电工程环评、电磁、噪声、生态和水土保持设施验收等专项规则；对于其他技术报告，则聚焦事实、计算、抄录、引用、证据及跨文件一致性。

## 核心能力

- 保全原件：写入前记录 SHA-256、原有批注和内容快照，所有操作均在临时副本上进行。
- 原生批注：DOCX 使用非空文字锚点的 Word 批注，XLSX 使用目标单元格批注/备注。
- 证据化审核：每个问题必须包含精确位置、原值或原文、原因、正确结果以及明确修改动作。
- 跨文件复核：对正文、专题、批复、设计、监测报告、原始记录、附图和附件进行一致性核查。
- 原子写入与验收：同一文件的批注一次性预检、批量写入和回读，任一验收项失败则不发布该项目。
- 分级结论：按“重大、较大、一般、文字”记录问题，并根据重大退回条件给出审核结论。
- 多档验收：提供 `fast`、`balanced`、`strict` 三种档位，默认使用 `fast`。

> 本技能按设计不渲染 DOCX/XLSX，也不生成 PDF、PNG 或页面截图进行视觉验收。PDF、图片和扫描件仅在支撑重大结论时读取必要内容，不生成对应批注版。

## 项目结构

```text
.
├── skills/
│   └── reviewing-reports/
│       ├── SKILL.md                         # 技能入口、工作流和交付规范
│       ├── AGENTS.md                        # 仓库协作约定
│       ├── agents/openai.yaml               # 技能名称、提示词和调用策略
│       ├── references/
│       │   ├── evidence-and-crosscheck.md   # 证据定位与跨文件复核规则
│       │   ├── return-criteria.md            # 严重程度与退回条件
│       │   ├── review-template.md            # 问题 JSON 与批注规范
│       │   ├── transmission-ai-review-checklist.md
│       │   └── water-conservation-acceptance-checklist.md
│       ├── scripts/review_artifact_guard.py # 准备、写入、验收与发布守护脚本
│       └── tests/test_review_artifact_guard.py
├── .gitignore
└── README.md
```

## 环境要求

- Python 3.9 或更高版本；守护脚本仅使用 Python 标准库。
- `officecli` 可执行文件已加入 `PATH`，用于读取和写入 DOCX/XLSX 原生批注；具体安装方式以所用运行环境的 OfficeCLI 文档为准。
- 支持 Skills 的 AI 代理环境，用于加载 `skills/reviewing-reports` 并执行专业审核流程。

可先检查本地命令是否可用：

```bash
python3 --version
officecli help
```

## 使用方式

将 `skills/reviewing-reports` 目录复制或软链接到所用代理的 Skills 目录，然后在任务中显式调用：

```text
使用 $reviewing-reports 审核我指定的 DOCX/XLSX 报告，保全原件并生成原生批注版。
```

技能默认执行以下流程：

1. 清点材料并锁定待审核的 DOCX/XLSX 文件。
2. 运行 `prepare`，记录原件哈希、原有批注及相应档位的内容快照。
3. 读取报告并按专项清单审查，生成结构化临时问题清单。
4. 运行 `write`，预检所有锚点并将问题原子批量写入临时副本。
5. 运行 `verify --publish`，回读批注并完成验收；全部通过后发布审核副本。

输出文件默认与原件位于同一目录，依次命名为：

```text
原文件名-审核批注版.docx
原文件名-审核批注版-v2.docx
原文件名-审核批注版-v3.docx
```

XLSX 文件采用相同的命名规则。

## 守护脚本

守护脚本负责文件安全和批注交付，不负责独立完成专业内容判断。问题清单格式见 [`review-template.md`](skills/reviewing-reports/references/review-template.md)。

```bash
python3 skills/reviewing-reports/scripts/review_artifact_guard.py prepare \
  --profile fast \
  --state /tmp/review-task/state.json \
  --input /path/to/report.docx \
  --input /path/to/data.xlsx

python3 skills/reviewing-reports/scripts/review_artifact_guard.py write \
  --state /tmp/review-task/state.json \
  --issues /tmp/review-task/issues.json

python3 skills/reviewing-reports/scripts/review_artifact_guard.py verify \
  --state /tmp/review-task/state.json \
  --issues /tmp/review-task/issues.json \
  --publish
```

若任务失败，可清理脚本创建的临时副本和状态：

```bash
python3 skills/reviewing-reports/scripts/review_artifact_guard.py cleanup \
  --state /tmp/review-task/state.json
```

`state.json` 和 `issues.json` 必须存放在本次任务的临时目录中，不应放入原件目录。脚本支持包含空格和中文的文件路径。

## 验收档位

| 档位 | 检查内容 | 适用场景 |
| --- | --- | --- |
| `fast` | 原件 SHA、原批注、正文快照、批注与锚点回读、ZIP 完整性 | 默认快速审核 |
| `balanced` | `fast` 的全部内容，加 outline 结构快照 | 需要额外结构校验 |
| `strict` | `balanced` 的全部内容，加 annotated 快照、OfficeCLI validate、真实页数及可用时的 `unzip -tq` | 高风险或明确要求全面验收 |

`write` 和 `verify` 会继承 `prepare` 选择的档位，流程中不能切换。所有档位的视觉检查均标记为 `not_performed_by_design`。

## 专项规则

- 输变电环评、电磁、噪声和生态报告使用 [`transmission-ai-review-checklist.md`](skills/reviewing-reports/references/transmission-ai-review-checklist.md)。
- 水土保持设施验收报告使用 [`water-conservation-acceptance-checklist.md`](skills/reviewing-reports/references/water-conservation-acceptance-checklist.md)。
- 多文件或涉及附件、附图和外部证据时使用 [`evidence-and-crosscheck.md`](skills/reviewing-reports/references/evidence-and-crosscheck.md)。
- 问题严重程度和退回判断以 [`return-criteria.md`](skills/reviewing-reports/references/return-criteria.md) 为准。

联网核验仅用于会改变结论的法规、标准版本、地理保护目标、资质或外部事实，并应优先采用政府、标准发布机构等权威一手来源。

## 运行测试

```bash
python3 -m unittest discover \
  -s skills/reviewing-reports/tests \
  -p 'test_*.py'
```

现有测试覆盖输出文件递增命名、路径空格、必填字段、禁用措辞、重复批注、面积/体积单位角标、DOCX 非空锚点，以及章节/表格汇总问题的边界校验。

## 安全约束

- 不直接修改原件的正文、表格、公式或格式。
- 不将原件或不可变基线传给可能重新封装文档的命令。
- 不用模糊措辞代替可定位、可执行的问题描述，也不为凑数而添加问题。
- 零问题文件仍发布零新增批注副本，以保证输出数与待审核文件数一致。
- 发布前再次确认原件未变化；任一文件验收失败时，不发布该项目。
