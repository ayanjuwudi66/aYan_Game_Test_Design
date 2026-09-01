---
name: game-test-design
description: "为游戏功能、玩法和实时交互设计可执行测试方案：先确认被测对象与需求材料，再按生命周期七问推演主链路，补充 L1-L7 打断、影响面和边界确认，最终生成结构化用例。适用于功能/玩法测试设计，不替代性能、安全或兼容性专项方案。"
---

# Game Test Design

## 使用前提

不要猜测未给出的规则。先明确功能名称、被测类型、范围和测试目的；再索取 PRD/策划案/原型/接口/配置表，或让用户按结构化模板补充。用户材料必须原样保留，推导出的内容单列并标记为“待确认”。

## SOP

1. **明确被测对象**：确认功能、角色/权限、范围（含不覆盖项）和目的。
2. **读取需求材料**：从材料提取前置、输入、规则、状态流转、输出、异常和非功能要求。提供配置表时才检查类型、空值、引用、热更、关联与极值；未提供则声明配置风险未覆盖。
3. **边界与规则确认**：输出 Part A（材料原样提取）和 Part B（主动补充，全部待确认）。主动补充按七问、L1-L7、影响面、状态机、边界、并发时序、失败回滚七条来源推导，并标注级别、来源、问题、依据。

用户拍板后再生成正式用例。循环无天然终态时，读取 [references/baseline-table.md](references/baseline-table.md) 采用或确认 N/T 截断值。

## 核心推演

- 把功能视为状态机，沿时间轴覆盖“节点、边、异常注入、分支”，每个节点按七问重新检查。
- 主链路先输出骨架，再询问是否允许生成 Excel；获准后展开 L1-L7 与影响面扫描。
- L1/L2/L5 用例备注建议叠加弱网复测；重点检查状态残留、动画卡死、本地预表现与服务器回滚不一致。
- 影响面是与七问正交的横向维度：自身、队友、敌对玩家、中立单位、环境场景、观战者、服务器全局、AI/召唤物。
- 绝不把推断、概率、仓库内部逻辑或未定版策划细节写成确定预期。

## 交付物

1. [templates/主链路骨架推演图.md](templates/主链路骨架推演图.md)
2. [templates/边界值确认清单.md](templates/边界值确认清单.md)
3. 获准后使用 [templates/用例Excel骨架.md](templates/用例Excel骨架.md)；需要文件时运行 `scripts/gen_testcase_xlsx.py`。

按需读取：

- [references/seven-questions.md](references/seven-questions.md)：七问详解和品类实例。
- [references/interruption-L1-L7.md](references/interruption-L1-L7.md)：打断分层与 FPS/MOBA/RPG/SLG 实例库。
- [references/impact-scan.md](references/impact-scan.md)：8 类作用对象 × 4 问。
- [references/derivation-7.md](references/derivation-7.md)：主动补充的推导来源、三级分级和三要素。
- [references/baseline-table.md](references/baseline-table.md)：循环截断基线与截断后必查项。
- [references/s1-game-scenarios.md](references/s1-game-scenarios.md)：按功能性质动态判定 S1 的游戏场景。
