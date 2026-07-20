# 客户特殊规则 Agent：技术待支持与映射设计说明

日期：2026-07-08
适用范围：南亚营销转码 Agent 的客户特殊规则读取、映射设计、技术待支持规则处理。
关联草稿表：`客户特殊规则结构化草稿_按原表行_20260707.xlsx`

## 2026-07-13 技术状态补充

```text
原63条未出码中，18条不需要新业务口径的基础解析问题已修复。
余下45条不再按“无码”处理：候选码完整时显示编码，红色标注并保留人工确认原因。
当前业务原表无“订单备注”列，后台已将其作为可选字段预留；后续新增后无需改页面即可被语义模块识别。
模型的职责只是将口语/备注标准化为白名单JSON字段；标准代码映射、条件校验、生码和红绿显示由程序负责。
```

---

## 0. 当前实现状态补充（2026-07-09）

以下原“技术待支持”能力已经进入营销转码 Agent 后台运行时：

```text
完整尺寸映射
单边尺寸映射
尺寸加大算法
客户专用厚度映射
物料编码口径 631/632
无新美亚外部尺寸表读取
珠海乐健/赣州乐健 CCL 基板级别规则
```

当前 active Agent 规则版本：

```text
transcode_agent_rules_20260709_150049
机器规则=200
辅助映射启用=78
待接入规则Sheet=15
```

仍未进入当前运行时的重点项：

```text
赣州景旺 J0J0F0 后缀/扩展字段
订单语义模块：订单备注、含铜/不含铜、汽车板/电源板/能源板等
组合结构、铜箔厂商、玻布厂商、配方代码本轮仍不参与转码
```

---

## 1. 本文目标

本文用于说明两件事：

1. `技术待支持` 部分应该如何处理。
2. 客户特殊需求转成 Agent 可识别规则时，映射应该放在 Python 代码里，还是放在 Excel / 数据表里。

核心结论：

> Agent 当前执行“可执行草稿”规则，以及已经完成技术接入并启用的辅助映射规则；未启用的 `技术待支持` 仍必须保留下来，作为后续技术能力实现清单。
> 映射不要二选一，应该“业务变化放 Excel / 数据表，系统边界放 Python 枚举和校验”。

---

## 2. 技术待支持部分怎么处理

### 2.1 技术待支持的定义

`技术待支持` 不是业务待确认，也不是废弃规则。

它表示：

```text
业务规则已经清楚
触发条件也基本明确
但当前 Agent / 规则引擎还不具备执行能力
所以暂时不能自动转码
```

例如：

```text
完整尺寸映射：37*49 = 37.3*49.3
单边尺寸映射：1888 = 74.3
外部表查找：新美亚尺寸转换表
健鼎/超颖 mil 板厚换算
物料编码条件：631 = 芯厚，632 = 总厚
编码后缀位映射：J0J0F0 对应 * 后 24-30 位
客户专用厚度映射：003 = 3mil = 0.079
```

---

### 2.2 技术待支持不直接进入当前 Agent 执行表

当前 Agent 执行规则表只应读取：

```text
enabled = TRUE
status = approved
结构化处理状态 = 可执行草稿
```

技术待支持应排除在当前执行之外：

```text
enabled = FALSE
status = technical_pending
结构化处理状态 = 技术待支持
```

换句话说：

> 当前 Agent 不执行未启用的技术待支持规则；当某类技术能力已经开发、回归通过并在辅助映射表中启用后，才允许进入运行时。

---

### 2.3 技术待支持不能丢弃

技术待支持应该单独形成一张表或 Sheet：

```text
agent_technical_pending_rules
```

或：

```text
技术待支持清单
```

它的作用是：

```text
给开发排期
给后续规则引擎能力扩展
给回归测试准备样本
给业务说明“这个规则不是不做，而是技术上还没支持”
```

建议字段：

| 字段 | 说明 |
|---|---|
| `pending_id` | 技术待支持编号 |
| `customer_code` | 客户代码 |
| `customer_name` | 客户简称 |
| `source_row` | 来源行号 |
| `source_column` | 来源字段，例如 CCL特殊规则 |
| `source_text` | 原始规则文本 |
| `structured_text` | 当前结构化草稿文本 |
| `missing_capability` | 缺失的技术能力 |
| `suggested_rule_type` | 后续建议实现的规则类型 |
| `blocked_reason` | 当前不能自动执行的原因 |
| `priority` | 技术实现优先级 |
| `status` | `technical_pending` / `implemented` / `ignored` |
| `remark` | 备注 |

---

## 3. 技术待支持的分类处理

### 3.1 完整尺寸映射

示例：

```text
37*49 = 37.3*49.3
725.4*1244.6 = 28.56*49
```

含义：

```text
客户下单尺寸明确对应一个厂内尺寸
```

后续建议实现成：

```text
rule_type = size_mapping
```

示例结构：

```json
{
  "rule_type": "size_mapping",
  "condition_size": "37*49",
  "target_field": "size_code",
  "target_value": "37304930"
}
```

建议优先级：`P1`。

原因：规则明确、实现难度较低、业务价值高。

---

### 3.2 单边尺寸映射

示例：

```text
1888 = 74.3
2091 = 82.3
```

含义：

```text
客户订单里某一边尺寸要被替换为厂内尺寸，再组合生成完整 size_code
```

后续建议实现成：

```text
rule_type = single_side_size_mapping
```

执行逻辑：

```text
识别客户订单尺寸
    ↓
判断其中一边是否命中特殊映射
    ↓
替换单边尺寸
    ↓
重新组合完整尺寸
    ↓
生成 size_code
```

建议优先级：`P2`。

---

### 3.3 外部转换表 lookup

示例：

```text
新美亚尺寸转换表
健鼎/超颖板厚换算表
```

含义：

```text
规则不是简单的一条 A -> B，而是要查询外部规则源
```

后续建议实现成：

```text
rule_type = external_lookup
```

例如健鼎/超颖板厚换算：

```text
客户规格 mil 板厚 + 铜厚 -> 厂内规格 mm
mil < 31  -> 芯厚
mil >= 31 -> 总厚
```

注意：

```text
外部转换表应作为受控规则源进入项目
不能继续依赖个人桌面硬编码路径
```

建议优先级：健鼎/超颖板厚换算建议 `P1`，新美亚尺寸转换表建议 `P2`。

---

### 3.4 物料编码条件

示例：

```text
631 = 芯厚
632 = 总厚
```

含义：

```text
规则不是看客户规格文本，而是要读取客户物料编码或订单物料号中的特定位
```

后续建议实现成：

```text
rule_type = material_code_condition
```

前提条件：

```text
系统输入中必须稳定提供 customer_material_code / customer_product_code / order_material_no
```

建议优先级：`P3`。

---

### 3.5 编码后缀位映射

示例：

```text
J0J0F0 对应 * 后 24-30 位数
```

含义：

```text
规则直接影响内部编码后段位数
```

后续建议实现成：

```text
rule_type = code_suffix_mapping
```

建议优先级：`P4`。

原因：直接影响最终品号/品名后段，风险较高，需要更多样本和业务确认。

---

### 3.6 客户专用厚度映射

示例：

```text
003 = 3mil = 0.079
```

含义：

```text
客户自有厚度写法需要先转成厂内厚度，再参与总厚/芯厚判断和编码
```

后续建议实现成：

```text
rule_type = customer_thickness_mapping
```

执行逻辑：

```text
客户规格厚度文本
    ↓
命中客户专用厚度映射
    ↓
转换为 thickness_mm
    ↓
继续判断总厚/芯厚
    ↓
生成厚度相关代码
```

建议优先级：`P3`。

---

### 3.7 订单语义类规则

示例：

```text
订单备注字段作为铜箔规格来源
订单备注第 5 码识别汽车板
未标 TG 或 TG130 默认 NY2140
订单有双面字样 -> RTF，否则 HTE
备注中有 HW / 688 / 华为字样
```

当前处理口径：

```text
暂不进入客户特殊规则执行表
后续单独建立订单语义模块
```

建议后续表：

```text
order_semantic_rules
```

建议优先级：`P4`。

---

## 4. 技术待支持的运行时表现

如果订单命中了技术待支持规则，系统不应该直接报“转码失败”。

建议输出：

```json
{
  "result": "success_with_warning",
  "warning_type": "technical_pending_rule",
  "message": "该客户存在特殊规则，但当前 Agent 尚未支持该类规则，本次未自动执行。",
  "manual_review_required": true
}
```

执行口径：

```text
可执行规则命中 -> 自动执行
技术待支持命中 -> 提醒人工，不自动执行
无客户特殊规则命中 -> 按普通规则继续转码
```

---

## 5. 技术实现优先级建议

| 优先级 | 能力 | 原因 |
|---:|---|---|
| P1 | 完整尺寸映射 `size_mapping` | 规则明确，容易实现 |
| P1 | 健鼎/超颖板厚换算 `external_lookup` | 业务口径明确，客户重要 |
| 已完成 | 新美亚尺寸表 lookup | 2026-07-09 已接入，631按表、632按英寸*25.4 |
| P2 | 单边尺寸映射 | 技术稍复杂，但可控 |
| P3 | 客户专用厚度映射 | 需要更多样本验证 |
| P3 | 物料编码条件 | 依赖订单字段稳定性 |
| P4 | 编码后缀位映射 | 直接影响品号后段，风险较高 |
| P4 | 订单语义规则 | 后续单独模块 |

---

## 6. 映射部分放 Python 还是 Excel

### 6.1 总体建议

不要二选一。

建议：

> 业务会变的映射放 Excel / 数据表；系统不能乱变的枚举和校验逻辑放 Python。

也就是：

```text
Excel / 数据表 = 业务配置源
Python = 系统安全边界和校验器
```

---

## 7. 应该放 Excel / 数据表的内容

以下属于业务规则或主数据，后续可能会调整，应该放在 Excel 或数据库里。

### 7.1 客户特殊规则映射

例如：

```text
客户A：NY2150 -> 2B
客户B：R/R -> F/F
客户C：全部芯厚 -> C
客户D：TFT -> AT
```

建议放在：

```text
agent_customer_special_rules.xlsx
```

原因：客户会新增、业务会维护、规则会变化，不应该每次改代码。

---

### 7.2 胶系代码映射

例如：

```text
NY2150 -> 2B
NY3170HF -> 3C
NY3150HC -> RV
```

建议来源：

```text
胶系代码表
客户下单与胶系基板转换表
agent_customer_special_rules
```

处理原则：

```text
标准胶系映射 -> 胶系代码表
客户专属覆盖 -> agent_customer_special_rules
```

---

### 7.3 铜箔规格映射

例如：

```text
H/H -> HH
1/1 -> 11
R/R -> FF
H/H -> JJ
1-/1- -> KK
```

建议来源：

```text
编码规则表
客户特殊规则表
```

---

### 7.4 基板级别映射

例如：

```text
汽车板 -> AC
TFT -> AT
电源板 -> AP
能源板 -> AN
```

建议来源：

```text
编码规则表
agent_customer_special_rules
```

---

### 7.5 总厚/芯厚映射与转换规则

例如：

```text
芯厚 -> C
总厚 -> T
某客户全部芯厚
某客户默认总厚
<0.8 需要转换
>=0.8 需要转换
```

建议分工：

| 内容 | 建议放哪里 |
|---|---|
| `芯厚=C`、`总厚=T` | Python 枚举 + Excel 字典均可 |
| 某客户全部芯厚 | Excel / Agent 规则表 |
| 某客户厚度范围转换 | Excel / Agent 规则表 |
| 总芯厚换算值 | 总芯厚转换规则表 |

---

### 7.6 尺寸映射和外部表 lookup

例如：

```text
37*49 -> 37.3*49.3
1888 -> 74.3
健鼎/超颖 mil + 铜厚 -> 厂内 mm
新美亚尺寸转换表
```

建议放在：

```text
受控规则表
外部 lookup 表
agent_technical_pending_rules
```

不要写死到 Python 代码里。

---

## 8. 应该放 Python 枚举的内容

Python 中应该放系统框架级枚举和安全边界。

这些东西不能让业务随便在 Excel 中修改。

### 8.1 允许执行的 target_field 枚举

当前只允许这些字段参与 Agent 转码：

```python
ALLOWED_TARGET_FIELDS = {
    "glue_code",
    "thickness_rule",
    "thickness_code",
    "copper_code",
    "size_code",
    "glue_category_code",
    "copper_type_code",
    "grade_code",
    "tc_code",
}
```

对应业务字段：

```text
胶系
基板厚度
铜箔规格
基板尺寸
胶水类别
铜箔类型+印字/非印字
基板级别
总/芯厚
```

---

### 8.2 禁止执行的 target_field 枚举

```python
BLOCKED_TARGET_FIELDS = {
    "struct_code",
    "copper_vendor_code",
    "glass_vendor_code",
    "formula_code",
    "pp_rule",
    "pp_sheet_rule",
    "order_semantic_rule",
    "reference_note",
}
```

对应本轮不执行内容：

```text
组合结构
铜箔厂商
玻布厂商
配方代码
PP规则
PP小片规则
订单语义规则
非影响转码备注
```

即使 Excel 里误写了这些字段，Python 也必须拦截。

---

### 8.3 rule_type 枚举

```python
ALLOWED_RULE_TYPES = {
    "mapping",
    "default_set",
    "conditional_set",
    "range_rule",
    "size_mapping",
    "single_side_size_mapping",
    "external_lookup",
    "material_code_condition",
    "code_suffix_mapping",
    "customer_thickness_mapping",
}
```

当前第一阶段可执行的规则类型建议只启用：

```python
EXECUTABLE_RULE_TYPES = {
    "mapping",
    "default_set",
    "conditional_set",
    "range_rule",
}
```

暂不执行的技术待支持类型：

```python
PENDING_RULE_TYPES = {
    "size_mapping",
    "single_side_size_mapping",
    "external_lookup",
    "material_code_condition",
    "code_suffix_mapping",
    "customer_thickness_mapping",
}
```

---

### 8.4 状态枚举

```python
RULE_STATUS = {
    "approved",
    "technical_pending",
    "disabled",
    "reference_only",
}
```

Agent 只执行：

```python
enabled is True and status == "approved"
```

---

### 8.5 优先级枚举

当前确认的优先级：

```text
客户订单明示 > CCL特殊规则 > 通用特殊规则 > 普通规则 > 默认规则
```

建议 Python 中固定：

```python
PRIORITY = {
    "order_explicit": 0,
    "ccl_special": 10,
    "general_special": 20,
    "normal_rule": 50,
    "default_rule": 90,
}
```

Excel 可以存 `priority` 值，但 Python 要校验是否在允许范围内。

---

## 9. 推荐的表结构

### 9.1 Agent 执行规则表

表名：

```text
agent_customer_special_rules
```

示例字段：

| 字段 | 说明 |
|---|---|
| `rule_id` | 规则编号 |
| `enabled` | 是否启用 |
| `status` | 状态，只执行 `approved` |
| `priority` | 优先级 |
| `version` | 版本 |
| `customer_code` | 客户代码 |
| `customer_name` | 客户简称 |
| `material_scope` | `ALL` / `CCL` |
| `source_row` | 来源行号 |
| `source_column` | 来源字段 |
| `source_text` | 来源原文 |
| `business_feedback` | 业务反馈 |
| `rule_type` | 规则类型 |
| `target_field` | 目标字段 |
| `target_value` | 目标值 |
| `condition_glue` | 胶系条件 |
| `condition_thickness_op` | 厚度比较符 |
| `condition_thickness_min` | 厚度下限 |
| `condition_thickness_max` | 厚度上限 |
| `condition_copper` | 铜箔条件 |
| `condition_size` | 尺寸条件 |
| `condition_text` | 条件说明 |
| `rule_explanation` | 规则解释 |
| `validation_status` | 校验状态 |
| `confidence` | 置信度 |
| `remark` | 备注 |

---

### 9.2 字段映射字典表

表名：

```text
agent_rule_field_mapping
```

示例：

| business_field | target_field | enabled | current_scope | remark |
|---|---|---|---|---|
| 胶系 | glue_code | TRUE | CCL | 当前启用 |
| 基板厚度 | thickness_rule | TRUE | CCL | 当前启用 |
| 铜箔规格 | copper_code | TRUE | CCL | 当前启用 |
| 基板尺寸 | size_code | TRUE | CCL | 当前启用 |
| 胶水类别 | glue_category_code | TRUE | CCL | 当前启用 |
| 铜箔类型+印字/非印字 | copper_type_code | TRUE | CCL | 当前启用 |
| 基板级别 | grade_code | TRUE | CCL | 当前启用 |
| 总/芯厚 | tc_code | TRUE | CCL | 当前启用 |
| 组合结构 | struct_code | FALSE | future | 本轮不启用 |
| 铜箔厂商 | copper_vendor_code | FALSE | future | 本轮不启用 |
| 玻布厂商 | glass_vendor_code | FALSE | future | 本轮不启用 |
| 配方代码 | formula_code | FALSE | future | 本轮不启用 |

说明：这张表可以给人看，也可以给转换脚本参考。但 Python 仍然必须有一份白名单校验。

---

### 9.3 规则值字典表

表名：

```text
agent_rule_value_dictionary
```

示例：

| target_field | business_value | system_value | enabled | source |
|---|---|---|---|---|
| tc_code | 芯厚 | C | TRUE | 固定枚举 |
| tc_code | 总厚 | T | TRUE | 固定枚举 |
| grade_code | 汽车板 | AC | TRUE | 编码规则 |
| grade_code | TFT | AT | TRUE | 编码规则 |
| copper_code | H/H | HH | TRUE | 编码规则 |
| copper_code | 1/1 | 11 | TRUE | 编码规则 |

说明：如果项目已有胶系代码、编码规则、总芯厚转换、客户下单与胶系基板转换等表，可以先复用已有表，不一定重复建完整字典。

---

## 10. 推荐读取和校验流程

```text
读取 agent_customer_special_rules
    ↓
只保留 enabled=true
    ↓
只保留 status=approved
    ↓
校验 material_scope in {ALL, CCL}
    ↓
校验 rule_type 是否可执行
    ↓
校验 target_field 是否在 ALLOWED_TARGET_FIELDS
    ↓
拦截 BLOCKED_TARGET_FIELDS
    ↓
校验 target_value 是否在对应字典中
    ↓
按 customer_code 建索引
    ↓
按 material_scope 过滤
    ↓
按 priority 排序
    ↓
判断 condition 是否命中
    ↓
执行 target_field = target_value
    ↓
记录命中来源和解释
```

伪代码：

```python
ALLOWED_TARGET_FIELDS = {
    "glue_code",
    "thickness_rule",
    "thickness_code",
    "copper_code",
    "size_code",
    "glue_category_code",
    "copper_type_code",
    "grade_code",
    "tc_code",
}

BLOCKED_TARGET_FIELDS = {
    "struct_code",
    "copper_vendor_code",
    "glass_vendor_code",
    "formula_code",
    "pp_rule",
    "pp_sheet_rule",
    "order_semantic_rule",
    "reference_note",
}

EXECUTABLE_RULE_TYPES = {
    "mapping",
    "default_set",
    "conditional_set",
    "range_rule",
}

def validate_rule(rule):
    if not rule["enabled"]:
        return False

    if rule["status"] != "approved":
        return False

    if rule["material_scope"] not in {"ALL", "CCL"}:
        return False

    if rule["rule_type"] not in EXECUTABLE_RULE_TYPES:
        return False

    if rule["target_field"] not in ALLOWED_TARGET_FIELDS:
        return False

    if rule["target_field"] in BLOCKED_TARGET_FIELDS:
        return False

    return True
```

---

## 11. 为什么不建议全部写在 Python

如果把客户特殊规则都写在 Python 里，例如：

```python
if customer == "方正F7" and glue == "NY2150":
    glue_code = "2B"
```

会有这些问题：

```text
每次业务改规则都要改代码
业务人员无法复核
无法版本化维护
无法导出确认
无法保留规则来源
客户规则会越来越难维护
```

因此：

> 客户特殊规则、胶系/铜箔/等级/尺寸等业务映射不应硬编码在 Python 里。

---

## 12. 为什么也不建议全部放 Excel

如果全部放 Excel，一旦有人误填：

```text
target_field = struct_code
status = approved
```

Agent 可能会执行本轮不该执行的组合结构规则。

因此 Python 必须控制：

```text
只读 enabled=true
只读 status=approved
只允许 CCL / ALL
只允许当前 8 个 target_field
禁止 struct_code / vendor / formula / PP / 订单语义
校验 target_value 是否在字典中
校验 priority 是否符合规则
```

Excel 是配置源，但 Python 是安全门。

---

## 13. 最终建议

映射设计采用三层结构最稳：

### 第一层：Python 固定枚举

控制系统边界：

```text
哪些字段能执行
哪些字段不能执行
哪些 rule_type 能执行
哪些状态能执行
优先级顺序
```

### 第二层：Excel / 数据表维护业务映射

维护业务变化：

```text
客户特殊规则
胶系代码
铜箔规格
基板级别
总芯厚转换
尺寸映射
客户专属覆盖
```

### 第三层：Python 校验 Excel

保证配置不乱：

```text
字段白名单校验
值字典校验
状态校验
优先级校验
冲突检测
来源解释保留
```

最终口径：

> 字段枚举、规则类型、状态、优先级、安全边界放 Python；客户特殊规则、胶系/铜箔/等级/尺寸/总芯厚等业务映射放 Excel 或数据库；Python 读取 Excel 后必须做白名单和字典校验。

## 14. 模型语义规则的运行映射

P2-2已将审批后的39条TSR规则接入影子解释器。该映射不写在客户`if/else`代码中：

```text
客户、条件、标准语义和证据：transcode_semantic_rules.xlsx/json
操作符、输入字段适配和安全边界：transcode_semantic_shadow.py
版本、哈希、active和回滚：transcode_semantic_rules.py
订单批量链路与证据Sheet：transcode_agent_service.py
```

这一层只输出语义判断及证据，不能直接写入制造码或修改评分。当规则依赖订单备注但订单没有该列时，必须输出“缺少输入”，不得用其他字段猜测。
