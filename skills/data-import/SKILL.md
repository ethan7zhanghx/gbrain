---
name: data-import
version: 1.0.0
description: |
  Import Excel data into gbrain. Agent reads the Excel structure,
  infers a field mapping, then executes the import. No LLM calls
  inside the tools — the Agent itself handles all inference.
triggers:
  - "导入数据"
  - "导入 Excel"
  - "import excel"
  - "import data"
  - "把这个表导进去"
tools:
  - read_excel
  - import_with_mapping
  - search
  - get_page
mutating: true
writes_pages: true
writes_to:
  - companies/
  - people/
  - products/
  - activities/
---

# Data Import Skill

## 概述

将 Excel 文件中的数据导入 gbrain，创建实体页面和关系边。

## 工作流程

### Phase 1: 读取 Excel 结构

调用 `read_excel` tool，获取表头和样本数据。

```
read_excel({ file_path: "<Excel 文件路径>" })
```

返回结果包含每个 sheet 的：
- sheet_name: sheet 名称
- total_rows: 总行数
- headers: 所有列名
- sample_rows: 前 5 行样本数据

### Phase 2: 推断映射方案

根据 read_excel 返回的表头和样本数据，推断映射方案。需要判断：

1. **主实体类型和名称列**：这张表的每一行代表什么实体？哪一列是它的名称？
   - 如果每行是一家企业 → type: "company", slug_prefix: "companies"
   - 如果每行是一个人 → type: "person", slug_prefix: "people"
   - 如果每行是一个产品 → type: "product", slug_prefix: "products"

2. **字段映射**：每列映射为什么英文字段名？
   - 用简洁有意义的英文名，如 industry, status, region, phone
   - 跳过明显无用的列（纯序号、空列）

3. **关联实体抽取**：哪些列代表独立实体需要拆出来？
   - "对接人" → 拆出 person 实体，关系 contacted_by
   - "百度对接人" → 拆出 person 实体，关系 managed_by
   - "开发企业" → 拆出 company 实体，关系 developed_by
   - extra_fields 定义关联实体的附加属性（如手机号、职位属于对接人）

4. **compiled_truth_template**：用 {列名} 占位符构造一段摘要文字

关系类型参考：
- contacted_by: 对方联系人
- managed_by: 我方对接人
- works_at: 任职于
- source_from: 关系来源
- has_product: 拥有产品
- developed_by: 产品的开发企业
- attended: 参加活动

### Phase 3: 执行导入

调用 `import_with_mapping` tool，传入文件路径、sheet 名和映射方案。

```
import_with_mapping({
  file_path: "<Excel 文件路径>",
  sheet: "<sheet 名>",
  mapping: {
    primary_entity: { type: "company", slug_prefix: "companies", name_column: "企业全称" },
    field_mappings: { "所属行业": "industry", "伙伴状态": "status", ... },
    entity_extractions: [
      { column: "对接人", entity_type: "person", slug_prefix: "people",
        relation_to_primary: "contacted_by",
        extra_fields: { "手机号": "phone", "对接人职位": "position" } }
    ],
    compiled_truth_template: "{企业全称}，{所属行业}行业，当前状态{伙伴状态}。{企业简介}"
  }
})
```

### Phase 4: 报告结果

导入完成后，向用户报告：
- 创建了多少个实体
- 更新了多少个已有实体
- 建立了多少条关系
- 有哪些错误

如果 Excel 有多个 sheet，对每个有效 sheet 重复 Phase 2-4。

## 注意事项

- 同名实体不会重复创建，会合并 frontmatter（first-write-wins，已有字段不覆盖）
- 空字段不存储
- 原始行数据通过 put_raw_data 保留，可追溯来源文件、sheet 和行号
- 中文企业名会自动转为拼音 slug（去掉"有限公司"等常见后缀）
- 每个 sheet 需要独立推断 mapping，因为不同 sheet 的列可能不同
