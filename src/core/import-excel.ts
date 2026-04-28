/**
 * Excel 导入模块
 *
 * 读取任意 Excel 文件，通过 LLM 推断字段映射方案，
 * 然后按映射逐行创建 gbrain page + link。
 *
 * 暴露为 import_excel operation（CLI + MCP 对等）。
 */

import { readFileSync, writeFileSync, existsSync, mkdirSync } from 'fs';
import { join, basename } from 'path';
import { createHash } from 'crypto';
import * as XLSX from 'xlsx';
import { pinyin } from 'pinyin-pro';
import Anthropic from '@anthropic-ai/sdk';
import type { BrainEngine } from './engine.ts';

// ============================================================
// 类型定义
// ============================================================

/** 单个实体抽取规则：从某一列拆出独立实体 */
export interface EntityExtraction {
  /** 来源列名 */
  column: string;
  /** 实体类型 */
  entity_type: string;
  /** slug 前缀 */
  slug_prefix: string;
  /** 与主实体的关系类型 */
  relation_to_primary: string;
  /** 附加字段：列名 → frontmatter 字段名 */
  extra_fields: Record<string, string>;
}

/** LLM 推断出的映射方案 */
export interface MappingSchema {
  /** 来源文件名 */
  source_file: string;
  /** sheet 名 */
  source_sheet: string;
  /** 表头 hash（用于匹配已有 mapping） */
  header_hash: string;
  /** 主实体配置 */
  primary_entity: {
    type: string;
    slug_prefix: string;
    name_column: string;
  };
  /** 列 → frontmatter 字段映射 */
  field_mappings: Record<string, string>;
  /** 需要拆出独立实体的列 */
  entity_extractions: EntityExtraction[];
  /** compiled truth 生成模板 */
  compiled_truth_template: string;
}

/** 导入报告 */
export interface ImportReport {
  source_file: string;
  sheets_processed: number;
  pages_created: number;
  pages_updated: number;
  pages_skipped: number;
  links_created: number;
  errors: string[];
}

// ============================================================
// Excel 读取
// ============================================================

export interface SheetData {
  sheet_name: string;
  headers: string[];
  rows: Record<string, unknown>[];
}

/**
 * 读取 Excel 文件的所有 sheet（或指定 sheet），返回结构化数据。
 * 自动跳过空行和纯统计行。
 */
export function readExcel(filePath: string, sheetName?: string): SheetData[] {
  const buf = readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });

  const sheetsToProcess = sheetName
    ? [sheetName]
    : wb.SheetNames;

  const results: SheetData[] = [];

  for (const name of sheetsToProcess) {
    const ws = wb.Sheets[name];
    if (!ws) continue;

    // 转为 JSON，header 行自动识别
    const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
      defval: '',
      raw: false,
    });

    if (rawRows.length === 0) continue;

    // 提取表头
    const headers = Object.keys(rawRows[0]).filter(h => h && h.trim());

    // 过滤空行（所有字段都为空的行）
    const rows = rawRows.filter(row =>
      headers.some(h => {
        const v = row[h];
        return v !== '' && v !== null && v !== undefined;
      })
    );

    if (rows.length === 0) continue;

    results.push({ sheet_name: name, headers, rows });
  }

  return results;
}

// ============================================================
// 中文名 → slug
// ============================================================

/** 中文企业/人名常见后缀，生成 slug 时去掉以缩短长度 */
const SLUG_SUFFIXES = [
  '有限责任公司', '有限公司', '股份有限公司', '股份公司',
  '集团有限公司', '集团公司', '集团',
  '科技有限公司', '科技公司', '科技',
  '技术有限公司', '技术公司',
  '信息技术有限公司', '信息科技有限公司',
  '网络科技有限公司', '网络技术有限公司',
  '智能科技有限公司', '智能技术有限公司',
  '数字科技有限公司',
];

/**
 * 将中文名转为 gbrain slug 格式。
 * 例如：厦门五卓未来科技有限公司 → xiamen-wuzhuo-weilai
 */
export function chineseToSlug(name: string, prefix: string): string {
  let cleaned = name.trim();

  // 去掉括号及内容
  cleaned = cleaned.replace(/[（(][^）)]*[）)]/g, '');

  // 去掉常见后缀（从长到短匹配）
  const sortedSuffixes = [...SLUG_SUFFIXES].sort((a, b) => b.length - a.length);
  for (const suffix of sortedSuffixes) {
    if (cleaned.endsWith(suffix)) {
      cleaned = cleaned.slice(0, -suffix.length);
      break;
    }
  }

  if (!cleaned) cleaned = name.trim();

  // 转拼音
  const py = pinyin(cleaned, { toneType: 'none', type: 'array' });
  const slug = py
    .map(s => s.toLowerCase().replace(/[^a-z0-9]/g, ''))
    .filter(s => s.length > 0)
    .join('-');

  if (!slug) {
    // fallback: hash
    const hash = createHash('md5').update(name).digest('hex').slice(0, 8);
    return `${prefix}/${prefix.slice(0, -1)}-${hash}`;
  }

  return `${prefix}/${slug}`;
}

// ============================================================
// 表头 hash（用于匹配已有 mapping）
// ============================================================

export function hashHeaders(headers: string[]): string {
  const normalized = headers
    .map(h => h.trim().toLowerCase())
    .sort()
    .join('|');
  return createHash('md5').update(normalized).digest('hex').slice(0, 12);
}

// ============================================================
// Mapping 持久化
// ============================================================

const MAPPINGS_DIR = join(process.cwd(), 'data', 'mappings');

export function saveMappingSchema(schema: MappingSchema): string {
  if (!existsSync(MAPPINGS_DIR)) {
    mkdirSync(MAPPINGS_DIR, { recursive: true });
  }
  const filename = `${schema.header_hash}.json`;
  const filepath = join(MAPPINGS_DIR, filename);
  writeFileSync(filepath, JSON.stringify(schema, null, 2), 'utf-8');
  return filepath;
}

export function loadMappingSchema(headerHash: string): MappingSchema | null {
  const filepath = join(MAPPINGS_DIR, `${headerHash}.json`);
  if (!existsSync(filepath)) return null;
  const content = readFileSync(filepath, 'utf-8');
  return JSON.parse(content) as MappingSchema;
}

// ============================================================
// LLM 推断映射方案
// ============================================================

let anthropicClient: Anthropic | null = null;

function getAnthropicClient(): Anthropic {
  if (!anthropicClient) {
    anthropicClient = new Anthropic();
  }
  return anthropicClient;
}

/**
 * 用 LLM 推断 Excel 的字段映射方案。
 * 输入表头 + 前几行样本，输出 MappingSchema。
 */
export async function inferMapping(
  fileName: string,
  sheetName: string,
  headers: string[],
  sampleRows: Record<string, unknown>[],
): Promise<MappingSchema> {
  const headerHash = hashHeaders(headers);

  // 先检查有没有已保存的 mapping
  const existing = loadMappingSchema(headerHash);
  if (existing) {
    return existing;
  }

  const sampleText = sampleRows.slice(0, 5).map((row, i) =>
    `行${i + 1}: ${JSON.stringify(row, null, 0)}`
  ).join('\n');

  const prompt = `你是一个数据分析专家。我有一个 Excel 表格，需要你分析它的结构并生成映射方案。

文件名: ${fileName}
Sheet 名: ${sheetName}
表头: ${JSON.stringify(headers)}

前 5 行样本数据:
${sampleText}

请分析这张表，输出一个 JSON 映射方案。要求：

1. 判断主实体是什么（通常是企业/公司），找出哪一列是名称
2. 将每列映射为英文字段名（用于存储），保持语义清晰
3. 如果某些列代表独立实体（如"对接人"代表一个人），指出需要拆分，并定义与主实体的关系类型
4. 给出 compiled_truth_template：一个用 {列名} 占位符的模板，用于生成实体的文字摘要
5. 跳过明显无用的列（如纯序号列）
6. 关系类型使用: contacted_by（对方联系人）, managed_by（我方对接人）, works_at（任职于）, source_from（关系来源）, has_product（有产品）, attended（参加活动）

输出严格的 JSON 格式如下：
{
  "primary_entity": {
    "type": "company 或 person 或 product 或 activity",
    "slug_prefix": "companies 或 people 或 products 或 activities",
    "name_column": "列名"
  },
  "field_mappings": {
    "列名": "英文字段名",
    ...
  },
  "entity_extractions": [
    {
      "column": "列名",
      "entity_type": "person",
      "slug_prefix": "people",
      "relation_to_primary": "contacted_by",
      "extra_fields": { "相关列名": "英文字段名" }
    }
  ],
  "compiled_truth_template": "{企业全称}是一家{所属行业}行业的企业..."
}

只输出 JSON，不要其他文字。`;

  const response = await getAnthropicClient().messages.create({
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 4096,
    messages: [{ role: 'user', content: prompt }],
  });

  // 提取 JSON
  const text = response.content
    .filter(b => b.type === 'text')
    .map(b => (b as { type: 'text'; text: string }).text)
    .join('');

  // 从回复中提取 JSON（可能被包裹在 ```json ``` 中）
  const jsonMatch = text.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    throw new Error('LLM 未返回有效的 JSON 映射方案');
  }

  const parsed = JSON.parse(jsonMatch[0]);

  const schema: MappingSchema = {
    source_file: fileName,
    source_sheet: sheetName,
    header_hash: headerHash,
    primary_entity: parsed.primary_entity,
    field_mappings: parsed.field_mappings || {},
    entity_extractions: parsed.entity_extractions || [],
    compiled_truth_template: parsed.compiled_truth_template || '',
  };

  // 保存 mapping
  saveMappingSchema(schema);

  return schema;
}

// ============================================================
// 执行导入
// ============================================================

/**
 * 根据 mapping 方案，将一个 sheet 的数据导入 gbrain。
 */
export async function executeImport(
  engine: BrainEngine,
  sheetData: SheetData,
  schema: MappingSchema,
  fileName: string,
  dryRun: boolean = false,
): Promise<ImportReport> {
  const report: ImportReport = {
    source_file: fileName,
    sheets_processed: 1,
    pages_created: 0,
    pages_updated: 0,
    pages_skipped: 0,
    links_created: 0,
    errors: [],
  };

  const { primary_entity, field_mappings, entity_extractions, compiled_truth_template } = schema;

  for (let rowIdx = 0; rowIdx < sheetData.rows.length; rowIdx++) {
    const row = sheetData.rows[rowIdx];

    try {
      // 1. 主实体名称
      const entityName = String(row[primary_entity.name_column] || '').trim();
      if (!entityName) {
        report.pages_skipped++;
        continue;
      }

      // 2. 生成 slug
      const slug = chineseToSlug(entityName, primary_entity.slug_prefix);

      // 3. 构建 frontmatter
      const frontmatter: Record<string, unknown> = {};
      for (const [colName, fieldName] of Object.entries(field_mappings)) {
        const value = row[colName];
        if (value !== '' && value !== null && value !== undefined) {
          frontmatter[fieldName] = value;
        }
      }

      // 4. 生成 compiled truth
      let compiledTruth = compiled_truth_template;
      for (const [colName] of Object.entries(row)) {
        const value = String(row[colName] || '');
        compiledTruth = compiledTruth.replace(`{${colName}}`, value);
      }
      // 清理未替换的占位符
      compiledTruth = compiledTruth.replace(/\{[^}]+\}/g, '').replace(/\s+/g, ' ').trim();

      if (dryRun) {
        console.log(`[dry-run] 将创建: ${slug} (${entityName})`);
        report.pages_created++;
        continue;
      }

      // 5. 检查是否已存在
      const existing = await engine.getPage(slug);

      if (existing) {
        // 合并 frontmatter: first-write-wins，只补空字段
        const mergedFm = { ...frontmatter };
        for (const [k, v] of Object.entries(existing.frontmatter)) {
          if (v !== '' && v !== null && v !== undefined) {
            mergedFm[k] = v; // 已有值保留
          }
        }
        await engine.putPage(slug, {
          type: primary_entity.type as any,
          title: entityName,
          compiled_truth: existing.compiled_truth || compiledTruth,
          frontmatter: mergedFm,
        });
        report.pages_updated++;
      } else {
        await engine.putPage(slug, {
          type: primary_entity.type as any,
          title: entityName,
          compiled_truth: compiledTruth,
          frontmatter,
        });
        report.pages_created++;
      }

      // 6. 存原始数据（溯源）
      await engine.putRawData(slug, `excel:${fileName}:${sheetData.sheet_name}:row${rowIdx + 1}`, row);

      // 7. 处理关联实体
      for (const extraction of entity_extractions) {
        const relatedName = String(row[extraction.column] || '').trim();
        if (!relatedName) continue;

        const relatedSlug = chineseToSlug(relatedName, extraction.slug_prefix);

        // 创建关联实体 page（如果不存在）
        const relatedExists = await engine.getPage(relatedSlug);
        if (!relatedExists) {
          const relatedFm: Record<string, unknown> = {};
          for (const [colName, fieldName] of Object.entries(extraction.extra_fields)) {
            const v = row[colName];
            if (v !== '' && v !== null && v !== undefined) {
              relatedFm[fieldName] = v;
            }
          }

          await engine.putPage(relatedSlug, {
            type: extraction.entity_type as any,
            title: relatedName,
            compiled_truth: `${relatedName}。`,
            frontmatter: relatedFm,
          });
        }

        // 建立关系边
        try {
          await engine.addLink(slug, relatedSlug, '', extraction.relation_to_primary);
          report.links_created++;

          // 如果是 contacted_by，同时建反向 works_at
          if (extraction.relation_to_primary === 'contacted_by') {
            await engine.addLink(relatedSlug, slug, '', 'works_at');
            report.links_created++;
          }
        } catch {
          // link 已存在，跳过
        }
      }
    } catch (e: any) {
      report.errors.push(`行 ${rowIdx + 1}: ${e.message}`);
    }
  }

  return report;
}

// ============================================================
// 完整导入流程（一个 Excel 文件）
// ============================================================

export async function importExcel(
  engine: BrainEngine,
  filePath: string,
  options: { sheet?: string; dryRun?: boolean } = {},
): Promise<ImportReport> {
  const fileName = basename(filePath);
  const sheets = readExcel(filePath, options.sheet);

  if (sheets.length === 0) {
    return {
      source_file: fileName,
      sheets_processed: 0,
      pages_created: 0,
      pages_updated: 0,
      pages_skipped: 0,
      links_created: 0,
      errors: ['没有找到有效数据的 sheet'],
    };
  }

  const totalReport: ImportReport = {
    source_file: fileName,
    sheets_processed: 0,
    pages_created: 0,
    pages_updated: 0,
    pages_skipped: 0,
    links_created: 0,
    errors: [],
  };

  for (const sheet of sheets) {
    console.log(`\n处理 sheet: ${sheet.sheet_name} (${sheet.rows.length} 行)`);

    try {
      // LLM 推断映射（或复用已有）
      const schema = await inferMapping(
        fileName,
        sheet.sheet_name,
        sheet.headers,
        sheet.rows.slice(0, 5),
      );

      console.log(`映射方案: 主实体=${schema.primary_entity.type}, 名称列=${schema.primary_entity.name_column}`);
      console.log(`字段映射: ${Object.keys(schema.field_mappings).length} 个字段`);
      console.log(`关联实体: ${schema.entity_extractions.length} 个`);

      // 执行导入
      const report = await executeImport(engine, sheet, schema, fileName, options.dryRun);

      totalReport.sheets_processed++;
      totalReport.pages_created += report.pages_created;
      totalReport.pages_updated += report.pages_updated;
      totalReport.pages_skipped += report.pages_skipped;
      totalReport.links_created += report.links_created;
      totalReport.errors.push(...report.errors);

      console.log(`完成: 创建 ${report.pages_created}, 更新 ${report.pages_updated}, 跳过 ${report.pages_skipped}, 关系 ${report.links_created}`);
      if (report.errors.length > 0) {
        console.log(`错误: ${report.errors.length} 条`);
      }
    } catch (e: any) {
      totalReport.errors.push(`Sheet [${sheet.sheet_name}] 失败: ${e.message}`);
      console.error(`Sheet [${sheet.sheet_name}] 失败: ${e.message}`);
    }
  }

  return totalReport;
}
