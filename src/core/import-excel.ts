/**
 * Excel 导入模块
 *
 * 两个确定性 tool：
 *   1. read_excel — 读 Excel，返回表头 + 样本数据
 *   2. import_with_mapping — 接受 mapping 方案 + 文件路径，执行导入
 *
 * LLM 推断由 Agent 侧完成，不在此模块内。
 */

import { readFileSync } from 'fs';
import { basename } from 'path';
import { createHash } from 'crypto';
import * as XLSX from 'xlsx';
import { pinyin } from 'pinyin-pro';
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

/** 映射方案（由 Agent 推断生成，传入 import_with_mapping） */
export interface MappingSchema {
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
  /** compiled truth 生成模板（用 {列名} 占位） */
  compiled_truth_template: string;
}

/** read_excel 的返回结果 */
export interface ExcelReadResult {
  file_name: string;
  sheets: SheetPreview[];
}

/** 单个 sheet 的预览 */
export interface SheetPreview {
  sheet_name: string;
  total_rows: number;
  headers: string[];
  sample_rows: Record<string, unknown>[];
}

/** 导入报告 */
export interface ImportReport {
  source_file: string;
  sheet_name: string;
  pages_created: number;
  pages_updated: number;
  pages_skipped: number;
  links_created: number;
  errors: string[];
}

// ============================================================
// read_excel：读 Excel，返回表头 + 样本数据
// ============================================================

/**
 * 读取 Excel 文件，返回每个 sheet 的表头和前 N 行样本。
 * 纯确定性操作，不涉及 LLM。
 */
export function readExcel(filePath: string, options?: { sheet?: string; sampleSize?: number }): ExcelReadResult {
  const buf = readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const sampleSize = options?.sampleSize ?? 5;

  const sheetsToProcess = options?.sheet
    ? [options.sheet]
    : wb.SheetNames;

  const sheets: SheetPreview[] = [];

  for (const name of sheetsToProcess) {
    const ws = wb.Sheets[name];
    if (!ws) continue;

    const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
      defval: '',
      raw: false,
    });

    if (rawRows.length === 0) continue;

    const headers = Object.keys(rawRows[0]).filter(h => h && h.trim());

    // 过滤空行
    const validRows = rawRows.filter(row =>
      headers.some(h => {
        const v = row[h];
        return v !== '' && v !== null && v !== undefined;
      })
    );

    if (validRows.length === 0) continue;

    sheets.push({
      sheet_name: name,
      total_rows: validRows.length,
      headers,
      sample_rows: validRows.slice(0, sampleSize),
    });
  }

  return {
    file_name: basename(filePath),
    sheets,
  };
}

// ============================================================
// 中文名 → slug
// ============================================================

const SLUG_SUFFIXES = [
  '有限责任公司', '有限公司', '股份有限公司', '股份公司',
  '集团有限公司', '集团公司', '集团',
  '科技有限公司', '科技公司',
  '技术有限公司', '技术公司',
  '信息技术有限公司', '信息科技有限公司',
  '网络科技有限公司', '网络技术有限公司',
  '智能科技有限公司', '智能技术有限公司',
  '数字科技有限公司',
];

export function chineseToSlug(name: string, prefix: string): string {
  let cleaned = name.trim();
  cleaned = cleaned.replace(/[（(][^）)]*[）)]/g, '');

  const sortedSuffixes = [...SLUG_SUFFIXES].sort((a, b) => b.length - a.length);
  for (const suffix of sortedSuffixes) {
    if (cleaned.endsWith(suffix)) {
      cleaned = cleaned.slice(0, -suffix.length);
      break;
    }
  }

  if (!cleaned) cleaned = name.trim();

  const py = pinyin(cleaned, { toneType: 'none', type: 'array' });
  const slug = py
    .map(s => s.toLowerCase().replace(/[^a-z0-9]/g, ''))
    .filter(s => s.length > 0)
    .join('-');

  if (!slug) {
    const hash = createHash('md5').update(name).digest('hex').slice(0, 8);
    return `${prefix}/${prefix.replace(/s$/, '')}-${hash}`;
  }

  return `${prefix}/${slug}`;
}

// ============================================================
// import_with_mapping：按映射方案执行导入
// ============================================================

/**
 * 根据 Agent 提供的 mapping 方案，将 Excel 数据导入 gbrain。
 * 纯确定性操作：给什么 mapping 就按什么执行。
 */
export async function importWithMapping(
  engine: BrainEngine,
  filePath: string,
  sheetName: string,
  mapping: MappingSchema,
): Promise<ImportReport> {
  const fileName = basename(filePath);

  // 读取完整数据
  const buf = readFileSync(filePath);
  const wb = XLSX.read(buf, { type: 'buffer', cellDates: true });
  const ws = wb.Sheets[sheetName];

  if (!ws) {
    return {
      source_file: fileName,
      sheet_name: sheetName,
      pages_created: 0, pages_updated: 0, pages_skipped: 0, links_created: 0,
      errors: [`Sheet [${sheetName}] 不存在`],
    };
  }

  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(ws, {
    defval: '',
    raw: false,
  });

  const headers = rawRows.length > 0 ? Object.keys(rawRows[0]).filter(h => h && h.trim()) : [];
  const rows = rawRows.filter(row =>
    headers.some(h => {
      const v = row[h];
      return v !== '' && v !== null && v !== undefined;
    })
  );

  const report: ImportReport = {
    source_file: fileName,
    sheet_name: sheetName,
    pages_created: 0,
    pages_updated: 0,
    pages_skipped: 0,
    links_created: 0,
    errors: [],
  };

  const { primary_entity, field_mappings, entity_extractions, compiled_truth_template } = mapping;

  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx];

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
      for (const colName of Object.keys(row)) {
        const value = String(row[colName] || '');
        compiledTruth = compiledTruth.replace(`{${colName}}`, value);
      }
      compiledTruth = compiledTruth.replace(/\{[^}]+\}/g, '').replace(/\s+/g, ' ').trim();

      // 5. 检查是否已存在
      const existing = await engine.getPage(slug);

      if (existing) {
        const mergedFm = { ...frontmatter };
        for (const [k, v] of Object.entries(existing.frontmatter)) {
          if (v !== '' && v !== null && v !== undefined) {
            mergedFm[k] = v;
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
      await engine.putRawData(slug, `excel:${fileName}:${sheetName}:row${rowIdx + 1}`, row);

      // 7. 处理关联实体
      for (const extraction of entity_extractions) {
        const relatedName = String(row[extraction.column] || '').trim();
        if (!relatedName) continue;

        // 内外部人员区分：
        //   managed_by → 百度内部人员（staff/），slug 只用人名，不拼企业名
        //   contacted_by → 外部联系人（people/），slug 拼人名+企业名消歧
        const isInternal = extraction.relation_to_primary === 'managed_by';
        let relatedSlug: string;
        let relatedType: string;

        if (isInternal) {
          // 百度内部人员：slug 前缀 staff/，不拼企业名（同一个人对接多家企业）
          relatedSlug = chineseToSlug(relatedName, 'staff');
          relatedType = 'person';
        } else if (extraction.relation_to_primary === 'contacted_by') {
          // 外部联系人：slug 拼上企业名消歧（不同企业的同名人是不同的人）
          const companySuffix = entityName
            ? '-' + chineseToSlug(entityName, '').replace(/^\//, '')
            : '';
          relatedSlug = chineseToSlug(relatedName, 'people') + companySuffix;
          relatedType = 'person';
        } else {
          // 其他关系类型：保持原有逻辑
          relatedSlug = chineseToSlug(relatedName, extraction.slug_prefix);
          relatedType = extraction.entity_type;
        }

        const relatedExists = await engine.getPage(relatedSlug);
        if (!relatedExists) {
          const relatedFm: Record<string, unknown> = {};
          for (const [colName, fieldName] of Object.entries(extraction.extra_fields)) {
            const v = row[colName];
            if (v !== '' && v !== null && v !== undefined) {
              relatedFm[fieldName] = v;
            }
          }

          // 内部人员标记
          if (isInternal) {
            relatedFm.is_internal = true;
          }
          // 外部联系人标记所属企业
          if (extraction.relation_to_primary === 'contacted_by') {
            relatedFm.company = entityName;
          }

          await engine.putPage(relatedSlug, {
            type: relatedType as any,
            title: relatedName,
            compiled_truth: isInternal
              ? `${relatedName}，百度内部对接人。`
              : `${relatedName}，${entityName}。`,
            frontmatter: relatedFm,
          });
        }

        try {
          await engine.addLink(slug, relatedSlug, '', extraction.relation_to_primary);
          report.links_created++;

          if (extraction.relation_to_primary === 'contacted_by') {
            await engine.addLink(relatedSlug, slug, '', 'works_at');
            report.links_created++;
          }
        } catch {
          // link 已存在
        }
      }
    } catch (e: any) {
      report.errors.push(`行 ${rowIdx + 1}: ${e.message}`);
    }
  }

  return report;
}
