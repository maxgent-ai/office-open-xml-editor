import { expect, it } from 'vitest';
import { OoxmlResourceMetricsSession } from './resource-debug.js';
import { normalizeXlsxWorksheetPolicy } from './xlsx-worksheet-policy.js';
import { normalizeResourcePolicy } from './resource-policy.js';

it('reports a frozen content-free XLSX worksheet policy snapshot only for XLSX', () => {
  const limits = { maxCells: 300_000 };
  const worksheetPolicy = normalizeXlsxWorksheetPolicy({ xlsxWorksheetLimits: limits });
  const options = { enabled: true, mode: 'node' as const, policy: normalizeResourcePolicy({}),
    xlsxWorksheetPolicy: worksheetPolicy, emitToConsole: false };
  limits.maxCells = 999_999;
  const xlsx = new OoxmlResourceMetricsSession({ ...options, format: 'xlsx' }).succeed();
  expect(xlsx?.policy.xlsxWorksheetLimits).toEqual(worksheetPolicy.worksheet);
  expect(Object.isFrozen(xlsx?.policy.xlsxWorksheetLimits)).toBe(true);
  expect(Object.keys(xlsx?.policy.xlsxWorksheetLimits ?? {})).toEqual([
    'maxRows', 'maxCells', 'maxOwnedUtf8Bytes', 'maxJsonBytes',
  ]);
  const docx = new OoxmlResourceMetricsSession({ ...options, format: 'docx' }).succeed();
  expect(docx?.policy).not.toHaveProperty('xlsxWorksheetLimits');
});
