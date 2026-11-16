'use strict';

const UNSAFE_KEYS = new Set(['__proto__', 'constructor', 'prototype']);
const PLACEHOLDER_RE = /\{\{\s*([^{}]+?)\s*\}\}/g;
const EXACT_PLACEHOLDER_RE = /^\{\{\s*([^{}]+?)\s*\}\}$/;

function isSafeKey(key) {
  return Boolean(key) && !UNSAFE_KEYS.has(key);
}

function resolvePath(context, rawPath) {
  if (!context || typeof rawPath !== 'string') {
    return undefined;
  }
  const cleanPath = rawPath.trim().replace(/\[\]/g, '');
  if (!cleanPath) {
    return undefined;
  }
  const parts = cleanPath.split('.');
  let current = context;
  for (let i = 0; i < parts.length; i++) {
    const key = parts[i];
    if (!isSafeKey(key) || current === null || current === undefined) {
      return undefined;
    }
    if (
      typeof current === 'object' &&
      !Object.prototype.hasOwnProperty.call(current, key)
    ) {
      return undefined;
    }
    current = current[key];
  }
  return current;
}

function extractPlaceholdersFromCell(cell) {
  const tokens = [];
  if (!cell || cell.value === null || cell.value === undefined) {
    return tokens;
  }
  const val = cell.value;
  if (typeof val === 'string') {
    let match;
    PLACEHOLDER_RE.lastIndex = 0;
    while ((match = PLACEHOLDER_RE.exec(val)) !== null) {
      tokens.push(match[1].trim());
    }
  } else if (val && typeof val === 'object' && Array.isArray(val.richText)) {
    val.richText.forEach(part => {
      if (part && typeof part.text === 'string') {
        let match;
        PLACEHOLDER_RE.lastIndex = 0;
        while ((match = PLACEHOLDER_RE.exec(part.text)) !== null) {
          tokens.push(match[1].trim());
        }
      }
    });
  }
  return tokens;
}

function findLoopArrayKey(row, data) {
  let foundKey = null;
  row.eachCell({includeEmpty: false}, cell => {
    if (foundKey) {
      return;
    }
    const tokens = extractPlaceholdersFromCell(cell);
    for (let i = 0; i < tokens.length; i++) {
      const token = tokens[i].replace(/\[\]/g, '');
      const dotIdx = token.indexOf('.');
      if (dotIdx > 0) {
        const rootKey = token.slice(0, dotIdx);
        if (isSafeKey(rootKey)) {
          const candidate = resolvePath(data, rootKey);
          if (Array.isArray(candidate)) {
            foundKey = rootKey;
            return;
          }
        }
      }
    }
  });
  return foundKey;
}

function interpolateString(str, primaryCtx, fallbackCtx) {
  PLACEHOLDER_RE.lastIndex = 0;
  return str.replace(PLACEHOLDER_RE, (_, expr) => {
    let val = resolvePath(primaryCtx, expr);
    if (val === undefined && fallbackCtx) {
      val = resolvePath(fallbackCtx, expr);
    }
    if (val === undefined || val === null) {
      return '';
    }
    if (val instanceof Date) {
      return val.toISOString().slice(0, 10);
    }
    return String(val);
  });
}

function applyContextToCell(cell, primaryCtx, fallbackCtx) {
  if (!cell || cell.value === null || cell.value === undefined) {
    return;
  }
  const val = cell.value;
  if (typeof val === 'string') {
    const exact = EXACT_PLACEHOLDER_RE.exec(val.trim());
    if (exact) {
      let resolved = resolvePath(primaryCtx, exact[1]);
      if (resolved === undefined && fallbackCtx) {
        resolved = resolvePath(fallbackCtx, exact[1]);
      }
      if (resolved !== undefined) {
        cell.value = resolved;
        return;
      }
      cell.value = '';
      return;
    }
    if (PLACEHOLDER_RE.test(val)) {
      cell.value = interpolateString(val, primaryCtx, fallbackCtx);
    }
  } else if (val && typeof val === 'object' && Array.isArray(val.richText)) {
    let changed = false;
    const nextRuns = val.richText.map(part => {
      if (part && typeof part.text === 'string' && PLACEHOLDER_RE.test(part.text)) {
        changed = true;
        return Object.assign({}, part, {
          text: interpolateString(part.text, primaryCtx, fallbackCtx),
        });
      }
      return part;
    });
    if (changed) {
      cell.value = {richText: nextRuns};
    }
  }
}

function applyContextToRow(row, primaryCtx, fallbackCtx) {
  row.eachCell({includeEmpty: false}, cell => {
    applyContextToCell(cell, primaryCtx, fallbackCtx);
  });
}

function renderWorksheetTemplate(worksheet, data, options = {}) {
  const rootData = data && typeof data === 'object' ? data : {};
  const removeEmptyLoopRows =
    options.removeEmptyLoopRows !== undefined
      ? Boolean(options.removeEmptyLoopRows)
      : true;

  for (let r = 1; r <= worksheet.rowCount; r++) {
    const row = worksheet.findRow(r);
    if (!row) {
      continue;
    }

    const loopKey = findLoopArrayKey(row, rootData);
    if (loopKey) {
      const items = resolvePath(rootData, loopKey);
      if (items.length === 0) {
        if (removeEmptyLoopRows) {
          worksheet.spliceRows(r, 1);
          r--;
        } else {
          applyContextToRow(row, Object.create(null), rootData);
        }
        continue;
      }

      if (items.length > 1) {
        worksheet.duplicateRow(r, items.length - 1, true);
      }

      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const scoped = Object.create(null);
        const itemScope =
          item && typeof item === 'object' && !Array.isArray(item)
            ? Object.assign(Object.create(null), item, {$index: i, $number: i + 1})
            : item;
        scoped[loopKey] = itemScope;
        scoped.$index = i;
        scoped.$number = i + 1;
        if (item && typeof item === 'object' && !Array.isArray(item)) {
          Object.keys(item).forEach(k => {
            if (isSafeKey(k) && !Object.prototype.hasOwnProperty.call(scoped, k)) {
              scoped[k] = item[k];
            }
          });
        }
        const targetRow = worksheet.getRow(r + i);
        applyContextToRow(targetRow, scoped, rootData);
      }

      r += items.length - 1;
    } else {
      applyContextToRow(row, rootData, null);
    }
  }

  return worksheet;
}

module.exports = {
  renderWorksheetTemplate,
};
