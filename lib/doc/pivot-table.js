const {objectFromProps, range, toSortedArray} = require('../utils/utils');

const ALLOWED_METRICS = new Set([
  'sum',
  'count',
  'average',
  'max',
  'min',
  'product',
  'countNums',
  'stdDev',
  'stdDevP',
  'var',
  'varP',
]);

// TK(2023-10-10): turn this into a class constructor.

function makePivotTable(worksheet, model) {
  // Example `model`:
  // {
  //   // Source of data: the entire sheet range is taken,
  //   // akin to `worksheet1.getSheetValues()`.
  //   // Alternatively, `sourceTable: table1` can be passed.
  //   sourceSheet: worksheet1,
  //
  //   // Pivot table fields: values indicate field names;
  //   // they come from the first row in `worksheet1` or `sourceTable` columns.
  //   rows: ['A', 'B'],
  //   columns: ['C'],
  //   values: ['E'],
  //   metric: 'sum', // 'sum' | 'count' | 'average' | 'max' | 'min' | ...
  // }

  validate(worksheet, model);

  const sourceSheet = model.sourceSheet || model.sourceTable.worksheet;
  const sourceTable = model.sourceTable || null;
  let {rows, columns = [], values} = model;
  const metric = model.metric || 'sum';

  const cacheFields = makeCacheFields(sourceSheet, [...rows, ...columns]);

  const nameToIndex = cacheFields.reduce((result, cacheField, index) => {
    result[cacheField.name] = index;
    return result;
  }, {});
  rows = rows.map(row => nameToIndex[row]);
  columns = columns.map(column => nameToIndex[column]);
  values = values.map(value => nameToIndex[value]);

  const tableNumber = (worksheet.workbook.pivotTables || []).length + 1;

  // form pivot table object
  return {
    sourceSheet,
    sourceTable,
    rows,
    columns,
    values,
    metric,
    cacheFields,
    // defined in <pivotTableDefinition> of xl/pivotTables/pivotTableN.xml;
    // also used in xl/workbook.xml
    cacheId: String(9 + tableNumber),
    tableNumber,
  };
}

function validate(worksheet, model) {
  const sourceSheet = model.sourceSheet || (model.sourceTable && model.sourceTable.worksheet);
  if (!sourceSheet) {
    throw new Error('Must specify either sourceSheet or sourceTable for pivot table.');
  }

  if (model.metric && !ALLOWED_METRICS.has(model.metric)) {
    throw new Error(`Unsupported pivot table metric "${model.metric}".`);
  }

  if (!model.rows || !model.rows.length) {
    throw new Error('No pivot table rows specified.');
  }

  const columns = model.columns || [];
  if (!columns.length && (!model.values || model.values.length <= 1)) {
    throw new Error('No pivot table columns specified.');
  }

  if (!model.values || !model.values.length) {
    throw new Error('At least 1 value needs to be specified.');
  }

  const headerNames = model.sourceTable
    ? model.sourceTable.table.columns.map(col => col.name)
    : sourceSheet.getRow(1).values.slice(1);
  const isInHeaderNames = objectFromProps(headerNames, true);
  for (const name of [...model.rows, ...columns, ...model.values]) {
    if (!isInHeaderNames[name]) {
      throw new Error(`The header name "${name}" was not found in ${sourceSheet.name}.`);
    }
  }
}

function makeCacheFields(worksheet, fieldNamesWithSharedItems) {
  // Cache fields are used in pivot tables to reference source data.
  //
  // Example
  // -------
  // Turn
  //
  //  `worksheet` sheet values [
  //    ['A', 'B', 'C', 'D', 'E'],
  //    ['a1', 'b1', 'c1', 4, 5],
  //    ['a1', 'b2', 'c1', 4, 5],
  //    ['a2', 'b1', 'c2', 14, 24],
  //    ['a2', 'b2', 'c2', 24, 35],
  //    ['a3', 'b1', 'c3', 34, 45],
  //    ['a3', 'b2', 'c3', 44, 45]
  //  ];
  //  fieldNamesWithSharedItems = ['A', 'B', 'C'];
  //
  // into
  //
  //  [
  //    { name: 'A', sharedItems: ['a1', 'a2', 'a3'] },
  //    { name: 'B', sharedItems: ['b1', 'b2'] },
  //    { name: 'C', sharedItems: ['c1', 'c2', 'c3'] },
  //    { name: 'D', sharedItems: null },
  //    { name: 'E', sharedItems: null }
  //  ]

  const names = worksheet.getRow(1).values;
  const nameToHasSharedItems = objectFromProps(fieldNamesWithSharedItems, true);

  const aggregate = columnIndex => {
    const columnValues = worksheet
      .getColumn(columnIndex)
      .values.splice(2)
      .filter(v => v !== null && v !== undefined);
    const columnValuesAsSet = new Set(columnValues);
    return toSortedArray(columnValuesAsSet);
  };

  // make result
  const result = [];
  for (const columnIndex of range(1, names.length)) {
    const name = names[columnIndex];
    const sharedItems = nameToHasSharedItems[name] ? aggregate(columnIndex) : null;
    result.push({name, sharedItems});
  }
  return result;
}

module.exports = {makePivotTable};
