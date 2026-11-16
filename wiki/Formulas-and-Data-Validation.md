# 9. Formulas & Data Validation

## Regular Formulas

```js
// Standard formula with optional cached result
ws.getCell('C5').value = { formula: 'SUM(C2:C4)', result: 150000 };
ws.getCell('C6').value = { formula: 'AVERAGE(C2:C4)', result: 50000 };
```

---

## Shared Formulas (`fillFormula`)

Efficiently populate a range with a shared formula:

```js
// Vertical shared formula: D2..D10 = B2*C2 .. B10*C10
ws.fillFormula('D2:D10', 'B2*C2', [10, 20, 30, 40, 50, 60, 70, 80, 90]);
```

---

## Data Validation (Dropdown Lists & Input Rules)

```js
// Dropdown list
ws.getCell('B2').dataValidation = {
  type: 'list',
  allowBlank: true,
  formulae: ['"Pending,Approved,Rejected"'],
  showErrorMessage: true,
  errorStyle: 'stop',
  errorTitle: 'Invalid Status',
  error: 'Please choose Pending, Approved, or Rejected.',
};

// Whole number range (1..100)
ws.getCell('C2').dataValidation = {
  type: 'whole',
  operator: 'between',
  allowBlank: false,
  formulae: [1, 100],
  showInputMessage: true,
  promptTitle: 'Quantity',
  prompt: 'Enter a number between 1 and 100',
};
```
