const UNSAFE_KEYS = new Set(['__proto__', 'prototype', 'constructor']);

class DataValidations {
  constructor(model) {
    this.model = model || {};
  }

  add(address, validation) {
    if (UNSAFE_KEYS.has(address)) return undefined;
    return (this.model[address] = validation);
  }

  find(address) {
    if (UNSAFE_KEYS.has(address)) return undefined;
    return Object.prototype.hasOwnProperty.call(this.model, address)
      ? this.model[address]
      : undefined;
  }

  remove(address) {
    if (UNSAFE_KEYS.has(address)) return;
    this.model[address] = undefined;
  }
}

module.exports = DataValidations;
