const _ = verquire('utils/under-dash');
const util = require('util');

describe('under-dash', () => {
  describe('isEqual', () => {
    const values = [
      0,
      1,
      true,
      false,
      'string',
      'foobar',
      'other string',
      [],
      ['array'],
      ['array', 'foobar'],
      ['array2'],
      ['array2', 'foobar'],
      {},
      {object: 1},
      {object: 2},
      {object: 1, foobar: 'quux'},
      {object: 2, foobar: 'quux'},
      null,
      undefined,
      () => {},
      () => {},
      Symbol('foo'),
      Symbol('foo'),
      Symbol('bar'),
    ];

    function showVal(o) {
      return util.inspect(o, {compact: true});
    }

    it('works on simple values', () => {
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values.length; j++) {
          const a = values[i];
          const b = values[j];

          const assertion = `${showVal(a)} ${i === j ? '==' : '!='} ${showVal(
            b
          )}`;
          expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
        }
      }
    });

    it('works on complex arrays', () => {
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values.length; j++) {
          const a = [values[i]];
          const b = [values[j]];

          const assertion = `${showVal(a)} ${i === j ? '==' : '!='} ${showVal(
            b
          )}`;
          expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
        }
      }
    });

    it('works on complex objects', () => {
      for (let i = 0; i < values.length; i++) {
        for (let j = 0; j < values.length; j++) {
          const a = {key: values[i]};
          const b = {key: values[j]};

          const assertion = `${showVal(a)} ${i === j ? '==' : '!='} ${showVal(
            b
          )}`;
          expect(_.isEqual(a, b)).to.equal(i === j, `expected ${assertion}`);
        }
      }
    });
  });

  describe('deepMerge prototype pollution', () => {
    afterEach(() => {
      delete Object.prototype.polluted;
    });

    it('ignores __proto__, constructor, and prototype keys in deepMerge', () => {
      const malicious = JSON.parse(
        '{"__proto__": {"polluted": "yes"}, "constructor": {"prototype": {"polluted": "yes"}}, "safe": 1}'
      );
      const target = _.deepMerge({}, malicious);
      expect({}.polluted).to.equal(undefined);
      expect(target.safe).to.equal(1);
    });

    it('ignores __proto__ in keyBy', () => {
      const items = [
        {id: '__proto__', val: 1},
        {id: 'a', val: 2},
      ];
      const keyed = _.keyBy(items, 'id');
      expect({}.val).to.equal(undefined);
      expect(keyed.a).to.deep.equal({id: 'a', val: 2});
    });
  });
});
