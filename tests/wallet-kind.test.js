const { test } = require('node:test');
const assert = require('node:assert/strict');
const { classifyWalletEntryKind } = require('../vinted_scrape.js');

test('refund entries are recognized before purchase entries', () => {
  assert.equal(classifyWalletEntryKind({ kind_label: 'Zwrot środków', title: 'Pantaloni' }), 'zwrot');
  assert.equal(classifyWalletEntryKind({ kind_label: 'Zakup', title: 'Pantaloni' }), 'zakup');
  assert.equal(classifyWalletEntryKind({ kind_label: 'Zakup', title: 'Ekspozycja' }), 'usluga');
  assert.equal(classifyWalletEntryKind({ kind_label: 'Sprzedane', title: 'Bluza' }), 'sprzedaz');
});
