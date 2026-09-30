import test from 'node:test';
import assert from 'node:assert/strict';
import { calculateLoan, parseAmount, monthsBetween } from '../js/calc.js';

test('Test 1: 150,000 @5% x12', () => {
  const r = calculateLoan(150000, 5, 12);
  assert.equal(r.monthlyInterest, 7500);
  assert.equal(r.totalInterest, 90000);
  assert.equal(r.totalRepayment, 240000);
  assert.equal(r.monthlyRepayment, 20000);
});
test('Test 2: 315,375 @5% x12', () => {
  const r = calculateLoan(315375, 5, 12);
  assert.equal(r.monthlyInterest, 15768.75);
  assert.equal(r.totalInterest, 189225);
  assert.equal(r.totalRepayment, 504600);
  assert.equal(r.monthlyRepayment, 42050);
});
test('parseAmount', () => {
  assert.equal(parseAmount('₦1,500.50'), 1500.5);
  assert.ok(Number.isNaN(parseAmount('abc')));
  assert.equal(parseAmount('-5'), -5);
});
test('monthsBetween', () => {
  assert.equal(monthsBetween({ y: 2026, m: 10, d: 1 }, { y: 2027, m: 9, d: 30 }), 12);
});
