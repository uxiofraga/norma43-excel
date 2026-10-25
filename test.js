// Pruebas mínimas: node test.js
const fs = require("fs");
const assert = require("assert");
const { n43Leer, n43LeerXls } = require("./norma43.js");
const leer = (f) => new TextDecoder("windows-1252").decode(fs.readFileSync(f));

const a = n43Leer(leer("ejemplos/dos_sociedades.n43"), "dos_sociedades.n43");
assert.ok(a.ok);
assert.strictEqual(a.cuentas.length, 2);
assert.deepStrictEqual(a.cuentas.map((c) => c.titular), ["ACME HOLDING SL", "BETA INMUEBLES SL"]);
assert.strictEqual(a.cuentas[0].saldoFin, 17300.75);
assert.ok(a.cuentas.every((c) => c.cuadra));
assert.strictEqual(a.cuentas[0].movs[1].concepto, "TRANSFERENCIA DE MARIA LOPEZ ALQUILER OCT");

const b = n43Leer(leer("ejemplos/sin_saltos_no_cuadra.n43"), "sin_saltos_no_cuadra.n43");
assert.ok(b.ok);
assert.strictEqual(b.cuentas[0].cuadra, false);

assert.strictEqual(n43Leer("esto no es un norma 43", "x").ok, false);

const x = n43LeerXls([
  ["", "Movimientos desde : 01/10/2026 hasta: 24/10/2026"], [],
  ["", "Número de cuenta", "Oficina", "Divisa", "F. Operación", "F. Valor", "Ingreso (+)", "Gasto (-)", "Saldo (+)", "Saldo (-)",
   "Concepto común", "Concepto propio", "Referencia 1", "Referencia 2", "Concepto complementario 1"],
  ["", "ES0021000000000000001234", "0001", "EUR", "20/10/2026", "20/10/2026", 500, "", 1500, "", "04", "000", "", "",
   "MARIA LOPEZ".padEnd(38) + "ALQUILER OCTUBRE"],
  ["", "ES0021000000000000001234", "0001", "EUR", "05/10/2026", "05/10/2026", "", 100, 1000, "", "03", "000", "", "", "RECIBO LUZ"],
], "caixa.xls");
assert.ok(x.ok);
assert.strictEqual(x.cuentas[0].saldoIni, 1100);
assert.strictEqual(x.cuentas[0].saldoFin, 1500);
assert.strictEqual(x.cuentas[0].cuadra, true);
assert.strictEqual(x.cuentas[0].movs[0].concepto, "MARIA LOPEZ · ALQUILER OCTUBRE");
console.log("OK: todas las pruebas pasan");
