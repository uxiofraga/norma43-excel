# Norma 43 a Excel

Lector de extractos bancarios **Norma 43** (también llamado Cuaderno 43, AEB 43 o CSB 43) y del Excel **«C43» de CaixaBank Now**, escrito en JavaScript sin dependencias. Funciona en el navegador y en Node.

**Úsalo online, gratis y sin subir tus datos: https://getailift.com/norma43/**

Arrastras los ficheros de varios bancos, cuentas y sociedades y descargas un Excel con todos los movimientos, el resumen por cuenta y el saldo por titular. Todo se procesa en el navegador: los extractos no salen de tu ordenador.

## Qué hace

- Lee los registros 11 (cabecera de cuenta), 22 (movimiento), 23 (conceptos complementarios) y 33 (final de cuenta).
- Admite ficheros con saltos de línea y sin ellos (bloques de 80 caracteres), en codificación Windows-1252.
- Lee el Excel que CaixaBank entrega como «C43» desde CaixaBank Now, donde el ordenante y el concepto van pegados en la misma celda, y los separa.
- Comprueba que el saldo inicial más los movimientos da el saldo final que declara el banco, y marca la cuenta si no cuadra.
- Traduce el código de entidad al nombre del banco y el concepto común AEB a texto.

## Uso en Node

```js
const fs = require("fs");
const { n43Leer } = require("./norma43.js");

const texto = new TextDecoder("windows-1252").decode(fs.readFileSync("extracto.n43"));
const r = n43Leer(texto, "extracto.n43");
for (const c of r.cuentas) {
  console.log(c.banco, c.cuenta, c.titular, c.saldoIni, c.saldoFin, c.cuadra);
}
```

`n43LeerXls(filas, nombre)` recibe las filas del Excel de CaixaBank ya leídas (por ejemplo con SheetJS, `sheet_to_json(hoja, {header: 1, raw: true})`) y devuelve la misma estructura.

## Pruebas

```
node test.js
```

Los ficheros de `ejemplos/` son inventados.

## English

Parser for Spanish bank statements in **Norma 43** format (Cuaderno 43 / AEB 43 / CSB 43) and CaixaBank's "C43" Excel export. Plain JavaScript, no dependencies, runs in the browser and in Node. Free online converter to Excel that never uploads your files: https://getailift.com/norma43/

## Autor

Uxío Fraga, [AI Lift](https://getailift.com/): sistemas a medida para que grupos de empresas y family offices vean la caja y la deuda de todas sus sociedades sin rehacer Excel.

Licencia MIT.
