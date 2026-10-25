/* norma43.js: lector de extractos bancarios Norma 43 (Cuaderno 43, AEB 43, CSB 43) y del Excel «C43» de CaixaBank Now.
   Sin dependencias. Funciona en el navegador y en Node. Licencia MIT. Versión online: https://getailift.com/norma43/ */
var BANCOS={"0049":"Santander","0182":"BBVA","2100":"CaixaBank","0081":"Sabadell","0128":"Bankinter","2080":"Abanca","1465":"ING","0073":"Openbank","2085":"Ibercaja","2095":"Kutxabank","3058":"Cajamar","2038":"Bankia","0239":"EVO Banco","1491":"Triodos","0061":"Banca March","0186":"Mediolanum","0487":"Banco Mare Nostrum","0075":"Banco Popular","2048":"Liberbank","2103":"Unicaja","3035":"Caja Laboral","3025":"Caixa d'Enginyers","0019":"Deutsche Bank","0030":"Banco Español de Crédito","0238":"Banco Pastor","0216":"Targobank","0083":"Renta 4","0198":"Banco Cooperativo","0131":"Novo Banco","0232":"Banco Inversis","1544":"Andbank","2086":"Banco Grupo Cajatres","3183":"Arquia","0234":"Banco Caminos","6813":"Wise","1583":"Revolut","1568":"Qonto"};
var CONCEPTOS={"01":"Talones y reintegros","02":"Ingresos y entregas","03":"Recibos y domiciliaciones","04":"Transferencias y traspasos","05":"Préstamos y créditos","06":"Remesas de efectos","07":"Suscripciones y dividendos","08":"Cupones y amortizaciones","09":"Bolsa y valores","10":"Cheques gasolina","11":"Cajero automático","12":"Tarjetas","13":"Operaciones con el extranjero","14":"Devoluciones e impagados","15":"Nóminas y seguros sociales","16":"Timbres y corretajes","17":"Intereses, comisiones e impuestos","98":"Anulaciones y correcciones","99":"Varios"};
var DIVISAS={"978":"EUR","840":"USD","826":"GBP","756":"CHF","784":"AED"};
function n43Lineas(txt){
  txt=txt.replace(/\r/g,"");
  if(txt.indexOf("\n")<0&&txt.length>=160){var o=[];for(var i=0;i<txt.length;i+=80)o.push(txt.slice(i,i+80));return o}
  return txt.split("\n").filter(function(l){return l.trim().length});
}
function n43Fecha(s){var a=+s.slice(0,2),m=+s.slice(2,4),d=+s.slice(4,6);if(!m||!d||m>12||d>31)return null;return (2000+a)+"-"+("0"+m).slice(-2)+"-"+("0"+d).slice(-2)}
function n43Imp(clave,s,deudor){var v=parseInt(s,10);if(isNaN(v))return null;v=v/100;return clave===deudor?-v:v}
function n43Leer(txt,nombreFichero){
  var L=n43Lineas(txt),ctas=[],c=null,m=null,errores=[];
  if(!L.length||L[0].slice(0,2)!=="11")return {ok:false,error:"No parece un fichero Norma 43 (la primera línea debería empezar por 11)."};
  function cerrarMov(){if(m&&c){m.concepto=m.extra.join(" ").replace(/\s+/g," ").trim();delete m.extra;c.movs.push(m)}m=null}
  L.forEach(function(l,i){
    var t=l.slice(0,2);
    if(t==="11"){
      cerrarMov();
      c={fichero:nombreFichero,entidad:l.slice(2,6),banco:BANCOS[l.slice(2,6)]||("Entidad "+l.slice(2,6)),oficina:l.slice(6,10),cuenta:l.slice(10,20),
         desde:n43Fecha(l.slice(20,26)),hasta:n43Fecha(l.slice(26,32)),saldoIni:n43Imp(l.charAt(32),l.slice(33,47),"1"),
         divisa:DIVISAS[l.slice(47,50)]||l.slice(47,50),titular:l.slice(51,77).trim()||"(sin nombre)",movs:[],saldoFin:null,cuadra:null};
      ctas.push(c);
    }else if(t==="22"&&c){
      cerrarMov();
      var imp=n43Imp(l.charAt(27),l.slice(28,42),"1");
      if(imp===null){errores.push("Línea "+(i+1)+": importe ilegible");return}
      m={fecha:n43Fecha(l.slice(10,16)),valor:n43Fecha(l.slice(16,22)),comun:l.slice(22,24),propio:l.slice(24,27),
         tipo:CONCEPTOS[l.slice(22,24)]||"",importe:imp,documento:l.slice(42,52).trim(),ref1:l.slice(52,64).trim(),ref2:l.slice(64,80).trim(),extra:[]};
    }else if(t==="23"&&m){
      m.extra.push(l.slice(4,42).trim(),l.slice(42,80).trim());
    }else if(t==="33"&&c){
      cerrarMov();
      c.saldoFin=n43Imp(l.charAt(58),l.slice(59,73),"1");
      var calc=c.saldoIni+c.movs.reduce(function(s,x){return s+x.importe},0);
      c.cuadra=c.saldoFin===null?null:Math.abs(calc-c.saldoFin)<0.005;
      c.saldoCalc=Math.round(calc*100)/100;
    }
  });
  cerrarMov();
  ctas.forEach(function(c){if(c.saldoFin===null){c.saldoCalc=Math.round((c.saldoIni+c.movs.reduce(function(s,x){return s+x.importe},0))*100)/100;c.saldoFin=c.saldoCalc}});
  return {ok:ctas.length>0,cuentas:ctas,errores:errores,error:ctas.length?null:"No se encontró ninguna cuenta."};
}
/* Excel «C43» de CaixaBank Now: una fila por movimiento, del más reciente al más antiguo, con el saldo tras cada uno.
   El complementario 1 son los dos conceptos del registro 23 pegados (38 + 38): el primero suele ser quien paga. */
function n43XlsFecha(v){
  if(typeof v==="number"&&v>20000){var d=new Date(Math.round((v-25569)*864e5));return d.toISOString().slice(0,10)}
  var p=String(v).trim().split("/");if(p.length!==3)return null;
  return p[2]+"-"+("0"+p[1]).slice(-2)+"-"+("0"+p[0]).slice(-2);
}
function n43XlsNum(v){
  if(v===""||v===null||v===undefined)return null;if(typeof v==="number")return v;
  var n=parseFloat(String(v).replace(/\./g,"").replace(",","."));return isNaN(n)?null:n;
}
function n43LeerXls(filas,nombreFichero){
  var h=-1,col={};
  for(var i=0;i<Math.min(filas.length,30);i++){
    var t=filas[i].map(function(c){return String(c).trim().toLowerCase()});
    if(t.indexOf("concepto complementario 1")>=0&&t.some(function(c){return c.indexOf("mero de cuenta")>=0})){h=i;t.forEach(function(c,j){col[c]=j});break}
  }
  if(h<0)return {ok:false,error:"No es un Norma 43 ni el Excel de movimientos de CaixaBank."};
  var k=function(pref){for(var c in col)if(c.indexOf(pref)===0)return col[c];return -1};
  var cCta=k("n"),cOf=k("oficina"),cDiv=k("divisa"),cFo=k("f. operaci"),cFv=k("f. valor"),cIn=k("ingreso"),cGa=k("gasto"),cSm=k("saldo (+)"),cSn=k("saldo (-)"),
      cCom=k("concepto com"),cPro=k("concepto propio"),cR1=k("referencia 1"),cR2=k("referencia 2"),cC1=k("concepto complementario 1");
  for(var c in col)if(c.indexOf("mero de cuenta")>=0)cCta=col[c];
  var g={},orden=[];
  for(var r=h+1;r<filas.length;r++){
    var f=filas[r],cuenta=String(f[cCta]||"").replace(/\D/g,"");
    var fecha=n43XlsFecha(f[cFo]);if(!cuenta||!fecha)continue;
    var ing=n43XlsNum(f[cIn]),gas=n43XlsNum(f[cGa]);if(ing===null&&gas===null)continue;
    var imp=ing!==null?ing:-Math.abs(gas);
    var sm=n43XlsNum(f[cSm]),sn=n43XlsNum(f[cSn]),saldo=sm!==null?sm:(sn!==null?-Math.abs(sn):null);
    var comp=[];for(var j=cC1;j>=0&&j<cC1+10&&j<f.length;j++)if(String(f[j]).trim())comp.push(String(f[j]));
    var c1=comp[0]||"",nombre=c1.slice(0,38).trim(),resto=[c1.slice(38)].concat(comp.slice(1)).join(" ").replace(/\s+/g," ").trim();
    var com=("0"+String(f[cCom]||"").trim()).slice(-2);
    if(!g[cuenta]){g[cuenta]={fichero:nombreFichero,entidad:"2100",banco:"CaixaBank",oficina:String(f[cOf]||"").trim(),cuenta:cuenta.slice(-10),
      divisa:String(f[cDiv]||"EUR").trim()||"EUR",titular:"Cuenta ··"+cuenta.slice(-4),movs:[],saldos:[]};orden.push(cuenta)}
    g[cuenta].movs.push({fecha:fecha,valor:n43XlsFecha(f[cFv]),comun:com,propio:String(f[cPro]||"").trim(),tipo:CONCEPTOS[com]||"",importe:imp,
      documento:"",ref1:String(f[cR1]||"").trim(),ref2:String(f[cR2]||"").trim(),concepto:(nombre&&resto?nombre+" · "+resto:nombre||resto)});
    g[cuenta].saldos.push(saldo);
  }
  var ctas=orden.map(function(k){
    var c=g[k],n=c.movs.length;
    var fechas=c.movs.map(function(m){return m.fecha}).sort();c.desde=fechas[0];c.hasta=fechas[n-1];
    var recienteArriba=c.movs[0].fecha>=c.movs[n-1].fecha;
    var primero=recienteArriba?n-1:0,ultimo=recienteArriba?0:n-1;
    c.saldoFin=c.saldos[ultimo];
    c.saldoIni=c.saldos[primero]===null?null:Math.round((c.saldos[primero]-c.movs[primero].importe)*100)/100;
    var calc=c.saldoIni===null?null:c.saldoIni+c.movs.reduce(function(s,x){return s+x.importe},0);
    if(c.saldoFin===null||c.saldoIni===null){c.saldoIni=c.saldoIni||0;c.saldoCalc=Math.round(calc*100)/100||0;c.saldoFin=c.saldoFin===null?c.saldoCalc:c.saldoFin;c.cuadra=null}
    else{c.saldoCalc=Math.round(calc*100)/100;c.cuadra=Math.abs(calc-c.saldoFin)<0.005}
    delete c.saldos;return c;
  });
  return {ok:ctas.length>0,cuentas:ctas,errores:[],error:ctas.length?null:"El Excel no trae movimientos."};
}
if(typeof module!=="undefined")module.exports={n43Leer:n43Leer,n43LeerXls:n43LeerXls,BANCOS:BANCOS,CONCEPTOS:CONCEPTOS};
