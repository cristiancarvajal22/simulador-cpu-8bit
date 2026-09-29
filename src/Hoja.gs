/** @OnlyCurrentDoc */
// Interfaz nativa de Google Sheets: no depende de HtmlService.
const STATE_KEY = 'cpu8.native.v1', RUN_KEY = 'cpu8.run';
const REGIONES = {
  PC:'U20:W22', IR:'Q16:W18', OP:'Q20:S22', MAR:'V32:AD34',
  MDR:'AK32:AS34', AX:'AD21:AF23', BX:'AH21:AJ23', AC:'AD10:AJ12',
  REN1:'AP20:AS22', REN2:'AU20:AX22', FLAGS:'AD16:AJ18',
  DEC:'Q10:W12', SEQ:'I10:N15', CLOCK:'C10:F12', ALU:'AS12:AU14'
};
const C = {fondo:'#808080',uc:'#D8EDF3',alu:'#ECF1DC',mem:'#FFEAD4',valor:'#C65A49',activo:'#FFD166'};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Simulador CPU')
    .addItem('Elegir operación y datos','elegirOperacion').addItem('Cargar operación elegida','cargarOperacion')
    .addItem('Cargar demo: multiplicación','cargarDemo').addSeparator()
    .addItem('PASO · una microoperación','paso').addItem('EJECUTAR','ejecutar')
    .addItem('PAUSAR','pausar').addItem('RESET · conservar RAM','reiniciar')
    .addItem('CARGAR · bytes de Programa','cargarPrograma').addSeparator()
    .addItem('Leer memoria','leerMemoria').addItem('Escribir memoria','escribirMemoria')
    .addSeparator().addItem('Preparar diagrama','prepararDiagrama').addToUi();
}
function conBloqueo_(callback) {
  const lock=LockService.getDocumentLock();lock.waitLock(10000);
  try{return callback(PropertiesService.getDocumentProperties());}finally{lock.releaseLock();}
}
function leerPrograma_() {
  return SpreadsheetApp.getActive().getSheetByName('Programa').getRange('B7:B262').getDisplayValues().map((row,i)=>{
    try{return Cpu.parseHex(row[0]);}catch(error){throw Error('Programa!B'+(i+7)+': '+error.message);}
  });
}
function sesion_(props) {const raw=props.getProperty(STATE_KEY);return raw?JSON.parse(raw):Cpu.create(leerPrograma_());}
function paso() {
  conBloqueo_(props=>{
    if(props.getProperty(RUN_KEY))throw Error('Pulsa PAUSAR antes de avanzar manualmente.');
    const s=sesion_(props),e=Cpu.step(s);pintar_(s,e,false);props.setProperty(STATE_KEY,JSON.stringify(s));
  });
}
function ejecutar() {
  const token=Utilities.getUuid();
  conBloqueo_(props=>{if(props.getProperty(RUN_KEY))throw Error('Ya está en ejecución.');props.setProperty(RUN_KEY,token);});
  const start=Date.now();
  try {
    while(Date.now()-start<240000){
      const done=conBloqueo_(props=>{
        if(props.getProperty(RUN_KEY)!==token)return true;
        const s=sesion_(props),e=Cpu.step(s);pintar_(s,e,false);props.setProperty(STATE_KEY,JSON.stringify(s));return s.halted;
      });
      if(done)return;
      const delay=Number(SpreadsheetApp.getActive().getSheetByName('Diagrama').getRange('AQ4').getValue());
      Utilities.sleep(Math.max(100,Math.min(2000,Number.isFinite(delay)?delay:500)));
    }
    SpreadsheetApp.getActive().toast('Pausa automática tras 4 minutos. Pulsa EJECUTAR para continuar.');
  }finally{conBloqueo_(props=>{if(props.getProperty(RUN_KEY)===token)props.deleteProperty(RUN_KEY);});}
}
function pausar(){conBloqueo_(props=>props.deleteProperty(RUN_KEY));SpreadsheetApp.getActive().toast('Pausado al terminar el paso actual.');}
function reiniciar(){sustituir_(false);}
function cargarPrograma(){sustituir_(true);}
function sustituir_(load){conBloqueo_(props=>{
  const s=load?Cpu.create(leerPrograma_()):Cpu.reset(sesion_(props));
  props.deleteProperty(RUN_KEY);pintar_(s,null,true);props.setProperty(STATE_KEY,JSON.stringify(s));
});}
function leerMemoria(){conBloqueo_(props=>{
  const sh=SpreadsheetApp.getActive().getSheetByName('Memoria'),a=Cpu.parseHex(sh.getRange('C25').getDisplayValue()),v=Cpu.read(sesion_(props),a);
  sh.getRange('F25').setNumberFormat('@').setValue(Cpu.hex(v));sh.getRange('I25').setValue(v);
  sh.getRange('L25').setNumberFormat('@').setValue(v.toString(2).padStart(8,'0'));
});}
function escribirMemoria(){conBloqueo_(props=>{
  if(props.getProperty(RUN_KEY))throw Error('Pausa antes de editar la RAM.');
  const sh=SpreadsheetApp.getActive().getSheetByName('Memoria'),a=Cpu.parseHex(sh.getRange('C25').getDisplayValue()),v=Cpu.parseHex(sh.getRange('F25').getDisplayValue()),s=sesion_(props);
  Cpu.write(s,a,v);pintar_(s,{step:s.steps,phase:'EDICIÓN',from:'USUARIO',to:'RAM',value:v,address:a,label:'RAM['+Cpu.hex(a)+'h] ← '+Cpu.hex(v)+'h'},false);
  props.setProperty(STATE_KEY,JSON.stringify(s));
});}

function pintar_(s,e,clearLog){
  const book=SpreadsheetApp.getActive(),sh=book.getSheetByName('Diagrama');
  if(e) {
    sh.getRange('B30:R33').setValue(e.label);
    sh.getRange('B50:AX51').setValue('EN TRANSFERENCIA · '+e.from+' → '+e.to+'     '+Cpu.hex(e.value)+'h = '+e.value+' decimal = '+e.value.toString(2).padStart(8,'0'));
    [e.from,e.to].filter(k=>REGIONES[k]).forEach(k=>sh.getRange(REGIONES[k]).setBackground(C.activo).setFontColor('#17212B'));
  }
  if(e && e.before) animarTransferencia_(sh,e);
  Object.keys(s.r).forEach(k=>sh.getRange(REGIONES[k]).setValue(Cpu.hex(s.r[k])+'h').setBackground(C.valor).setFontColor('#FFFFFF'));
  ['DEC','SEQ','CLOCK'].forEach(k=>sh.getRange(REGIONES[k]).setBackground(C.uc).setFontColor('#17212B'));
  ['FLAGS','ALU'].forEach(k=>sh.getRange(REGIONES[k]).setBackground(C.alu).setFontColor('#17212B'));
  sh.getRange(REGIONES.FLAGS).setValue('ZF '+s.flags.ZF+'   CF '+s.flags.CF+'   SF '+s.flags.SF);
  sh.getRange(REGIONES.DEC).setValue(s.instruction.split(' ')[0]);sh.getRange(REGIONES.SEQ).setValue(s.steps+'\n'+s.phase);
  sh.getRange(REGIONES.CLOCK).setValue(s.halted?'■':s.steps%2?'●':'○');sh.getRange(REGIONES.ALU).setValue(Cpu.ISA[s.r.IR]?Cpu.ISA[s.r.IR].op:'');
  sh.getRange('B25:AX26').setValue('Entradas ALU: REN1 = '+s.r.REN1+'   REN2 = '+s.r.REN2+'       AC = '+s.r.AC+'   (valores decimales)');
  sh.getRange('V38:AB40').setValue(Cpu.hex(s.r.MAR)+'h');
  sh.getRange('B27:R28').setValue(s.error?'ERROR':s.halted?'DETENIDO · HLT':s.phase);
  sh.getRange('B30:R33').setValue(s.error||(e?e.label:'Pulsa PASO para comenzar.'));
  sh.getRange('B35:R37').setValue(s.instruction+(s.hasOperand?'\nOperando: '+Cpu.hex(s.r.OP)+'h':''));
  const list=s.queue.map((q,i)=>(i===s.index-1?'▶ ':'   ')+(i+1)+'. '+q.label);
  sh.getRange('B39:R48').setValue(list.slice(Math.max(0,s.index-4),s.index+2).join('\n'));
  sh.getRange('B50:AX51').setValue(e?'Paso '+e.step+': '+e.from+' → '+e.to+'     Valor: '+Cpu.hex(e.value)+'h / '+e.value+' / '+e.value.toString(2).padStart(8,'0'):'PASO realiza una microoperación. RESET conserva RAM; CARGAR restaura Programa.');
  [e&&e.from,e&&e.to].filter(k=>REGIONES[k]).forEach(k=>sh.getRange(REGIONES[k]).setBackground(C.activo).setFontColor('#17212B'));
  const start=Math.min(248,Math.max(0,s.r.MAR-3));
  for(let i=0;i<8;i++){
    sh.getRange(39+i,32,1,3).setValue(Cpu.hex(start+i));
    sh.getRange(39+i,35,1,9).setValue(Cpu.hex(s.ram[start+i])+'h    ('+s.ram[start+i]+')').setBackground(start+i===s.r.MAR?C.activo:C.mem);
  }
  const mem=book.getSheetByName('Memoria');
  mem.getRange('U6:X15').setValues(Object.entries(s.r).map(([k,v])=>[k,Cpu.hex(v),v,v.toString(2).padStart(8,'0')]));
  mem.getRange('C6:R21').setNumberFormat('@').setValues(Array.from({length:16},(_,r)=>s.ram.slice(r*16,r*16+16).map(Cpu.hex)));
  mem.getRange('C6:R21').setBackgrounds(Array.from({length:16},(_,r)=>Array.from({length:16},(_,c)=>r*16+c===s.r.MAR?C.activo:r<8?C.uc:C.alu)));
  const log=book.getSheetByName('Registro');
  if(clearLog&&log.getLastRow()>=5)log.getRange(5,1,log.getLastRow()-4,10).clearContent();
  if(e){const row=Math.max(5,log.getLastRow()+1);if(row>log.getMaxRows())log.insertRowsAfter(log.getMaxRows(),500);
    log.getRange(row,1,1,10).setValues([[e.step,e.phase,e.label,Cpu.hex(e.value)+'h',...['PC','IR','MAR','MDR','AX','BX'].map(k=>Cpu.hex(s.r[k])+'h')]]);
  }
  SpreadsheetApp.flush();
}

function mostrarALU(){SpreadsheetApp.getActive().toast('La ALU opera con REN1 y REN2; deposita el resultado en AC y actualiza ZF, CF y SF.');}
function mostrarBuses(){SpreadsheetApp.getActive().toast('Las flechas muestran la dirección de transferencia de datos y direcciones.');}
function mostrarTransferencia(){SpreadsheetApp.getActive().toast('El indicador recorre la conexión de la microoperación actual.');}

function animarTransferencia_(sh,e){
  const token=sh.getDrawings().find(d=>d.getOnAction()==='mostrarTransferencia');
  if(!token)return;
  function center(k){
    if(k==='RAM')return [770,(38+e.address-Math.min(248,Math.max(0,e.address-3))+0.5)*16];
    if(k==='ONE'||k==='ZERO')return [770,384];
    if(!REGIONES[k])return null;
    const r=sh.getRange(REGIONES[k]);
    return [(r.getColumn()-1+r.getNumColumns()/2)*20,(r.getRow()-1+r.getNumRows()/2)*16];
  }
  const a=center(e.from),b=center(e.to);
  if(!a||!b)return;
  // Los registros internos también tienen recorrido: no dependen de flechas decorativas.
  const path=[a,[(a[0]+b[0])/2,(a[1]+b[1])/2],b];
  for(const point of path){
    const p=[point[0]-7,point[1]-7];
    token.setPosition(Math.floor(p[1]/16)+1,Math.floor(p[0]/20)+1,Math.round(p[0]%20),Math.round(p[1]%16));
    SpreadsheetApp.flush();Utilities.sleep(140);
  }
}

function actualizarVista(){
  const sh=SpreadsheetApp.getActive().getSheetByName('Diagrama');
  // Únicamente se retira el dibujo decorativo de buses; se conservan ALU y controles.
  sh.getDrawings().filter(d=>d.getOnAction()==='mostrarBuses').forEach(d=>d.remove());
  sh.getRange('B4:AL5').clearContent();
  sh.getRange('B25:AX26').breakApart().merge().setBackground(C.fondo).setFontColor('#FFFFFF').setFontSize(12).setWrap(true);
  conBloqueo_(props=>{props.deleteProperty(RUN_KEY);pintar_(sesion_(props),null,false);});
}

function elegirOperacion(){SpreadsheetApp.getActive().getSheetByName('Operación').activate();}
function cargarOperacion(){
  conBloqueo_(props=>{
    const book=SpreadsheetApp.getActive(),sh=book.getSheetByName('Operación');
    const code=Cpu.parseHex(sh.getRange('C5').getDisplayValue().split(' · ')[0]);
    const values=sh.getRange('C7:C9').getValues().flat().map(v=>Cpu.byte(Number(v)));
    instalarPrograma_(props,Ejemplos.crear(code,...values));
    book.getSheetByName('Diagrama').activate();
  });
}
function cargarDemo(){conBloqueo_(props=>{
  instalarPrograma_(props,{ram:Cpu.demo(),notas:{128:'Multiplicando: 5',129:'Contador: 3',130:'Resultado'}});
  SpreadsheetApp.getActive().getSheetByName('Diagrama').activate();
});}
function instalarPrograma_(props,programa){
  const book=SpreadsheetApp.getActive();
  book.getSheetByName('Programa').getRange('B7:C262').setNumberFormat('@').setValues(programa.ram.map((v,i)=>[Cpu.hex(v),programa.notas[i]||'']));
  const s=Cpu.create(programa.ram);props.deleteProperty(RUN_KEY);pintar_(s,null,true);props.setProperty(STATE_KEY,JSON.stringify(s));
}
function prepararControles(){
  const book=SpreadsheetApp.getActive();book.setSpreadsheetTimeZone('America/La_Paz');
  const sh=book.getSheetByName('Operación')||book.insertSheet('Operación');
  sh.setHiddenGridlines(true).setColumnWidths(1,1,22).setColumnWidths(2,1,220).setColumnWidths(3,1,245).setColumnWidths(4,3,100);
  sh.getRange('A1:F23').setFontFamily('Arial').setFontSize(12).setBackground('#F3F6FA').setVerticalAlignment('middle');sh.setRowHeights(1,23,30);
  function title(r,t){sh.getRange(r).merge().setValue(t).setWrap(true);}
  title('B2:F3','ELIGE UNA INSTRUCCIÓN · DATOS DE 8 BITS');sh.getRange('B2:F3').setBackground('#16324F').setFontColor('#FFFFFF').setFontSize(18);
  sh.getRange('B5').setValue('Instrucción');sh.getRange('C5:F5').merge().setValue('11 · ADD AX,BX');
  sh.getRange('C5').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(Object.values(Cpu.ISA).map(d=>Cpu.hex(d.code)+' · '+d.mnemonic),true).setAllowInvalid(false).build());
  sh.getRange('B7:C9').setValues([['Dato A (RAM 80h → AX)',5],['Dato B (RAM 81h → BX)',3],['Inmediato (si dice imm)',3]]);
  sh.getRange('C7:C9').setNumberFormat('0').setBackground('#FFF0C2').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(0,255).setAllowInvalid(false).build());
  title('B11:F12','1. Elige una instrucción y escribe enteros de 0 a 255.\n2. Menú Simulador CPU → Cargar operación elegida.\n3. Usa PASO o EJECUTAR en el diagrama.');
  title('B14:F16','ADD suma; SUB resta; AND/OR/XOR operan bit a bit. CMP compara sin modificar AX/BX. INC, DEC y NOT necesitan un solo dato. REN1 y REN2 son las entradas internas de la ALU; AC recoge su resultado.');
  title('B18:F20','El programa lee primero A y B desde RAM. En las variantes imm, el segundo operando es el inmediato. Al terminar, AX queda también en RAM[82h] y BX en RAM[83h]. LOAD usa 80h; STORE escribe en 84h.');
  title('B22:F23','JMP/JZ/JNZ incluyen CMP AX,BX y un salto de prueba: si no se toma, AX recibe EEh. HLT se detiene sin almacenamiento final. Para un bucle: menú → Cargar demo: multiplicación.');
  const mem=book.getSheetByName('Memoria');mem.getRange('U5:X5').setValues([['Registro','HEX','DEC','BIN']]).setFontWeight('bold');mem.getRange('V6:V15').setNumberFormat('@');mem.getRange('X6:X15').setNumberFormat('@');mem.setColumnWidths(21,3,75).setColumnWidth(24,105);
  onOpen();conBloqueo_(props=>{props.deleteProperty(RUN_KEY);pintar_(sesion_(props),null,false);});
  sh.activate();
}

function instalarFiguraNueva(){
  const sh=SpreadsheetApp.getActive().getSheetByName('Diagrama');
  const specs=[['mostrarALU',10,41,160,128],['mostrarTransferencia',10,3,14,14],
    ['paso',4,2,105,30],['ejecutar',4,8,115,30],['pausar',4,15,105,30],['reiniciar',4,21,105,30],['cargarPrograma',4,27,115,30]];
  const drawings=sh.getDrawings(),pending=drawings.filter(d=>!d.getOnAction());
  if(pending.length!==1)throw Error('Debe existir exactamente una figura nueva sin función asignada.');
  const spec=specs.find(s=>!drawings.some(d=>d.getOnAction()===s[0]));
  if(!spec)throw Error('Todas las figuras están registradas.');
  pending[0].setOnAction(spec[0]).setPosition(spec[1],spec[2],0,spec[0]==='mostrarBuses'?8:0).setWidth(spec[3]).setHeight(spec[4]);
  SpreadsheetApp.flush();
}

function prepararDiagrama(){
  const book=SpreadsheetApp.getActive(),sh=book.getSheetByName('Diagrama');
  if(sh.getMaxColumns()<50)sh.insertColumnsAfter(sh.getMaxColumns(),50-sh.getMaxColumns());
  sh.getRange('A1:AX52').breakApart().clear();
  sh.setHiddenGridlines(true).setColumnWidths(1,50,20).setRowHeights(1,52,16);
  sh.getRange('A1:AX52').setBackground(C.fondo).setFontFamily('Arial').setFontSize(10).setFontColor('#17212B').setVerticalAlignment('middle');
  function text(range,value,size,color){sh.getRange(range).merge().setValue(value).setFontSize(size||10).setFontColor(color||'#17212B').setWrap(true);}
  function box(range,color){sh.getRange(range).setBackground(color).setBorder(true,true,true,true,false,false,'#111111',SpreadsheetApp.BorderStyle.SOLID_MEDIUM);}
  text('B1:AX2','SIMULADOR CPU · CICLO DE INSTRUCCIÓN',16,'#FFFFFF');
  text('B4:AL5','',11,'#FFFFFF');text('B25:AX26','Entradas ALU',12,'#FFFFFF');
  text('AN3:AX3','Pausa entre pasos (100–2000 ms)',9,'#FFFFFF');text('AQ4:AX5',500,11);
  sh.getRange('AQ4:AX5').setBackground('#FFFFFF').setDataValidation(SpreadsheetApp.newDataValidation().requireNumberBetween(100,2000).setAllowInvalid(false).build());
  box('B7:X24',C.uc);box('AB7:AX24',C.alu);box('T29:AT48',C.mem);
  text('B6:X6','Unidad de Control',11,'#FFFFFF');text('AB6:AX6','Unidad Aritmético Lógica',11,'#FFFFFF');text('T28:AT28','Memoria Central',11,'#FFFFFF');
  const labels={CLOCK:'Reloj',SEQ:'Secuenciador',DEC:'Decodificador',IR:'R.I. · opcode',OP:'Operando',PC:'C.P.',AC:'Acumulador',FLAGS:'R. Estado',AX:'AX',BX:'BX',REN1:'REN 1',REN2:'REN 2',MAR:'RDM / MAR',MDR:'RIM / MDR',ALU:'C. OP.'};
  Object.entries(REGIONES).forEach(([k,range])=>{
    const r=sh.getRange(range);r.merge().setHorizontalAlignment('center').setFontSize(['FLAGS','DEC','SEQ'].includes(k)?10:14).setWrap(true);
    if(k!=='ALU')box(range,['FLAGS','DEC','SEQ','CLOCK'].includes(k)?(r.getColumn()<25?C.uc:C.alu):C.valor);
    sh.getRange(r.getRow()-1,r.getColumn(),1,r.getNumColumns()).merge().setValue(labels[k]).setFontSize(10);
  });
  text('I17:N19','↓ ↓ ↓ ↓ ↓ ↓\nMicroórdenes',10);text('V37:AB37','Selector',10);text('V38:AB40','Dirección\nseleccionada',10);box('V38:AB40',C.mem);
  text('AF37:AQ38','Memoria · dirección / contenido',10);
  for(let i=0;i<8;i++){sh.getRange(39+i,32,1,3).merge();sh.getRange(39+i,35,1,9).merge();box('AF'+(39+i)+':AQ'+(39+i),C.mem);}
  ['B27:R28','B30:R33','B35:R37','B39:R48','B50:AX51'].forEach(r=>text(r,'',r==='B27:R28'?14:10,'#FFFFFF'));
  sh.getRange('B39:R48').setVerticalAlignment('top');
  const mem=book.getSheetByName('Memoria');
  mem.getRange('B24').setValue('Dirección HEX');mem.getRange('F24').setValue('Dato HEX');mem.getRange('I24').setValue('Decimal');mem.getRange('L24').setValue('Binario');
  mem.getRange('C25').setNumberFormat('@').setValue('80');mem.getRange('F25').setNumberFormat('@').setValue('05');
  mem.getRange('B27:R28').breakApart().merge().setValue('Menú Simulador CPU → Leer / Escribir memoria. Pausa antes de escribir.').setWrap(true);
  onOpen();conBloqueo_(props=>{props.deleteProperty(RUN_KEY);pintar_(sesion_(props),null,false);});sh.activate();
}
