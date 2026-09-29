/** Motor de 8 bits. Independiente de la hoja y de la animacion. */
var Cpu = (() => {
  const ISA = {};
  function def(code, op, dst, src, size) { ISA[code] = {code,op,dst,src,size, mnemonic: op==='HLT'?'HLT':op+' '+(dst||'')+(src?','+src:'')}; }
  def(1,'MOV','AX','imm',2); def(2,'MOV','BX','imm',2);
  def(3,'MOV','AX','BX',1); def(4,'MOV','BX','AX',1);
  def(5,'LOAD','AX','[dir]',2); def(6,'LOAD','BX','[dir]',2);
  def(7,'STORE','[dir]','AX',2); def(8,'STORE','[dir]','BX',2);
  [['ADD',16],['SUB',18],['CMP',24],['AND',32],['OR',34],['XOR',36]].forEach(([op,n])=>{def(n,op,'AX','imm',2);def(n+1,op,'AX','BX',1);});
  def(20,'INC','AX','',1);def(21,'INC','BX','',1);def(22,'DEC','AX','',1);def(23,'DEC','BX','',1);def(38,'NOT','AX','',1);
  def(48,'JMP','dir','',2);def(49,'JZ','dir','',2);def(50,'JNZ','dir','',2);def(255,'HLT','','',1);
  [['ADD',64],['SUB',66],['CMP',68],['AND',70],['OR',72],['XOR',74]].forEach(([op,n])=>{def(n,op,'BX','imm',2);def(n+1,op,'BX','AX',1);});
  def(76,'NOT','BX','',1);
  const hex = n => n.toString(16).toUpperCase().padStart(2,'0');
  function byte(n) {if(!Number.isInteger(n)||n<0||n>255)throw Error('Se requiere un byte entre 0 y 255.');return n;}
  function read(s,a){return s.ram[byte(a)];}
  function write(s,a,v){s.ram[byte(a)]=byte(v);}
  function parseHex(v){const t=String(v).trim();if(!/^[\da-f]{1,2}$/i.test(t))throw Error('Byte hexadecimal inválido: '+t);return parseInt(t,16);}
  function demo(){const ram=Array(256).fill(0);[5,129,24,0,49,25,6,128,5,130,17,7,130,5,129,22,7,129,24,0,50,8,48,25,0,255].forEach((v,i)=>ram[i]=v);ram[128]=5;ram[129]=3;return ram;}
  function create(ram){if(!Array.isArray(ram)||ram.length!==256)throw Error('La RAM debe contener 256 bytes.');return reset({ram:ram.map(byte)});}
  function reset(s){s.r=Object.fromEntries(['PC','IR','OP','MAR','MDR','AX','BX','AC','REN1','REN2'].map(k=>[k,0]));s.flags={ZF:0,CF:0,SF:0};s.queue=[];s.index=0;s.steps=0;s.phase='LISTO';s.instruction='Sin decodificar';s.instructionAddress=0;s.hasOperand=false;s.halted=false;s.error='';s.last=null;return s;}
  function alu(op,a,b){byte(a);byte(b);let n=0,c=0;switch(op){case'ADD':n=a+b;c=+(n>255);break;case'SUB':case'CMP':n=a-b;c=+(n<0);break;case'INC':n=a+1;c=+(n>255);break;case'DEC':n=a-1;c=+(n<0);break;case'AND':n=a&b;break;case'OR':n=a|b;break;case'XOR':n=a^b;break;case'NOT':n=(~a)&255;break;default:throw Error('Operacion ALU desconocida.');}const result=n&255;return{result,flags:{ZF:+(result===0),CF:c,SF:+((result&128)!==0)}};}
  function add(s,kind,src,dst,phase,label){s.queue.push({kind,src,dst,phase,label:label||dst+' ← '+src});}
  function copy(s,src,dst,phase){add(s,'copy',src,dst,phase);}
  function fetch(s){s.queue=[];s.index=0;s.instructionAddress=s.r.PC;copy(s,'PC','MAR','FETCH');copy(s,'RAM','MDR','FETCH');copy(s,'MDR','IR','FETCH');add(s,'inc','PC','PC','FETCH','PC ← PC + 1');add(s,'decode','IR','DEC','DECODE','Interpretar IR y preparar microórdenes');}
  function decode(s){const d=ISA[s.r.IR];s.r.OP=0;s.hasOperand=false;if(!d){s.halted=true;s.error='Opcode '+hex(s.r.IR)+'h inválido en '+hex(s.instructionAddress)+'h. Pulsa RESET.';return;}s.instruction=d.mnemonic;
    if(d.size===2){copy(s,'PC','MAR','DECODE');copy(s,'RAM','MDR','DECODE');copy(s,'MDR','OP','DECODE');add(s,'inc','PC','PC','DECODE','PC ← PC + 1 (operando)');}
    switch(d.op){
      case'MOV':copy(s,d.src==='imm'?'OP':d.src,'AC','EXECUTE');copy(s,'AC',d.dst,'STORE');break;
      case'LOAD':copy(s,'OP','MAR','EXECUTE');copy(s,'RAM','MDR','EXECUTE');copy(s,'MDR',d.dst,'STORE');break;
      case'STORE':copy(s,'OP','MAR','EXECUTE');copy(s,d.src,'MDR','EXECUTE');copy(s,'MDR','RAM','STORE');break;
      case'JMP':case'JZ':case'JNZ':add(s,'jump','OP','PC','EXECUTE','Evaluar '+d.op);add(s,'none','SEQ','SEQ','STORE','Salto completado; sin escritura de datos');break;
      case'HLT':add(s,'none','DEC','SEQ','EXECUTE','HLT ordena detener el reloj');add(s,'halt','SEQ','CLOCK','STORE','HLT: procesador detenido');break;
      default:copy(s,d.dst,'REN1','EXECUTE');copy(s,['INC','DEC'].includes(d.op)?'ONE':d.op==='NOT'?'ZERO':d.src==='imm'?'OP':d.src,'REN2','EXECUTE');add(s,'alu','ALU','AC','EXECUTE','AC ← '+d.op+'(REN1, REN2); actualizar banderas');if(d.op==='CMP')add(s,'none','FLAGS','FLAGS','STORE','CMP conserva AX y BX');else copy(s,'AC',d.dst,'STORE');
    }
  }
  function value(s,k){return k==='RAM'?read(s,s.r.MAR):k==='ONE'?1:k==='ZERO'?0:(s.r[k]??0);}
  function step(s){if(s.halted)return null;if(s.index>=s.queue.length)fetch(s);const q=s.queue[s.index];const before={r:{...s.r},flags:{...s.flags},instruction:s.instruction};let v=value(s,q.src),from=q.src,to=q.dst,label=q.label;const address=s.r.MAR;s.phase=q.phase;
    switch(q.kind){case'copy':if(q.dst==='RAM')write(s,s.r.MAR,v);else s.r[q.dst]=v;if(q.dst==='OP')s.hasOperand=true;break;
    case'inc':s.r.PC=(s.r.PC+1)&255;v=s.r.PC;break;
    case'decode':decode(s);label=s.error||'IR = '+hex(s.r.IR)+'h → '+s.instruction;break;
    case'alu':{const r=alu(ISA[s.r.IR].op,s.r.REN1,s.r.REN2);s.r.AC=r.result;s.flags=r.flags;v=r.result;break;}
    case'jump':{const op=ISA[s.r.IR].op;const taken=op==='JMP'||op==='JZ'&&s.flags.ZF===1||op==='JNZ'&&s.flags.ZF===0;if(taken)s.r.PC=s.r.OP;else{from='FLAGS';to='SEQ';}v=taken?s.r.OP:s.flags.ZF;label=op+(taken?' tomado':' no tomado')+'; PC = '+hex(s.r.PC)+'h';break;}
    case'halt':s.halted=true;break;}
    s.index++;s.steps++;const event={step:s.steps,phase:s.phase,from,to,value:v,label,address,before};s.last={...event};delete s.last.before;return event;
  }
  return{ISA,hex,byte,parseHex,read,write,create,reset,demo,alu,step};
})();
if(typeof module!=='undefined')module.exports=Cpu;
