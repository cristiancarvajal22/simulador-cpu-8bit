const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const ctx={module:{exports:{}}};vm.runInNewContext(fs.readFileSync('src/Cpu.gs','utf8'),ctx);const C=ctx.module.exports;
const make=bytes=>C.create([...bytes,...Array(256-bytes.length).fill(0)]);
function finish(s){for(let n=0;!s.halted&&n<10000;n++)C.step(s);assert.equal(s.halted,true);assert.equal(s.error,'');return s;}
let s=make([1,5,255]);C.step(s);assert.equal(s.r.PC,0);C.step(s);assert.equal(s.r.MDR,1);assert.equal(s.r.IR,0);C.step(s);assert.equal(s.r.IR,1);C.step(s);assert.equal(s.r.PC,1);finish(s);assert.equal(s.r.AX,5);assert.equal(s.phase,'STORE');let n=s.steps;assert.equal(C.step(s),null);assert.equal(s.steps,n);
// Exhaustive unsigned arithmetic boundaries and flag semantics.
for(let a=0;a<256;a++)for(let b=0;b<256;b++){for(const op of ['ADD','SUB','CMP']){let out=C.alu(op,a,b),raw=op==='ADD'?a+b:a-b;assert.equal(out.result,raw&255);assert.equal(out.flags.CF,+(op==='ADD'?raw>255:raw<0));assert.equal(out.flags.ZF,+((raw&255)===0));assert.equal(out.flags.SF,+((raw&128)!==0));}}
for(const [bytes,reg,result] of [[[1,255,16,1,255],'AX',0],[[2,0,66,1,255],'BX',255],[[1,9,24,9,255],'AX',9],[[2,255,21,23,255],'BX',255],[[1,240,32,15,34,3,36,1,38,255],'AX',253],[[1,7,7,128,6,128,255],'BX',7]])assert.equal(finish(make(bytes)).r[reg],result);
for(const [a,expect] of [[0,0],[1,99]])assert.equal(finish(make([1,a,24,0,49,8,1,99,255])).r.AX,expect);
s=make([254]);for(let i=0;i<5;i++)C.step(s);assert.match(s.error,/00h/);assert.equal(s.halted,true);
s=make([42,255]);s.r.PC=255;s.ram[255]=1;finish(s);assert.equal(s.r.AX,42);assert.equal(s.r.PC,2);
for(const [a,b] of [[5,3],[7,0],[0,5],[255,2],[12,9]]){s=C.create(C.demo());s.ram[128]=a;s.ram[129]=b;finish(s);assert.equal(s.ram[130],(a*b)&255);assert.equal(s.ram[129],0);}
// Every opcode is exercised, including both destination registers.
for(const d of Object.values(C.ISA)){s=make(d.size===2?[d.code,5,255,0,0,255]:[d.code,255]);s.r.AX=9;s.r.BX=3;for(let i=0;i<30;i++){const e=C.step(s);if(e&&e.phase==='STORE')break;}assert.equal(s.error,'');assert.equal(s.phase,'STORE',d.mnemonic);}
s=make([1,255,16,1,255]);let event;while(!s.halted){const e=C.step(s);if(e.from==='ALU')event=e;}assert.equal(event.value,0);assert.equal(event.before.r.AX,255);
assert.throws(()=>C.parseHex('100'));assert.throws(()=>C.read(s,256));assert.throws(()=>C.write(s,0,-1));assert.equal(C.parseHex('ff'),255);
console.log('PASS: 196608 casos ALU; '+Object.keys(C.ISA).length+' opcodes; microoperaciones, saltos, RAM, wrap, HLT y 5 casos demo.');
