const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const context={};
for(const file of ['Cpu.gs','Ejemplos.gs','Hoja.gs'])vm.runInNewContext(fs.readFileSync('src/'+file,'utf8'),context);
for(const code of [17,19,25,20,22,38,33,35,37]){
  const s=context.Cpu.create(context.Ejemplos.crear(code,5,3,3).ram);
  let aluEvent;
  while(!s.halted){const e=context.Cpu.step(s);if(e.from==='ALU')aluEvent=e;}
  const paths=context.transferenciasVisuales_(aluEvent);
  const expected=code===38?['REN1:ALU','ALU:AC']:['REN1:ALU','REN2:ALU','ALU:AC'];
  assert.deepEqual(Array.from(paths,p=>p.from+':'+p.to),expected);
  assert.equal(paths[0].value,5);
  assert.equal(paths.at(-1).value,aluEvent.value);
  if(code!==38)assert.equal(paths[1].value,[20,22].includes(code)?1:3);
}
const transfer={from:'RAM',to:'MDR',value:128};
assert.equal(context.transferenciasVisuales_(transfer)[0],transfer);
console.log('PASS: orden REN1 → ALU, REN2 → ALU, ALU → AC; NOT usa una entrada; INC/DEC usan constante 1.');
