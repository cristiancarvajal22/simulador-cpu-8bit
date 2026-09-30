/** Programas pequeños para explorar cada instrucción sin omitir el ciclo CPU. */
var Ejemplos = (() => {
  function crear(code, a, b, inmediato) {
    const d = Cpu.ISA[code];
    if (!d) throw Error('Selecciona una instrucción válida.');
    [a,b,inmediato].forEach(Cpu.byte);
    const ram = Array(256).fill(0);
    ram[128] = a; ram[129] = b;
    // Todos los ejemplos comienzan leyendo los dos datos desde memoria.
    const bytes = [5,128,6,129];
    const notas = {0:'LOAD AX,[80h] · leer dato A',2:'LOAD BX,[81h] · leer dato B'};
    if (['JMP','JZ','JNZ'].includes(d.op)) {
      // CMP AX,BX prepara ZF. El salto omite MOV AX,EEh cuando se toma.
      notas[4] = 'CMP AX,BX · preparar condición'; bytes.push(25);
      notas[5] = d.op+' 09h · saltar a almacenamiento'; bytes.push(code,9);
      notas[7] = 'MOV AX,EEh · solo si no se toma el salto'; bytes.push(1,238);
    } else {
      const at = bytes.length;
      notas[at] = d.mnemonic;
      bytes.push(code);
      if (d.size === 2) bytes.push(d.op==='LOAD'?128:d.op==='STORE'?132:inmediato);
    }
    if (d.op !== 'HLT') {
      notas[bytes.length] = 'STORE [82h],AX · conservar AX final'; bytes.push(7,130);
      notas[bytes.length] = 'STORE [83h],BX · conservar BX final'; bytes.push(8,131);
      notas[bytes.length] = 'HLT'; bytes.push(255);
    }
    bytes.forEach((v,i)=>ram[i]=v);
    notas[128]='Dato A'; notas[129]='Dato B'; notas[130]='AX final'; notas[131]='BX final'; notas[132]='Destino del ejemplo STORE';
    return {ram,notas};
  }
  return {crear};
})();
