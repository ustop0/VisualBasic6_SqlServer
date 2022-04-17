'use strict';

/*const fs = require('fs');

//convertendo arquivo (pdf) para base 64
let buff = fs.readFileSync('./Temp/abc.pdf');
let base64data = buff.toString('base64');

console.log('PDF converted to base 64 is:\n\n' + base64data);*/

   
let pdf = process.stdin.read();

let buff2 = new Buffer.from(pdf, 'base64');
                            
//recebe o nome da filial por parametro
let nome_filial = "3"
let buff2 = new Buffer.from(pdf, 'base64');
//fs.writeFileSync('./pdfGerado/arquivo.pdf', buff2);
fs.writeFileSync(`./pdfGerado/arquivo${nome}.pdf`, buff2);

console.log('PDF converted to base 64 is convertido');