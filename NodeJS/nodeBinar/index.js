/** Essa api foi criada para o envio de email com arquivos anexos, foi criada para ser integrada ao Visual Basic 6
	, ela recebe os dados e o arquivo, em base 64, e decodifica o arquivo recebido gerando um pdf, pasta pdfGerado, 
	em seguida é disparado o email através dos dados recebidos via parametro **/
	
const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const morgan = require('morgan');
const cors = require('cors');
const emails = require('./config/enviarEmail'); //diretório do modulo que enviar emails

//inicia aexpress
const app = express();

app.use(morgan('dev'));
app.use(bodyParser.urlencoded({ extended: false }));
app.use(express.json());
app.use(cors());
app.use(emails); //permite usar o endpoint do modulo de emails na página principal /enviarEmail



app.listen(2300, () => {
    console.log('Servidor rodando -> http://localhost:2300');
});
