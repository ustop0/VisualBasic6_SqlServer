const express = require('express');
const enviarEmail = express.Router();
const nodemailer = require('nodemailer');//enviar email


//a req é a requisição gerada pelo cliente(vb6), res é a interface de resposta http
enviarEmail.post("/enviarEmail/:user/:pass/:from/:to/:subject/:mensagem/:anexo", (req,res) => { //get pelo node e post pelo vb6
        //pegando parametros recebidos da requisição do vb6
        let usuario = req.params.user; //dados de email do usuario
        let senha = req.params.pass;

        //let recebidoDe = usuario; //dados de envio
        let recebidoDe = req.params.from;
        let enviarPara = req.params.to;
        let assunto = req.params.subject;
        let mensagem = req.params.mensagem;
        let anexo = req.params.anexo;

        //testando  se os parametros armazenados nas variáveis estão retornando
        //res.send("Os dados retornados são: "+usuario+senha+recebidoDe+enviarPara+assunto+mensagem);
        /*conteudo do retorno
        console.log("usr: "+ usuario);
        console.log("psw: "+ senha);
        console.log("recebidoDe: "+ recebidoDe);
        console.log("enviarPara: "+ enviarPara);        
        console.log("assunto: "+ assunto);
        console.log("mensagem: "+ mensagem);*/

        //enviando email    
        let transporter = nodemailer.createTransport({
            host: 'smtp.gmail.com',
            port: 587,
            secure: false,
            auth:{
                user: usuario,
                pass: senha
            },
            tls: {
                ciphers:'SSLv3'
            }
        });

        let mailOptions = {
            from: recebidoDe,
            to: enviarPara,
            subject: assunto,
            text: mensagem,
            
            attachments: [
                {
                    //filename: 'abc.pdf',
                    path: anexo, // stream this file
                    contentType: 'application/pdf'
                }
            ]
        }
        transporter.sendMail(mailOptions, function(err, info){
            if(err){
                console.log(" \n\n\n ERRO:" + err + "Houve uma falha ao enviar a mensagem de e-mail");
                res.send("email não foi enviado. Erro: \n" + err);
            } else{
                console.log("Mensagem enviada com sucesso");
                res.send("Email enviado com sucesso e com html.");
            }
        });
        
    });

module.exports = enviarEmail;