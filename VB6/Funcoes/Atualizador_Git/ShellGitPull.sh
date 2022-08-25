#!/bin/sh

cd /c/Fontes

#Data e hora atual
hoje=$(date +%Y%m%d)

#renomeando o arquivo antigo com a data da operação
mv "Taura Embalagem.exe" "Taura Embalagem_${hoje}.exe"

#copiando os módulos para a pasta de backup_modulos
cp "Taura Embalagem_${hoje}.exe" /c/Fontes/backup_modulos       

#log do git pull
echo "Data e hora do pull: $(date +%Y%m%d)" > status_git_pull.txt
git pull >> status_git_pull.txt

exit

#echo Press Enter...
#read
