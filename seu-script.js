

// -----------------
//const fs = require('fs')

document.getElementById('divTxt').innerHTML = 'incío'

function mostraTxt() {
    // Caminho para o arquivo Texto.txt

    //document.getElementById('divTxt').innerHTML = 'DURANTE'

    caminhoArquivo = 'Texto.txt';
  
    // Leitura do conteúdo do arquivo
    fs.readFile(caminhoArquivo, 'utf8', (err, data) => {
      if (err) {
        document.getElementById('divTxt').innerHTML = '**** ERRO ****'
        return;
      }
  
      // Exibir conteúdo na div com id "divTxt"
      eledivTxt = document.getElementById('divTxt');

      eledivTxt.textContent = data;

    });
  }
