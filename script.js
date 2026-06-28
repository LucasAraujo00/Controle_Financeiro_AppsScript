var planilha = SpreadsheetApp.getActiveSpreadsheet();

var cadastro = planilha.getSheetByName("Cadastro"); // variavel para identificar a aba cadastro
var auxiliar = planilha.getSheetByName("Auxiliar"); // variavel para identificar a aba Auxiliar
var movimentacoes = planilha.getSheetByName("Movimentações"); // variavel para identificar a aba Movimentações
var relatorio = planilha.getSheetByName("Relatório"); // variavel para identificar a aba Relatório
var gerador = planilha.getSheetByName("Gerador de relatórios") //variavel para acessar a aba gerador de relatorios

// Cadastrar a movimentação financeira na aba "Auxiliar".

function cadastrar(){

  var data = cadastro.getRange("C3:G3").getValue(); // Vai armazenar o valor do campo "data"

  var tipo = cadastro.getRange("C5").getValue(); // Vai armazenar o valor do campo "tipo"

  var categoria = cadastro.getRange("F5:G5").getValue(); // Vai armazenar o valor do campo "categoria"

  var descricao = cadastro.getRange("C7:G7").getValue(); // Vai armazenar o valor do campo "descrição"

  var valor = cadastro.getRange("C9:G9").getValue(); // Vai armazenar o valor do campo "valor"

  
  


  var ultimaLinha = auxiliar.getLastRow()+1; // Seleciona a linha que fica após a ultima linha registrada da aba auxiliar

  auxiliar.getRange(ultimaLinha,1).setValue(data); //Atribui o valor da variável "data" para celula especificada
  auxiliar.getRange(ultimaLinha,2).setFormula('=SPLIT(A'+ultimaLinha+';"/")'); //Atribui a formula "SPLIT"
  auxiliar.getRange(ultimaLinha,5).setValue(tipo);//Atribui o valor da variável "tipo" para celula especificada
  auxiliar.getRange(ultimaLinha,6).setValue(categoria);//Atribui o valor da variável "categoria" para celula especificada
  auxiliar.getRange(ultimaLinha,7).setValue(descricao);//Atribui o valor da variável "descricao" para celula especificada

  if(tipo == 'Entrada'){
    auxiliar.getRange(ultimaLinha,8).setValue(valor);//Atribui o valor positivo da variável "valor" para celula especificada
  }else{
    auxiliar.getRange(ultimaLinha,8).setValue(-valor); //Atribui o valor negativo da variável "valor" para celula especificada
  }
  

  //Inserção de formulas de saldo atual na aba Movimentações
  if (ultimaLinha == 2 ){
    movimentacoes.getRange(ultimaLinha,9).setFormula('=H2');
  }else{
    movimentacoes.getRange(ultimaLinha,9).setFormula("I" + (ultimaLinha - 1) +"+H" +ultimaLinha+'');
  }
  

  limpar()
}

//Essa função vai limpar os campos da aba "Cadastro".

function limpar(){
  cadastro.getRange("C3:G3").clearContent(); // Limpa o intervalo especificado
  cadastro.getRange("C5").clearContent(); // Limpa o intervalo especificado
  cadastro.getRange("F5:G5").clearContent(); // Limpa o intervalo especificado
  cadastro.getRange("C7:G7").clearContent(); // Limpa o intervalo especificado
  cadastro.getRange("C9:G9").clearContent(); // Limpa o intervalo especificado

}



//função para gerar o relatório de movimentações

function gerar(){

  
  relatorio.getRange("F2:F").clearContent();
  relatorio.getRange("F2").setFormula("E2");

  for (var i = 3; i <= relatorio.getLastRow(); i++){
    relatorio.getRange(i,6).setFormula("=F"+(i-1)+"+E"+i)
  };

  limparGerar()
}

function limparGerar(){
  gerador.getRange("C3").clearContent();
  gerador.getRange("F3").clearContent();
}


function limparEnviar(){

  
  gerador.getRange("C5:F5").clearContent();
  gerador.getRange("I4:I6").clearContent();
  gerador.getRange("K4:K5").clearContent();

}

// ENVIAR RELATÓRIO POR EMAIL

function enviar(){

  var destinatario = gerador.getRange("K4:K5").getValue();
  var mensagem = gerador.getRange("I4:I6").getValue();
  var email ={
    to: destinatario,
    subject: "Relatório Financeiro",
    body: mensagem,
    name: "Lukas Araujo",
    attachments: [planilha.getAs(MimeType.PDF).setName("Relatório Financeiro.pdf")] //transformar planilha em pdf
    }

  cadastro.hideSheet();
  movimentacoes.hideSheet();
  gerador.hideSheet();

  MailApp.sendEmail(email)

  cadastro.showSheet();
  movimentacoes.showSheet();
  gerador.showSheet();

  limparEnviar()

}

