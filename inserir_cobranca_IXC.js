// Definição da função verificador_de_registros_de_cobranca
function verificador_de_registros_de_cobranca() {
  // Obtenção da planilha ativa e especificação da aba "Base de cobrança"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base de cobrança");
  // Determinação da última linha preenchida na planilha
  var ultima_linha = sheet.getLastRow();
  // Obtenção do último ID de registro de cobrança na coluna A
  var ultimo_id = sheet.getRange(`A${ultima_linha}`).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  // Obtenção da linha da última data de cobrança na coluna D
  var ultima_data = sheet.getRange(`D${ultima_linha}`).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  // Impressão do último ID e última data para verificação
  console.log(`${ultimo_id}\n${ultima_data}`);

  // Incremento do último ID para começar a iteração
  ultimo_id++;
  // Laço for para iterar sobre os registros a partir do último ID
  for (var i = ultimo_id; i < ultima_data; i++) {
    // Extração dos dados do registro atual
    var id_atendimento = sheet.getRange(i, 1).getValue();
    var data_cobranca = sheet.getRange(i, 4).getValue();
    // Verificação se o registro de cobrança está pendente e a data de cobrança está preenchida
    if (id_atendimento == "" && data_cobranca != "") {
      // Chamada da função para inserir o registro no sistema IXC
      Inserir_registro_no_IXC(i);
    }
  }
}

// Definição da função para inserir registro no sistema IXC
function Inserir_registro_no_IXC(lst) {
  // Obtenção da planilha e dos valores da última linha
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base de cobrança");
  var lastCell = sheet.getRange(lst, 4, lst, 45).getValues();
  // Extração dos valores do registro atual
  var form = lastCell[0];
  
  // Formatação dos dados do registro para envio ao sistema IXC
  var data_cobranca = Utilities.formatDate(form[0], "Brazil/Sao Paulo", "dd/MM/yyyy");
  var analista = form[1];
  var id_boleto = form[2] == "" ? "Não informado" : form[2];
  var id_contrato = form[3];
  var id_cliente = form[39];
  var cliente = form[4];
  var bairro = form[5] == "" ? "Não informado" : form[5];
  var parcelas = form[6] == "" ? "Não informado" : form[6];
  var valor = form[7] == "" ? "Não informado" : form[7];
  var data_vencimento_boleto = form[8] == "" ? "Não informado" : Utilities.formatDate(form[8], "Brazil/Sao Paulo", "dd/MM/yyyy");
  var valor_renegociado = form[9] == "" ? "Não informado" : form[9];
  var data_prevista_de_pagamento = form[10] == "" ? "Não informado" : Utilities.formatDate(form[10], "Brazil/Sao Paulo", "dd/MM/yyyy");
  var id_boleto_renegociado = form[11] == "" ? "Não informado" : form[11];
  var tentativas_telefone = form[12];
  // Continuação das atribuições dos valores...

  // Construção da mensagem a ser enviada ao sistema IXC
  var mensagem = `Data: ${data_cobranca}  |  Atendente: ${analista}\nID Boleto: ${id_boleto}  |  ID Contrato: ${id_contrato}  |  ID Cliente: ${id_cliente}\nNome: ${cliente}  |  Bairro: ${bairro}\nData Vencimento: ${data_vencimento_boleto}  |  Parcelas: ${parcelas}  |  Valor total: ${valor}\nTentativas: ${tentativas_resumo}  |  Respondeu? ${respondeu_resumo}  |  Renegociou? ${renegociou_resumo}\nValor Renegociado: ${valor_renegociado}  |  Id Boleto Renegociado: ${id_boleto_renegociado}\nData Prevista de Pagamento: ${data_prevista_de_pagamento}\n`;

  // Continuação da construção da mensagem...

  // Configuração da requisição HTTP para enviar dados ao sistema IXC
  var token_ixc = seu_token; // chave token criada pelo sistema
  var encode = Utilities.base64Encode(token_ixc, Utilities.Charset.UTF_8);
  var url = url_seu_provedor + "/v1/su_ticket"; // su_ticket é o nome da tabela de atendimentos que receberá os registros
  var settings = {
    "method": "POST",
    "headers": {
      "ixcsoft": "",
      "Authorization": "Basic " + encode,
      "Content-Type": "application/json",
    },
    "payload": JSON.stringify({
      // Definição dos dados a serem enviados para o sistema IXC
    }),
    "json": true
  };

  // Envio da requisição HTTP e obtenção da resposta
  var response = UrlFetchApp.fetch(url, settings).getContentText();
  response = JSON.parse(response);
  var stt = response['type'];
  console.log(response);
  
  // Tratamento da resposta do sistema IXC
  if (stt == "success") {
    var idd = response['id'];
    console.log(idd);
    // Atualização do ID do registro na planilha
    sheet.getRange(lst, 1).setValue(idd);
  } else {
    // Registro de erro na planilha, caso a inserção falhe
    sheet.getRange(lst, 1).setValue("Erro");
  }
  // Formatação da célula na planilha
  sheet.getRange(lst, 1, lst, 45).setHorizontalAlignment('left');
}
