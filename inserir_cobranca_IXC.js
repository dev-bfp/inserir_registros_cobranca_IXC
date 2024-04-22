// Definição da função verificador_de_registros_de_cobranca
function verificador_de_registros_de_cobranca() {
  // Obtenção da planilha ativa e especificação da aba "Base de cobrança"
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("nome_da_planilha");
  
  var ultima_linha = sheet.getLastRow(); // Determinação da última linha preenchida na planilha
  var ultimo_id = sheet.getRange(`A${ultima_linha}`).getNextDataCell(SpreadsheetApp.Direction.UP).getRow(); // Obtenção do último ID de registro de cobrança na coluna A
  var ultima_data = sheet.getRange(`D${ultima_linha}`).getNextDataCell(SpreadsheetApp.Direction.UP).getRow(); // Obtenção da linha da última data de cobrança na coluna D
  console.log(`${ultimo_id}\n${ultima_data}`); // Impressão do último ID e última data para verificação
  
  ultimo_id++; // Incremento do último ID para começar a iteração
  // Laço for para iterar sobre os registros a partir do último ID
  for (var i = ultimo_id; i < ultima_data; i++) {
    // Extração dos dados do registro atual
    var id_atendimento = sheet.getRange(i, 1).getValue();
    var data_cobranca = sheet.getRange(i, 4).getValue();
    // Verificação se o registro de cobrança está pendente e a data de cobrança está preenchida
    if (id_atendimento == "" && data_cobranca != "") {
      Inserir_registro_no_IXC(i); // Chamada da função para inserir o registro no sistema IXC
    }
  }
}

// Definição da função para inserir registro no sistema IXC
function Inserir_registro_no_IXC(lst) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("nome_da_planilha"); // Obtenção da planilha e dos valores da última linha
  var lastCell = sheet.getRange(lst, 4, lst, 45).getValues();
  var form = lastCell[0]; // Extração dos valores do registro atual
  
  // Formatação dos dados do registro para envio ao sistema IXC
  data_cobranca = Utilities.formatDate(form[0], "Brazil/Sao Paulo", "dd/MM/yyyy")
  analista = form[1]
  id_boleto = form[2] === "" ? "Não informado" : form[2]
  id_contrato	= form[3]
  id_cliente = form[39]
  cliente	= form[4]
  id_boleto = form[2] === "" ? "Não informado" : form[2];
  id_contrato = form[3];
  id_cliente = form[39];
  cliente = form[4];
  bairro = form[5] === "" ? "Não informado" : form[5];
  parcelas = form[6] === "" ? "Não informado" : form[6];
  valor = form[7] === "" ? "Não informado" : form[7];
  data_vencimento_boleto = form[8] === "" ? "Não informado" : Utilities.formatDate(form[8], "Brazil/Sao Paulo", "dd/MM/yyyy");
  valor_renegociado = form[9] === "" ? "Não informado" : form[9];
  data_prevista_de_pagamento = form[10] === "" ? "Não informado" : Utilities.formatDate(form[10], "Brazil/Sao Paulo", "dd/MM/yyyy");
  id_boleto_renegociado = form[11] === "" ? "Não informado" : form[11];
  tentativas_telefone = form[12]
  respondeu_telefone = form[13]
  renegociou_telefone = form[14]
  numero_contatado_telefone = form[15]
  observacao_telefone = form[16]
  tentativa_whatsapp = form[17]
  respondeu_whatsapp = form[18]
  renegociou_whatsapp = form[19]
  numero_contatado_whatsapp = form[20]
  observacao_whatsapp = form[21]
  tentativas_email = form[22]
  respondeu_email	= form[23]
  renegociou_email = form[24]
  observacao_email = form[25]
  tentativas_sms = form[26]
  respondeu_sms = form[27]
  renegociou_sms = form[28]
  observacao  = form[29]
  tentativas_resumo = form[30]
  respondeu_resumo = form[31]
  renegociou_resumo = form[32]

  // Construção da mensagem a ser inserida no campo descrição
  mensagem = `Data: ${data_cobranca}  |  Atendente: ${analista}
ID Boleto: ${id_boleto}  |  ID Contrato: ${id_contrato}  |  ID Cliente: ${id_cliente}
Nome: ${cliente}  |  Bairro: ${bairro}
Data Vencimento: ${data_vencimento_boleto}  |  Parcelas: ${parcelas}  |  Valor total: ${valor}
Tentativas: ${tentativas_resumo}  |  Respondeu? ${respondeu_resumo}  |  Renegociou? ${renegociou_resumo}
Valor Renegociado: ${valor_renegociado}  |  Id Boleto Renegociado: ${id_boleto_renegociado}
Data Prevista de Pagamento: ${data_prevista_de_pagamento}

  Telefone
- Tentativas: ${tentativas_telefone}
- Respondeu? ${respondeu_telefone} | Renegociou? ${renegociou_telefone} 
- Número contatado: ${numero_contatado_telefone} | Observação: ${observacao_telefone}

  WhatsApp
- Tentativas: ${tentativa_whatsapp}
- Respondeu? ${respondeu_whatsapp} | Renegociou? ${renegociou_whatsapp} 
- Número contatado: ${numero_contatado_whatsapp} | Observação: ${observacao_whatsapp}
  
  E-mail
- Tentativas: ${tentativas_email}
- Respondeu? ${respondeu_email} | Renegociou? ${renegociou_email}
- Observação: ${observacao_email}
  
  SMS
- Tentativas: ${tentativas_sms}
- Respondeu? ${respondeu_sms} | Renegociou? ${renegociou_sms}
`
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
      
      'tipo': 'C',
      'id_estrutura': '',
      'protocolo': 'cobrança',
      'id_circuito': '',
      'id_cliente': id_cliente,
      'id_login': '',
      'id_contrato': id_contrato,
      'id_filial': '',
      'id_assunto': '45',
      'titulo': 'Cobrança Ativa',
      'origem_endereco': 'C',
      'origem_endereco_estrutura': 'E',
      'endereco': '',
      'latitude': '',
      'longitude': '',
      'id_wfl_processo': '',
      'id_ticket_setor': '1',
      'id_responsavel_tecnico': '',
      'data_criacao': '',
      'data_ultima_alteracao': '',
      'prioridade': 'M',
      'data_reservada': '',
      'melhor_horario_reserva': 'Q',
      'id_ticket_origem': 'I',
      'id_usuarios': '',
      'id_resposta': '',
      'menssagem': mensagem,
      'interacao_pendente': 'N',
      'su_status': 'N',
      'id_evento_status_processo': '',
      'id_canal_atendimento': '',
      'status': 'S',
      'mensagens_nao_lida_cli': '0',
      'mensagens_nao_lida_sup': '0',
      'token': '',
      'finalizar_atendimento': 'S',
      'id_su_diagnostico': '',
      'status_sla': '',
      'origem_cadastro': 'P',
      'ultima_atualizacao': 'CURRENT_TIMESTAMP',
      'cliente_fone': '',
      'cliente_telefone_comercial': '',
      'cliente_id_operadora_celular': '',
      'cliente_telefone_celular': '',
      'cliente_whatsapp': '',
      'cliente_ramal': '',
      'cliente_email': '',
      'cliente_contato': '',
      'cliente_website': '',
      'cliente_skype': '',
      'cliente_facebook': '',
      'atualizar_cliente': 'N',
      'latitude_cli': '',
      'longitude_cli': '',
      'atualizar_login': 'N',
      'latitude_login': '',
      'longitude_login': ''
   
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
