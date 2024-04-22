# inserir_registros_cobranca_IXC
O código é uma função em JavaScript que visa verificar registros de cobrança em uma planilha do Google Sheets e, se necessário, inserir esses registros em um sistema externo chamado IXC. Aqui está um resumo das principais funcionalidades:

verificador_de_registros_de_cobranca(): Esta função verifica os registros de cobrança na planilha. Ele determina o último ID e a última data de cobrança registrada. Em seguida, itera sobre os registros para verificar se há registros de cobrança pendentes de inserção no sistema IXC. Se encontrar registros pendentes, chama a função Inserir_registro_no_IXC().
Inserir_registro_no_IXC(lst): Esta função é chamada para inserir um registro de cobrança no sistema IXC. Ele extrai os dados relevantes da linha especificada da planilha, formata esses dados conforme necessário e os envia para o sistema IXC por meio de uma solicitação HTTP POST. Se a inserção for bem-sucedida, atualiza o ID do registro na planilha com o ID atribuído pelo sistema IXC. Se houver algum erro durante o processo, registra "Erro" na planilha.
O código também inclui a definição de variáveis ​​e a configuração da solicitação HTTP para interagir com o sistema IXC, bem como tratamento de resposta para lidar com casos de sucesso e erro durante a inserção dos registros.
