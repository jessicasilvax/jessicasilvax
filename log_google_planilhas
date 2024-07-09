function onEdit(e) {
  // Verifica se há um objeto de evento (e) passado para a função
  if (!e) {
    return;
  }

  // Obtém a planilha ativa onde ocorreu a edição
  var sheet = e.source.getActiveSheet();
  
  // Obtém o intervalo (célula ou grupo de células) que foi editado
  var range = e.range;
  
  // Obtém o número da linha onde ocorreu a edição
  var row = range.getRow();
  
  // Obtém o número da coluna onde ocorreu a edição
  var col = range.getColumn();
  
  // Obtém os cabeçalhos (nomes das colunas) da primeira linha da planilha
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Verifica se a linha alterada é a linha 1 (cabeçalhos)
  if (row === 1) {
    return;
  }

  // Define as colunas que serão monitoradas para alterações
  var watchColumns = ["COMO (PLANO DE AÇÃO)", "STATUS AUTOMÁTICO", "DATA STATUS","Responsável","Impedimento?","Data Fim Real"];
  
  // Mapeia os índices das colunas monitoradas com base nos cabeçalhos da planilha
  var watchColumnsIndices = watchColumns.map(function(colName) {
    return headers.indexOf(colName) + 1; // +1 porque getRange e getValues usam indexação baseada em 1
  });

  // Verifica se a coluna editada está entre as colunas monitoradas
  if (watchColumnsIndices.indexOf(col) !== -1) {
    // Obtém o valor antigo da célula editada (ou define como "vazio" se não houver valor antigo)
    var oldValue = e.oldValue || "vazio";
    
    // Obtém o novo valor formatado da célula editada
    var newValue = range.getDisplayValue();
    
    // Obtém o timestamp atual em milissegundos
    var timestamp = new Date().getTime();
    
    // Formata o timestamp para o formato "dd/MM/yyyy HH:mm:ss"
    var timestampFormatted = Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    // Obtém o e-mail do usuário ativo que realizou a edição
    var user = Session.getActiveUser().getEmail();
    
    // Obtém o nome da coluna editada com base nos cabeçalhos da planilha
    var columnName = headers[col - 1]; // -1 porque headers é indexado a partir de 0
    
    // Monta a mensagem de histórico com informações sobre a alteração
    var historyValue = `Alteração em ${timestampFormatted} por ${user}: ${columnName} alterada para "${newValue}".`;

    // Obtém o índice da coluna "OBSERVAÇÃO"
    var observationCol = headers.indexOf("OBSERVAÇÃO") + 1;
    
    // Obtém a célula específica na coluna "OBSERVAÇÃO" e na linha editada
    var observationCell = sheet.getRange(row, observationCol);
    
    // Obtém o valor atual da célula de OBSERVAÇÃO
    var observationValue = observationCell.getValue();

    // Adiciona a nova informação ao final do valor atual da célula de OBSERVAÇÃO, separando por uma nova linha
    var updatedObservationValue = observationValue + "\n" + historyValue;

    // Define o valor da célula de OBSERVAÇÃO com o novo valor concatenado
    observationCell.setValue(updatedObservationValue);
  }
}
