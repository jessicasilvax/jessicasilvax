function sendEmails() {
  var sheetName = 'ATIVIDADES';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if (!sheet) {
    console.error('Planilha "' + sheetName + '" não encontrada.');
    return;
  }
  
  var dataRange = sheet.getDataRange();
  var data = dataRange.getValues();
  var headers = data[0];
  
  // Obtém o índice da coluna 'Status Envio E-mail'
  var statusEnvioIndex = headers.indexOf('Status Envio E-mail');
  
  // Verifica se a coluna 'Status Envio E-mail' existe
  if (statusEnvioIndex === -1) {
    console.error('Coluna "Status Envio E-mail" não encontrada.');
    return;
  }
  
  // E-mails para cópia
  var ccEmails = ["jessica.alexandre@totvs.com.br", "pmo.integrador@totvs.com.br", "carlos.junior@emive.com.br"];
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var email = row[headers.indexOf('Email')];
    var responsavel = formatName(row[headers.indexOf('Responsável')]);
    var item = row[headers.indexOf('ITEM')];
    var atividade = row[headers.indexOf('ATIVIDADE')];
    var status = row[headers.indexOf('Status Automático')];
    var dataFim = row[headers.indexOf('Nova Data Fim Planejada')];
    var ondaImplantacao = row[headers.indexOf('ONDA DE IMPLANTAÇÃO')];
    var dataStatus = row[headers.indexOf('DATA STATUS')];
    var planoAcao = row[headers.indexOf('COMO (PLANO DE AÇÃO)')];
    var nomeTabela = row[headers.indexOf('NOME DA TABELA')]; // Obtém o valor da coluna NOME DA TABELA
    
    // Verifica se o conteúdo da coluna 'ONDA DE IMPLANTAÇÃO' é igual a 1
    if (ondaImplantacao !== 1) {
      continue;
    }
    
    // Verifica se o status é CONCLUÍDO ou CANCELADA
    if (status === "CONCLUÍDO" || status === "CANCELADA") {
      sheet.getRange(i + 1, statusEnvioIndex + 1).setValue('Não Enviado');
      continue;
    }
    
    // Formata a data para DD/MM/AAAA
    var dataFimFormatada = Utilities.formatDate(new Date(dataFim), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var dataStatusFormatada = Utilities.formatDate(new Date(dataStatus), Session.getScriptTimeZone(), "dd/MM/yyyy");
    
    // Constrói o item no formato "ITEM + ATIVIDADE"
    var itemCompleto = item + " - " + atividade;
    
    // Constrói o corpo do e-mail com quebras de linha utilizando HTML
    var corpoEmail = "Olá " + responsavel + ",<br><br>" +
                    "Espero que esteja bem.<br><br>" +
                    "Está no radar do PMO Integrador o acompanhamento da atividade abaixo:<br><br>" +
                    "<b>Item:</b> " + itemCompleto + "<br><br>";
    
    // Inclui a variável 'NOME DA TABELA' se não estiver vazia
    if (nomeTabela) {
      corpoEmail += "<b>Tabela vinculada à atividade:</b> " + nomeTabela + "<br><br>";
    }

    corpoEmail += "<b>Status:</b> " + status + "<br><br>" +
                  "<b>Prazo:</b> " + dataFimFormatada + "<br><br>" +
                  "<b>Ultima atualização:</b> " + dataStatusFormatada + "<br><br>";
    
    // Inclui a variável 'COMO (PLANO DE AÇÃO)' se não estiver vazia
    if (planoAcao) {
      corpoEmail += "<b>Evolução da atividade:</b> " + planoAcao + "<br><br>";
    }
    
    corpoEmail += "Favor informar se houve alguma atualização na atividade, previsão de início, bem como plano de ação e/ou previsão para conclusão.<br><br>" +
                  "Conto com seu apoio para atualização do status diário.<br><br>" +
                  "Jessica Alexandre | PMO Integrador";
    
    // Envio do e-mail com cópia para os e-mails desejados
    try {
      MailApp.sendEmail({
        to: email,
        cc: ccEmails.join(","),
        subject: "[PROTHEUS EMIVE] Atualização de atividade | " + responsavel + " - " + itemCompleto,
        htmlBody: corpoEmail
      });
      
      // Registra o status de envio com data e hora na planilha
      var agora = new Date();
      var dataHoraEnvio = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      sheet.getRange(i + 1, statusEnvioIndex + 1).setValue('Enviado em ' + dataHoraEnvio);
    } catch (e) {
      // Em caso de erro ao enviar o e-mail, registra como "Não Enviado" na planilha
      sheet.getRange(i + 1, statusEnvioIndex + 1).setValue('Não Enviado');
      console.error('Erro ao enviar e-mail para ' + email + ': ' + e.message);
    }
  }
}

function formatName(name) {
  if (!name) return '';
  return name.charAt(0).toUpperCase() + name.slice(1).toLowerCase();
}
