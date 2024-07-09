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

  var statusEnvioIndex = headers.indexOf('Status Envio E-mail');

  if (statusEnvioIndex === -1) {
    console.error('Coluna "Status Envio E-mail" não encontrada.');
    return;
  }

  var ccEmails = ["jessica.alexandre@totvs.com.br", "pmo.integrador@totvs.com.br"];
  var emailActivities = {};

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
    var nomeTabela = row[headers.indexOf('NOME DA TABELA')];

    if (ondaImplantacao !== 1) {
      continue;
    }

    if (status === "CONCLUÍDO" || status === "CANCELADA") {
      sheet.getRange(i + 1, statusEnvioIndex + 1).setValue('Não Enviado');
      continue;
    }

    var dataFimFormatada = Utilities.formatDate(new Date(dataFim), Session.getScriptTimeZone(), "dd/MM/yyyy");
    var dataStatusFormatada = Utilities.formatDate(new Date(dataStatus), Session.getScriptTimeZone(), "dd/MM/yyyy");

    if (!emailActivities[email]) {
      emailActivities[email] = {
        responsavel: responsavel,
        atividades: []
      };
    }

    emailActivities[email].atividades.push({
      itemCompleto: item + " - " + atividade,
      status: status,
      dataFimFormatada: dataFimFormatada,
      dataStatusFormatada: dataStatusFormatada,
      planoAcao: planoAcao || '-',
      nomeTabela: nomeTabela || '-'
    });
  }

  for (var email in emailActivities) {
    var responsavel = emailActivities[email].responsavel;
    var atividades = emailActivities[email].atividades;

    var htmlTemplate = HtmlService.createTemplateFromFile('EmailTemplate');
    htmlTemplate.responsavel = responsavel;
    htmlTemplate.atividades = atividades;
    var corpoEmail = htmlTemplate.evaluate().getContent();

    try {
      var agora = new Date();
      var dataHoraEnvio = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
      var dataAtual = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd.MM.yy");
      MailApp.sendEmail({
        to: email,
        cc: ccEmails.join(","),
        subject: "[IMPLANTAÇÃO PROTHEUS EMIVE] Atualização de atividade | MIGRAÇÃO" + dataAtual + " " + responsavel,
        htmlBody: corpoEmail
      });

      for (var i = 1; i < data.length; i++) {
        if (data[i][headers.indexOf('Email')] === email && atividades.some(a => a.itemCompleto.includes(data[i][headers.indexOf('ITEM')] + " - " + data[i][headers.indexOf('ATIVIDADE')]))) {
          sheet.getRange(i + 1, statusEnvioIndex + 1).setValue('Enviado em ' + dataHoraEnvio);
        }
      }
    } catch (e) {
      console.error('Erro ao enviar e-mail para ' + email + ': ' + e.message);
    }
  }
}

function formatName(name) {
  if (!name) return '';
  return name.charAt(0).toUpperCase() + name.slice(1).toLowerCase();
}
