// ============================================================
// PORTAL DE ENVIO DE E-MAILS - Google Apps Script
// Ficheiro: Code.gs (Backend)
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Portal de Envio de E-mails')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── 1. DESTINATÁRIOS ─────────────────────────────────────────

function loadRecipientsFromSheet(sheetUrl) {
  try {
    var ss = SpreadsheetApp.openByUrl(sheetUrl.trim());
    var sheet = ss.getActiveSheet();
    var data = sheet.getDataRange().getValues();

    if (data.length < 2) {
      return { error: 'A folha está vazia ou só tem cabeçalho.' };
    }

    var headers = data[0].map(function(h) {
      return h.toString().toLowerCase().trim();
    });

    var emailCol = -1, nameCol = -1;
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].indexOf('mail') !== -1) emailCol = i;
      if (headers[i].indexOf('nome') !== -1 || headers[i].indexOf('name') !== -1) nameCol = i;
    }

    if (emailCol === -1) {
      return { error: 'Coluna de e-mail não encontrada. O cabeçalho deve conter a palavra "email" ou "mail".' };
    }

    var recipients = [];
    for (var i = 1; i < data.length; i++) {
      var email = data[i][emailCol].toString().trim();
      var nome  = nameCol >= 0 ? data[i][nameCol].toString().trim() : '';
      if (email && email.indexOf('@') !== -1) {
        recipients.push({ Nome: nome, Email: email });
      }
    }

    if (recipients.length === 0) {
      return { error: 'Nenhum e-mail válido encontrado na folha.' };
    }

    return { recipients: recipients, count: recipients.length };
  } catch (e) {
    return { error: 'Erro ao aceder à folha: ' + e.message };
  }
}

function parseCSVContent(csvContent) {
  try {
    var lines = csvContent.split('\n').filter(function(l) { return l.trim() !== ''; });
    if (lines.length === 0) return { error: 'CSV vazio.' };

    var firstCells = lines[0].split(',').map(function(h) {
      return h.trim().replace(/^"|"$/g, '').toLowerCase();
    });

    var emailCol = -1, nameCol = -1;
    for (var i = 0; i < firstCells.length; i++) {
      if (firstCells[i].indexOf('mail') !== -1) emailCol = i;
      if (firstCells[i].indexOf('nome') !== -1 || firstCells[i].indexOf('name') !== -1) nameCol = i;
    }

    var hasHeader = emailCol !== -1;
    var startLine = hasHeader ? 1 : 0;
    var recipients = [];

    for (var i = startLine; i < lines.length; i++) {
      var parts = lines[i].split(',').map(function(p) { return p.trim().replace(/^"|"$/g, ''); });

      if (hasHeader) {
        var email = parts[emailCol] || '';
        var nome  = nameCol >= 0 ? (parts[nameCol] || '') : '';
        if (email && email.indexOf('@') !== -1) {
          recipients.push({ Nome: nome, Email: email });
        }
      } else {
        // Sem cabeçalho: tentar "Nome,Email" ou só "Email"
        if (parts.length >= 2 && parts[1].indexOf('@') !== -1) {
          recipients.push({ Nome: parts[0], Email: parts[1] });
        } else if (parts[0].indexOf('@') !== -1) {
          recipients.push({ Nome: '', Email: parts[0] });
        }
      }
    }

    if (recipients.length === 0) {
      return { error: 'Nenhum e-mail válido encontrado no CSV.' };
    }

    return { recipients: recipients, count: recipients.length };
  } catch (e) {
    return { error: 'Erro ao processar CSV: ' + e.message };
  }
}

// ── 2. ANEXOS ────────────────────────────────────────────────

function verifyDriveFolder(folderUrl) {
  try {
    var match = folderUrl.match(/[-\w]{25,}/);
    if (!match) return { error: 'URL da pasta inválido. Copie o link completo da pasta do Google Drive.' };

    var folder = DriveApp.getFolderById(match[0]);
    var files  = folder.getFiles();
    var fileList = [];

    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      fileList.push({
        id  : file.getId(),
        name: name,
        key : name.replace(/\.[^/.]+$/, '').toLowerCase().trim()
      });
    }

    return { folderName: folder.getName(), files: fileList, count: fileList.length };
  } catch (e) {
    return { error: 'Erro ao aceder à pasta do Drive: ' + e.message };
  }
}

// ── 3. ENVIO ─────────────────────────────────────────────────

function sendAllEmails(params) {
  var results = { sent: 0, errors: 0, errorDetails: [] };

  try {
    var recipients    = params.recipients;
    var subject       = params.subject;
    var body          = params.body;
    var driveFiles    = params.driveFiles    || [];
    var uploadedFiles = params.uploadedFiles || [];

    for (var i = 0; i < recipients.length; i++) {
      var pessoa = recipients[i];
      var nome   = pessoa.Nome  || '';
      var email  = pessoa.Email;

      try {
        var subjectFinal = subject.replace(/\{Nome\}/gi, nome);
        var bodyFinal    = body.replace(/\{Nome\}/gi, nome);

        var options     = {};
        var attachments = [];

        // Procurar anexo no Google Drive
        for (var j = 0; j < driveFiles.length; j++) {
          var df = driveFiles[j];
          if (df.key === email.toLowerCase() || (nome && df.key === nome.toLowerCase())) {
            attachments.push(DriveApp.getFileById(df.id).getBlob());
            break;
          }
        }

        // Procurar nos ficheiros carregados (base64)
        if (attachments.length === 0) {
          for (var j = 0; j < uploadedFiles.length; j++) {
            var uf    = uploadedFiles[j];
            var ufKey = uf.name.replace(/\.[^/.]+$/, '').toLowerCase().trim();
            if (ufKey === email.toLowerCase() || (nome && ufKey === nome.toLowerCase())) {
              attachments.push(
                Utilities.newBlob(Utilities.base64Decode(uf.data), uf.mimeType, uf.name)
              );
              break;
            }
          }
        }

        if (attachments.length > 0) options.attachments = attachments;

        GmailApp.sendEmail(email, subjectFinal, bodyFinal, options);
        results.sent++;
        Utilities.sleep(1000); // Pausa de segurança (limite Gmail)

      } catch (e) {
        results.errors++;
        results.errorDetails.push(email + ': ' + e.message);
      }
    }
  } catch (e) {
    results.fatalError = e.message;
  }

  return results;
}
