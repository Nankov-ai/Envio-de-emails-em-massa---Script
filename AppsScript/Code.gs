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

    var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });

    // Detectar colunas fixas: nome e email
    var emailCol = -1, nameCol = -1;
    for (var i = 0; i < headers.length; i++) {
      if (headers[i].indexOf('mail') !== -1) emailCol = i;
      if (headers[i].indexOf('nome') !== -1 || headers[i].indexOf('name') !== -1) nameCol = i;
    }

    if (emailCol === -1) {
      return { error: 'Coluna de e-mail não encontrada. O cabeçalho deve conter "email" ou "mail".' };
    }

    // Colunas extra: todas as que não são nome nem email, com o seu label do cabeçalho
    var extraCols = [];
    for (var i = 0; i < headers.length; i++) {
      if (i !== emailCol && i !== nameCol) {
        extraCols.push({ col: i, label: headers[i] });
      }
    }

    var recipients = [];

    for (var i = 1; i < data.length; i++) {
      var row   = data[i];
      var email = row[emailCol] ? row[emailCol].toString().trim() : '';
      if (!email || email.indexOf('@') === -1) continue;

      var rec = { Email: email };
      rec.Nome = nameCol >= 0 && row[nameCol] ? row[nameCol].toString().trim() : '';

      // Colunas extra → chave = nome do cabeçalho (ex: rec['empresa'])
      for (var j = 0; j < extraCols.length; j++) {
        var label = extraCols[j].label;
        var col   = extraCols[j].col;
        rec[label] = row[col] ? row[col].toString().trim() : '';
      }

      recipients.push(rec);
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

    // Converter todas as linhas em arrays de células (sem limite de colunas)
    var allRows = lines.map(function(line) {
      return line.split(',').map(function(cell) {
        return cell.trim().replace(/^"|"$/g, '');
      });
    });

    // Verificar se a primeira linha é cabeçalho (nenhuma célula tem @)
    var firstRow  = allRows[0];
    var hasHeader = firstRow.every(function(c) { return c.indexOf('@') === -1; });
    var startLine = hasHeader ? 1 : 0;

    var emailCol = -1, nameCol = -1;

    if (hasHeader) {
      for (var i = 0; i < firstRow.length; i++) {
        var h = firstRow[i].toLowerCase();
        if (h.indexOf('mail') !== -1) emailCol = i;
        if (h.indexOf('nome') !== -1 || h.indexOf('name') !== -1) nameCol = i;
      }
    }

    // Fallback: detectar coluna de e-mail pelos dados
    if (emailCol === -1) {
      outer: for (var i = startLine; i < allRows.length; i++) {
        for (var j = 0; j < allRows[i].length; j++) {
          if (allRows[i][j].indexOf('@') !== -1) { emailCol = j; break outer; }
        }
      }
    }

    if (emailCol === -1) return { error: 'Coluna de e-mail não encontrada no CSV.' };

    // Colunas extra: com cabeçalho → usa o nome; sem cabeçalho → {Coluna A/B/C/D}
    var extraCols = [];
    var colLabels = ['A', 'B', 'C', 'D'];
    if (hasHeader) {
      for (var i = 0; i < firstRow.length; i++) {
        if (i !== emailCol && i !== nameCol) {
          extraCols.push({ col: i, label: firstRow[i].toLowerCase().trim() });
        }
      }
    } else {
      var idx = 0;
      for (var i = 0; i < firstRow.length && idx < 4; i++) {
        if (i !== emailCol) {
          extraCols.push({ col: i, label: 'coluna ' + colLabels[idx].toLowerCase() });
          idx++;
        }
      }
    }

    var recipients = [];

    for (var i = startLine; i < allRows.length; i++) {
      var row   = allRows[i];
      var email = (row[emailCol] || '').trim();
      if (!email || email.indexOf('@') === -1) continue;

      var rec = { Email: email };

      // Nome: da coluna detectada, ou primeira coluna não-email
      if (nameCol >= 0) {
        rec.Nome = row[nameCol] || '';
      } else {
        rec.Nome = '';
        for (var j = 0; j < row.length; j++) {
          if (j !== emailCol) { rec.Nome = row[j] || ''; break; }
        }
      }

      for (var j = 0; j < extraCols.length; j++) {
        rec[extraCols[j].label] = row[extraCols[j].col] || '';
      }

      recipients.push(rec);
    }

    if (recipients.length === 0) return { error: 'Nenhum e-mail válido encontrado no CSV.' };
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

        // Substituir variáveis das colunas extra pelo nome do cabeçalho (case-insensitive)
        var keys = Object.keys(pessoa);
        for (var k = 0; k < keys.length; k++) {
          var key = keys[k];
          if (key === 'Nome' || key === 'Email') continue;
          var val        = pessoa[key] || '';
          var escapedKey = key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          var regex      = new RegExp('\\{' + escapedKey + '\\}', 'gi');
          subjectFinal   = subjectFinal.replace(regex, val);
          bodyFinal      = bodyFinal.replace(regex, val);
        }

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
