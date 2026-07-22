// ============================================================
// PLANNER DA PRODUTIVIDADE — Google Apps Script Backend
// Deploy como: Web App > Executar como: Eu > Acesso: Qualquer pessoa
// Copiar a URL gerada e colar em index.html na variável API_URL
// ============================================================

const SHEET_ID = '1N-xXFVDyRxkwWIlrM__vJJ-0G1eGY55G-QrzzuSAxIg';
const ABAS = ['PESSOAL', 'OBRA', 'WORK', 'ODOO PROJETO', 'Projeto Ideias'];

// ⚙️ E-mail da impressora Brother (preencha com o e-mail do seu equipamento)
const PRINTER_EMAIL = '82315484369@print.brother.com';

// ⚠️ TEMPORÁRIA: só serve pra forçar a tela de permissão do Google Calendar.
// Selecione ESSA função no dropdown do editor, clique em Executar, aceite a permissão,
// e depois pode apagar essa função (ou deixar, não atrapalha nada).
function autorizarCalendario() {
  const cal = CalendarApp.getCalendarById('leidi.pessoaco@gmail.com') || CalendarApp.getDefaultCalendar();
  Logger.log('Calendário autorizado: ' + cal.getName());
}

// ⚠️ TEMPORÁRIA: força a permissão de UrlFetchApp (usada pra mandar WhatsApp) e do Google Contacts.
// Selecione "autorizarLembrete" no dropdown, execute, aceite a permissão. Pode deixar aqui depois.
function autorizarLembrete() {
  const resp = UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('UrlFetchApp autorizado: ' + resp.getResponseCode());
  const contatos = People.People.searchContacts({ query: 'Daniel', readMask: 'names,phoneNumbers' });
  Logger.log('People API autorizado: ' + JSON.stringify(contatos));
}

// Lista todas as pastas (abas) que são de tarefas de verdade — identifica pelo cabeçalho
// ter as colunas "ID" e "Tarefa", assim ignora sozinho abas técnicas tipo BUFFER/FILA.
// Isso é o que permite pastas novas (criadas pelo botão "Nova Pasta") aparecerem no app sem
// precisar mexer em código nenhum.
function listarPastas() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ss.getSheets()
    .filter(sh => {
      const lastCol = sh.getLastColumn();
      if (lastCol < 1) return false;
      // Trim + minúsculo pra não perder a pasta por causa de espaço extra ou "id"/"tarefa" com case diferente
      const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
      return headers.includes('id') && headers.includes('tarefa');
    })
    .map(sh => sh.getName());
}

function doGet(e) {
  const action = e.parameter.action;
  const aba = e.parameter.aba;

  if (action === 'getData' && aba) return jsonResponse(getData(aba));
  if (action === 'getAll') {
    const result = {};
    listarPastas().forEach(a => { result[a] = getData(a); });
    return jsonResponse(result);
  }
  if (action === 'getPastas') return jsonResponse({ ok: true, pastas: listarPastas() });
  if (action === 'fillMissingIds') return jsonResponse(fillMissingIds());
  if (action === 'printSheet' && aba) return jsonResponse(printSheet(aba));
  if (action === 'searchTema') return jsonResponse(searchTema(e.parameter.tema));
  if (action === 'getRecorrentes') return jsonResponse(getRecorrentes());
  if (action === 'bufferHasWaiting') return jsonResponse(bufferHasWaiting(e.parameter.numero));
  if (action === 'getHeaders' && aba) return jsonResponse(getHeaders(aba));
  if (action === 'adicionarColunaRecorrente') return jsonResponse(adicionarColunaRecorrente());
  if (action === 'migrarAltaParaUrgente') return jsonResponse(migrarAltaParaUrgente());

  return jsonResponse({ error: 'Ação inválida' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch(err) { return jsonResponse({ error: 'Body inválido: ' + err.message }); }
  const { action, aba, id, status, keyword, fields, numero, texto } = body;

  if (action === 'updateStatus') return jsonResponse(updateStatus(aba, id, status));
  if (action === 'updateFields') return jsonResponse(updateFields(aba, keyword, fields));
  if (action === 'deleteTask') return jsonResponse(deleteTask(aba, id));
  if (action === 'addTask') return jsonResponse(addTask(aba, body.fields || {}));
  if (action === 'editTask') return jsonResponse(editTask(aba, id, body.fields || {}));
  if (action === 'enviarLembrete') return jsonResponse(enviarLembrete(aba, id));
  if (action === 'bufferAdd') return jsonResponse(bufferAdd(numero, texto));
  if (action === 'bufferReadAndClear') return jsonResponse(bufferReadAndClear(numero));
  if (action === 'criarPasta') return jsonResponse(criarPasta(body.origem, body.nomeNova));

  return jsonResponse({ error: 'Ação inválida' });
}

// Cria uma pasta nova. Se "origem" for passado, duplica a aba inteira (com todas as tarefas)
// e limpa o que é "da obra antiga": Observação, Data/Hora (início e fim), Data conclusão,
// Envolvido e o ID do evento do Calendar. Deixa: Tarefa, Divisão e Prioridade (é o "molde").
// Todo Status volta pra "Aberto", mesmo que a tarefa original estivesse concluída.
// Sem "origem", cria uma pasta vazia (só com o cabeçalho certo, copiado de uma pasta existente).
function criarPasta(nomeOrigem, nomeNova) {
  try {
    nomeNova = String(nomeNova || '').trim();
    if (!nomeNova) return { ok: false, error: 'Digite um nome pra pasta nova' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    if (ss.getSheetByName(nomeNova)) return { ok: false, error: 'Já existe uma pasta com esse nome' };

    let nova;

    if (nomeOrigem) {
      const origem = ss.getSheetByName(nomeOrigem);
      if (!origem) return { ok: false, error: 'Pasta de origem não encontrada' };

      nova = origem.copyTo(ss);
      nova.setName(nomeNova);
      ss.setActiveSheet(nova);
      ss.moveActiveSheet(ss.getSheets().length); // manda pro final, junto das outras pastas

      const headers = nova.getRange(1, 1, 1, nova.getLastColumn()).getValues()[0];
      const lastRow = nova.getLastRow();
      if (lastRow >= 2) {
        const idCol = headers.indexOf('ID');
        const statusCol = headers.indexOf('Status');
        const dataCol = headers.indexOf('Data Prevista');
        const horaCol = headers.indexOf('Hora');
        const dataFimCol = headers.indexOf('Data Fim');
        const horaFimCol = headers.indexOf('Hora Fim');
        const dataConclusaoCol = headers.indexOf('Data conclusão');
        const obsCol = headers.indexOf('Observação');
        const envolvidoCol = headers.indexOf('Envolvido');
        const eventoIdCol = headers.indexOf('ID Evento Calendário');
        const dataCriacaoCol = headers.indexOf('Data criação');
        const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
        const agora = new Date().getTime();

        const dados = nova.getRange(2, 1, lastRow - 1, headers.length).getValues();
        dados.forEach((linha, i) => {
          if (idCol >= 0) linha[idCol] = agora + i;
          if (statusCol >= 0) linha[statusCol] = 'Aberto';
          if (dataCol >= 0) linha[dataCol] = '';
          if (horaCol >= 0) linha[horaCol] = '';
          if (dataFimCol >= 0) linha[dataFimCol] = '';
          if (horaFimCol >= 0) linha[horaFimCol] = '';
          if (dataConclusaoCol >= 0) linha[dataConclusaoCol] = '';
          if (obsCol >= 0) linha[obsCol] = '';
          if (envolvidoCol >= 0) linha[envolvidoCol] = '';
          if (eventoIdCol >= 0) linha[eventoIdCol] = '';
          if (dataCriacaoCol >= 0) linha[dataCriacaoCol] = hoje;
        });
        nova.getRange(2, 1, lastRow - 1, headers.length).setValues(dados);
      }
    } else {
      nova = ss.insertSheet(nomeNova);
      // Copia o cabeçalho de uma pasta de tarefas existente, pra já nascer com a estrutura certa
      const modelo = ss.getSheets().find(sh => {
        const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
        return h.includes('ID') && h.includes('Tarefa');
      });
      if (modelo) {
        const headers = modelo.getRange(1, 1, 1, modelo.getLastColumn()).getValues()[0];
        nova.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
    }

    return { ok: true, pasta: nomeNova };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function printSheet(abaName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    if (!sheet) return { ok: false, error: 'Aba não encontrada: ' + abaName };

    const sheetId = sheet.getSheetId();
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export`
      + `?format=pdf`
      + `&gid=${sheetId}`
      + `&portrait=true`
      + `&fitw=true`
      + `&size=A4`
      + `&gridlines=false`
      + `&sheetnames=false`
      + `&fzr=false`;

    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token }
    });

    const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    const pdfBlob = response.getBlob().setName(`Planner ${abaName} - ${hoje}.pdf`);

    MailApp.sendEmail({
      to: PRINTER_EMAIL,
      subject: `Planner ${abaName} - ${hoje}`,
      body: '',
      attachments: [pdfBlob]
    });

    return { ok: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function bufferAdd(numero, texto) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('BUFFER');
    if (!sheet) {
      sheet = ss.insertSheet('BUFFER');
      sheet.getRange(1, 1, 1, 4).setValues([['Numero', 'Texto', 'Status', 'Timestamp']]);
    }
    const now = new Date().getTime();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(numero) && data[i][2] === 'waiting') {
        if (now - data[i][3] < 120000) {
          sheet.getRange(i + 1, 2).setValue(data[i][1] + '\n' + texto);
          return { ok: true, isFirst: false };
        }
      }
    }
    sheet.appendRow([numero, texto, 'waiting', now]);
    return { ok: true, isFirst: true };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function getHeaders(abaName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    if (!sheet) return { ok: false, error: 'Aba não encontrada: ' + abaName };
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return { ok: true, headers };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function bufferHasWaiting(numero) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('BUFFER');
    if (!sheet) return { ok: true, waiting: false };
    const now = new Date().getTime();
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(numero) && data[i][2] === 'waiting' && (now - data[i][3] < 120000)) {
        return { ok: true, waiting: true };
      }
    }
    return { ok: true, waiting: false };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function bufferReadAndClear(numero) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName('BUFFER');
    if (!sheet) return { ok: true, texto: '' };
    const data = sheet.getDataRange().getValues();
    let texto = '';
    for (let i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(numero)) {
        if (!texto) texto = String(data[i][1]);
        sheet.deleteRow(i + 1);
      }
    }
    return { ok: true, texto };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function fillMissingIds() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const hoje = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy');
    let filled = 0;
    listarPastas().forEach(abaName => {
      const sheet = ss.getSheetByName(abaName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const idCol = headers.indexOf('ID');
      const tarefaCol = headers.indexOf('Tarefa');
      const dataCriacaoCol = headers.indexOf('Data criação');
      for (let i = 1; i < data.length; i++) {
        const tarefa = String(data[i][tarefaCol] || '').trim();
        if (!tarefa) continue;
        if (idCol >= 0 && (!data[i][idCol] || String(data[i][idCol]).trim() === '')) {
          sheet.getRange(i + 1, idCol + 1).setValue(new Date().getTime() + i);
          filled++;
        }
        if (dataCriacaoCol >= 0 && (!data[i][dataCriacaoCol] || String(data[i][dataCriacaoCol]).trim() === '')) {
          sheet.getRange(i + 1, dataCriacaoCol + 1).setValue(hoje);
          filled++;
        }
      }
    });
    return { ok: true, filled };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function enviarLembrete(abaName, id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    if (!sheet) return { ok: false, error: 'Aba não encontrada: ' + abaName };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const tarefaCol = headers.indexOf('Tarefa');
    const categoriaCol = headers.indexOf('Divisão');
    const envolvidoCol = headers.indexOf('Envolvido');
    const dataCol = headers.indexOf('Data Prevista');
    const horaCol = headers.indexOf('Hora');
    const obsCol = headers.indexOf('Observação');

    let tarefa = null, categoria = '', envolvido = '', dataPrevista = null, hora = '', observacao = '';
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        tarefa = data[i][tarefaCol];
        categoria = String(data[i][categoriaCol] || '').trim();
        envolvido = String(data[i][envolvidoCol] || '').trim();
        dataPrevista = data[i][dataCol];
        hora = horaCol >= 0 ? String(data[i][horaCol] || '').trim() : '';
        observacao = obsCol >= 0 ? String(data[i][obsCol] || '').trim() : '';
        break;
      }
    }
    if (!tarefa) return { ok: false, error: 'Tarefa não encontrada' };
    if (!envolvido) return { ok: false, error: 'Essa tarefa não tem um envolvido definido' };

    const numero = resolverNumeroEnvolvido(envolvido);
    if (!numero) return { ok: false, error: 'Não encontrei o número de "' + envolvido + '" nos seus contatos do Google' };

    enviarWhatsapp(numero, montarTextoLembrete(categoria, tarefa, dataPrevista, hora, observacao));
    return { ok: true, numero };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Daniel: fala só no grupo específico do WhatsApp, não no número pessoal. Outros: busca no Google Contacts.
function resolverNumeroEnvolvido(envolvido) {
  const GRUPO_DANIEL = '120363329699178980@g.us';
  if (String(envolvido).toLowerCase().includes('daniel')) return GRUPO_DANIEL;
  return buscarNumeroContato(envolvido);
}

// Monta o texto do lembrete: categoria, tarefa, data+hora, observação (se tiver)
function montarTextoLembrete(categoria, tarefa, dataPrevista, hora, observacao) {
  const dataTxt = dataPrevista
    ? (dataPrevista instanceof Date ? Utilities.formatDate(dataPrevista, 'America/Sao_Paulo', 'dd/MM/yyyy') : String(dataPrevista))
    : '';
  const horaTxt = String(hora || '').trim().substring(0, 5);

  let msg = '🔔 Lembrete';
  if (categoria) msg += `\n📁 ${categoria}`;
  msg += `\n📋 ${tarefa}`;
  if (dataTxt) msg += `\n📅 ${dataTxt}` + (horaTxt ? ` às ${horaTxt}` : '');
  if (observacao) msg += `\n💬 ${observacao}`;
  return msg;
}

// Busca o telefone salvo nos Contatos do Google pelo nome (People API)
function buscarNumeroContato(nome) {
  try {
    const resp = People.People.searchContacts({ query: nome, readMask: 'names,phoneNumbers' });
    if (!resp.results || !resp.results.length) return null;

    const nomeBusca = String(nome).trim().toLowerCase();

    // Prioriza um contato cujo nome completo bate EXATAMENTE com o que foi digitado
    // (ex: "Daniel Silva" acha certo, mesmo se tiver vários "Daniel" nos contatos).
    // Se só o primeiro nome foi digitado e houver mais de um contato, pega o primeiro resultado
    // (mesma limitação de sempre — recomenda-se usar nome completo quando há homônimos).
    let pessoa = resp.results.find(r => {
      const displayName = (r.person.names && r.person.names[0] && r.person.names[0].displayName || '').toLowerCase();
      return displayName === nomeBusca;
    })?.person;

    if (!pessoa) pessoa = resp.results[0].person;
    if (!pessoa.phoneNumbers || !pessoa.phoneNumbers.length) return null;
    let numero = pessoa.phoneNumbers[0].value.replace(/\D/g, '');
    if (numero.length <= 11) numero = '55' + numero; // adiciona código do Brasil se faltar
    return numero;
  } catch (err) {
    return null;
  }
}

function enviarWhatsapp(numero, texto) {
  const url = 'https://api.ialink.com.br/message/sendText/Zap Leidi';
  UrlFetchApp.fetch(encodeURI(url), {
    method: 'post',
    contentType: 'application/json',
    headers: { apikey: '27137EEDBCE5-4534-A40D-440B61AA9F46' },
    payload: JSON.stringify({ number: numero, text: texto })
  });
}

function addTask(abaName, fields) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    if (!sheet) return { ok: false, error: 'Aba não encontrada: ' + abaName };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const idCol = headers.indexOf('ID');

    if (fields.ID && idCol >= 0) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === String(fields.ID)) {
            return { ok: true, id: fields.ID, duplicate: true };
          }
        }
      }
    }

    const now = new Date();
    const id = fields.ID || now.getTime();
    const hoje = Utilities.formatDate(now, 'America/Sao_Paulo', 'dd/MM/yyyy');

    const row = headers.map(h => {
      if (h === 'ID') return id;
      if (h === 'Data criação') return hoje;
      if (h === 'Status') return fields['Status'] || 'Aberto';
      return fields[h] !== undefined ? fields[h] : '';
    });

    sheet.appendRow(row);

    // Se tem data E hora, cria o evento no Google Calendar e guarda o ID dele na planilha
    // (pra próxima vez que a data/hora mudar, dá pra EDITAR esse mesmo evento em vez de duplicar)
    if (fields['Data Prevista'] && fields['Hora']) {
      const eventoId = criarOuAtualizarEventoCalendario(fields.Tarefa, fields['Data Prevista'], fields['Hora'], fields.Envolvido, null, fields['Hora Fim'], fields['Data Fim']);
      const eventoIdCol = headers.indexOf('ID Evento Calendário');
      if (eventoId && eventoIdCol >= 0) {
        sheet.getRange(sheet.getLastRow(), eventoIdCol + 1).setValue(eventoId);
      }
    }

    // Tarefa nova já nasce com envolvido definido → avisa direto, sem precisar tocar no sino
    if (fields.Envolvido) {
      avisarEnvolvidoSeNovo(fields['Divisão'], fields.Tarefa, fields['Data Prevista'], fields['Hora'], fields['Observação'], fields.Envolvido, fields.Prioridade);
    }

    return { ok: true, id };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Monta o texto de "nova tarefa" no mesmo formato usado pelo fluxo do WhatsApp
function montarTextoNovaTarefa(tarefa, prioridade, envolvido) {
  const emojiPrioridade = { 'Urgente': '🔴', 'Média': '🟡', 'Baixa': '🟢' }[prioridade] || '';
  let msg = `🆕 Nova tarefa para você, ${envolvido}!\n\n・${tarefa}`;
  if (prioridade) msg += `\n${emojiPrioridade} Prioridade: ${prioridade}`;
  msg += `\n\nResponda:\n1️⃣ Ciente\n2️⃣ Explique melhor`;
  return msg;
}

// Manda o aviso de WhatsApp pro envolvido, sem travar o salvamento se der erro
function avisarEnvolvidoSeNovo(categoria, tarefa, dataPrevista, hora, observacao, envolvido, prioridade) {
  try {
    const numero = resolverNumeroEnvolvido(envolvido);
    if (!numero) return;
    enviarWhatsapp(numero, montarTextoNovaTarefa(tarefa, prioridade, envolvido));
  } catch (err) {
    // Não interrompe o salvamento da tarefa se o aviso falhar
  }
}

function editTask(abaName, id, fields) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    if (!sheet) return { ok: false, error: 'Aba não encontrada: ' + abaName };

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');

    const dataCol = headers.indexOf('Data Prevista');
    const horaCol = headers.indexOf('Hora');
    const envolvidoCol = headers.indexOf('Envolvido');
    const categoriaCol = headers.indexOf('Divisão');
    const tarefaCol = headers.indexOf('Tarefa');
    const obsCol = headers.indexOf('Observação');
    const prioridadeCol = headers.indexOf('Prioridade');
    const eventoIdCol = headers.indexOf('ID Evento Calendário');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        const rowIndex = i + 1;

        // Guarda o estado ANTES da edição, pra saber se data/hora realmente mudou
        const dataAntes = dataCol >= 0 ? String(data[i][dataCol] || '').trim() : '';
        const horaAntes = horaCol >= 0 ? String(data[i][horaCol] || '').trim() : '';
        const envolvidoAntes = envolvidoCol >= 0 ? String(data[i][envolvidoCol] || '').trim() : '';
        const eventoIdAntes = eventoIdCol >= 0 ? String(data[i][eventoIdCol] || '').trim() : '';

        // Uma escrita só na planilha (em vez de uma chamada por campo) — bem mais rápido
        const colunas = [];
        const valores = [];
        for (const [campo, valor] of Object.entries(fields)) {
          const col = headers.indexOf(campo);
          if (col < 0) continue;
          colunas.push(col);
          valores.push(valor);
        }
        if (colunas.length) {
          const minCol = Math.min(...colunas);
          const maxCol = Math.max(...colunas);
          const linhaAtual = sheet.getRange(rowIndex, minCol + 1, 1, maxCol - minCol + 1).getValues()[0];
          colunas.forEach((col, i) => { linhaAtual[col - minCol] = valores[i]; });
          sheet.getRange(rowIndex, minCol + 1, 1, maxCol - minCol + 1).setValues([linhaAtual]);
        }

        const envolvidoAtual = fields.Envolvido !== undefined ? fields.Envolvido : envolvidoAntes;
        const dataAtual = fields['Data Prevista'] !== undefined ? String(fields['Data Prevista']).trim() : dataAntes;
        const horaAtual = fields['Hora'] !== undefined ? String(fields['Hora']).trim() : horaAntes;
        const tarefaAtual = fields.Tarefa !== undefined ? fields.Tarefa : data[i][tarefaCol];
        const tarefaMudou = fields.Tarefa !== undefined && fields.Tarefa !== data[i][tarefaCol];

        // Cria/atualiza o evento quando data, hora OU o nome da tarefa mudaram de verdade (não só na primeira vez).
        // Passa o ID do evento antigo (se tiver) pra EDITAR o mesmo evento, nunca duplicar.
        if (dataAtual && horaAtual && (dataAtual !== dataAntes || horaAtual !== horaAntes || tarefaMudou)) {
          const novoEventoId = criarOuAtualizarEventoCalendario(tarefaAtual, dataAtual, horaAtual, envolvidoAtual, eventoIdAntes, fields['Hora Fim'], fields['Data Fim']);
          if (novoEventoId && eventoIdCol >= 0 && novoEventoId !== eventoIdAntes) {
            sheet.getRange(rowIndex, eventoIdCol + 1).setValue(novoEventoId);
          }
        }

        // Só avisa quando o Envolvido PASSA a existir/mudar agora (não tinha antes ou é outra pessoa)
        if (envolvidoAtual && envolvidoAtual.trim().toLowerCase() !== envolvidoAntes.toLowerCase()) {
          avisarEnvolvidoSeNovo(
            fields['Divisão'] !== undefined ? fields['Divisão'] : data[i][categoriaCol],
            fields.Tarefa !== undefined ? fields.Tarefa : data[i][tarefaCol],
            fields['Data Prevista'] !== undefined ? fields['Data Prevista'] : data[i][dataCol],
            fields['Hora'] !== undefined ? fields['Hora'] : data[i][horaCol],
            fields['Observação'] !== undefined ? fields['Observação'] : (obsCol >= 0 ? data[i][obsCol] : ''),
            envolvidoAtual,
            fields.Prioridade !== undefined ? fields.Prioridade : (prioridadeCol >= 0 ? data[i][prioridadeCol] : '')
          );
        }

        return { ok: true, row: rowIndex };
      }
    }
    return { ok: false, error: 'ID não encontrado' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// E-mails conhecidos pra convidar como participante do evento, por nome do envolvido
const EMAILS_ENVOLVIDOS = {
  'daniel': 'daniel.vendas.adm@gmail.com'
};

// Converte dd/MM/yyyy OU yyyy-MM-dd num objeto {dd,mm,yyyy}, ou null se não reconhecer
function partesDeData_(raw) {
  const s = String(raw || '').trim();
  const br = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (br) return { dd: br[1], mm: br[2], yyyy: br[3] };
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return { dd: iso[3], mm: iso[2], yyyy: iso[1] };
  return null;
}

// Cria (ou edita, se já existir eventoIdExistente) um evento no Google Calendar.
// horaFim/dataFim são opcionais: sem eles, o evento dura 30min (bom pra tarefa rápida, não trava a agenda à toa).
// Com dataFim diferente da dataPrevista, vira um evento de vários dias (período, tipo viagem).
// Retorna o ID do evento — quem chamar deve guardar esse ID na planilha (coluna "ID Evento Calendário")
// pra da próxima vez conseguir EDITAR o mesmo evento em vez de criar um novo.
function criarOuAtualizarEventoCalendario(titulo, dataPrevista, hora, envolvido, eventoIdExistente, horaFim, dataFim) {
  try {
    if (!titulo) return null;

    const partesInicio = partesDeData_(dataPrevista);
    if (!partesInicio) return null;

    const partesHora = String(hora).trim().split(':');
    if (partesHora.length < 2) return null;
    const [hh, min] = partesHora;

    const inicio = new Date(Number(partesInicio.yyyy), Number(partesInicio.mm) - 1, Number(partesInicio.dd), Number(hh), Number(min));
    if (isNaN(inicio)) return null;

    // Data/hora de término: usa o que foi informado; sem nada, cai pra 30min de duração no mesmo dia
    let fim;
    if (horaFim) {
      const partesHoraFim = String(horaFim).trim().split(':');
      const partesFimData = dataFim ? partesDeData_(dataFim) : partesInicio;
      if (partesHoraFim.length >= 2 && partesFimData) {
        fim = new Date(Number(partesFimData.yyyy), Number(partesFimData.mm) - 1, Number(partesFimData.dd), Number(partesHoraFim[0]), Number(partesHoraFim[1]));
      }
    }
    if (!fim || isNaN(fim) || fim <= inicio) fim = new Date(inicio.getTime() + 30 * 60000);

    const calendario = CalendarApp.getCalendarById('leidi.pessoaco@gmail.com') || CalendarApp.getDefaultCalendar();

    // Tenta achar o evento pelo ID guardado na planilha — só assim dá pra ter certeza de editar o certo
    // (título sozinho não serve: pode ter "Dentista" várias vezes no ano, em datas diferentes)
    let evento = null;
    if (eventoIdExistente) {
      try { evento = calendario.getEventById(eventoIdExistente); } catch (e) { evento = null; }
    }

    if (evento) {
      evento.setTime(inicio, fim);
      if (evento.getTitle() !== titulo) evento.setTitle(titulo); // sincroniza o nome se a tarefa foi renomeada
    } else {
      evento = calendario.createEvent(titulo, inicio, fim);
      evento.addPopupReminder(60);   // 1h antes
      evento.addPopupReminder(1440); // 1 dia antes
    }

    // Convida o envolvido como participante do evento, se tiver e-mail conhecido
    const nomeEnvolvido = String(envolvido || '').trim().toLowerCase();
    for (const chave in EMAILS_ENVOLVIDOS) {
      if (nomeEnvolvido.includes(chave)) {
        const jaConvidado = evento.getGuestList().some(g => g.getEmail() === EMAILS_ENVOLVIDOS[chave]);
        if (!jaConvidado) evento.addGuest(EMAILS_ENVOLVIDOS[chave]);
        break;
      }
    }

    return evento.getId();
  } catch (err) {
    // Não interrompe o salvamento da tarefa se o calendário falhar
    return null;
  }
}

function deleteTask(abaName, id) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'ID não encontrado' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function updateFields(abaName, keyword, fields) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { ok: false, error: 'Planilha vazia' };

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const tarefaCol = headers.indexOf('Tarefa');
    const statusCol = headers.indexOf('Status');
    const prioCol = headers.indexOf('Prioridade');

    const tarefas = sheet.getRange(2, tarefaCol + 1, lastRow - 1, 1).getValues();
    const statuses = sheet.getRange(2, statusCol + 1, lastRow - 1, 1).getValues();
    const prios = sheet.getRange(2, prioCol + 1, lastRow - 1, 1).getValues();

    const kw = semAcento_(keyword.toLowerCase());
    const candidatos = [];
    for (let i = 0; i < tarefas.length; i++) {
      const tarefa = semAcento_(String(tarefas[i][0]).toLowerCase());
      const status = String(statuses[i][0]);
      if (tarefa.includes(kw) && status !== 'Concluído' && status !== 'Cancelado') {
        candidatos.push({ i, prio: String(prios[i][0]), tarefa: tarefas[i][0] });
      }
    }

    if (!candidatos.length) {
      return { ok: false, error: 'Tarefa não encontrada com a palavra-chave: ' + keyword };
    }

    const prioOrdem = { 'Urgente': 0, 'Média': 1, 'Media': 1, 'Baixa': 2 };
    candidatos.sort((a, b) => (prioOrdem[a.prio] ?? 9) - (prioOrdem[b.prio] ?? 9));

    const melhor = candidatos[0];
    const rowIndex = melhor.i + 2;
    for (const [campo, valor] of Object.entries(fields)) {
      const col = headers.indexOf(campo);
      if (col < 0) continue;
      if ((campo === 'Observação' || campo === 'Observacao') && melhor.tarefa.toLowerCase() === 'mercado') {
        const obsCol = headers.indexOf('Observação') >= 0 ? headers.indexOf('Observação') : headers.indexOf('Observacao');
        const obsAtual = obsCol >= 0 ? String(sheet.getRange(rowIndex, obsCol + 1).getValue()).trim() : '';
        const novoValor = obsAtual ? obsAtual + ', ' + valor : valor;
        sheet.getRange(rowIndex, obsCol + 1).setValue(novoValor);
      } else {
        sheet.getRange(rowIndex, col + 1).setValue(valor);
      }
    }
    return { ok: true, tarefa: melhor.tarefa, row: rowIndex };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function getData(abaName) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0];
    return data.slice(1).map((row, i) => {
      const obj = { _row: i + 2 };
      headers.forEach((h, j) => {
        const val = row[j];
        if (val instanceof Date) {
          obj[h] = val.getFullYear() > 1900
            ? Utilities.formatDate(val, 'America/Sao_Paulo', 'dd/MM/yyyy')
            : Utilities.formatDate(val, 'America/Sao_Paulo', 'HH:mm');
        } else {
          obj[h] = val !== undefined ? String(val) : '';
        }
      });
      return obj;
    }).filter(r => Object.keys(r).some(k => k !== '_row' && String(r[k]).trim() !== ''));
  } catch (err) {
    return { error: err.message };
  }
}

function updateStatus(abaName, id, novoStatus) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(abaName);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idCol = headers.indexOf('ID');
    const statusCol = headers.indexOf('Status');
    const conclusaoCol = headers.indexOf('Data conclusão');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id)) {
        sheet.getRange(i + 1, statusCol + 1).setValue(novoStatus);
        if (novoStatus === 'Concluído' && conclusaoCol >= 0) {
          sheet.getRange(i + 1, conclusaoCol + 1).setValue(
            new Date().toLocaleDateString('pt-BR')
          );
        }
        return { ok: true, row: i + 1 };
      }
    }
    return { ok: false, error: 'ID não encontrado' };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// Lista explícita de sinônimos/variações — evita falso positivo de raiz genérica
// (ex: 'comp' bateria em 'compartilhável' se usássemos prefixo de 4 letras)
const SINONIMOS_TEMA_ = {
  'compras': ['comprar'],
  'comprar': ['compras'],
  'compra': ['comprar', 'compras'],
  'orçar': ['orçamento', 'orçamentos'],
  'orçamento': ['orçar'],
  'orçamentos': ['orçar']
};

function semAcento_(s) {
  return String(s || '')
    .replace(/[áàâã]/g, 'a').replace(/[éê]/g, 'e').replace(/[í]/g, 'i')
    .replace(/[óôõ]/g, 'o').replace(/[ú]/g, 'u').replace(/[ç]/g, 'c');
}

function bateTexto_(kw, texto) {
  const lower = semAcento_(texto.toLowerCase());
  const kwSemAcento = semAcento_(kw);
  if (lower.includes(kwSemAcento)) return true;
  const sinonimos = SINONIMOS_TEMA_[kw] || [];
  return sinonimos.some(s => lower.includes(semAcento_(s)));
}

function searchTema(tema) {
  try {
    if (!tema) return { ok: false, error: 'Tema não informado' };
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const kw = tema.toLowerCase().trim();
    const prioOrdem = { 'Urgente': 0, 'Média': 1, 'Media': 1, 'Baixa': 2 };
    const resultado = {};

    ABAS.forEach(abaName => {
      const sheet = ss.getSheetByName(abaName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      const headers = data[0];
      const tarefaCol = headers.indexOf('Tarefa');
      const categoriaCol = headers.indexOf('Divisão');
      const statusCol = headers.indexOf('Status');
      const prioCol = headers.indexOf('Prioridade');
      const horaCol = headers.indexOf('Hora');
      const dataCol = headers.indexOf('Data Prevista');
      const obsCol = headers.indexOf('Observação');
      const contextoCol = headers.indexOf('Contexto');

      const encontradas = [];
      for (let i = 1; i < data.length; i++) {
        const tarefa = String(data[i][tarefaCol] || '').trim();
        const categoria = String(data[i][categoriaCol] || '').trim();
        const contexto = contextoCol >= 0 ? String(data[i][contextoCol] || '').trim() : '';
        const status = String(data[i][statusCol] || '').trim();
        if (!tarefa) continue;
        if (status === 'Concluído' || status === 'Cancelado') continue;
        const bate = bateTexto_(kw, categoria) || bateTexto_(kw, tarefa) || bateTexto_(kw, contexto);
        if (bate) {
          encontradas.push({
            tarefa,
            categoria,
            prioridade: String(data[i][prioCol] || ''),
            status,
            hora: horaCol >= 0 ? String(data[i][horaCol] || '').trim() : '',
            dataPrevista: dataCol >= 0 ? String(data[i][dataCol] || '').trim() : '',
            observacao: obsCol >= 0 ? String(data[i][obsCol] || '').trim() : ''
          });
        }
      }

      if (encontradas.length > 0) {
        encontradas.sort((a, b) => (prioOrdem[a.prioridade] ?? 9) - (prioOrdem[b.prioridade] ?? 9));
        resultado[abaName] = encontradas;
      }
    });

    return { ok: true, tema, resultado };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function getRecorrentes() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const recorrentes = [];
    ABAS.forEach(abaName => {
      const sheet = ss.getSheetByName(abaName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      const headers = data[0];
      const tarefaCol = headers.indexOf('Tarefa');
      const recorrenteCol = headers.indexOf('Recorrente');
      const categoriaCol = headers.indexOf('Divisão');
      if (recorrenteCol < 0) return;
      for (let i = 1; i < data.length; i++) {
        const tarefa = String(data[i][tarefaCol] || '').trim();
        const recorrente = String(data[i][recorrenteCol] || '').trim();
        if (!tarefa || !recorrente) continue;
        recorrentes.push({
          aba: abaName,
          tarefa,
          categoria: String(data[i][categoriaCol] || '').trim(),
          horario: recorrente
        });
      }
    });
    return { ok: true, recorrentes };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function adicionarColunaRecorrente() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const criadas = [];
    ABAS.forEach(abaName => {
      const sheet = ss.getSheetByName(abaName);
      if (!sheet) return;
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (headers.indexOf('Recorrente') >= 0) return;
      sheet.getRange(1, lastCol + 1).setValue('Recorrente');
      criadas.push(abaName);
    });
    return { ok: true, criadas };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function migrarAltaParaUrgente() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const migradas = {};
    ABAS.forEach(abaName => {
      const sheet = ss.getSheetByName(abaName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      const headers = data[0];
      const prioCol = headers.indexOf('Prioridade');
      if (prioCol < 0) return;
      let count = 0;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][prioCol]).trim() === 'Alta') {
          sheet.getRange(i + 1, prioCol + 1).setValue('Urgente');
          count++;
        }
      }
      if (count > 0) migradas[abaName] = count;
    });
    return { ok: true, migradas };
  } catch (err) {
    return { ok: false, error: err.message };
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function onEdit(e) {
  try {
    const sheet = e.range.getSheet();
    const abaName = sheet.getName();
    if (!ABAS.includes(abaName)) return;

    const row = e.range.getRow();
    if (row < 2) return;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const tarefaCol = headers.indexOf('Tarefa') + 1;
    const dataCriacaoCol = headers.indexOf('Data criação') + 1;
    const idCol = headers.indexOf('ID') + 1;

    if (tarefaCol < 1 || dataCriacaoCol < 1) return;

    const tarefa = sheet.getRange(row, tarefaCol).getValue();
    if (!tarefa || String(tarefa).trim() === '') return;

    // Preenche Data criação se estiver vazia
    const dataCriacao = sheet.getRange(row, dataCriacaoCol).getValue();
    if (!dataCriacao || String(dataCriacao).trim() === '') {
      sheet.getRange(row, dataCriacaoCol).setValue(
        Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy')
      );
    }

    // Preenche ID se estiver vazio
    if (idCol >= 1) {
      const id = sheet.getRange(row, idCol).getValue();
      if (!id || String(id).trim() === '') {
        sheet.getRange(row, idCol).setValue(new Date().getTime());
      }
    }
  } catch (err) {}
}
