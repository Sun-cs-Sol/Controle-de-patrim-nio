//Aba onde estao os itens do patrimonio
const SHEET_NAME = "Patrimonio";
const cache = CacheService.getScriptCache();

const STATUS_SHEET_NAME = "Status"; // Aba de controle
const STATUS_CELL = "A2"; 
const EDITOR_EMAIL_CELL = "B2"; 
const CATEGORIA_OPTIONS_HEADER = "Opcoes_Categoria"; 
const SETOR_OPTIONS_HEADER = "Opcoes_Setor";         
const STATUSPROD_OPTIONS_HEADER = "Opcoes_StatusProduto"; 

//usa os valores da aba status para popular os menus suspensos
function getDropdownOptions() { //
  try {
    //confere a existencia da aba Status
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    if (!sheet) {
      throw new Error(`Aba "${STATUS_SHEET_NAME}" não encontrada.`);
    }
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    
    const catColIndex = headers.indexOf(CATEGORIA_OPTIONS_HEADER); //opçoes de categoria
    const setColIndex = headers.indexOf(SETOR_OPTIONS_HEADER); // opçoes de setor
    const staColIndex = headers.indexOf(STATUSPROD_OPTIONS_HEADER); //opçoes de status do produto

    const options = {
      categorias: [],
      setores: [],
      statusProduto: []
    };

    // Mensagens caso de erro
    if (catColIndex === -1) Logger.log(`Aviso: Coluna "${CATEGORIA_OPTIONS_HEADER}" não encontrada na aba Status.`);
    if (setColIndex === -1) Logger.log(`Aviso: Coluna "${SETOR_OPTIONS_HEADER}" não encontrada na aba Status.`);
    if (staColIndex === -1) Logger.log(`Aviso: Coluna "${STATUSPROD_OPTIONS_HEADER}" não encontrada na aba Status.`);

    for (let i = 1; i < data.length; i++) {
      if (catColIndex !== -1 && data[i][catColIndex] && data[i][catColIndex].toString().trim() !== "") {
        options.categorias.push(data[i][catColIndex].toString().trim());
      }
      if (setColIndex !== -1 && data[i][setColIndex] && data[i][setColIndex].toString().trim() !== "") {
        options.setores.push(data[i][setColIndex].toString().trim());
      }
      if (staColIndex !== -1 && data[i][staColIndex] && data[i][staColIndex].toString().trim() !== "") {
        options.statusProduto.push(data[i][staColIndex].toString().trim());
      }
    }
    
    //ordena
    options.categorias.sort();
    options.setores.sort();
    options.statusProduto.sort();

    
    return { success: true, options: options };

  } catch (e) {
    Logger.log("Erro ao buscar opções de dropdown: " + e.message);
    return { success: false, message: e.message, options: { categorias: [], setores: [], statusProduto: [] } };
  }
}


function _setDirtyFlag(userEmail = 'Sistema') { 
  try {
    // detecta mudanças na planilha e troca os valores na aba Status para ativar o sistema de sinconização
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    sheet.getRange(STATUS_CELL + ":" + EDITOR_EMAIL_CELL).setValues([[1, userEmail]]); 
  } catch (e) {
    Logger.log("Erro ao setar dirty flag: " + e.message);
  }
}

//Procura o status da dirty flag na aba Status para saber se esta desatualizado
function getDirtyFlagStatus() { 
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    const data = sheet.getRange(STATUS_CELL + ":" + EDITOR_EMAIL_CELL).getValues(); 
    const status = data[0][0]; // 0 ou 1
    const lastEditor = data[0][1]; // Pessoa que editou
    return { status: status, lastEditor: lastEditor }; 
  } catch (e) {
    Logger.log("Erro ao ler dirty flag: " + e.message);
    return { status: 0, lastEditor: "" }; 
  }
}

//Quando alguem atualizar o botao de sincronizar, ele seta 0 e apaga o nome de quem editou por ultimo
function resetDirtyFlag() { 
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(STATUS_SHEET_NAME);
    sheet.getRange(STATUS_CELL + ":" + EDITOR_EMAIL_CELL).setValues([[0, ""]]); 
    return { success: true };
  } catch (e) {
    Logger.log(`Erro ao resetar dirty flag: ${e.message}`); 
    return { success: false, message: e.message };
  }
}


//funcao pra bloquear edições simultaneas
function acquireEditLock(itemId, userEmail) {
    // guarda o id do item que esta seno editado
  const lockKey = 'lock_' + itemId;
  try {

    const currentLock = cache.get(lockKey);
    //se a pessoa que esta editando for diferente de quem esta editando, retorna que espere para editar e bloqueia a ação
    if (currentLock && currentLock !== userEmail) {
      return { success: false, message: 'Este item está sendo editado por: ' + currentLock + ". Tente novamente em alguns minutos." };
    }
    //aguarda 30 segundos para liberar o lock
    cache.put(lockKey, userEmail, 300); 
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Erro de servidor ao tentar travar o item: ' + e.message };
  }
}

//libera a trava de edição
function releaseEditLock(itemId, userEmail) {
  const lockKey = 'lock_' + itemId;
  try {
    const currentLock = cache.get(lockKey);
    //Se for igual ele libera
    if (currentLock === userEmail) {
      cache.remove(lockKey);
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Erro ao liberar trava: ' + e.message };
  }
}



//função que carrega a página inicial do app
function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Gestão de Patrimônio")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function onEdit(e) {
  const editedSheetName = e.range.getSheet().getName();
  //verifica se a edição foi na aba de status, se for ele nao faz nada
  if (editedSheetName === STATUS_SHEET_NAME) {
     return; 
  }
  //se for em outra aba ele marca como ediçao e ativa o sistema de sincronização
  clearDashboardCache();
  _setDirtyFlag('Edição Manual'); 
}

//processa o login
function handleLogin(email, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios"); // procura pela aba de usuários no arquivo
    if (!sheet) return { success: false, message: 'Erro crítico: Planilha de usuários não configurada.' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toLowerCase().trim());
    //pega as colunas de email, senha e nivel
    const emailIndex = headers.indexOf('email');
    const passwordIndex = headers.indexOf('senha');
    const levelIndex = headers.indexOf('nivel');

    if (emailIndex === -1 || passwordIndex === -1 || levelIndex === -1) {
        return { success: false, message: 'Configuração da planilha de usuários está incorreta.' };
    }

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[emailIndex].toLowerCase() === email.toLowerCase() && row[passwordIndex] === password) {
        return { success: true, message: "Login bem-sucedido!", role: row[levelIndex] };
      }
    }
    return { success: false, message: "Usuário ou senha inválidos." };
  } catch (e) {
    return { success: false, message: `Ocorreu um erro no servidor: ${e.message}` };
  }
}

function clearDashboardCache() {
  cache.remove('dashboard_data');
  return { success: true, message: "Cache limpo." };
}

//logica para gerar o próximo ID baseado na categoria
function getNextIdForCategory(sheet, category) {
    // Define os prefixos para cada categoria prara o ID e que se quiser adicionar outras categorias e pra adicionar aqui
  const CATEGORY_PREFIXES = {
    'TI': '1',
    'Mobiliário': '2',
    'Equipamentos Eletrônicos': '3'
  };
  const prefix = CATEGORY_PREFIXES[category] || '9'; 
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) { 
    //gera o primero id de uma categoria
    return prefix + "01"; 
  }

  const headers = data[0].map(h => h.toString().toLowerCase().trim());
  const idIndex = headers.indexOf("id");
  const categoryIndex = headers.indexOf("categoria");
  if (idIndex === -1 || categoryIndex === -1) {
    throw new Error("Planilha 'Patrimonio' deve conter colunas 'ID' e 'Categoria'.");
  }
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const rowCategory = data[i][categoryIndex];
    const rowId = data[i][idIndex] ? data[i][idIndex].toString() : "";
    if (rowCategory === category && rowId.startsWith(prefix)) {
      const seq = parseInt(rowId.substring(prefix.length), 10);
      if (seq > maxSeq) {
        maxSeq = seq;
      }
    }
  }
  //ve o maior id da categoria e adiciona 1
  const nextSeq = (maxSeq + 1).toString().padStart(2, '0');
  return prefix + nextSeq; //retorna o proximo id
}


//funcao de adicionar item
function addItem(item, userEmail) { 
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: `Planilha "${SHEET_NAME}" não encontrada.` };

    // cria o id pra categoria
    item.id = getNextIdForCategory(sheet, item.categoria);

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const normalizedHeaders = headers.map(h => h.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g,""));
    
    const newRow = [];
    for (const header of normalizedHeaders) {
      newRow.push(item[header] || "");
    }

    sheet.appendRow(newRow);
    clearDashboardCache();
    _setDirtyFlag(userEmail); 
    return { success: true, message: "Item adicionado com sucesso.", newItem: item };
  } catch (e) {
    Logger.log(`Erro em addItem: ${e.message}`); 
    return { success: false, message: `Erro ao adicionar item: ${e.message}` };
  }
}

//gera as descrições das mudanças feitas no item quando editados
function generateChangesDescription(oldItem, newItem, headers) {
  let changes = [];
  const specialFormatting = {
    'valor unitario': (v) => parseFloat(v || 0).toFixed(2)
  };
  for (const key of headers) {
    if (key === 'id' || key === 'categoria') continue;  // nao serao mudados a categoria e o item
    const formatter = specialFormatting[key] || ((v) => (v || "").toString().trim());
    const oldValue = formatter(oldItem[key]);
    const newValue = formatter(newItem[key]);
    if (oldValue !== newValue) {
        //cri a a descrição da mudança
      const fieldName = key.charAt(0).toUpperCase() + key.slice(1);
      changes.push(`${fieldName} alterado de '${oldValue || "vazio"}' para '${newValue || "vazio"}'`);
    }
  }
  return changes.join('; ');
}

function editItem(item, userEmail = 'Sistema') {
//edita o item na planilha
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: `Planilha "${SHEET_NAME}" não encontrada.` };

    const range = sheet.getDataRange();
    const data = range.getValues();
    const headers = data[0].map(h => h.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g,""));
    const idIndex = headers.indexOf('id');
    const catIndex = headers.indexOf('categoria'); 

    if (idIndex === -1) return { success: false, message: "Coluna 'ID' não encontrada." };
    if (catIndex === -1) return { success: false, message: "Coluna 'Categoria' não encontrada." };

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] && data[i][idIndex].toString() === item.id.toString()) {
        
        const oldRowData = data[i];
        const oldItem = {};
        headers.forEach((h, j) => {
          let cellValue = oldRowData[j]; 
          if (cellValue instanceof Date) {
            cellValue = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          oldItem[h] = cellValue !== undefined ? cellValue.toString() : "";
        });

       
        item.categoria = oldRowData[catIndex]; 

        const itemInSheetOrder = headers.map(header => item[header] || "");

        sheet.getRange(i + 1, 1, 1, headers.length).setValues([itemInSheetOrder]);
        clearDashboardCache();

        const changesDescription = generateChangesDescription(oldItem, item, headers);
        //cria o histórico se houver mudança de edicao
        if (changesDescription) {
          const historyEntry = {
            id: item.id,
            data: new Date(),
            tipo: 'Edição',
            descricao: changesDescription,
            custo: 0, 
            responsavel: userEmail
          };
          addHistoryEntry(historyEntry, userEmail); 
        } else {
          _setDirtyFlag(userEmail); 
        }

        return { success: true, message: "Item atualizado com sucesso." };
      }
    }
    return { success: false, message: "Item não encontrado." };
  } catch (e) {
    Logger.log(`Erro em editItem: ${e.message}`); 
    return { success: false, message: `Erro ao editar item: ${e.message}` };
  }
}

//função para deletar item
function deleteItem(id, userEmail) { 
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: `Planilha "${SHEET_NAME}" não encontrada.` };

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g,""));
    const idIndex = headers.indexOf('id');
    
    if (idIndex === -1) return { success: false, message: "Coluna 'ID' não encontrada." };

    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idIndex] && data[i][idIndex].toString() === id.toString()) {
        sheet.deleteRow(i + 1);
        clearDashboardCache();
        _setDirtyFlag(userEmail); 
        return { success: true, message: "Item excluído com sucesso." };
      }
    }
    return { success: false, message: "Item não encontrado." };
  } catch (e) {
    Logger.log(`Erro em deleteItem: ${e.message}`); 
    return { success: false, message: `Erro ao deletar item: ${e.message}` };
  }
}

//procura o item pelo id
function getItemById(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return { success: false, message: `Planilha "${SHEET_NAME}" não encontrada.` };
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return { success: false, message: "Planilha está vazia." };
    const headers = data[0].map(h => h.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, ""));
    const idIndex = headers.indexOf("id");
    if (idIndex === -1) return { success: false, message: "Coluna 'ID' não encontrada." };
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[idIndex] && row[idIndex].toString() === id.toString()) {
        const item = {};
        headers.forEach((h, j) => {
          if (h) {
            let cellValue = row[j]; 
            if (cellValue instanceof Date) {
              cellValue = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
            }
            item[h] = cellValue !== undefined ? cellValue : "";
          }
        });
        return { success: true, item: item };
      }
    }
    return { success: false, message: "Item não encontrado." };
  } catch (e) {
    return { success: false, message: `Erro ao buscar item: ${e.message}` };
  }
}

//função que carrega os dados do dashboard

function getDashboardData(forceReload = false) {
  const cache = CacheService.getScriptCache();
  if (!forceReload) {
    const cached = cache.get('dashboard_data');
    if (cached != null) {
      return { success: true, data: JSON.parse(cached), source: 'cache' };
    }
  }
  try {
    const allItems = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) { 
        const data = sheet.getDataRange().getValues();
        const headers = data[0].map(h => {
          if (!h) return "";
          return h.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        });
        const rows = data.slice(1);
        rows.forEach(row => {
          if (row.every(cell => cell === "" || cell === null)) return;
          let obj = {};
          headers.forEach((h, i) => {
            if (h) {
              let cellValue = row[i];
              if (cellValue instanceof Date) {
                cellValue = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "yyyy-MM-dd");
              }
              obj[h] = cellValue;
            }
          });
          allItems.push(obj);
        });
      }
    } else {
      Logger.log(`Aviso: A planilha '${SHEET_NAME}' não foi encontrada.`);
      return { success: false, message: `Planilha principal '${SHEET_NAME}' não encontrada.` };
    }
    cache.put('dashboard_data', JSON.stringify(allItems), 3600); 
    return { success: true, data: allItems, source: 'spreadsheet' };
  } catch (e) {
    Logger.log(`Erro em getDashboardData: ${e.message}`);
    return { success: false, message: e.message };
  }
}

//gera o relatorio em pdf
function generatePdfReport(filteredData, tableColumns, logoUrl) {
  try {
    let folder;
    const folders = DriveApp.getFoldersByName("Relatorios_Patrimonio");
    folder = folders.hasNext() ? folders.next() : DriveApp.createFolder("Relatorios_Patrimonio");
    const docName = `Relatorio_Patrimonio_${new Date().toISOString().split('T')[0]}`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    body.setMarginLeft(20);
    body.setMarginRight(20);
    body.setMarginTop(30);
    body.setMarginBottom(30);
    if (logoUrl) {
      try {
        const response = UrlFetchApp.fetch(logoUrl);
        const imgBlob = response.getBlob();
        const image = body.insertImage(0, imgBlob);
        image.setWidth(200);
      } catch (e) {
        Logger.log("Falha ao inserir logo: " + e.message);
      }
    }
    const title = body.appendParagraph(docName);
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setFontSize(12);
    body.appendParagraph(`Gerado em: ${new Date().toLocaleString('pt-BR')}`).setFontSize(8);
    body.appendParagraph(`Total de itens: ${filteredData.length}`).setFontSize(8);
    const headers = tableColumns.filter(c => c.key !== 'actions').map(c => c.label);
    const keys = tableColumns.filter(c => c.key !== 'actions').map(c => c.key);
    const tableData = [headers];
    let totalValue = 0;
    filteredData.forEach(item => {
      const row = [];
      keys.forEach(key => {
        let value = item[key] || '--';
        const colFormat = tableColumns.find(c => c.key === key)?.format;
        if (colFormat === 'currency') {
          const num = parseFloat(String(value).replace(/[^\d,.-]/g, '').replace(',', '.')) || 0;
          totalValue += num;
          value = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(num);
        }
        if (colFormat === 'date' && value) {
          const date = new Date(value);
          if (!isNaN(date)) value = date.toLocaleDateString('pt-BR');
        }
        row.push(String(value));
      });
      tableData.push(row);
    });
    const table = body.appendTable(tableData);
    const pageWidth = body.getPageWidth() - body.getMarginLeft() - body.getMarginRight();
    try { table.setWidth(pageWidth); } catch (_) {}
    const headerRow = table.getRow(0);
    for (let c = 0; c < headerRow.getNumCells(); c++) {
      const cell = headerRow.getCell(c);
      cell.setBackgroundColor("#CCCCCC");
      const p = cell.getChild(0).asParagraph();
      p.setFontSize(8).setBold(true).setSpacingAfter(0).setSpacingBefore(0);
      p.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      try { p.setWordBreak(DocumentApp.WordBreak.NORMAL); } catch (_) {}
    }
    for (let r = 1; r < table.getNumRows(); r++) {
      const row = table.getRow(r);
      for (let c = 0; c < row.getNumCells(); c++) {
        const cell = row.getCell(c);
        const p = cell.getChild(0).asParagraph();
        p.setFontSize(7).setSpacingAfter(0).setSpacingBefore(0);
        try { p.setWordBreak(DocumentApp.WordBreak.NORMAL); } catch (_) {}
      }
    }
    const total = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(totalValue);
    const totalLine = body.appendParagraph(`\nValor Total (Filtrado): ${total}`);
    totalLine.setFontSize(9).setBold(true);
    doc.saveAndClose();
    const pdfBlob = doc.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob).setName(`${docName}.pdf`);
    DriveApp.getFileById(doc.getId()).setTrashed(true);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, message: "Relatório PDF gerado com sucesso!", url: pdfFile.getUrl() };
  } catch (e) {
    return { success: false, message: `Erro ao gerar PDF: ${e.message}` };
  }
}

//função para buscar o histórico pelo id do patrimonio
function getHistoryByPatrimonyId(patrimonyId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Historico");
    if (!sheet) return { success: false, message: "Aba 'Historico' não encontrada." };
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase().trim().normalize("NFD")
                 .replace(/[\u0300-\u036f]/g, ""));
    const idPatrimonioIndex = headers.indexOf("id");
    if (idPatrimonioIndex === -1) return { success: false, message: "Coluna 'ID' não encontrada na aba Histórico." };
    const historyRecords = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][idPatrimonioIndex] && data[i][idPatrimonioIndex].toString() == patrimonyId.toString()) {
        const record = {};
        headers.forEach((h, index) => {
          let cellValue = data[i][index];
          if (cellValue instanceof Date) {
            cellValue = Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
          }
          record[h] = cellValue;
        });
        historyRecords.push(record);
      }
    }
    historyRecords.sort((a, b) => {
        const partsA = a.data.split(' ')[0].split('/');
        const dateA = new Date(+partsA[2], partsA[1] - 1, +partsA[0]);
        const partsB = b.data.split(' ')[0].split('/');
        const dateB = new Date(+partsB[2], partsB[1] - 1, +partsB[0]);
        return dateB - dateA;
    });
    return { success: true, history: historyRecords };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

//função para adicionar entrada no histórico
function addHistoryEntry(entryData, userEmail) { 
  const emailToLog = userEmail || entryData.responsavel; 
  Logger.log(`addHistoryEntry: Iniciado por ${emailToLog} para item ${entryData.id}`); //
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Historico");
    if (!sheet) return { success: false, message: "Aba 'Historico' não encontrada." };
    
    const newRow = [
      entryData.id,
      new Date(entryData.data),
      entryData.tipo,
      entryData.descricao,
      entryData.custo || 0, 
      emailToLog 
    ];
    
    sheet.appendRow(newRow);
    _setDirtyFlag(emailToLog); 
    return { success: true, message: "Registro de histórico adicionado com sucesso." };
  } catch (e) {
    Logger.log(`Erro em addHistoryEntry: ${e.message}`); 
    return { success: false, message: `Erro ao adicionar histórico: ${e.message}` };
  }
}