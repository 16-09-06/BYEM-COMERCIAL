const SHEET_HEADERS = {
  Solicitacoes: ['Data', 'Colaborador', 'Setor', 'Item', 'Qtd', 'Justificativa', 'Prioridade', 'Status', 'ID', 'PedidoID'],
  Pedidos: ['ID', 'Data', 'Fornecedor', 'Responsavel', 'FormaPagto', 'Total', 'Status', 'Itens'],
  Estoque: ['Nome', 'Categoria', 'QtdAtual', 'QtdMinima', 'Localizacao', 'UltimaMov'],
  MovimentacoesEstoque: ['Data', 'Tipo', 'Material', 'Quantidade', 'Anterior', 'Atual', 'Responsavel', 'Motivo'],
  Catalogo: ['ID', 'Nome', 'Categoria', 'Variacoes'],
  Usuarios: ['id', 'nome', 'senha', 'perfil']
};

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : '';
    if (action === 'getInitialData') {
      return jsonOutput_(getInitialData());
    }
    return jsonOutput_({ error: 'Ação não reconhecida' });
  } catch (error) {
    return jsonOutput_({ error: error.message || String(error) });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'verifyLogin') {
      const user = verifyUserLogin(data.userId, data.password);
      return jsonOutput_(user);
    }

    switch (data.action) {
      case 'salvarPedido':
        salvarPedido(data);
        break;
      case 'atualizarStatus':
        atualizarStatus(data.idPedido, data.novoStatus);
        break;
      case 'salvarSolicitacao':
        salvarSolicitacao(data);
        break;
      case 'atualizarStatusSolicitacao':
        atualizarStatusSolicitacao(data.id, data.novoStatus);
        break;
      case 'vincularSolicitacoesAoPedido':
        vincularSolicitacoesAoPedido(data.solicitacaoIds, data.idPedido);
        break;
      case 'excluirSolicitacao':
        excluirSolicitacao(data.id);
        break;
      case 'salvarEstoque':
        salvarEstoque(data);
        break;
      case 'atualizarQuantidadeEstoque':
        atualizarQuantidadeEstoque(data.nome, data.campo, data.valor, data.responsavel);
        break;
      case 'salvarItemCatalogo':
        salvarItemCatalogo(data);
        break;
      case 'adicionarVariacao':
        adicionarVariacao(data.itemId, data.novaVariacao);
        break;
      default:
        throw new Error('Ação não reconhecida');
    }

    return jsonOutput_({ success: true });
  } catch (error) {
    return jsonOutput_({ error: error.message || String(error) });
  }
}

function getInitialData() {
  const usuarios = sheetToJSON_('Usuarios').map(user => ({
    id: user.id,
    name: user.nome,
    role: normalizarPerfil_(user.perfil)
  }));

  return {
    solicitacoes: ensureSolicitacaoIds_(),
    pedidos: sheetToJSON_('Pedidos'),
    estoque: sheetToJSON_('Estoque'),
    movimentacoesEstoque: obterHistoricoEstoque_(100),
    catalogo: ensureCatalogIds_(),
    usuarios: usuarios
  };
}

function verifyUserLogin(userId, password) {
  const users = sheetToJSON_('Usuarios');
  const user = users.find(row =>
    String(row.id) === String(userId) &&
    String(row.senha == null ? '' : row.senha) === String(password == null ? '' : password)
  );

  if (!user) return null;

  return {
    id: String(user.id),
    name: String(user.nome || ''),
    role: normalizarPerfil_(user.perfil)
  };
}

function normalizarPerfil_(perfil) {
  const normalizedProfile = String(perfil || '').trim().toLowerCase();
  return ['admin', 'administrador', 'compras', 'compras (admin)'].includes(normalizedProfile)
    ? 'admin'
    : 'colab';
}

function salvarPedido(data) {
  if (!data.idPedido || !data.fornecedor) {
    throw new Error('O pedido precisa de ID e fornecedor.');
  }

  const items = Array.isArray(data.itens) ? data.itens : [];
  const existing = sheetToJSON_('Pedidos').some(row => String(row.ID) === String(data.idPedido));
  if (existing) throw new Error('Já existe um pedido com este ID.');

  appendObject_('Pedidos', {
    ID: String(data.idPedido),
    Data: new Date(),
    Fornecedor: data.fornecedor,
    Responsavel: data.responsavel || '',
    FormaPagto: data.formaPagto || '',
    Total: Number(data.total) || 0,
    Status: 'Aguardando Compra',
    Itens: JSON.stringify(items)
  });
}

function atualizarStatus(idPedido, novoStatus) {
  if (!idPedido || !novoStatus) throw new Error('ID do pedido e novo status são obrigatórios.');

  const sheet = getSheet_('Pedidos');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idColumn = findHeaderIndex_(headers, 'ID');
  const statusColumn = findHeaderIndex_(headers, 'Status');

  for (let index = 1; index < values.length; index++) {
    if (String(values[index][idColumn]) === String(idPedido)) {
      sheet.getRange(index + 1, statusColumn + 1).setValue(novoStatus);
      return;
    }
  }

  throw new Error('Pedido não encontrado: ' + idPedido);
}

function salvarSolicitacao(data) {
  if (!data.colaborador) {
    throw new Error('O nome de quem está solicitando é obrigatório.');
  }

  const items = Array.isArray(data.itens)
    ? data.itens
    : data.item
      ? [{ nome: data.item, qtd: data.qtd }]
      : [];

  if (items.length === 0) {
    throw new Error('Adicione pelo menos um item à solicitação.');
  }

  items.forEach(item => {
    const nome = String(item.nome || item.item || '').trim();
    const variacao = String(item.variacao || '').trim();
    const qtd = Number(item.qtd);

    if (!nome) throw new Error('Todos os itens precisam ter um nome.');
    if (!Number.isInteger(qtd) || qtd < 1) {
      throw new Error('A quantidade de cada item deve ser um número inteiro maior que zero.');
    }

    appendObject_('Solicitacoes', {
      Data: new Date(),
      Colaborador: data.colaborador,
      Setor: data.setor || '',
      Item: variacao ? nome + ' - ' + variacao : nome,
      Qtd: qtd,
      Justificativa: data.justificativa || '',
      Prioridade: data.prioridade || 'Baixa',
      Status: 'Pendente',
      ID: Utilities.getUuid()
    });
  });
}

function atualizarStatusSolicitacao(id, novoStatus) {
  const solicitationId = String(id || '').trim();
  const allowedStatuses = [
    'Pendente',
    'Aprovada',
    'Aguardando Compra',
    'VERIFICANDO',
    'COMPRADO',
    'RECUSADO'
  ];

  if (!solicitationId) throw new Error('ID da solicitação obrigatório.');
  if (!allowedStatuses.includes(novoStatus)) throw new Error('Status da solicitação inválido.');

  const sheet = getSheet_('Solicitacoes');
  const values = sheet.getDataRange().getValues();
  const idColumn = findHeaderIndex_(values[0], 'ID');
  const statusColumn = findHeaderIndex_(values[0], 'Status');
  const rowIndex = values.findIndex((row, index) =>
    index > 0 && String(row[idColumn]) === solicitationId
  );

  if (rowIndex < 1) throw new Error('Solicitação não encontrada.');
  sheet.getRange(rowIndex + 1, statusColumn + 1).setValue(novoStatus);
}

function vincularSolicitacoesAoPedido(ids, idPedido) {
  const requestIds = Array.isArray(ids)
    ? [...new Set(ids.map(id => String(id || '').trim()).filter(Boolean))]
    : [];
  const orderId = String(idPedido || '').trim();

  if (requestIds.length === 0) throw new Error('Selecione solicitações aprovadas.');
  if (!orderId) throw new Error('ID do pedido obrigatório.');

  const sheet = getSheet_('Solicitacoes');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idColumn = findHeaderIndex_(headers, 'ID');
  const statusColumn = findHeaderIndex_(headers, 'Status');
  const orderColumn = findHeaderIndex_(headers, 'PedidoID');
  const rows = requestIds.map(requestId => values.findIndex((row, index) =>
    index > 0 && String(row[idColumn]) === requestId
  ));

  if (rows.some(rowIndex => rowIndex < 1)) {
    throw new Error('Uma ou mais solicitações selecionadas não foram encontradas.');
  }
  if (rows.some(rowIndex => String(values[rowIndex][statusColumn]) !== 'Aprovada')) {
    throw new Error('Todas as solicitações precisam continuar com status Aprovada.');
  }

  rows.forEach(rowIndex => {
    sheet.getRange(rowIndex + 1, statusColumn + 1).setValue('Aguardando Compra');
    sheet.getRange(rowIndex + 1, orderColumn + 1).setValue(orderId);
  });
}

function ensureSolicitacaoIds_() {
  const sheet = getSheet_('Solicitacoes');
  const data = sheet.getDataRange().getValues();
  const idColumn = findHeaderIndex_(data[0], 'ID');

  for (let index = 1; index < data.length; index++) {
    const hasContent = data[index].some(value => value !== '' && value != null);
    if (hasContent && !data[index][idColumn]) {
      sheet.getRange(index + 1, idColumn + 1).setValue(Utilities.getUuid());
    }
  }

  return sheetToJSON_('Solicitacoes');
}

function excluirSolicitacao(id) {
  const solicitationId = String(id || '').trim();
  if (!solicitationId) throw new Error('ID da solicitação obrigatório.');

  const sheet = getSheet_('Solicitacoes');
  const data = sheet.getDataRange().getValues();
  const idColumn = findHeaderIndex_(data[0], 'ID');
  const rowIndex = data.findIndex((row, index) =>
    index > 0 && String(row[idColumn]) === solicitationId
  );

  if (rowIndex < 1) throw new Error('Solicitação não encontrada.');
  sheet.deleteRow(rowIndex + 1);
}

function salvarEstoque(data) {
  const nome = String(data.nome || '').trim();
  const qtd = Number(data.qtd);
  const qtdMinima = data.qtdMinima === '' || data.qtdMinima == null
    ? null
    : Number(data.qtdMinima);

  if (!nome) throw new Error('O nome do material é obrigatório.');
  if (!Number.isFinite(qtd) || qtd <= 0) {
    throw new Error('A quantidade a adicionar deve ser maior que zero.');
  }
  if (qtdMinima !== null && (!Number.isFinite(qtdMinima) || qtdMinima < 0)) {
    throw new Error('A quantidade mínima não pode ser negativa.');
  }

  const sheet = getSheet_('Estoque');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const nameColumn = findHeaderIndex_(headers, 'Nome');
  const quantityColumn = findHeaderIndex_(headers, 'QtdAtual');
  const categoryColumn = findHeaderIndex_(headers, 'Categoria');
  const minimumColumn = findHeaderIndex_(headers, 'QtdMinima');
  const locationColumn = findHeaderIndex_(headers, 'Localizacao');
  const movementColumn = findHeaderIndex_(headers, 'UltimaMov');

  const existingIndex = values.findIndex((row, index) =>
    index > 0 && String(row[nameColumn] || '').trim().toLowerCase() === nome.toLowerCase()
  );

  if (existingIndex > 0) {
    const sheetRow = existingIndex + 1;
    const currentQuantity = Number(values[existingIndex][quantityColumn]) || 0;
    const updatedQuantity = currentQuantity + qtd;
    sheet.getRange(sheetRow, quantityColumn + 1).setValue(updatedQuantity);
    if (data.categoria) sheet.getRange(sheetRow, categoryColumn + 1).setValue(data.categoria);
    if (qtdMinima !== null) sheet.getRange(sheetRow, minimumColumn + 1).setValue(qtdMinima);
    if (data.localizacao) sheet.getRange(sheetRow, locationColumn + 1).setValue(data.localizacao);
    sheet.getRange(sheetRow, movementColumn + 1).setValue(new Date());
    registrarMovimentacaoEstoque_('Entrada', nome, qtd, currentQuantity, updatedQuantity, data.responsavel, data.motivo || 'Entrada manual');
    return;
  }

  const initialMinimum = qtdMinima === null ? 0 : qtdMinima;
  appendObject_('Estoque', {
    Nome: nome,
    Categoria: data.categoria || '',
    QtdAtual: qtd,
    QtdMinima: initialMinimum,
    Localizacao: data.localizacao || '',
    UltimaMov: new Date()
  });
  registrarMovimentacaoEstoque_('Entrada', nome, qtd, 0, qtd, data.responsavel, data.motivo || 'Cadastro manual');
}

function atualizarQuantidadeEstoque(nome, campo, quantidade, responsavel) {
  const normalizedName = String(nome || '').trim().toLowerCase();
  const newQuantity = Number(quantidade);
  const allowedFields = ['QtdAtual', 'QtdMinima'];

  if (!normalizedName) throw new Error('O nome do material é obrigatório.');
  if (!allowedFields.includes(campo)) {
    throw new Error('Campo de estoque inválido.');
  }
  if (quantidade === '' || quantidade == null || !Number.isInteger(newQuantity) || newQuantity < 0) {
    throw new Error('O valor deve ser um número inteiro igual ou maior que zero.');
  }

  const sheet = getSheet_('Estoque');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const nameColumn = findHeaderIndex_(headers, 'Nome');
  const quantityColumn = findHeaderIndex_(headers, campo);
  const movementColumn = findHeaderIndex_(headers, 'UltimaMov');
  const rowIndex = values.findIndex((row, index) =>
    index > 0 && String(row[nameColumn] || '').trim().toLowerCase() === normalizedName
  );

  if (rowIndex < 1) throw new Error('Material não encontrado no estoque: ' + nome);

  const previousQuantity = Number(values[rowIndex][quantityColumn]) || 0;
  sheet.getRange(rowIndex + 1, quantityColumn + 1).setValue(newQuantity);
  sheet.getRange(rowIndex + 1, movementColumn + 1).setValue(new Date());
  if (campo === 'QtdAtual' && previousQuantity !== newQuantity) {
    registrarMovimentacaoEstoque_(
      'Ajuste',
      nome,
      newQuantity - previousQuantity,
      previousQuantity,
      newQuantity,
      responsavel,
      'Ajuste manual de quantidade'
    );
  }
}

function registrarMovimentacaoEstoque_(tipo, material, quantidade, anterior, atual, responsavel, motivo) {
  appendObject_('MovimentacoesEstoque', {
    Data: new Date(),
    Tipo: tipo,
    Material: material,
    Quantidade: quantidade,
    Anterior: anterior,
    Atual: atual,
    Responsavel: responsavel || '',
    Motivo: motivo || ''
  });
}

function obterHistoricoEstoque_(limit) {
  const sheet = getSheet_('MovimentacoesEstoque');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastColumn = sheet.getLastColumn();
  const startRow = Math.max(2, lastRow - limit + 1);
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastColumn)
    .getValues()
    .reverse();

  return rows.map(row => {
    const movement = {};
    headers.forEach((header, index) => {
      if (header !== '' && header != null) movement[header] = row[index];
    });
    return movement;
  });
}

function salvarItemCatalogo(data) {
  if (!data.nome) throw new Error('O nome do item é obrigatório.');

  const id = String(data.id || ('CAT-' + Utilities.getUuid().slice(0, 8)));
  const duplicate = sheetToJSON_('Catalogo').some(row => String(row.ID) === id);
  if (duplicate) throw new Error('Já existe um item com este ID.');

  appendObject_('Catalogo', {
    ID: id,
    Nome: data.nome,
    Categoria: data.categoria || '',
    Variacoes: data.variacoes || ''
  });
}

function adicionarVariacao(itemId, novaVariacao) {
  if (!itemId || !novaVariacao) throw new Error('Item e variação são obrigatórios.');

  const sheet = getSheet_('Catalogo');
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idColumn = findHeaderIndex_(headers, 'ID');
  const variationsColumn = findHeaderIndex_(headers, 'Variacoes');

  for (let index = 1; index < values.length; index++) {
    if (String(values[index][idColumn]) === String(itemId)) {
      const current = String(values[index][variationsColumn] || '');
      const variations = current.split(',').map(value => value.trim()).filter(Boolean);
      if (!variations.some(value => value.toLowerCase() === String(novaVariacao).trim().toLowerCase())) {
        variations.push(String(novaVariacao).trim());
        sheet.getRange(index + 1, variationsColumn + 1).setValue(variations.join(', '));
      }
      return;
    }
  }

  throw new Error('Item do catálogo não encontrado: ' + itemId);
}

function sheetToJSON_(sheetName) {
  const sheet = getSheet_(sheetName);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0];
  const required = SHEET_HEADERS[sheetName] || [];
  return data.slice(1).filter(row => row.some(value => value !== '' && value != null)).map(row => {
    const result = {};
    headers.forEach((header, index) => {
      if (header === '' || header == null) return;
      const canonical = required.find(name => normalizeHeader_(name) === normalizeHeader_(header)) || header;
      result[canonical] = row[index];
    });

    if (sheetName === 'Pedidos') {
      result.Itens = parseItems_(result.Itens);
    }
    return result;
  });
}

function ensureCatalogIds_() {
  const sheet = getSheet_('Catalogo');
  const data = sheet.getDataRange().getValues();
  const idColumn = findHeaderIndex_(data[0], 'ID');

  for (let index = 1; index < data.length; index++) {
    const hasContent = data[index].some(value => value !== '' && value != null);
    if (hasContent && !data[index][idColumn]) {
      const id = 'CAT-' + Utilities.getUuid().slice(0, 8);
      sheet.getRange(index + 1, idColumn + 1).setValue(id);
    }
  }

  return sheetToJSON_('Catalogo');
}

function appendObject_(sheetName, object) {
  const sheet = getSheet_(sheetName);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = headers.map(header => {
    const key = Object.keys(object).find(name => normalizeHeader_(name) === normalizeHeader_(header));
    return key === undefined ? '' : object[key];
  });
  sheet.appendRow(values);
}

function getSheet_(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('Este script precisa estar vinculado à planilha.');

  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);
  ensureHeaders_(sheet, SHEET_HEADERS[sheetName] || []);
  return sheet;
}

function ensureHeaders_(sheet, requiredHeaders) {
  if (requiredHeaders.length === 0) return;

  const lastColumn = sheet.getLastColumn();
  const current = lastColumn > 0 ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
  const hasHeader = current.some(value => String(value || '').trim() !== '');

  if (!hasHeader) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  let nextColumn = lastColumn;
  requiredHeaders.forEach(required => {
    const exists = current.some(value => normalizeHeader_(value) === normalizeHeader_(required));
    if (!exists) {
      nextColumn++;
      sheet.getRange(1, nextColumn).setValue(required);
      current.push(required);
    }
  });
}

function findHeaderIndex_(headers, name) {
  const index = headers.findIndex(header => normalizeHeader_(header) === normalizeHeader_(name));
  if (index < 0) throw new Error('Cabeçalho não encontrado: ' + name);
  return index;
}

function normalizeHeader_(value) {
  return String(value == null ? '' : value).trim().toLowerCase();
}

function parseItems_(value) {
  if (Array.isArray(value)) return value;
  if (!value) return [];
  try {
    const parsed = JSON.parse(String(value));
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function jsonOutput_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}