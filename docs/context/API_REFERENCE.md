# API REFERENCE - Google Apps Script
> **Versão:** 1.0 | **Atualizado:** 2026-01-20

Referência rápida das APIs mais utilizadas no projeto.

---

## SpreadsheetApp

### Obter Objetos
```javascript
// Planilha ativa
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = SpreadsheetApp.getActiveSheet();
const range = SpreadsheetApp.getActiveRange();
const cell = SpreadsheetApp.getCurrentCell();

// Por ID
const ss = SpreadsheetApp.openById('spreadsheet-id');

// Por URL
const ss = SpreadsheetApp.openByUrl('https://docs.google.com/...');
```

### Manipulação de Abas
```javascript
// Obter aba
const sheet = ss.getSheetByName('NomeAba');
const sheets = ss.getSheets(); // Array de todas

// Criar aba
const newSheet = ss.insertSheet('NovaAba');
const newSheet = ss.insertSheet('NovaAba', 0); // Na posição 0

// Deletar aba
ss.deleteSheet(sheet);

// Renomear
sheet.setName('NovoNome');

// Duplicar
const copy = sheet.copyTo(ss);
copy.setName('Cópia');

// Ocultar/Mostrar
sheet.hideSheet();
sheet.showSheet();
```

### Manipulação de Ranges
```javascript
// Obter range
const range = sheet.getRange('A1');           // Célula única
const range = sheet.getRange('A1:B10');       // Intervalo
const range = sheet.getRange(1, 1);           // Linha 1, Col 1
const range = sheet.getRange(1, 1, 10, 2);    // 10 linhas, 2 colunas

// Range com dados
const dataRange = sheet.getDataRange();
const lastRow = sheet.getLastRow();
const lastCol = sheet.getLastColumn();
```

### Leitura de Dados
```javascript
// Valor único
const value = range.getValue();

// Múltiplos valores (array 2D)
const values = range.getValues();
// [[A1, B1], [A2, B2], ...]

// Fórmulas
const formula = range.getFormula();
const formulas = range.getFormulas();

// Display values (formatados)
const display = range.getDisplayValues();
```

### Escrita de Dados
```javascript
// Valor único
range.setValue('Hello');
range.setValue(123);
range.setValue(new Date());

// Múltiplos valores
range.setValues([
  ['A1', 'B1'],
  ['A2', 'B2']
]);

// Fórmulas
range.setFormula('=SUM(A1:A10)');
range.setFormulas([['=A1+B1'], ['=A2+B2']]);
```

### Formatação
```javascript
// Background
range.setBackground('#FF0000');
range.setBackgrounds([['#FF0000', '#00FF00']]);

// Fonte
range.setFontColor('#000000');
range.setFontSize(12);
range.setFontWeight('bold');
range.setFontFamily('Arial');

// Alinhamento
range.setHorizontalAlignment('center');
range.setVerticalAlignment('middle');

// Bordas
range.setBorder(true, true, true, true, true, true);

// Número
range.setNumberFormat('#,##0.00');
range.setNumberFormat('dd/MM/yyyy');

// Merge
range.merge();
range.breakApart();

// Wrap
range.setWrap(true);
```

### Data Validation
```javascript
const rule = SpreadsheetApp.newDataValidation()
  .requireValueInList(['Opção 1', 'Opção 2', 'Opção 3'])
  .setAllowInvalid(false)
  .setHelpText('Selecione uma opção')
  .build();

range.setDataValidation(rule);

// Remover validação
range.clearDataValidations();
```

### Operações em Massa
```javascript
// Inserir linhas
sheet.insertRowsBefore(1, 5);    // 5 linhas antes da linha 1
sheet.insertRowsAfter(10, 5);    // 5 linhas depois da linha 10

// Deletar linhas
sheet.deleteRows(1, 5);          // 5 linhas a partir da linha 1

// Colunas
sheet.insertColumnsBefore(1, 2);
sheet.deleteColumns(1, 2);

// Limpar
range.clear();                    // Tudo
range.clearContent();             // Só valores
range.clearFormat();              // Só formatação
```

### UI
```javascript
const ui = SpreadsheetApp.getUi();

// Menu
ui.createMenu('Meu Menu')
  .addItem('Item 1', 'funcao1')
  .addSeparator()
  .addSubMenu(ui.createMenu('SubMenu')
    .addItem('SubItem', 'funcao2'))
  .addToUi();

// Alert
ui.alert('Mensagem');
ui.alert('Título', 'Mensagem', ui.ButtonSet.OK_CANCEL);

// Prompt
const response = ui.prompt('Digite algo:');
const text = response.getResponseText();

// Toast
ss.toast('Mensagem', 'Título', 5); // 5 segundos
```

---

## DriveApp

### Arquivos
```javascript
// Por ID
const file = DriveApp.getFileById('file-id');

// Por nome (retorna iterator)
const files = DriveApp.getFilesByName('nome.pdf');
while (files.hasNext()) {
  const file = files.next();
}

// Criar arquivo
const file = DriveApp.createFile('nome.txt', 'conteúdo');
const file = DriveApp.createFile(blob);

// Deletar (mover para lixeira)
file.setTrashed(true);
```

### Pastas
```javascript
// Por ID
const folder = DriveApp.getFolderById('folder-id');

// Root
const root = DriveApp.getRootFolder();

// Criar pasta
const newFolder = DriveApp.createFolder('NovaPasta');
const subFolder = folder.createFolder('SubPasta');

// Criar arquivo na pasta
const file = folder.createFile('nome.txt', 'conteúdo');
```

### Exportar para PDF
```javascript
// Exportar planilha inteira
const blob = ss.getAs('application/pdf');
const file = folder.createFile(blob.setName('relatorio.pdf'));

// Exportar aba específica (via URL)
function exportSheetAsPDF(ss, sheet, folder, fileName) {
  const url = ss.getUrl()
    .replace(/\/edit.*$/, '')
    + '/export?format=pdf'
    + '&gid=' + sheet.getSheetId()
    + '&size=A4'
    + '&portrait=true';

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  return folder.createFile(response.getBlob().setName(fileName + '.pdf'));
}
```

---

## CacheService

### Básico
```javascript
const cache = CacheService.getScriptCache();    // Compartilhado
const cache = CacheService.getUserCache();       // Por usuário
const cache = CacheService.getDocumentCache();   // Por documento

// Set (max 6h = 21600s, max 100KB)
cache.put('key', 'value', 21600);

// Get
const value = cache.get('key'); // null se expirado

// Remove
cache.remove('key');
```

### Múltiplos
```javascript
// Set múltiplos
cache.putAll({
  'key1': 'value1',
  'key2': 'value2'
}, 21600);

// Get múltiplos
const values = cache.getAll(['key1', 'key2']);
// { key1: 'value1', key2: 'value2' }

// Remove múltiplos
cache.removeAll(['key1', 'key2']);
```

### JSON
```javascript
// Salvar objeto
cache.put('data', JSON.stringify(objeto), 21600);

// Recuperar objeto
const data = JSON.parse(cache.get('data') || 'null');
```

---

## PropertiesService

### Tipos
```javascript
// Script (compartilhado)
const props = PropertiesService.getScriptProperties();

// Usuário (por usuário)
const props = PropertiesService.getUserProperties();

// Documento (por documento)
const props = PropertiesService.getDocumentProperties();
```

### Operações
```javascript
// Set
props.setProperty('key', 'value');
props.setProperties({ key1: 'val1', key2: 'val2' });

// Get
const value = props.getProperty('key');
const all = props.getProperties(); // objeto

// Delete
props.deleteProperty('key');
props.deleteAllProperties();
```

---

## LockService

```javascript
const lock = LockService.getScriptLock();    // Global
const lock = LockService.getUserLock();       // Por usuário
const lock = LockService.getDocumentLock();   // Por documento

// Aguardar lock
try {
  lock.waitLock(30000); // 30 segundos
  // Operação crítica
} finally {
  lock.releaseLock();
}

// Tentar lock (não-bloqueante)
if (lock.tryLock(30000)) {
  try {
    // Operação crítica
  } finally {
    lock.releaseLock();
  }
} else {
  // Não conseguiu o lock
}

// Verificar se tem lock
const hasLock = lock.hasLock();
```

---

## HtmlService

### Criar Output
```javascript
// De string
const html = HtmlService.createHtmlOutput('<h1>Hello</h1>');

// De arquivo
const html = HtmlService.createHtmlOutputFromFile('Sidebar');

// De template (com scriptlets)
const template = HtmlService.createTemplateFromFile('Template');
template.data = { nome: 'Valor' };
const html = template.evaluate();
```

### Configurar
```javascript
html.setTitle('Título');
html.setWidth(300);
html.setHeight(400);
html.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
```

### Mostrar
```javascript
// Sidebar
SpreadsheetApp.getUi().showSidebar(html);

// Modal dialog
SpreadsheetApp.getUi().showModalDialog(html, 'Título');

// Modeless dialog
SpreadsheetApp.getUi().showModelessDialog(html, 'Título');
```

### Include (CSS/JS externos)
```javascript
// No .gs
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// No HTML
<?!= include('Styles'); ?>
<?!= include('Scripts'); ?>
```

---

## UrlFetchApp

### GET
```javascript
const response = UrlFetchApp.fetch('https://api.example.com/data');
const content = response.getContentText();
const json = JSON.parse(content);
const code = response.getResponseCode();
```

### POST
```javascript
const options = {
  method: 'post',
  contentType: 'application/json',
  payload: JSON.stringify({ key: 'value' }),
  headers: {
    'Authorization': 'Bearer token'
  },
  muteHttpExceptions: true
};

const response = UrlFetchApp.fetch('https://api.example.com', options);
```

### Múltiplas Requisições
```javascript
const requests = [
  { url: 'https://api.example.com/1' },
  { url: 'https://api.example.com/2' }
];

const responses = UrlFetchApp.fetchAll(requests);
```

---

## Utilities

```javascript
// Base64
const encoded = Utilities.base64Encode('texto');
const decoded = Utilities.base64Decode(encoded);

// UUID
const uuid = Utilities.getUuid();

// Sleep
Utilities.sleep(1000); // 1 segundo

// Formatação
const formatted = Utilities.formatString('Hello %s!', 'World');

// Data
const formatted = Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy HH:mm');

// Hash
const hash = Utilities.computeDigest(
  Utilities.DigestAlgorithm.SHA_256,
  'texto'
);

// Zip
const blob = Utilities.zip([blob1, blob2], 'arquivo.zip');
const blobs = Utilities.unzip(zipBlob);

// JSON (alternativa)
const json = Utilities.jsonStringify(obj);
const obj = Utilities.jsonParse(json);
```

---

## Triggers

### Criar Triggers
```javascript
// Por tempo
ScriptApp.newTrigger('funcao')
  .timeBased()
  .everyMinutes(5)
  .create();

// Diário
ScriptApp.newTrigger('funcao')
  .timeBased()
  .atHour(9)
  .everyDays(1)
  .create();

// onEdit instalável
ScriptApp.newTrigger('funcao')
  .forSpreadsheet(ss)
  .onEdit()
  .create();

// onChange
ScriptApp.newTrigger('funcao')
  .forSpreadsheet(ss)
  .onChange()
  .create();
```

### Gerenciar Triggers
```javascript
// Listar
const triggers = ScriptApp.getProjectTriggers();

// Deletar
triggers.forEach(t => ScriptApp.deleteTrigger(t));

// Deletar específico
ScriptApp.deleteTrigger(trigger);
```

---

## Logger / Console

```javascript
// Logger (clássico)
Logger.log('Mensagem');
Logger.log('Valor: %s', valor);
const logs = Logger.getLog();

// Console (V8 - recomendado)
console.log('Info');
console.info('Info');
console.warn('Warning');
console.error('Error');
console.time('timer');
console.timeEnd('timer');
```
