# Biblioteca de Snippets - Google Apps Script

Funções reutilizáveis organizadas por categoria.

## Estrutura

```
Snippets/
├── cache/
│   └── cache_manager.gs    # Gerenciador de cache com chunks
├── sheets/
│   └── sheet_utils.gs      # Utilitários de planilha
├── ui/
│   └── sidebar_utils.gs    # Criação de sidebars e dialogs
├── utils/
│   └── string_utils.gs     # Manipulação de strings
└── drive/
    └── pdf_export.gs       # Exportação para PDF
```

## Como Usar

Os snippets são namespaces globais que podem ser usados diretamente:

```javascript
// Cache
CacheManager.set('key', data, 300);
const data = CacheManager.get('key');

// Sheets
const obj = SheetUtils.getDataAsObjects(sheet);
SheetUtils.formatHeader(sheet);

// UI
UIUtils.showSidebar('MySidebar', 'Título');
UIUtils.toast('Operação concluída!');

// Strings
const slug = StringUtils.slugify('Meu Texto');
const clean = StringUtils.clean(dirtyString);

// PDF
PDFExport.exportSheet(ss, sheet, folder, 'relatorio');
```

## Adicionando Novos Snippets

1. Criar arquivo na categoria apropriada
2. Usar padrão namespace (objeto com métodos)
3. Documentar com JSDoc
4. Atualizar este README
