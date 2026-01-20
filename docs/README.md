# Documentação - Ferramentas PLB Sheets

Este diretório contém toda a documentação do projeto organizada para facilitar o desenvolvimento com assistência de IA.

## Estrutura

```
docs/
├── README.md              # Este arquivo
├── context/               # CRÍTICO para IA
│   ├── DIRETRIZES.md      # Padrões de código obrigatórios
│   ├── PROJECT_CONTEXT.md # Contexto do projeto
│   └── API_REFERENCE.md   # Referência rápida de APIs
└── guides/                # Guias e tutoriais
    └── (futuros guias)
```

## Ordem de Consulta para IA

1. **CLAUDE.md** (raiz) - Referência rápida e completa
2. **docs/context/DIRETRIZES.md** - Padrões de código
3. **docs/context/PROJECT_CONTEXT.md** - Estrutura do projeto
4. **docs/context/API_REFERENCE.md** - APIs do Google Apps Script

## Notas

- Os arquivos em `context/` são essenciais para a IA entender o projeto
- Mantenha a documentação atualizada ao adicionar novas funcionalidades
- Use JSDoc nos arquivos `.gs` para documentação inline
