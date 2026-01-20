# Templates - Google Apps Script

Templates base para criar novos scripts e sidebars.

## Arquivos Disponíveis

### `new_script_template.gs`
Template completo para novos módulos .gs com:
- Estrutura de configuração
- Função principal com tratamento de erros
- Funções auxiliares privadas
- Documentação JSDoc
- Função de teste

**Como usar:**
1. Copie o arquivo para `src/`
2. Renomeie conforme a funcionalidade
3. Substitua os placeholders `[...]`
4. Implemente a lógica específica

### `sidebar_template.html`
Template completo para sidebars com:
- CSS estilizado (Google Add-ons style)
- Estrutura de seções
- Loading spinner
- Mensagens de status
- Comunicação com backend

**Como usar:**
1. Copie para `html/`
2. Renomeie conforme a sidebar
3. Adapte os campos e opções
4. Implemente as funções backend chamadas

## Checklist para Novos Scripts

- [ ] Copiar template apropriado
- [ ] Renomear arquivo
- [ ] Preencher cabeçalho (descrição, autor, versão)
- [ ] Implementar lógica principal
- [ ] Adicionar tratamento de erros
- [ ] Usar batch operations para planilhas
- [ ] Documentar com JSDoc
- [ ] Testar com dados reais
- [ ] Adicionar ao menu (se necessário)
