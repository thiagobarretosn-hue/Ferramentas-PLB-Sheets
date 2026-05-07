# Request Generator — Design Spec
> **Data:** 2026-05-07 | **Status:** Aprovado | **Versão:** 1.0

---

## 1. Contexto

O projeto `Ferramentas-PLB-Sheets` já possui um **Gerador de BOM** (`BOM.gs`) que processa dados de origem e gera relatórios agrupados por Nivel 1/2/3. O **Request Generator** é uma ferramenta **independente** e **separada** que gera uma planilha de requisição de materiais no formato KOJO, com coluna de arredondamento automático por tipo de material.

**Motivação:** O KOJO exige uma ordem de colunas específica (`QTY (ROUND UP) | UOM | DESC | UPC`) com lógica de arredondamento para pipes (PVC Foam Core → 20ft, CPVC → 10ft). Essa lógica não existe no BOM atual.

---

## 2. Escopo

### Incluído
- Ferramenta nova e independente: `Request.gs` + `RequestSidebar.html`
- Configuração própria (sem dependência de `BOM.gs` ou `ConfigService`)
- Sidebar com config de colunas, seleção de combinação, regras de rounding editáveis
- Geração de uma aba de output por vez
- Campos de header: PROJECT, REQUEST, BOM KOJO, ENG, VERSION, Requisition #, Need By
- Linha automática `GENERATED FROM` no header mostrando a combinação usada
- Coluna `ROUND UP` editável + coluna `QTY (ROUND UP)` com fórmula viva

### Excluído
- Geração de múltiplas abas simultâneas (one-at-a-time by design)
- Exportação para PDF (pode ser adicionada numa v2)
- Compartilhamento de config com o BOM

---

## 3. Arquivos

| Arquivo | Ação | Descrição |
|---------|------|-----------|
| `Request.gs` | Criar | Backend completo do Request Generator |
| `RequestSidebar.html` | Criar | Sidebar HTML dedicada |
| `Menu.gs` | Modificar | Adicionar item `📋 Gerador de Request` |

---

## 4. Configuração

**Chave PropertiesService:** `REQUEST_SETTINGS_V1` (DocumentProperties)

```json
{
  "sourceSheet": "REVIT DES CRPLB RISER REQ",
  "colDesc": "J - DESC",
  "colUpc": "M - UPC",
  "colUom": "L - UOM",
  "colQty": "O - QTY",
  "groupL1": "I - FLOOR",
  "groupL2": "H - PHASE",
  "groupL3": "",
  "project": "MOTION 1067",
  "kojoPrefix": "MTN.PLB.RGH.JS.B1067",
  "engineer": "THIAGO",
  "version": "01",
  "roundingRules": [
    { "pattern": "FOAM CORE", "roundUp": 20 },
    { "pattern": "CPVC", "roundUp": 10 }
  ]
}
```

**Fallback de arredondamento:** qualquer item que não matchear nenhuma regra recebe `ROUND UP = 1`.

**Avaliação das regras:** em ordem, primeira que matchear vence. Match via `desc.toUpperCase().includes(pattern.toUpperCase())`.

---

## 5. Sidebar — Estrutura da UI

### Seção 1 — ⚙️ Configuração *(salva uma vez)*
- Dropdown: **Aba Fonte** (lista todas as abas da planilha)
- Dropdowns de colunas (populados com headers da aba fonte):
  - Coluna DESC
  - Coluna UPC
  - Coluna UOM
  - Coluna QTY (raw)
- Dropdowns de grupo (até 3 níveis):
  - Agrupar por Nível 1
  - Agrupar por Nível 2
  - Agrupar por Nível 3 *(opcional)*
- Botão **Salvar Configuração**

### Seção 2 — 📋 Header do Request *(preenchido a cada geração)*
- PROJECT
- REQUEST *(ex: "RISERS 6th Floor")*
- BOM KOJO suffix *(ex: "CA.RSR.F6")*  
  → label mostra o KOJO completo: `{kojoPrefix}.{suffix}`
- ENG
- VERSION
- Requisition # *(ex: REQ-A4799)*
- Need By *(data, ex: 05/29/2026)*

### Seção 3 — 🔽 Seleção da Combinação
- Dropdowns dinâmicos por nível (populados com valores únicos da aba fonte)
- Preview inline: *"→ 73 itens encontrados para esta combinação"*
- Atualiza automaticamente ao mudar a seleção

### Seção 4 — 🔢 Regras de Arredondamento *(editável, salvo com config)*
Tabela com duas colunas:

| PADRÃO (texto no DESC) | ROUND UP |
|------------------------|----------|
| FOAM CORE | 20 |
| CPVC | 10 |
| *(linha vazia para adicionar)* | |

- Botão **+ Adicionar Regra**
- Botão **✕** para remover linha
- Fallback automático (não editável): *"Qualquer outro item → 1"*

### Rodapé
- Botão principal: **`📋 GERAR REQUEST`**
- Status / feedback de sucesso ou erro

---

## 6. Output — Aba Gerada

**Nome da aba:** BOM KOJO suffix digitado pelo usuário.

### Header (linhas 1–7, protegido)

| Col A | Col B (merge B:D) | Col E | Col F |
|-------|-------------------|-------|-------|
| PROJECT: | MOTION 1067 | Requisition # | Need By |
| REQUEST: | RISERS 6th Floor | | |
| BOM KOJO: | MTN.PLB.RGH.JS.B1067.CA.RSR.F6 | REQ-A4799 | 05/29/2026 |
| ENG.: | THIAGO | | |
| VERSION: | 01 | | |
| LAST UPDATE: | 05/07/2026 *(auto)* | | |
| GENERATED FROM: | FLOOR: 6th \| PHASE: Job Site *(auto)* | | |

- Linha 8: vazia
- Linha 9: cabeçalho de dados (negrito)

### Cabeçalho de Dados (linha 9)

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| QTY (ROUND UP) | UOM | DESC | UPC | QTY | ROUND UP |

### Dados (linha 10+)

| Coluna | Tipo | Conteúdo |
|--------|------|----------|
| A | Fórmula | `=ROUNDUP(E10/F10)*F10` — recalcula ao editar F |
| B | Valor | UOM |
| C | Valor | DESC |
| D | Valor | UPC |
| E | Valor | QTY raw (somado por grupo DESC+UPC+UOM) |
| F | Valor | ROUND UP (pré-preenchido pelas regras, **editável**) |

- Ordenado por DESC (A→Z)
- Banding LIGHT_GREY nas linhas de dados

### Larguras de Coluna

| Col | Largura |
|-----|---------|
| A (QTY ROUND UP) | 120 |
| B (UOM) | 80 |
| C (DESC) | 550 |
| D (UPC) | 100 |
| E (QTY) | 100 |
| F (ROUND UP) | 100 |

---

## 7. Fluxo de Processamento — `processRequestCore()`

```
1. Lê aba fonte → sheet.getDataRange().getValues() (batch)
2. Identifica índices de coluna a partir da config salva
3. Filtra linhas pela combinação selecionada (grupo L1 + L2 + L3)
4. Agrupa por chave composta: DESC + "|" + UPC + "|" + UOM
5. Soma QTY para itens iguais
6. Para cada item agrupado:
   a. Aplica regras de rounding em ordem (primeira que matchear)
   b. Fallback: roundUp = 1
7. Ordena por DESC (A→Z)
8. Cria ou limpa aba de destino (nome = KOJO suffix)
9. Escreve header (7 linhas) — protege contra edição
10. Escreve linha vazia + cabeçalho de dados
11. Escreve dados:
    - Cols B:F como setValues() (batch)
    - Col A: setFormula() por linha = "=ROUNDUP(En/Fn)*Fn"
12. Aplica banding, larguras, negrito no cabeçalho
13. Retorna { success: true, count: N }
```

---

## 8. Menu

Adicionar em `Menu.gs`, dentro do menu `🔧 Relatórios Dinâmicos`:

```javascript
.addSeparator()
.addItem('📋 Gerador de Request', 'openRequestSidebar')
```

---

## 9. Funções Públicas de `Request.gs`

| Função | Chamada por | Descrição |
|--------|-------------|-----------|
| `openRequestSidebar()` | Menu | Abre a sidebar |
| `getRequestInitData()` | HTML onload | Retorna abas, colunas, config salva |
| `saveRequestConfig(config)` | HTML | Salva config no PropertiesService |
| `getRequestCombinations(groupCols, sheetName)` | HTML | Retorna combinações + contagem |
| `processRequestCore(combination, settings)` | HTML | Gera a aba de output |

---

## 10. Constantes e Config Interna

```javascript
const REQUEST_CONFIG = {
  SETTINGS_KEY: 'REQUEST_SETTINGS_V1',
  CACHE_TTL: 180,
  HEADER_ROWS: 7,
  DATA_START_ROW: 9,
  COLORS: {
    HEADER_BG: '#2c3e50',
    FONT_LIGHT: '#ffffff',
  },
  DEFAULT_RULES: [
    { pattern: 'FOAM CORE', roundUp: 20 },
    { pattern: 'CPVC', roundUp: 10 }
  ]
};
```

---

## 11. Decisões de Design

| Decisão | Justificativa |
|---------|---------------|
| Ferramenta completamente independente do BOM | Mudanças em BOM.gs não afetam o Request; cada ferramenta evolui sozinha |
| Config própria no PropertiesService | Sem acoplamento de estado; pode ser usada em planilhas diferentes |
| Col F como valor editável (não fórmula) | Permite override manual por linha sem quebrar a fórmula de col A |
| Col A como fórmula viva | Recalcula automaticamente quando o usuário edita col F |
| Regras avaliadas em ordem (first match wins) | Comportamento previsível; regras mais específicas devem vir antes |
| One request at a time | Simplifica a UI; para requests em lote, usar o BOM existente |
