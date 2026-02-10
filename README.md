# Conversor Milheiro → Unidade

Aplicação web para converter preços de milheiro para unidade em planilhas.

## Estrutura

```
conversor-milheiro/
├── server.py              # Backend Flask (API)
├── static/
│   └── index.html         # Frontend (HTML/CSS/JS)
└── README.md
```

## Instalação e Execução

### 1. Instalar dependências

```bash
pip install flask pandas openpyxl
```

### 2. Executar

```bash
python server.py
pip install xlrd

```

### 3. Acessar

Abra o navegador em: **http://localhost:5000**

## Como usar

1. **Upload** — Arraste ou selecione sua planilha (.xlsx, .xls, .csv, .ods)
2. **Selecione** — Marque as colunas com preços em milheiro (apenas numéricas são selecionáveis)
3. **Converta** — Clique em "Converter Selecionadas" (divisor padrão: 1000)
4. **Download** — Baixe a planilha convertida

## API Endpoints

| Método | Rota | Descrição |
|--------|------|-----------|
| POST | `/api/upload` | Upload de planilha (multipart/form-data) |
| POST | `/api/convert` | Converte colunas selecionadas `{columns: [...], divisor: 1000}` |
| GET | `/api/download` | Download do arquivo convertido |

## Tecnologias

- **Backend:** Python 3, Flask, Pandas, OpenPyXL
- **Frontend:** HTML5, CSS3 (dark theme), JavaScript vanilla
