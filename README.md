# ComexStat Pine Chemicals BR

Automação de inteligência de mercado para o setor de **Pine Chemicals** no Brasil, com coleta e análise de dados de exportação e importação via API do [ComexStat MDIC](https://comexstat.mdic.gov.br).

Desenvolvido para **RJ Comércio e Extração de Resinas Ltda / Alliance Resin Partners**.

---

## Notebooks

### 1. `DADOS_GERAIS_POR_NCM` — Monitor por NCM (8 dígitos)

Consulta a API geral do ComexStat e retorna dados por país de destino/origem, desagregados por NCM completo.

**Produtos monitorados:**

| NCM | Produto |
|---|---|
| 3806.10.00 | Colofônia / Gum Rosin / Breu |
| 3805.10.10 | Terebintina de Goma |
| 3805.10.90 | Outras Terebintinas |
| 3806.20.00 | Sais de ácidos resinosos |
| 3806.30.00 | Gomas éster |

**Saídas geradas:**
- Excel com aba de resumo, aba por NCM e aba de preço médio mensal
- Gráfico PNG de preço médio FOB (USD/kg) e volume exportado

---

### 2. `DADOS_POR_MUNICIPIO` — Monitor por Município + Fabricante

Consulta o endpoint `/cities` do ComexStat (nível SH4 — 4 dígitos) e cruza os dados com um mapa de **município → fabricante** para identificar o produtor por trás de cada embarque.

**Produtos monitorados (SH4):**

| SH4 | Produto |
|---|---|
| 1301 | Goma Resina |
| 3805 | Terebintina |
| 3806 | Breu e Derivados |

**Fabricantes mapeados** (amostra): Resinas Jardim, OCQ, Harima, RB/Resipim, AS Resinas, Pinus Brasil, Resineves, Nohva Pine, Meneghel, Enova, Ambar, entre outros.

**Saídas geradas:**
- Excel com aba de resumo, abas detalhadas por SH4 e ranking Top 50 municípios por FOB

---

## Estrutura

```
.
├── DADOS_GERAIS_POR_NCM/
│   ├── ComexStat_PineChemicals_Colab_7_.ipynb   # Notebook principal (NCM)
│   ├── ComexStat_VBA_Module.bas                  # Módulo VBA para Excel
│   └── ComexStat_PineChemicals.xlsm              # Planilha macro-enabled
│
└── DADOS_POR_MUNICIPIO/
    ├── Untitled5_corrigido_1_.ipynb              # Notebook principal (municípios)
    ├── ComexStat_Municipios_VBA(3).bas           # Módulo VBA para Excel
    └── ComexStat_Municipios.xlsm                 # Planilha macro-enabled
```

---

## Como usar

Os notebooks são projetados para rodar no **Google Colab**.

1. Acesse [colab.research.google.com](https://colab.research.google.com) e faça upload do `.ipynb` desejado
2. Na **Célula 2**, configure os parâmetros:
   ```python
   ANO_INICIAL = '2022'   # ano de início da série
   ANO_FINAL   = '2025'   # ano final
   FLUXO       = 'export' # 'export' ou 'import'
   ```
3. Execute `Runtime > Run all`
4. O Excel e o gráfico serão gerados e o download iniciado automaticamente

> A API do ComexStat aplica rate limit. Os notebooks já incluem retry automático com backoff exponencial e intervalo entre requisições.

---

## Fonte de dados

**ComexStat — MDIC (Ministério do Desenvolvimento, Indústria, Comércio e Serviços)**

- API geral: `POST https://api-comexstat.mdic.gov.br/general`
- API municípios: `POST https://api-comexstat.mdic.gov.br/cities`

---

## Dependências

```
requests
pandas
openpyxl
matplotlib
```

Instaladas automaticamente pela Célula 1 do notebook (`!pip install -q ...`).
