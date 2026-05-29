# ⚡ Power BI Doc Fácil

Documentação técnica para arquivos `.pbix` **sem precisar abrir o Power BI Desktop**.  
Inspirado em [powerbidocfacil.com.br](https://www.powerbidocfacil.com.br).

---

## O que faz

- **Lê `.pbix` localmente** — extrai tabelas, colunas, medidas DAX, relacionamentos e fontes de dados
- **Gera um documento Word** completo e formatado, pronto para compartilhar
- **IA opcional** — pode usar Claude, OpenAI, Gemini, OpenRouter ou modelos locais (Ollama/LM Studio) para gerar um texto de visão geral
- **Análises interativas** — explore medidas, colunas e relacionamentos direto no browser

---

## Instalação rápida

```bash
# 1. Clone ou extraia os arquivos
cd pbi_doc_facil

# 2. Crie ambiente virtual (recomendado)
python -m venv .venv
source .venv/bin/activate        # Linux/Mac
# .venv\Scripts\activate         # Windows

# 3. Instale dependências
pip install -r requirements.txt

# 4. Execute
streamlit run app.py
```

O app abre automaticamente em `http://localhost:8501`.

---

## Estrutura

```
pbi_doc_facil/
├── app.py                  # Interface Streamlit
├── requirements.txt
├── README.md
└── core/
    ├── __init__.py
    ├── extractor.py        # Leitura do .pbix (ZIP + JSON)
    ├── docgen.py           # Geração do Word
    └── ai_providers.py     # Integração com provedores de IA
```

---

## Como funciona a extração

Um arquivo `.pbix` é um ZIP. Dentro dele existe um arquivo `DataModel` (formato ABF) ou `model.bim` / `DataModelSchema` (JSON puro). O app tenta ler o JSON diretamente.

**Para melhor compatibilidade:** salve o relatório como `.pbit` (Power BI Template) e faça upload do `.pbit` — ele sempre contém o JSON completo e legível.

---

## Provedores de IA suportados

| Provedor | Tipo | Modelo padrão |
|---|---|---|
| Claude (Anthropic) | `claude` | `claude-sonnet-4-20250514` |
| OpenAI | `openai` | `gpt-4o-mini` |
| Google Gemini | `gemini` | `gemini-2.0-flash` |
| OpenRouter | `openrouter` | `anthropic/claude-3-haiku` |
| Local (Ollama etc.) | `local` | `llama3` |

---

## Créditos

Inspirado no trabalho de **Lucas Ferraioli Curti** ([powerbidocfacil.com.br](https://www.powerbidocfacil.com.br)).
