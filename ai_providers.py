"""
Camada de abstração para provedores de IA.
Suporta: Claude (Anthropic), OpenAI, Gemini (Google), OpenRouter, Local (Ollama/LM Studio).
"""

import json
from typing import Optional
from .extractor import PBIXModel


SAMPLE_LIMIT = 5  # tabelas por amostra


def _build_metadata_text(model: PBIXModel, full_context: bool = False) -> str:
    tables = model.tables if full_context else model.tables[:SAMPLE_LIMIT]
    lines = [f"Arquivo: {model.file_name}", ""]

    for t in tables:
        hidden = " (oculta)" if t.is_hidden else ""
        lines.append(f"TABELA: {t.name}{hidden} | Fonte: {t.source_type}")
        if t.description:
            lines.append(f"  Descrição: {t.description}")
        col_names = [c.name for c in t.columns[:20]]
        if col_names:
            lines.append(f"  Colunas ({len(t.columns)}): {', '.join(col_names)}" +
                         (" ..." if len(t.columns) > 20 else ""))
        for m in t.measures[:10]:
            expr_preview = m.expression[:120].replace("\n", " ") if m.expression else ""
            lines.append(f"  [MEDIDA] {m.name}: {expr_preview}")
        lines.append("")

    if model.relationships:
        lines.append("RELACIONAMENTOS:")
        for r in model.relationships[:30]:
            lines.append(f"  {r.from_table}[{r.from_column}] → {r.to_table}[{r.to_column}] ({r.cardinality})")

    if model.data_sources:
        lines.append("\nFONTES DE DADOS:")
        for ds in model.data_sources:
            lines.append(f"  {ds.name} ({ds.kind})")

    return "\n".join(lines)


def _make_prompt(metadata: str, extra_instructions: str = "") -> str:
    extra = f"\n\nInstruções adicionais: {extra_instructions}" if extra_instructions else ""
    return f"""Você é um especialista em Power BI e documentação de dados.

Com base nos metadados extraídos de um arquivo .pbix abaixo, escreva um texto de visão geral (overview) em português do Brasil que descreva:
1. O propósito provável do relatório
2. As principais tabelas e o que representam
3. Medidas e cálculos relevantes
4. Fontes de dados identificadas
5. Considerações sobre o modelo de dados

Seja objetivo, técnico e útil. Use parágrafos, não listas.{extra}

---
{metadata}
---

Escreva apenas o texto de overview, sem cabeçalhos ou formatação Markdown."""


def generate_ai_overview(
    model: PBIXModel,
    provider: str,
    api_key: str,
    model_name: str = "",
    full_context: bool = False,
    extra_instructions: str = "",
    base_url: str = "",
) -> str:
    metadata = _build_metadata_text(model, full_context)
    prompt = _make_prompt(metadata, extra_instructions)

    try:
        if provider == "claude":
            return _call_claude(api_key, prompt, model_name or "claude-sonnet-4-20250514")
        elif provider == "openai":
            return _call_openai(api_key, prompt, model_name or "gpt-4o-mini")
        elif provider == "gemini":
            return _call_gemini(api_key, prompt, model_name or "gemini-2.0-flash")
        elif provider == "openrouter":
            return _call_openrouter(api_key, prompt, model_name or "anthropic/claude-3-haiku")
        elif provider == "local":
            return _call_local(base_url or "http://localhost:11434", prompt, model_name or "llama3")
        else:
            return f"[Provedor '{provider}' não reconhecido]"
    except Exception as e:
        return f"[Erro ao chamar IA: {e}]"


def _call_claude(api_key: str, prompt: str, model_name: str) -> str:
    import anthropic
    client = anthropic.Anthropic(api_key=api_key)
    msg = client.messages.create(
        model=model_name,
        max_tokens=2000,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


def _call_openai(api_key: str, prompt: str, model_name: str) -> str:
    from openai import OpenAI
    client = OpenAI(api_key=api_key)
    resp = client.chat.completions.create(
        model=model_name,
        messages=[{"role": "user", "content": prompt}],
        max_tokens=2000,
    )
    return resp.choices[0].message.content


def _call_gemini(api_key: str, prompt: str, model_name: str) -> str:
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    gen_model = genai.GenerativeModel(model_name)
    response = gen_model.generate_content(prompt)
    return response.text


def _call_openrouter(api_key: str, prompt: str, model_name: str) -> str:
    import requests
    resp = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        json={
            "model": model_name,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 2000,
        },
        timeout=60,
    )
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]


def _call_local(base_url: str, prompt: str, model_name: str) -> str:
    import requests
    url = base_url.rstrip("/") + "/api/chat"
    resp = requests.post(
        url,
        json={
            "model": model_name,
            "messages": [{"role": "user", "content": prompt}],
            "stream": False,
        },
        timeout=120,
    )
    resp.raise_for_status()
    data = resp.json()
    return data.get("message", {}).get("content", data.get("response", ""))
