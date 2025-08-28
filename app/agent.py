import os
import time
import argparse
from typing import TypedDict, List, Dict, Any, Optional

import pandas as pd
from dotenv import load_dotenv
from duckduckgo_search import DDGS
from langgraph.graph import StateGraph, END

# OpenAI es opcional: si tienes OPENAI_API_KEY, lo usamos para desambiguar
LLM_AVAILABLE = False
try:
    from langchain_openai import ChatOpenAI
    if os.getenv("OPENAI_API_KEY"):
        LLM_AVAILABLE = True
except Exception:
    LLM_AVAILABLE = False


# -----------------------------
# Definición de estado LangGraph
# -----------------------------
class Row(TypedDict, total=False):
    consulta: str
    url: Optional[str]
    candidatos: List[Dict[str, str]]
    notas: Optional[str]

class AgentState(TypedDict, total=False):
    rows: List[Row]
    i: int
    input_path: str
    query_column: str
    output_path: str


# -----------------------------
# Nodos del grafo
# -----------------------------
def cargar_excel(state: AgentState) -> AgentState:
    df = pd.read_excel(state["input_path"])
    col = state["query_column"]
    if col not in df.columns:
        raise ValueError(
            f"La columna '{col}' no existe en {state['input_path']}. "
            f"Columnas disponibles: {list(df.columns)}"
        )

    # Normalizamos filas
    rows: List[Row] = []
    for val in df[col].fillna(""):
        consulta = str(val).strip()
        rows.append({"consulta": consulta})

    state["rows"] = rows
    state["i"] = 0
    return state


def buscar_web(state: AgentState) -> AgentState:
    idx = state["i"]
    fila = state["rows"][idx]
    consulta = fila.get("consulta", "")

    if not consulta:
        fila["url"] = None
        fila["notas"] = "consulta vacía"
        return state

    # DuckDuckGo (resultados en español; puedes ajustar 'region')
    with DDGS() as ddg:
        resultados = list(ddg.text(
            consulta,
            region="en-uk",
            safesearch="moderate",
            max_results=8,
            backend="html"
        ))

    candidatos = []
    for r in resultados:
        # La lib retorna 'title','href','body'
        href = r.get("href")
        if href:
            candidatos.append({
                "title": r.get("title", ""),
                "href": href,
                "snippet": r.get("body", "")
            })

    fila["candidatos"] = candidatos

    # Heurística rápida: nos quedamos con el 1º candidato si no hay LLM
    fila["url"] = candidatos[0]["href"] if candidatos else None
    if not candidatos:
        fila["notas"] = "sin resultados"

    # Respetar un pequeño delay para no ser agresivos con el buscador
    time.sleep(3)
    return state


def desambiguar_con_llm(state: AgentState) -> AgentState:
    if not LLM_AVAILABLE:
        return state

    idx = state["i"]
    fila = state["rows"][idx]
    consulta = fila.get("consulta", "")
    candidatos = fila.get("candidatos", [])

    if not candidatos:
        return state

    # Pedimos al LLM elegir un único URL oficial
    llm = ChatOpenAI(model="gpt-4o-mini", temperature=0)
    opciones = "\n".join(f"- {c['href']} :: {c['title']}" for c in candidatos[:6])
    prompt = (
        "Eres un asistente que selecciona el sitio web oficial de una entidad.\n"
        f"Consulta: {consulta}\n"
        "Candidatos (URL :: título):\n"
        f"{opciones}\n\n"
        "Devuelve SOLO el URL oficial más probable. Si dudas, elige el dominio primario "
        "(no redes sociales ni páginas internas). Responde con una única línea que sea el URL."
    )
    try:
        resp = llm.invoke(prompt).content.strip()
        # Tomamos el primer token que parece URL
        fila["url"] = resp.split()[0]
        fila["notas"] = "url seleccionada por LLM"
    except Exception as e:
        fila["notas"] = f"fallback heurístico (error LLM: {e})"
        # ya quedó la heurística del nodo anterior
    return state


def avanzar(state: AgentState) -> AgentState:
    state["i"] += 1
    return state


def deberia_continuar(state: AgentState) -> str:
    return "buscar" if state["i"] < len(state["rows"]) else "escribir"


def escribir_salida(state: AgentState) -> AgentState:
    df = pd.read_excel(state["input_path"])
    urls = [r.get("url") for r in state["rows"]]
    notas = [r.get("notas") for r in state["rows"]]
    df["url_encontrada"] = urls
    df["notas"] = notas
    os.makedirs(os.path.dirname(state["output_path"]) or ".", exist_ok=True)
    df.to_excel(state["output_path"], index=False)
    print(f"✅ Archivo generado: {state['output_path']}")
    return state


# -----------------------------
# Construcción del grafo
# -----------------------------
def construir_grafo():
    graph = StateGraph(AgentState)

    graph.add_node("cargar", cargar_excel)
    graph.add_node("buscar", buscar_web)
    graph.add_node("desambiguar", desambiguar_con_llm)
    graph.add_node("avanzar", avanzar)
    graph.add_node("escribir", escribir_salida)

    graph.set_entry_point("cargar")

    # Al terminar de cargar, si hay filas -> buscar, si no -> escribir
    def after_load(state: AgentState) -> str:
        return "buscar" if state.get("rows") else "escribir"

    graph.add_conditional_edges("cargar", after_load, {
        "buscar": "buscar",
        "escribir": "escribir",
    })

    # buscar -> desambiguar -> avanzar
    graph.add_edge("buscar", "desambiguar")
    graph.add_edge("desambiguar", "avanzar")

    # bucle mientras haya filas
    graph.add_conditional_edges("avanzar", deberia_continuar, {
        "buscar": "buscar",
        "escribir": "escribir",
    })

    graph.add_edge("escribir", END)

    return graph.compile()


def run(input_path: str, query_column: str, output_path: str):
    app = construir_grafo()
    initial_state: AgentState = {
        "input_path": input_path,
        "query_column": query_column,
        "output_path": output_path
    }
    app.invoke(initial_state)


if __name__ == "__main__":
    load_dotenv()

    parser = argparse.ArgumentParser(description="Agente LangGraph: Excel -> URL oficial")
    parser.add_argument("--input", default=os.getenv("INPUT_EXCEL", "app/entrada.xlsx"))
    parser.add_argument("--query-column", default=os.getenv("QUERY_COLUMN", "consulta"))
    parser.add_argument("--output", default=os.getenv("OUTPUT_EXCEL", "app/salida_urls.xlsx"))
    args = parser.parse_args()

    run(args.input, args.query_column, args.output)
