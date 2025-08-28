import os
import time
import random
import re
from urllib.parse import urlparse

import pandas as pd
import requests
from dotenv import load_dotenv

# Load API keys from .env
load_dotenv()
API_KEY = os.getenv("GOOGLE_API_KEY")
CSE_ID = os.getenv("GOOGLE_CSE_ID")

# -----------------------------
# Funciones de scoring
# -----------------------------
def generar_consultas_optimizadas(consulta_original: str):
    consulta_clean = consulta_original.strip()
    consultas = [
        f'"{consulta_clean}" official website',
        f'{consulta_clean} official site',
        f'{consulta_clean} homepage',
        f'{consulta_clean} company website',
        f'{consulta_clean} .com site:',
        f'{consulta_clean} software company',
        f'{consulta_clean} technology company'
    ]
    return consultas

def es_sitio_oficial(url: str, domain: str, title: str, snippet: str, consulta: str) -> int:
    score = 0
    consulta_lower = consulta.lower()
    domain_lower = domain.lower()
    title_lower = title.lower()
    snippet_lower = snippet.lower()
    
    consulta_base = re.sub(r'\b(software|hardware|inc|corp|ltd|llc|sa|srl|gmbh|ag)\b', '', consulta_lower).strip()
    palabras_consulta = [p for p in consulta_base.split() if len(p) > 2]

    for palabra in palabras_consulta:
        if palabra in domain_lower:
            score += 25
    if any(domain_lower.startswith(f"{palabra}.") or f".{palabra}." in domain_lower for palabra in palabras_consulta):
        score += 40
    if domain.endswith(('.com', '.net', '.org', '.io', '.tech')):
        score += 15
    social_platforms = [
        'facebook.com', 'twitter.com', 'linkedin.com', 'youtube.com', 'instagram.com',
        'wikipedia.org', 'crunchbase.com', 'bloomberg.com', 'reuters.com',
        'amazon.com', 'ebay.com', 'alibaba.com', 'github.com'
    ]
    for platform in social_platforms:
        if platform in domain_lower:
            score -= 30
            break
    official_words = ['official', 'homepage', 'home page', 'corporate', 'company']
    if any(word in title_lower for word in official_words):
        score += 10
    palabras_en_titulo = sum(1 for palabra in palabras_consulta if palabra in title_lower)
    score += palabras_en_titulo * 5
    subdomain_penalties = ['support.', 'help.', 'docs.', 'forum.', 'community.', 'blog.']
    if any(sub in domain_lower for sub in subdomain_penalties):
        score -= 10
    if url.count('/') > 3:
        score -= 5
    if url.startswith('https://'):
        score += 5
    return max(0, min(100, score))

def seleccionar_mejor_url_oficial(consulta: str, candidatos):
    if not candidatos:
        return None, "sin candidatos"
    scored_candidates = []
    for item in candidatos:
        url = item.get("href", "")
        title = item.get("title", "")
        snippet = item.get("snippet", "")
        domain = item.get("displayLink", "")
        if not url:
            continue
        score = es_sitio_oficial(url, domain, title, snippet, consulta)
        scored_candidates.append({"score": score, "url": url, "domain": domain})
    if not scored_candidates:
        return None, "sin candidatos válidos"
    scored_candidates.sort(key=lambda x: x['score'], reverse=True)
    best = scored_candidates[0]
    return best['url'], f"score {best['score']}, domain: {best['domain']}"

# -----------------------------
# Function for Google CSE
# -----------------------------
def buscar_con_google_cse_multiples(consultas):
    todos_candidatos = []
    urls_vistas = set()
    for query in consultas[:3]:  # hasta 3 consultas
        try:
            url = "https://www.googleapis.com/customsearch/v1"
            params = {
                'key': API_KEY,
                'cx': CSE_ID,
                'q': query,
                'num': 5,
                'safe': 'medium'
            }
            response = requests.get(url, params=params, timeout=15)
            if response.status_code == 200:
                data = response.json()
                for item in data.get('items', []):
                    href = item.get("link", "")
                    if href and href not in urls_vistas:
                        candidato = {
                            "title": item.get("title", ""),
                            "href": href,
                            "snippet": item.get("snippet", ""),
                            "displayLink": item.get("displayLink", "")
                        }
                        todos_candidatos.append(candidato)
                        urls_vistas.add(href)
            elif response.status_code == 429:
                print("Rate limit, esperando 30s...")
                time.sleep(30)
            time.sleep(1)
        except Exception as e:
            print(f"Error en consulta '{query}': {e}")
    return todos_candidatos

# -----------------------------
# Main flow
# -----------------------------
def main():
    input_excel = "./app/input.xlsx"
    output_excel = "./app/output_con_urls.xlsx"

    df = pd.read_excel(input_excel)
    df['url_oficial'] = None
    df['notas_busqueda'] = None

    for idx, row in df.iterrows():
        consulta = str(row[df.columns[0]]).strip()
        if not consulta:
            df.at[idx, 'notas_busqueda'] = "consulta vacía"
            continue

        print(f"\nBuscando sitio oficial de: {consulta}")
        consultas = generar_consultas_optimizadas(consulta)
        candidatos = buscar_con_google_cse_multiples(consultas)
        url, notas = seleccionar_mejor_url_oficial(consulta, candidatos)
        df.at[idx, 'url_oficial'] = url
        df.at[idx, 'notas_busqueda'] = notas
        print(f"→ {url} ({notas})")
        time.sleep(random.uniform(2,4))  # delay entre consultas

    df.to_excel(output_excel, index=False)
    print(f"\n✅ Resultados guardados en '{output_excel}'")

if __name__ == "__main__":
    main()
