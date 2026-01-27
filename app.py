"""
SERP Analyzer para Linkbuilding - Versión Playwright v2
Con selector de país, exportación XLSX y detección SERP.
"""

from flask import Flask, render_template, request, jsonify, send_file
import time
import random
import re
import io
from urllib.parse import urlparse, unquote, quote_plus
from datetime import datetime
from collections import defaultdict
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from playwright.sync_api import sync_playwright

app = Flask(__name__)

# Configuración de dispositivo móvil
MOBILE_CONFIG = {
    "viewport": {"width": 412, "height": 915},
    "user_agent": "Mozilla/5.0 (Linux; Android 13; SM-S918B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Mobile Safari/537.36",
    "device_scale_factor": 2.625,
    "is_mobile": True,
    "has_touch": True,
}

# Configuración de países
COUNTRIES = {
    "es": {
        "name": "España",
        "domain": "google.es",
        "gl": "es",
        "hl": "es",
        "geo": {"latitude": 40.4168, "longitude": -3.7038},
        "timezone": "Europe/Madrid"
    },
    "mx": {
        "name": "México",
        "domain": "google.com.mx",
        "gl": "mx",
        "hl": "es",
        "geo": {"latitude": 19.4326, "longitude": -99.1332},
        "timezone": "America/Mexico_City"
    },
    "ar": {
        "name": "Argentina",
        "domain": "google.com.ar",
        "gl": "ar",
        "hl": "es",
        "geo": {"latitude": -34.6037, "longitude": -58.3816},
        "timezone": "America/Argentina/Buenos_Aires"
    },
    "co": {
        "name": "Colombia",
        "domain": "google.com.co",
        "gl": "co",
        "hl": "es",
        "geo": {"latitude": 4.7110, "longitude": -74.0721},
        "timezone": "America/Bogota"
    },
    "cl": {
        "name": "Chile",
        "domain": "google.cl",
        "gl": "cl",
        "hl": "es",
        "geo": {"latitude": -33.4489, "longitude": -70.6693},
        "timezone": "America/Santiago"
    },
    "pe": {
        "name": "Perú",
        "domain": "google.com.pe",
        "gl": "pe",
        "hl": "es",
        "geo": {"latitude": -12.0464, "longitude": -77.0428},
        "timezone": "America/Lima"
    },
    "us": {
        "name": "Estados Unidos",
        "domain": "google.com",
        "gl": "us",
        "hl": "en",
        "geo": {"latitude": 40.7128, "longitude": -74.0060},
        "timezone": "America/New_York"
    },
    "uk": {
        "name": "Reino Unido",
        "domain": "google.co.uk",
        "gl": "uk",
        "hl": "en",
        "geo": {"latitude": 51.5074, "longitude": -0.1278},
        "timezone": "Europe/London"
    },
    "de": {
        "name": "Alemania",
        "domain": "google.de",
        "gl": "de",
        "hl": "de",
        "geo": {"latitude": 52.5200, "longitude": 13.4050},
        "timezone": "Europe/Berlin"
    },
    "fr": {
        "name": "Francia",
        "domain": "google.fr",
        "gl": "fr",
        "hl": "fr",
        "geo": {"latitude": 48.8566, "longitude": 2.3522},
        "timezone": "Europe/Paris"
    },
    "it": {
        "name": "Italia",
        "domain": "google.it",
        "gl": "it",
        "hl": "it",
        "geo": {"latitude": 41.9028, "longitude": 12.4964},
        "timezone": "Europe/Rome"
    },
    "pt": {
        "name": "Portugal",
        "domain": "google.pt",
        "gl": "pt",
        "hl": "pt",
        "geo": {"latitude": 38.7223, "longitude": -9.1393},
        "timezone": "Europe/Lisbon"
    },
    "br": {
        "name": "Brasil",
        "domain": "google.com.br",
        "gl": "br",
        "hl": "pt",
        "geo": {"latitude": -23.5505, "longitude": -46.6333},
        "timezone": "America/Sao_Paulo"
    },
}


def extract_domain(url):
    """Extrae el dominio de una URL"""
    try:
        parsed = urlparse(url)
        domain = parsed.netloc
        if domain.startswith("www."):
            domain = domain[4:]
        return domain
    except:
        return None


def clean_url(url):
    """Limpia y decodifica una URL de Google"""
    if not url:
        return None
    
    url = unquote(url)
    
    if "/url?q=" in url:
        match = re.search(r'/url\?q=([^&]+)', url)
        if match:
            url = unquote(match.group(1))
    
    if url.startswith("/search") or url.startswith("/"):
        return None
    if "google." in url and "/search" in url:
        return None
        
    return url


def detect_serp_features(soup):
    """Detecta las features presentes en la SERP."""
    features = {
        "google_ads": False,
        "ai_overview": False,
        "featured_snippet": False,
        "people_also_ask": False,
        "image_carousel": False,
        "video_carousel": False,
        "local_pack": False,
        "shopping_results": False,
        "knowledge_panel": False,
        "news": False,
        "other_features": []
    }
    
    html_str = str(soup).lower()
    
    # Anuncios
    if soup.find("span", string=re.compile(r"^(Patrocinado|Sponsored|Anuncio)$", re.I)) or soup.find("div", {"data-text-ad": True}):
        features["google_ads"] = True
    
    # AI Overview
    if soup.find("div", {"data-async-context-required": "aobs"}) or "ai overview" in html_str or "generado por ia" in html_str:
        features["ai_overview"] = True
    
    # Featured Snippet
    if soup.find("div", {"class": "xpdopen"}):
        features["featured_snippet"] = True
    
    # People Also Ask
    if soup.find("div", {"jsname": "Cpkphb"}) or "preguntas relacionadas" in html_str or "people also ask" in html_str:
        features["people_also_ask"] = True
    
    # Imágenes
    if soup.find("div", {"id": "iur"}) or soup.find("g-scrolling-carousel"):
        features["image_carousel"] = True
    
    # Videos
    if soup.find("video-voyager") or "youtube.com/watch" in html_str:
        features["video_carousel"] = True
    
    # Local Pack
    if soup.find("div", {"class": re.compile(r"VkpGBb")}) or soup.find("a", {"href": re.compile(r"maps\.google")}):
        features["local_pack"] = True
    
    # Shopping
    if soup.find("div", {"class": re.compile(r"commercial-unit|cu-container")}) or "google.es/shopping" in html_str or "google.com/shopping" in html_str:
        features["shopping_results"] = True
    
    # Knowledge Panel
    if soup.find("div", {"class": re.compile(r"kp-wholepage|liYKde")}):
        features["knowledge_panel"] = True
    
    # Noticias
    if soup.find("div", {"class": re.compile(r"ftSUBd|So9e7d")}) or "noticias principales" in html_str or "top stories" in html_str:
        features["news"] = True
    
    return features


def extract_organic_results(soup):
    """Extrae los resultados orgánicos del HTML - versión actualizada para nuevo formato Google."""
    organic_results = []
    seen_urls = set()
    position = 0
    
    print("[DEBUG] Iniciando extracción de resultados orgánicos...")
    
    # ESTRATEGIA 1: Buscar enlaces con clase rTyHce (formato móvil actual de Google)
    result_links = soup.find_all("a", class_="rTyHce")
    print(f"[DEBUG] Enlaces rTyHce encontrados: {len(result_links)}")
    
    for link in result_links:
        if position >= 10:
            break
        
        href = link.get("href", "")
        if not href:
            continue
        
        url = clean_url(href)
        if not url or url in seen_urls:
            continue
        
        domain = extract_domain(url)
        if not domain or "google." in domain:
            continue
        
        # Verificar que no es un anuncio
        is_ad = False
        parent = link
        for _ in range(8):
            parent = parent.find_parent()
            if not parent:
                break
            parent_text = parent.get_text()
            if any(ad_text in parent_text for ad_text in ["Patrocinado", "Sponsored", "Anuncio", "Resultado patrocinado"]):
                is_ad = True
                break
            if parent.get("data-text-ad") or parent.get("data-ae"):
                is_ad = True
                break
        
        if is_ad:
            continue
        
        # Buscar título dentro del enlace - nuevo formato usa role="heading"
        title = None
        
        # Buscar div con role="heading" dentro del enlace
        heading_div = link.find("div", attrs={"role": "heading"})
        if heading_div:
            # El título está en un span dentro
            title_span = heading_div.find("span")
            if title_span:
                title = title_span.get_text(strip=True)
            else:
                title = heading_div.get_text(strip=True)
        
        # Alternativa: buscar clase específica de títulos
        if not title:
            title_div = link.find("div", class_=re.compile(r"F0FGWb|v7jaNc|MBeuO"))
            if title_div:
                title = title_div.get_text(strip=True)
        
        # Alternativa: buscar h3 tradicional
        if not title:
            h3 = link.find("h3")
            if h3:
                title = h3.get_text(strip=True)
        
        if not title or len(title) < 5:
            continue
        
        seen_urls.add(url)
        position += 1
        
        organic_results.append({
            "position": position,
            "title": title,
            "url": url,
            "domain": domain
        })
        print(f"[DEBUG] Resultado {position}: {domain} - {title[:50]}...")
    
    # ESTRATEGIA 2: Si no encontramos con rTyHce, buscar en contenedores MjjYud
    if len(organic_results) == 0:
        print("[DEBUG] Intentando estrategia alternativa con MjjYud...")
        containers = soup.find_all("div", class_="MjjYud")
        print(f"[DEBUG] Contenedores MjjYud encontrados: {len(containers)}")
        
        for container in containers:
            if position >= 10:
                break
            
            # Buscar enlace principal
            link = container.find("a", href=True)
            if not link:
                continue
            
            href = link.get("href", "")
            url = clean_url(href)
            if not url or url in seen_urls:
                continue
            
            domain = extract_domain(url)
            if not domain or "google." in domain:
                continue
            
            # Buscar título
            title = None
            heading = container.find(attrs={"role": "heading"})
            if heading:
                title = heading.get_text(strip=True)
            
            if not title:
                h3 = container.find("h3")
                if h3:
                    title = h3.get_text(strip=True)
            
            if not title or len(title) < 5:
                continue
            
            # Verificar que no es anuncio
            container_text = container.get_text()
            if any(ad in container_text for ad in ["Patrocinado", "Sponsored", "Anuncio"]):
                continue
            
            seen_urls.add(url)
            position += 1
            
            organic_results.append({
                "position": position,
                "title": title,
                "url": url,
                "domain": domain
            })
            print(f"[DEBUG] Resultado {position}: {domain} - {title[:50]}...")
    
    print(f"[DEBUG] Total resultados extraídos: {len(organic_results)}")
    return organic_results


def scrape_google_serp_playwright(keyword, country_code="es", delay_min=4, delay_max=7):
    """Realiza scraping de la SERP de Google usando Playwright."""
    country = COUNTRIES.get(country_code, COUNTRIES["es"])
    
    results = {
        "keyword": keyword,
        "country": country["name"],
        "timestamp": datetime.now().isoformat(),
        "organic_results": [],
        "serp_features": {},
        "error": None
    }
    
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(
                headless=False,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--disable-dev-shm-usage",
                    "--no-sandbox",
                ]
            )
            
            context = browser.new_context(
                viewport=MOBILE_CONFIG["viewport"],
                user_agent=MOBILE_CONFIG["user_agent"],
                device_scale_factor=MOBILE_CONFIG["device_scale_factor"],
                is_mobile=MOBILE_CONFIG["is_mobile"],
                has_touch=MOBILE_CONFIG["has_touch"],
                geolocation=country["geo"],
                permissions=["geolocation"],
                locale=f"{country['hl']}-{country['gl'].upper()}",
                timezone_id=country["timezone"],
            )
            
            page = context.new_page()
            page.set_default_timeout(30000)
            
            encoded_keyword = quote_plus(keyword)
            search_url = f"https://www.{country['domain']}/search?q={encoded_keyword}&hl={country['hl']}&gl={country['gl']}"
            
            delay = random.uniform(delay_min, delay_max)
            time.sleep(delay)
            
            page.goto(search_url, wait_until="networkidle")
            page.wait_for_timeout(2000)
            
            # Scroll
            page.evaluate("window.scrollBy(0, 500)")
            page.wait_for_timeout(500)
            page.evaluate("window.scrollBy(0, -200)")
            page.wait_for_timeout(500)
            
            html_content = page.content()
            
            # Guardar HTML para debug
            with open("debug_serp.html", "w", encoding="utf-8") as f:
                f.write(html_content)
            print(f"[DEBUG] HTML guardado en debug_serp.html ({len(html_content)} chars)")
            
            if "captcha" in html_content.lower() or "unusual traffic" in html_content.lower():
                results["error"] = "CAPTCHA detectado."
                browser.close()
                return results
            
            soup = BeautifulSoup(html_content, "html.parser")
            
            results["serp_features"] = detect_serp_features(soup)
            results["organic_results"] = extract_organic_results(soup)
            
            print(f"[DEBUG] Resultados extraídos: {len(results['organic_results'])}")
            
            if not results["organic_results"]:
                results["error"] = "No se encontraron resultados orgánicos."
            
            browser.close()
    
    except Exception as e:
        results["error"] = f"Error: {str(e)}"
        import traceback
        traceback.print_exc()
    
    return results


def aggregate_domains(all_results):
    """Agrega los dominios de todos los resultados."""
    domain_stats = defaultdict(lambda: {
        "appearances": 0,
        "keywords": [],
        "positions": [],
        "urls": set()
    })
    
    for result in all_results:
        if result.get("error") and not result.get("organic_results"):
            continue
        
        keyword = result["keyword"]
        
        for organic in result.get("organic_results", []):
            domain = organic["domain"]
            domain_stats[domain]["appearances"] += 1
            domain_stats[domain]["keywords"].append(keyword)
            domain_stats[domain]["positions"].append(organic["position"])
            domain_stats[domain]["urls"].add(organic["url"])
    
    aggregated = []
    for domain, stats in domain_stats.items():
        avg_position = sum(stats["positions"]) / len(stats["positions"]) if stats["positions"] else 0
        aggregated.append({
            "domain": domain,
            "appearances": stats["appearances"],
            "avg_position": round(avg_position, 2),
            "keywords": list(set(stats["keywords"])),
            "unique_urls": len(stats["urls"]),
            "urls": list(stats["urls"])
        })
    
    aggregated.sort(key=lambda x: (-x["appearances"], x["avg_position"]))
    return aggregated


def parse_keywords_input(text):
    """Parsea el input de keywords."""
    keywords = []
    lines = text.strip().split("\n")
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        if "|" in line:
            parts = line.split("|", 1)
            keyword = parts[0].strip()
            tag = parts[1].strip() if len(parts) > 1 else ""
        else:
            keyword = line
            tag = ""
        
        if keyword:
            keywords.append({"keyword": keyword, "tag": tag})
    
    return keywords


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/countries", methods=["GET"])
def get_countries():
    countries_list = [{"code": code, "name": data["name"]} for code, data in COUNTRIES.items()]
    return jsonify(countries_list)


@app.route("/api/analyze", methods=["POST"])
def analyze():
    data = request.get_json()
    
    if not data or "keywords" not in data:
        return jsonify({"error": "No se proporcionaron keywords"}), 400
    
    keywords_input = data.get("keywords", "")
    country_code = data.get("country", "es")
    delay_min = data.get("delay_min", 4)
    delay_max = data.get("delay_max", 7)
    
    keywords_list = parse_keywords_input(keywords_input)
    
    if not keywords_list:
        return jsonify({"error": "No se encontraron keywords válidas"}), 400
    
    all_results = []
    
    for kw_data in keywords_list:
        keyword = kw_data["keyword"]
        tag = kw_data["tag"]
        
        result = scrape_google_serp_playwright(
            keyword,
            country_code=country_code,
            delay_min=delay_min,
            delay_max=delay_max
        )
        result["tag"] = tag
        all_results.append(result)
    
    domain_aggregation = aggregate_domains(all_results)
    
    return jsonify({
        "success": True,
        "results": all_results,
        "domain_aggregation": domain_aggregation,
        "total_keywords": len(keywords_list),
        "keywords_with_errors": sum(1 for r in all_results if r.get("error"))
    })


@app.route("/api/export/xlsx", methods=["POST"])
def export_xlsx():
    data = request.get_json()
    
    if not data:
        return jsonify({"error": "No hay datos para exportar"}), 400
    
    results = data.get("results", [])
    domain_aggregation = data.get("domain_aggregation", [])
    
    wb = Workbook()
    
    # Hoja 1: Resultados
    ws1 = wb.active
    ws1.title = "Resultados por Keyword"
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="2E7D32")
    alt_fill = PatternFill("solid", fgColor="F5F5F5")
    border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    
    headers1 = ["Keyword", "Tag", "País", "Posición", "Título", "URL", "Dominio",
                "Ads", "AI Overview", "Featured Snippet", "PAA", "Imágenes",
                "Vídeos", "Local Pack", "Shopping", "Knowledge Panel", "Noticias", "Error"]
    
    for col, header in enumerate(headers1, 1):
        cell = ws1.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
    
    row_num = 2
    for result in results:
        keyword = result.get("keyword", "")
        tag = result.get("tag", "")
        country = result.get("country", "")
        error = result.get("error", "")
        features = result.get("serp_features", {})
        organic = result.get("organic_results", [])
        
        if not organic:
            ws1.cell(row=row_num, column=1, value=keyword)
            ws1.cell(row=row_num, column=2, value=tag)
            ws1.cell(row=row_num, column=3, value=country)
            ws1.cell(row=row_num, column=8, value="Sí" if features.get("google_ads") else "No")
            ws1.cell(row=row_num, column=9, value="Sí" if features.get("ai_overview") else "No")
            ws1.cell(row=row_num, column=10, value="Sí" if features.get("featured_snippet") else "No")
            ws1.cell(row=row_num, column=11, value="Sí" if features.get("people_also_ask") else "No")
            ws1.cell(row=row_num, column=12, value="Sí" if features.get("image_carousel") else "No")
            ws1.cell(row=row_num, column=13, value="Sí" if features.get("video_carousel") else "No")
            ws1.cell(row=row_num, column=14, value="Sí" if features.get("local_pack") else "No")
            ws1.cell(row=row_num, column=15, value="Sí" if features.get("shopping_results") else "No")
            ws1.cell(row=row_num, column=16, value="Sí" if features.get("knowledge_panel") else "No")
            ws1.cell(row=row_num, column=17, value="Sí" if features.get("news") else "No")
            ws1.cell(row=row_num, column=18, value=error)
            row_num += 1
        else:
            for i, org in enumerate(organic):
                ws1.cell(row=row_num, column=1, value=keyword if i == 0 else "")
                ws1.cell(row=row_num, column=2, value=tag if i == 0 else "")
                ws1.cell(row=row_num, column=3, value=country if i == 0 else "")
                ws1.cell(row=row_num, column=4, value=org.get("position", ""))
                ws1.cell(row=row_num, column=5, value=org.get("title", ""))
                ws1.cell(row=row_num, column=6, value=org.get("url", ""))
                ws1.cell(row=row_num, column=7, value=org.get("domain", ""))
                
                if i == 0:
                    ws1.cell(row=row_num, column=8, value="Sí" if features.get("google_ads") else "No")
                    ws1.cell(row=row_num, column=9, value="Sí" if features.get("ai_overview") else "No")
                    ws1.cell(row=row_num, column=10, value="Sí" if features.get("featured_snippet") else "No")
                    ws1.cell(row=row_num, column=11, value="Sí" if features.get("people_also_ask") else "No")
                    ws1.cell(row=row_num, column=12, value="Sí" if features.get("image_carousel") else "No")
                    ws1.cell(row=row_num, column=13, value="Sí" if features.get("video_carousel") else "No")
                    ws1.cell(row=row_num, column=14, value="Sí" if features.get("local_pack") else "No")
                    ws1.cell(row=row_num, column=15, value="Sí" if features.get("shopping_results") else "No")
                    ws1.cell(row=row_num, column=16, value="Sí" if features.get("knowledge_panel") else "No")
                    ws1.cell(row=row_num, column=17, value="Sí" if features.get("news") else "No")
                    ws1.cell(row=row_num, column=18, value=error)
                
                for col in range(1, 19):
                    cell = ws1.cell(row=row_num, column=col)
                    cell.border = border
                    if row_num % 2 == 0:
                        cell.fill = alt_fill
                
                row_num += 1
    
    col_widths = [30, 15, 12, 8, 50, 60, 25, 6, 10, 14, 6, 8, 8, 10, 8, 14, 8, 30]
    for i, width in enumerate(col_widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = width
    
    # Hoja 2: Dominios
    ws2 = wb.create_sheet("Dominios Agregados")
    
    headers2 = ["#", "Dominio", "Apariciones", "Posición Media", "URLs Únicas", "Keywords"]
    
    for col, header in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
    
    for idx, dom in enumerate(domain_aggregation, 1):
        row = idx + 1
        ws2.cell(row=row, column=1, value=idx)
        ws2.cell(row=row, column=2, value=dom.get("domain", ""))
        ws2.cell(row=row, column=3, value=dom.get("appearances", 0))
        ws2.cell(row=row, column=4, value=dom.get("avg_position", 0))
        ws2.cell(row=row, column=5, value=dom.get("unique_urls", 0))
        ws2.cell(row=row, column=6, value=", ".join(dom.get("keywords", [])))
        
        for col in range(1, 7):
            cell = ws2.cell(row=row, column=col)
            cell.border = border
            if row % 2 == 0:
                cell.fill = alt_fill
    
    col_widths2 = [5, 35, 12, 14, 12, 80]
    for i, width in enumerate(col_widths2, 1):
        ws2.column_dimensions[get_column_letter(i)].width = width
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"serp_analysis_{timestamp}.xlsx"
    
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )


if __name__ == "__main__":
    print("\n" + "="*60)
    print("  SERP Analyzer - Versión Playwright v2")
    print("  Abre http://localhost:5000 en tu navegador")
    print("="*60 + "\n")
    app.run(debug=True, host="0.0.0.0", port=5000)
