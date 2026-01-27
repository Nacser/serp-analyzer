# SERP Analyzer - Herramienta de Linkbuilding

Aplicación para analizar resultados de búsqueda de Google y descubrir oportunidades de linkbuilding.

**Versión**: Playwright v2.0

## Características

- **Selector de país**: España, México, Argentina, Colombia, Chile, Perú, USA, UK, Alemania, Francia, Italia, Portugal, Brasil
- **Simulación móvil**: Viewport, user-agent y touch como dispositivo Android
- **Geolocalización**: Ubicación automática según país seleccionado
- **Detección de SERP Features refinada**: 
  - Google Ads
  - AI Overview
  - Featured Snippet
  - People Also Ask
  - Carrusel de imágenes
  - Carrusel de vídeos
  - Local Pack
  - Shopping
  - Knowledge Panel
  - Noticias
- **Agregación de dominios**: Ranking por frecuencia de aparición
- **Sistema de tags**: Organiza keywords por intención o categoría
- **Exportación a Excel (.xlsx)**: Con formato profesional y dos hojas
- **Interfaz mejorada**: Keywords desplegables, URLs clicables

## Instalación

```bash
# Descomprimir el proyecto
cd serp-analyzer

# Crear entorno virtual (recomendado)
python -m venv venv
source venv/bin/activate  # Linux/Mac
# o: venv\Scripts\activate  # Windows

# Instalar dependencias Python
pip install -r requirements.txt

# IMPORTANTE: Instalar navegador Chromium
python -m playwright install chromium
```

## Uso

```bash
python app.py
```

Abre http://localhost:5000 en tu navegador.

### Formato de keywords

Una keyword por línea. Opcionalmente puedes añadir un tag separado por `|`:

```
comprar zapatillas running | transaccional
mejores zapatillas running 2024 | informacional
zapatillas running baratas
```

### Configuración

- **País**: Selecciona el país para la búsqueda (afecta dominio, idioma y geolocalización)
- **Delay mínimo/máximo**: Segundos entre peticiones (recomendado: 5-8+)

## Estructura del proyecto

```
serp-analyzer/
├── app.py                 # Backend Flask + Playwright
├── requirements.txt       # Dependencias Python
├── README.md             
└── templates/
    └── index.html         # Frontend
```

## Notas importantes

1. **Primera ejecución**: Playwright descargará Chromium (~150MB)
2. **Delays**: Usa delays generosos (5-8+ segundos) para evitar bloqueos
3. **Volumen**: Para análisis masivos (100+ keywords), considera usar una API de SERP

## Licencia

Uso interno / educativo.
