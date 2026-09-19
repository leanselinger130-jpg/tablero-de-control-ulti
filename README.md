# Centro de control · Cesim Global Challenge

Panel del equipo, con datos compartidos vía Netlify Blobs.

## Deploy

```bash
npm install
npm install -g netlify-cli
netlify login
netlify deploy --prod
```

La primera vez elegí "Create & configure a new project" y ponele un nombre.
Te va a dar una URL tipo `https://tu-equipo.netlify.app` — ese es el link para
compartir con el equipo y el coach. No hace falta login para usarlo.

**Importante:** como esta es una reconstrucción completa del código, si ya
tenían un sitio publicado antes, corran `netlify link` primero (dentro de la
carpeta del proyecto) para conectarse al MISMO sitio existente antes de hacer
`deploy --prod` — así no pierden la URL que ya compartieron, y como los datos
viven en Netlify Blobs (no en el código), van a seguir viendo las mismas
rondas, plan y premisas que ya cargaron.

## Activar el Asistente IA (opcional)

Necesita su propia clave de la API de Anthropic:

1. Creen cuenta en https://console.anthropic.com y generen una API key.
2. En Netlify: Project configuration → Environment variables → agreguen
   `ANTHROPIC_API_KEY` con esa clave.
3. `netlify deploy --prod` de nuevo.

## Qué incluye

- **General**: KPIs, ranking entre los 7 equipos, alertas de riesgo
  financiero y de desvío estratégico, margen sobre ventas, notas de
  decisiones, exportar a Excel/PDF.
- **Premisas**: repositorio de reglas/supuestos del profesor o de Cesim.
- **Competencia**: mapa de posicionamiento, mix de tecnología, distancia
  estratégica y rivales más cercanos, vista de todo el mercado.
- **Panel dinámico**: buscador de cualquier indicador del reporte, comparado
  entre los equipos que elijan.
- **Estrategia**: plan objetivo por región (precio, enfoque de marketing,
  cuota) + objetivos de compañía (I+D, apalancamiento), con chequeo
  automático de alineación cada ronda.
- **Análisis Financiero**: análisis vertical (cada línea como % de ingresos o
  de activos totales) y horizontal (variación % ronda a ronda) del estado de
  resultados y el balance.
- **Asistente IA**: preguntas en lenguaje natural sobre los datos, el plan y
  las premisas cargadas.
- Pequeños íconos "i" clickeables junto a los títulos explican qué muestra
  cada sección, sin ocupar espacio permanente en la pantalla.
