# Process Capability Shiny App

![R](https://img.shields.io/badge/R-4.5+-276DC3?logo=r&logoColor=white)
![Shiny](https://img.shields.io/badge/Shiny-App-75AADB?logo=rstudioide&logoColor=white)
![bslib](https://img.shields.io/badge/UI-bslib-4B5563)
![SixSigma](https://img.shields.io/badge/Capability-SixSigma-0F766E)
![OpenAI](https://img.shields.io/badge/AI-OpenAI-412991?logo=openai&logoColor=white)
![Excel](https://img.shields.io/badge/Input-Excel%20%2F%20CSV-217346?logo=microsoft-excel&logoColor=white)

Aplicacion Shiny para ejecutar analisis de capacidad de proceso con `SixSigma::ss.ca.cp`, `SixSigma::ss.ca.cpk`, `SixSigma::ss.ca.z` y `SixSigma::ss.study.ca` a partir de archivos CSV o Excel, visualizar indicadores clave, exportarlos a Excel y generar una interpretacion ejecutiva opcional con OpenAI.

## Que hace este proyecto

Este proyecto resuelve un problema practico frecuente en calidad: tener datos de una caracteristica medida en Excel o CSV y necesitar calcular la capacidad del proceso sin preparar manualmente el dataset en R.

La app:

- carga archivos de entrada desde interfaz grafica
- detecta automaticamente columnas numericas candidatas a medicion
- permite seleccionar una o varias columnas de medicion comparables
- ejecuta indicadores de capacidad de proceso con `SixSigma`
- muestra resumen estadistico e indices Cp, Cpk y Z
- genera un histograma con media y limites de especificacion
- genera un estudio grafico de capacidad tipo `Sixpack` con `SixSigma::ss.study.ca`
- exporta un archivo Excel con hojas separadas para resumen, capacidad, grafico, sixpack e interpretacion
- puede enviar el resultado numerico, el histograma y el sixpack a OpenAI para obtener una interpretacion resumida

En otras palabras, funciona como una capa de trabajo operativo sobre funciones de capacidad de proceso de `SixSigma`, pensada para reducir friccion al analizar datos reales desde archivos.

## Flujo del analisis

1. El usuario carga un archivo `.csv`, `.txt`, `.xls` o `.xlsx`.
2. La app inspecciona los tipos de columnas y propone automaticamente las numericas como mediciones.
3. El usuario define:
   - una o varias columnas de medicion
   - LSL y/o USL
   - Target
   - si desea intervalos de confianza
4. La app valida que la estructura sea compatible con el analisis.
5. Se ejecutan `SixSigma::ss.ca.cp`, `SixSigma::ss.ca.cpk`, `SixSigma::ss.ca.z` y `SixSigma::ss.study.ca`.
6. Los resultados se presentan en pestanas separadas:
   - vista previa de datos
   - resultados estadisticos
   - histograma
   - sixpack
   - interpretacion opcional con OpenAI
7. El usuario puede exportar un archivo Excel consolidado desde la pestana `Resultados`.

## Como trata multiples columnas de medicion

Si el archivo contiene mas de una columna numerica de medicion, la app permite seleccionarlas juntas.

Cuando se seleccionan varias columnas, la app concatena internamente esas mediciones en una sola serie numerica antes de calcular la capacidad. Este comportamiento es util cuando las columnas representan mediciones comparables de la misma caracteristica. Si representan caracteristicas distintas, no deberian mezclarse en un solo analisis.

## Funcionalidades

- carga de archivos `.csv`, `.txt`, `.xls` y `.xlsx`
- deteccion automatica de columnas numericas de medicion
- soporte para una o multiples columnas de medicion
- calculo de Cp, Cpk y Z
- opcion de calcular intervalos de confianza para Cp y Cpk
- parametro `Target` para el estudio grafico de capacidad
- histograma con media y limites de especificacion
- estudio grafico de capacidad tipo `Sixpack`
- exportacion del grafico en PNG
- exportacion de resultados a Excel con hojas para resumen, capacidad, grafico, sixpack e interpretacion LLM
- interpretacion opcional con OpenAI, solo bajo solicitud del usuario

## Que muestra el resultado

### Salida tecnica

- resumen de datos
- Cp
- Cpk
- Z
- intervalos de confianza opcionales para Cp y Cpk
- histograma en PNG
- estudio grafico de capacidad tipo sixpack en PNG
- archivo Excel exportable con:
  - hoja `Resumen`
  - hoja `Capacidad`
  - hoja `Grafico`
  - hoja `Sixpack`
  - hoja `Interpretacion`

### Salida interpretativa

De forma opcional, el usuario puede pedir una interpretacion con OpenAI. Esa interpretacion:

- no se ejecuta automaticamente
- consume tokens solo cuando el usuario pulsa el boton
- envia a OpenAI el resultado numerico del analisis
- envia tambien el histograma y el sixpack generados por la app
- devuelve un resumen corto orientado a decision

## Estructura

- `global.R`: utilidades compartidas y exportacion
- `ui.R`: interfaz Shiny
- `server.R`: logica reactiva y analisis
- `openai_helpers.R`: integracion con OpenAI
- `gage_rr_sample.xlsx`: archivo legado del repositorio

## Requisitos

Instala estos paquetes en R:

```r
install.packages(c(
  "shiny",
  "bslib",
  "readxl",
  "SixSigma",
  "httr2",
  "jsonlite",
  "base64enc",
  "openxlsx"
))
```

## Como ejecutar

Desde R en la carpeta del proyecto:

```r
shiny::runApp()
```

## OpenAI opcional

Si quieres usar la interpretacion automatica, configura:

- `OPENAI_API_KEY`
- opcionalmente `OPENAI_MODEL`

Puedes partir del archivo `.Renviron.example`.
