#' @title Interfaz de usuario para la aplicacion de capacidad de proceso
#'
#' @description
#' Define la interfaz Shiny basada en `bslib` para cargar archivos, configurar
#' un analisis de capacidad de proceso, revisar indicadores y solicitar una
#' interpretacion automatizada del resultado.
#'
#' @keywords internal
NULL

library(shiny)
library(bslib)

#' Interfaz principal de la aplicacion Shiny
#'
#' @description
#' Objeto de interfaz que organiza la captura de parametros, la visualizacion
#' de resultados del analisis y el panel de interpretacion asistida.
#'
#' @format Un objeto `shiny.tag.list`.
#' @export
ui <- page_sidebar(
  title = "Capacidad de proceso con SixSigma",
  tags$head(
    tags$style(HTML("
      .plot-toolbar {
        display: flex;
        align-items: center;
        gap: 0.75rem;
        margin-bottom: 1rem;
      }

      .copy-status {
        color: #4b5563;
        font-size: 0.9rem;
      }

      #capability_plot img,
      #study_plot img {
        display: block;
        width: 100%;
        max-width: 100%;
        height: auto;
        object-fit: contain;
      }
    ")),
    tags$script(HTML("
      window.copyPlotToClipboard = async function(imageId, statusId) {
        const image = document.querySelector(`#${imageId} img`);
        const status = document.getElementById(statusId);

        if (!image) {
          if (status) status.textContent = 'Genera el grafico antes de copiarlo.';
          return;
        }

        if (!navigator.clipboard || typeof window.ClipboardItem === 'undefined') {
          if (status) status.textContent = 'El navegador no permite copiar imagenes.';
          return;
        }

        try {
          if (!image.complete) {
            await new Promise((resolve, reject) => {
              image.addEventListener('load', resolve, { once: true });
              image.addEventListener('error', reject, { once: true });
            });
          }

          const width = image.naturalWidth || image.width;
          const height = image.naturalHeight || image.height;

          if (!width || !height) {
            throw new Error('La imagen no tiene dimensiones validas.');
          }

          const canvas = document.createElement('canvas');
          canvas.width = width;
          canvas.height = height;

          const context = canvas.getContext('2d');
          context.drawImage(image, 0, 0, width, height);

          const blob = await new Promise((resolve, reject) => {
            canvas.toBlob(function(result) {
              if (result) {
                resolve(result);
              } else {
                reject(new Error('No se pudo generar la imagen.'));
              }
            }, 'image/png');
          });

          await navigator.clipboard.write([new ClipboardItem({ [blob.type]: blob })]);

          if (status) status.textContent = 'Imagen copiada al portapapeles.';
        } catch (error) {
          if (status) status.textContent = 'No se pudo copiar la imagen. Revisa permisos del navegador.';
        }
      };
    "))
  ),
  sidebar = sidebar(
    width = 340,
    h4("Datos"),
    fileInput("file", "Archivo", accept = c(".csv", ".txt", ".xls", ".xlsx")),
    checkboxInput("header", "La primera fila es encabezado", TRUE),
    uiOutput("sheet_control"),
    conditionalPanel(
      condition = "input.file && !/\\.(xls|xlsx)$/i.test(input.file.name || '')",
      selectInput(
        "sep",
        "Separador",
        choices = c("Coma" = ",", "Punto y coma" = ";", "Tab" = "\t"),
        selected = ","
      )
    ),
    conditionalPanel(
      condition = "input.file && !/\\.(xls|xlsx)$/i.test(input.file.name || '')",
      selectInput(
        "dec",
        "Separador decimal",
        choices = c("Punto" = ".", "Coma" = ","),
        selected = "."
      )
    ),
    conditionalPanel(
      condition = "input.file && !/\\.(xls|xlsx)$/i.test(input.file.name || '')",
      selectInput(
        "quote",
        "Comillas",
        choices = c('Doble comilla' = '"', "Simple comilla" = "'", "Ninguna" = ""),
        selected = '"'
      )
    ),
    tags$hr(),
    h4("Columnas"),
    uiOutput("column_controls"),
    tags$hr(),
    h4("Especificaciones"),
    checkboxInput("use_lsl", "Usar LSL", TRUE),
    conditionalPanel(
      condition = "input.use_lsl",
      numericInput("lsl", "LSL", value = 9.5, step = 0.01)
    ),
    checkboxInput("use_usl", "Usar USL", TRUE),
    conditionalPanel(
      condition = "input.use_usl",
      numericInput("usl", "USL", value = 10.5, step = 0.01)
    ),
    numericInput("target", "Target", value = 10.0, step = 0.01),
    tags$hr(),
    h4("Parametros"),
    checkboxInput("remove_na", "Excluir NA", TRUE),
    checkboxInput("long_term", "Aplicar ajuste largo plazo (Z - 1.5)", FALSE),
    checkboxInput("include_ci", "Calcular intervalos de confianza", FALSE),
    numericInput("alpha", "Alpha", value = 0.05, min = 0.001, max = 0.2, step = 0.001),
    numericInput("digits", "Decimales", value = 4, min = 0, step = 1),
    textInput("plot_title", "Titulo del grafico", "Histograma de capacidad de proceso"),
    textInput("plot_subtitle", "Subtitulo", ""),
    actionButton("run_analysis", "Ejecutar analisis", class = "btn-primary")
  ),
  card(
    full_screen = TRUE,
    uiOutput("active_measurement_control"),
    tags$div(
      style = "padding: 1rem 1rem 0 1rem;",
      downloadButton("download_excel", "Exportar resultados Excel")
    ),
    navset_card_tab(
      nav_panel(
        "Vista previa",
        br(),
        uiOutput("data_status"),
        tableOutput("data_preview")
      ),
      nav_panel(
        "Resultados",
        br(),
        verbatimTextOutput("analysis_log"),
        h4("Resumen"),
        tableOutput("summary_table"),
        h4("Indices de capacidad"),
        tableOutput("capability_table")
      ),
      nav_panel(
        "Graficos",
        br(),
        tags$div(
          class = "plot-toolbar",
          tags$button(
            type = "button",
            class = "btn btn-outline-secondary",
            onclick = "window.copyPlotToClipboard('capability_plot', 'capability_plot_copy_status')",
            "Copiar al portapapeles"
          ),
          tags$span(id = "capability_plot_copy_status", class = "copy-status")
        ),
        imageOutput("capability_plot", width = "100%")
      ),
      nav_panel(
        "Sixpack",
        br(),
        tags$div(
          class = "plot-toolbar",
          tags$button(
            type = "button",
            class = "btn btn-outline-secondary",
            onclick = "window.copyPlotToClipboard('study_plot', 'study_plot_copy_status')",
            "Copiar al portapapeles"
          ),
          tags$span(id = "study_plot_copy_status", class = "copy-status")
        ),
        imageOutput("study_plot", width = "100%")
      ),
      nav_panel(
        "Interpretacion",
        br(),
        p("La interpretacion con OpenAI solo se ejecuta cuando pulses el boton."),
        textAreaInput(
          "interpretation_instructions",
          "Instrucciones adicionales",
          placeholder = "Ejemplo: enfocate en riesgo de incumplimiento de especificaciones y capacidad real del proceso.",
          rows = 4,
          width = "100%"
        ),
        tags$div(
          style = "margin-bottom: 1rem;",
          actionButton("run_interpretation", "Generar interpretacion", class = "btn-primary")
        ),
        verbatimTextOutput("interpretation_status"),
        uiOutput("interpretation_text")
      ),
      nav_panel(
        "Ayuda",
        br(),
        p("La app ejecuta las funciones ", code("SixSigma::ss.ca.cp"), ", ", code("SixSigma::ss.ca.cpk"), ", ", code("SixSigma::ss.ca.z"), " y ", code("SixSigma::ss.study.ca"), "."),
        p("Necesitas un archivo con al menos una columna numerica de medicion."),
        tags$ul(
          tags$li("puedes seleccionar una o varias columnas numericas comparables"),
          tags$li("si eliges varias columnas, la app las concatena en una sola serie"),
          tags$li("debes definir al menos un limite de especificacion: LSL o USL")
        ),
        p("Si el paquete no esta instalado, en R usa: ", code("install.packages('SixSigma')"))
      )
    )
  )
)
