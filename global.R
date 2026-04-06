#' @title Utilidades globales para la aplicacion de capacidad de proceso
#'
#' @description
#' Define helpers compartidos para cargar archivos de entrada, normalizar
#' valores opcionales, ejecutar analisis de capacidad de proceso con
#' `SixSigma` y exportar resultados.
#'
#' @keywords internal
NULL

library(shiny)

if (file.exists("openai_helpers.R")) {
  source("openai_helpers.R", local = TRUE)
}

#' Devuelve un valor por defecto cuando la entrada esta vacia
#'
#' @param x Valor a evaluar.
#' @param default Valor alternativo que se devuelve cuando `x` es `NULL`
#'   o una cadena vacia.
#'
#' @return El valor original o el valor por defecto.
or_default <- function(x, default) {
  if (is.null(x) || identical(x, "")) {
    default
  } else {
    x
  }
}

#' Lee archivos delimitados en texto plano
#'
#' @param path Ruta del archivo a leer.
#' @param header Indica si la primera fila contiene encabezados.
#' @param sep Separador de campos.
#' @param quote Caracter usado para comillas.
#' @param dec Separador decimal.
#'
#' @return Un `data.frame` con los datos cargados.
read_delimited_data <- function(path, header, sep, quote, dec) {
  utils::read.table(
    file = path,
    header = header,
    sep = sep,
    quote = quote,
    dec = dec,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

#' Lee datos de entrada desde CSV, TXT o Excel
#'
#' @param path Ruta del archivo a leer.
#' @param header Indica si la primera fila contiene encabezados.
#' @param sep Separador de campos para archivos delimitados.
#' @param quote Caracter usado para comillas en archivos delimitados.
#' @param dec Separador decimal para archivos delimitados.
#' @param sheet Hoja de Excel a leer cuando aplica.
#'
#' @return Un `data.frame` con los datos importados.
#' @export
read_input_data <- function(path, header, sep, quote, dec, sheet = NULL) {
  ext <- tolower(tools::file_ext(path))

  if (ext %in% c("xls", "xlsx")) {
    if (!requireNamespace("readxl", quietly = TRUE)) {
      stop(
        paste(
          "El paquete 'readxl' no esta instalado.",
          "Instalalo con install.packages('readxl')."
        ),
        call. = FALSE
      )
    }

    return(as.data.frame(readxl::read_excel(path = path, sheet = sheet)))
  }

  read_delimited_data(
    path = path,
    header = header,
    sep = sep,
    quote = quote,
    dec = dec
  )
}

#' Calcula limites de especificacion a partir de entradas Shiny
#'
#' @param input Lista reactiva de entradas Shiny.
#'
#' @return Una lista con `lsl` y `usl`.
get_spec_limits <- function(input) {
  list(
    lsl = if (isTRUE(input$use_lsl)) input$lsl else NA_real_,
    usl = if (isTRUE(input$use_usl)) input$usl else NA_real_
  )
}

#' Ejecuta el analisis de capacidad de proceso con `SixSigma`
#'
#' @param x Vector numerico con las mediciones.
#' @param input Lista reactiva de entradas Shiny.
#'
#' @return Una lista con resumen estadistico e indices de capacidad.
#' @export
run_process_capability <- function(x, input) {
  if (!requireNamespace("SixSigma", quietly = TRUE)) {
    stop(
      paste(
        "El paquete 'SixSigma' no esta instalado.",
        "Instalalo con install.packages('SixSigma')."
      ),
      call. = FALSE
    )
  }

  limits <- get_spec_limits(input)
  clean_x <- if (isTRUE(input$remove_na)) stats::na.omit(x) else x
  clean_x <- as.numeric(clean_x)

  if (!length(clean_x)) {
    stop("No hay datos numericos disponibles para el analisis.", call. = FALSE)
  }

  summary_table <- data.frame(
    Indicador = c("n", "media", "desv_est", "min", "max", "LSL", "USL"),
    Valor = c(
      length(clean_x),
      mean(clean_x, na.rm = TRUE),
      stats::sd(clean_x, na.rm = TRUE),
      min(clean_x, na.rm = TRUE),
      max(clean_x, na.rm = TRUE),
      limits$lsl,
      limits$usl
    ),
    check.names = FALSE,
    row.names = NULL
  )

  cp_value <- SixSigma::ss.ca.cp(
    x = clean_x,
    LSL = limits$lsl,
    USL = limits$usl,
    f.na.rm = isTRUE(input$remove_na),
    ci = FALSE,
    alpha = input$alpha
  )

  cpk_value <- SixSigma::ss.ca.cpk(
    x = clean_x,
    LSL = limits$lsl,
    USL = limits$usl,
    f.na.rm = isTRUE(input$remove_na),
    ci = FALSE,
    alpha = input$alpha
  )

  z_value <- SixSigma::ss.ca.z(
    x = clean_x,
    LSL = limits$lsl,
    USL = limits$usl,
    LT = isTRUE(input$long_term),
    f.na.rm = isTRUE(input$remove_na)
  )

  capability_rows <- list(
    data.frame(Indicador = "Cp", Valor = as.numeric(cp_value), check.names = FALSE),
    data.frame(Indicador = "Cpk", Valor = as.numeric(cpk_value), check.names = FALSE),
    data.frame(Indicador = if (isTRUE(input$long_term)) "Z (largo plazo)" else "Z", Valor = as.numeric(z_value), check.names = FALSE)
  )

  if (isTRUE(input$include_ci)) {
    cp_ci <- SixSigma::ss.ca.cp(
      x = clean_x,
      LSL = limits$lsl,
      USL = limits$usl,
      f.na.rm = isTRUE(input$remove_na),
      ci = TRUE,
      alpha = input$alpha
    )

    cpk_ci <- SixSigma::ss.ca.cpk(
      x = clean_x,
      LSL = limits$lsl,
      USL = limits$usl,
      f.na.rm = isTRUE(input$remove_na),
      ci = TRUE,
      alpha = input$alpha
    )

    capability_rows[[length(capability_rows) + 1]] <- data.frame(
      Indicador = c("Cp CI inferior", "Cp CI superior", "Cpk CI inferior", "Cpk CI superior"),
      Valor = c(cp_ci[[1]], cp_ci[[2]], cpk_ci[[1]], cpk_ci[[2]]),
      check.names = FALSE
    )
  }

  capability_table <- do.call(rbind, capability_rows)
  row.names(capability_table) <- NULL

  list(
    data = clean_x,
    limits = limits,
    summary = summary_table,
    capability = capability_table
  )
}

#' Convierte una tabla de resultados a formato exportable
#'
#' @param x Objeto tabular.
#'
#' @return Un `data.frame` listo para exportar a Excel o `NULL`.
as_export_table <- function(x) {
  if (is.null(x)) {
    return(NULL)
  }

  as.data.frame(x, check.names = FALSE)
}

#' Escribe un libro de Excel con resultados, grafico e interpretacion
#'
#' @param path Ruta del archivo de salida.
#' @param capability_result Resultado devuelto por `run_process_capability()`.
#' @param plot_path Ruta local del PNG con el grafico.
#' @param imr_plot_path Ruta local del PNG con el grafico IMR.
#' @param interpretation_text Texto de interpretacion generado por LLM.
#'
#' @return Invisiblemente, la ruta del archivo generado.
write_capability_export_workbook <- function(path, capability_result, plot_path, imr_plot_path = NULL, interpretation_text = NULL, study_plot_path = NULL) {
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop(
      paste(
        "El paquete 'openxlsx' no esta instalado.",
        "Instalalo con install.packages('openxlsx')."
      ),
      call. = FALSE
    )
  }

  wb <- openxlsx::createWorkbook()
  title_style <- openxlsx::createStyle(textDecoration = "bold", fgFill = "#DCE6F1")
  body_style <- openxlsx::createStyle(valign = "top", wrapText = TRUE)

  openxlsx::addWorksheet(wb, "Resumen")
  openxlsx::writeData(wb, "Resumen", x = as_export_table(capability_result$summary), startRow = 1, startCol = 1, rowNames = FALSE)
  openxlsx::addStyle(wb, "Resumen", title_style, rows = 1, cols = 1:ncol(capability_result$summary), gridExpand = TRUE, stack = TRUE)
  openxlsx::setColWidths(wb, "Resumen", cols = 1:ncol(capability_result$summary), widths = "auto")

  openxlsx::addWorksheet(wb, "Capacidad")
  openxlsx::writeData(wb, "Capacidad", x = as_export_table(capability_result$capability), startRow = 1, startCol = 1, rowNames = FALSE)
  openxlsx::addStyle(wb, "Capacidad", title_style, rows = 1, cols = 1:ncol(capability_result$capability), gridExpand = TRUE, stack = TRUE)
  openxlsx::setColWidths(wb, "Capacidad", cols = 1:ncol(capability_result$capability), widths = "auto")

  openxlsx::addWorksheet(wb, "Grafico")
  if (!is.null(plot_path) && file.exists(plot_path)) {
    openxlsx::insertImage(
      wb,
      sheet = "Grafico",
      file = plot_path,
      startRow = 2,
      startCol = 2,
      width = 8,
      height = 5,
      units = "in"
    )
  } else {
    openxlsx::writeData(
      wb,
      "Grafico",
      x = "No se pudo generar el grafico para la exportacion.",
      startRow = 2,
      startCol = 2
    )
  }

  if (!is.null(imr_plot_path) && file.exists(imr_plot_path)) {
    openxlsx::insertImage(
      wb,
      sheet = "Grafico",
      file = imr_plot_path,
      startRow = 30,
      startCol = 2,
      width = 8,
      height = 7.5,
      units = "in"
    )
  } else {
    openxlsx::writeData(
      wb,
      "Grafico",
      x = "No se pudo generar el grafico IMR para la exportacion.",
      startRow = 30,
      startCol = 2
    )
  }

  openxlsx::addWorksheet(wb, "Sixpack")
  if (!is.null(study_plot_path) && file.exists(study_plot_path)) {
    openxlsx::insertImage(
      wb,
      sheet = "Sixpack",
      file = study_plot_path,
      startRow = 2,
      startCol = 2,
      width = 10,
      height = 8,
      units = "in"
    )
  } else {
    openxlsx::writeData(
      wb,
      "Sixpack",
      x = "No se pudo generar el estudio grafico de capacidad para la exportacion.",
      startRow = 2,
      startCol = 2
    )
  }

  openxlsx::addWorksheet(wb, "Interpretacion")
  interpretation_value <- if (is.null(interpretation_text) || !nzchar(trimws(interpretation_text))) {
    "No se ha generado una interpretacion con LLM para este analisis."
  } else {
    interpretation_text
  }
  interpretation_df <- data.frame(Interpretacion = interpretation_value, check.names = FALSE)
  openxlsx::writeData(wb, "Interpretacion", x = interpretation_df, startRow = 1, startCol = 1, rowNames = FALSE)
  openxlsx::addStyle(wb, "Interpretacion", body_style, rows = 2, cols = 1, gridExpand = TRUE, stack = TRUE)
  openxlsx::setColWidths(wb, "Interpretacion", cols = 1, widths = 120)
  openxlsx::setRowHeights(wb, "Interpretacion", rows = 2, heights = 120)

  openxlsx::saveWorkbook(wb, path, overwrite = TRUE)
  invisible(path)
}
