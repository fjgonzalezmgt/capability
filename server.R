#' @title Logica del servidor para la aplicacion de capacidad de proceso
#'
#' @description
#' Implementa el flujo reactivo de la aplicacion Shiny: carga de datos,
#' seleccion de columnas, ejecucion del analisis de capacidad, generacion de
#' histograma y consulta opcional a OpenAI para interpretar resultados.
#'
#' @keywords internal
NULL

#' Logica del servidor principal de la aplicacion
#'
#' @param input Entradas reactivas de Shiny.
#' @param output Salidas reactivas de Shiny.
#' @param session Sesion activa de Shiny.
#'
#' @return No devuelve valor. Registra reactividad y salidas en la sesion.
#' @export
server <- function(input, output, session) {
  capability_plot_dims <- list(width = 2400, height = 1600)
  imr_plot_dims <- list(width = 2400, height = 2200)
  study_plot_dims <- list(width = 2600, height = 2200)

  #' Prepara una o varias columnas de medicion para capacidad de proceso
  #'
  #' @param df `data.frame` fuente.
  #' @param measurement_cols Nombres de columnas numericas de medicion.
  #'
  #' @return Una lista con vector numerico consolidado y etiqueta.
  prepare_capability_input <- function(df, measurement_cols) {
    measurement_cols <- unique(measurement_cols)

    if (length(measurement_cols) == 1) {
      return(list(
        values = as.numeric(df[[measurement_cols[[1]]]]),
        label = measurement_cols[[1]]
      ))
    }

    values <- unlist(lapply(measurement_cols, function(col_name) df[[col_name]]), use.names = FALSE)

    list(
      values = as.numeric(values),
      label = paste(measurement_cols, collapse = " + ")
    )
  }

  #' Genera un histograma de capacidad en un archivo PNG
  #'
  #' @param path Ruta del archivo de salida.
  #' @param width Ancho del PNG en pixeles.
  #' @param height Alto del PNG en pixeles.
  #' @param res Resolucion del PNG.
  #'
  #' @return Invisiblemente, la ruta del archivo generado.
  build_capability_plot <- function(
    path,
    width = capability_plot_dims$width,
    height = capability_plot_dims$height,
    res = 200
  ) {
    req(analysis_result())

    result <- current_result()
    x <- result$data
    limits <- result$limits
    legend_labels <- c("Media")
    legend_colors <- c("#1d4ed8")
    legend_types <- c(3)
    x_bounds <- c(x, limits$lsl, limits$usl)
    x_bounds <- x_bounds[is.finite(x_bounds)]
    x_range <- range(x_bounds, na.rm = TRUE)
    x_span <- diff(x_range)
    if (!is.finite(x_span) || x_span <= 0) {
      x_span <- max(abs(x_range[1]), 1)
    }
    x_padding <- x_span * 0.08
    x_limits <- c(x_range[1] - x_padding, x_range[2] + x_padding)
    hist_info <- graphics::hist(x, breaks = "FD", plot = FALSE)
    hist_max <- max(hist_info$counts)

    grDevices::png(filename = path, width = width, height = height, res = res)
    on.exit(grDevices::dev.off(), add = TRUE)
    old_par <- graphics::par(no.readonly = TRUE)
    on.exit(graphics::par(old_par), add = TRUE)

    hist_info <- graphics::hist(
      x,
      breaks = hist_info$breaks,
      col = "#93c5fd",
      border = "white",
      main = or_default(input$plot_title, "Histograma de capacidad de proceso"),
      sub = or_default(input$plot_subtitle, ""),
      xlab = analysis_result()$label,
      xlim = x_limits,
      ylim = c(0, hist_max * 1.18)
    )

    graphics::abline(v = mean(x, na.rm = TRUE), col = "#1d4ed8", lwd = 2, lty = 3)

    if (!is.na(limits$lsl)) {
      graphics::abline(v = limits$lsl, col = "#dc2626", lwd = 2, lty = 2)
      legend_labels <- c(legend_labels, "LSL")
      legend_colors <- c(legend_colors, "#dc2626")
      legend_types <- c(legend_types, 2)
    }

    if (!is.na(limits$usl)) {
      graphics::abline(v = limits$usl, col = "#dc2626", lwd = 2, lty = 2)
      legend_labels <- c(legend_labels, "USL")
      legend_colors <- c(legend_colors, "#dc2626")
      legend_types <- c(legend_types, 2)
    }

    graphics::legend(
      "topright",
      legend = legend_labels,
      col = legend_colors,
      lwd = 2,
      lty = legend_types,
      bty = "n"
    )

    box_stats <- grDevices::boxplot.stats(x)$stats
    box_outliers <- grDevices::boxplot.stats(x)$out
    box_center <- hist_max * 1.08
    box_height <- max(hist_max * 0.08, 0.6)
    whisker_height <- box_height * 0.45

    graphics::segments(
      x0 = box_stats[1],
      y0 = box_center,
      x1 = box_stats[2],
      y1 = box_center,
      col = "#2563eb",
      lwd = 2,
      xpd = NA
    )
    graphics::segments(
      x0 = box_stats[4],
      y0 = box_center,
      x1 = box_stats[5],
      y1 = box_center,
      col = "#2563eb",
      lwd = 2,
      xpd = NA
    )
    graphics::segments(
      x0 = box_stats[1],
      y0 = box_center - whisker_height / 2,
      x1 = box_stats[1],
      y1 = box_center + whisker_height / 2,
      col = "#2563eb",
      lwd = 2,
      xpd = NA
    )
    graphics::segments(
      x0 = box_stats[5],
      y0 = box_center - whisker_height / 2,
      x1 = box_stats[5],
      y1 = box_center + whisker_height / 2,
      col = "#2563eb",
      lwd = 2,
      xpd = NA
    )
    graphics::rect(
      xleft = box_stats[2],
      ybottom = box_center - box_height / 2,
      xright = box_stats[4],
      ytop = box_center + box_height / 2,
      col = "#bfdbfe",
      border = "#2563eb",
      lwd = 2,
      xpd = NA
    )
    graphics::segments(
      x0 = box_stats[3],
      y0 = box_center - box_height / 2,
      x1 = box_stats[3],
      y1 = box_center + box_height / 2,
      col = "#1d4ed8",
      lwd = 2,
      xpd = NA
    )

    graphics::segments(
      x0 = mean(x, na.rm = TRUE),
      y0 = box_center - box_height / 2,
      x1 = mean(x, na.rm = TRUE),
      y1 = box_center + box_height / 2,
      col = "#1d4ed8",
      lwd = 2,
      lty = 3,
      xpd = NA
    )

    if (length(box_outliers) > 0) {
      graphics::points(
        x = box_outliers,
        y = rep(box_center, length(box_outliers)),
        pch = 16,
        cex = 0.8,
        col = "#2563eb",
        xpd = NA
      )
    }

    invisible(path)
  }

  #' Genera el estudio grafico de capacidad tipo sixpack en un archivo PNG
  #'
  #' @param path Ruta del archivo de salida.
  #' @param width Ancho del PNG en pixeles.
  #' @param height Alto del PNG en pixeles.
  #' @param res Resolucion del PNG.
  #'
  #' @return Invisiblemente, la ruta del archivo generado.
  build_study_plot <- function(
    path,
    width = study_plot_dims$width,
    height = study_plot_dims$height,
    res = 200
  ) {
    req(analysis_result())

    result <- current_result()
    x <- result$data
    limits <- result$limits

    grDevices::png(filename = path, width = width, height = height, res = res)
    on.exit(grDevices::dev.off(), add = TRUE)

    SixSigma::ss.study.ca(
      xST = x,
      LSL = limits$lsl,
      USL = limits$usl,
      Target = input$target,
      alpha = input$alpha,
      f.na.rm = isTRUE(input$remove_na),
      f.main = or_default(input$plot_title, "Histograma de capacidad de proceso"),
      f.sub = or_default(input$plot_subtitle, "")
    )

    invisible(path)
  }

  #' Genera un grafico IMR (Individuals and Moving Range) en un archivo PNG
  #'
  #' @param path Ruta del archivo de salida.
  #' @param width Ancho del PNG en pixeles.
  #' @param height Alto del PNG en pixeles.
  #' @param res Resolucion del PNG.
  #'
  #' @return Invisiblemente, la ruta del archivo generado.
  build_imr_plot <- function(
    path,
    width = imr_plot_dims$width,
    height = imr_plot_dims$height,
    res = 200
  ) {
    req(analysis_result())

    result <- current_result()
    x <- result$data
    validate(
      need(length(x) >= 2, "Se requieren al menos dos mediciones para el grafico IMR.")
    )

    mr <- abs(diff(x))
    d2 <- SixSigma::ss.cc.getd2(2)
    d3 <- SixSigma::ss.cc.getd3(2)

    i_center <- mean(x, na.rm = TRUE)
    sigma_est <- mean(mr, na.rm = TRUE) / d2
    i_lcl <- i_center - 3 * sigma_est
    i_ucl <- i_center + 3 * sigma_est

    mr_center <- mean(mr, na.rm = TRUE)
    mr_lcl <- mr_center * (1 - 3 * (d3 / d2))
    mr_ucl <- mr_center * (1 + 3 * (d3 / d2))
    mr_lcl <- max(0, mr_lcl)

    i_index <- seq_along(x)
    mr_index <- seq_along(mr) + 1
    i_out <- x < i_lcl | x > i_ucl
    mr_out <- mr < mr_lcl | mr > mr_ucl

    grDevices::png(filename = path, width = width, height = height, res = res)
    on.exit(grDevices::dev.off(), add = TRUE)
    old_par <- graphics::par(no.readonly = TRUE)
    on.exit(graphics::par(old_par), add = TRUE)

    graphics::par(mfrow = c(2, 1), mar = c(4.5, 4.5, 3, 1.5))

    graphics::plot(
      i_index,
      x,
      type = "b",
      pch = 16,
      col = "#2563eb",
      xlab = "Observacion",
      ylab = analysis_result()$label,
      main = "IMR: Individuals",
      ylim = range(c(x, i_lcl, i_center, i_ucl), na.rm = TRUE)
    )
    graphics::abline(h = i_center, col = "#1d4ed8", lwd = 2, lty = 3)
    graphics::abline(h = i_lcl, col = "#dc2626", lwd = 2, lty = 2)
    graphics::abline(h = i_ucl, col = "#dc2626", lwd = 2, lty = 2)
    if (any(i_out, na.rm = TRUE)) {
      graphics::points(i_index[i_out], x[i_out], pch = 16, cex = 1.1, col = "#dc2626")
    }
    graphics::legend(
      "topright",
      legend = c("CL", "LCL/UCL", "Fuera de control"),
      col = c("#1d4ed8", "#dc2626", "#dc2626"),
      lwd = c(2, 2, NA),
      lty = c(3, 2, NA),
      pch = c(NA, NA, 16),
      bty = "n"
    )

    graphics::plot(
      mr_index,
      mr,
      type = "b",
      pch = 16,
      col = "#2563eb",
      xlab = "Observacion",
      ylab = "Moving Range",
      main = "IMR: Moving Range",
      ylim = range(c(mr, mr_lcl, mr_center, mr_ucl), na.rm = TRUE)
    )
    graphics::abline(h = mr_center, col = "#1d4ed8", lwd = 2, lty = 3)
    graphics::abline(h = mr_lcl, col = "#dc2626", lwd = 2, lty = 2)
    graphics::abline(h = mr_ucl, col = "#dc2626", lwd = 2, lty = 2)
    if (any(mr_out, na.rm = TRUE)) {
      graphics::points(mr_index[mr_out], mr[mr_out], pch = 16, cex = 1.1, col = "#dc2626")
    }

    invisible(path)
  }

  is_excel_file <- reactive({
    req(input$file)
    tolower(tools::file_ext(input$file$name)) %in% c("xls", "xlsx")
  })

  output$sheet_control <- renderUI({
    req(input$file)
    if (!is_excel_file()) {
      return(NULL)
    }

    if (!requireNamespace("readxl", quietly = TRUE)) {
      return(
        p("Para leer Excel instala el paquete ", code("readxl"), ".")
      )
    }

    sheets <- readxl::excel_sheets(input$file$datapath)
    selectInput("sheet", "Hoja", choices = sheets, selected = sheets[[1]])
  })

  data_reactive <- reactive({
    req(input$file)
    read_input_data(
      path = input$file$datapath,
      header = input$header,
      sep = or_default(input$sep, ","),
      quote = or_default(input$quote, '"'),
      dec = or_default(input$dec, "."),
      sheet = if (is_excel_file()) input$sheet else NULL
    )
  })

  output$column_controls <- renderUI({
    req(data_reactive())
    df <- data_reactive()
    cols <- names(df)
    numeric_cols <- cols[vapply(df, is.numeric, logical(1))]
    default_vars <- if (length(numeric_cols) > 0) numeric_cols else cols[[1]]

    labeled_cols <- vapply(
      cols,
      function(col) {
        cls <- class(df[[col]])[[1]]
        sprintf("%s (%s)", col, cls)
      },
      character(1)
    )
    choice_values <- stats::setNames(cols, labeled_cols)

    tagList(
      selectizeInput(
        "var_cols",
        "Columnas de medicion",
        choices = choice_values,
        selected = default_vars,
        multiple = TRUE,
        options = list(plugins = list("remove_button"))
      )
    )
  })

  prepared_analysis_input <- reactive({
    req(data_reactive(), input$var_cols)
    prepare_capability_input(
      df = data_reactive(),
      measurement_cols = input$var_cols
    )
  })

  analysis_result <- eventReactive(input$run_analysis, {
    req(data_reactive(), input$var_cols)

    df <- data_reactive()
    measurement_cols <- input$var_cols

    validate(
      need(length(measurement_cols) >= 1, "Selecciona al menos una columna de medicion."),
      need(all(measurement_cols %in% names(df)), "Las columnas seleccionadas no existen en el archivo."),
      need(
        all(vapply(df[measurement_cols], is.numeric, logical(1))),
        "Todas las columnas de medicion deben ser numericas."
      ),
      need(isTRUE(input$use_lsl) || isTRUE(input$use_usl), "Debes definir al menos un limite de especificacion."),
      need(!isTRUE(input$use_lsl) || !isTRUE(input$use_usl) || isTRUE(input$usl > input$lsl), "USL debe ser mayor que LSL."),
      need(is.finite(input$target), "Target debe ser numerico."),
      need(input$alpha > 0 && input$alpha < 1, "Alpha debe estar entre 0 y 1."),
      need(sum(vapply(df[measurement_cols], function(x) sum(!is.na(x)), numeric(1))) > 1, "Se requieren al menos dos mediciones no vacias.")
    )

    prepared <- prepare_capability_input(
      df = df,
      measurement_cols = measurement_cols
    )

    captured <- NULL
    result <- NULL

    captured <- capture.output({
      result <- run_process_capability(prepared$values, input)
    })

    list(
      result = result,
      log = paste(captured, collapse = "\n"),
      label = prepared$label,
      measurement_count = length(measurement_cols)
    )
  })

  output$active_measurement_control <- renderUI({
    req(analysis_result())

    tags$div(
      style = "padding: 1rem 1rem 0 1rem;",
      tags$div(
        style = "font-size: 0.95rem; color: #4b5563;",
        if (analysis_result()$measurement_count > 1) {
          sprintf(
            "Analisis concatenado de %s columnas de medicion: %s",
            analysis_result()$measurement_count,
            analysis_result()$label
          )
        } else {
          sprintf("Columna de medicion: %s", analysis_result()$label)
        }
      )
    )
  })

  current_result <- reactive({
    req(analysis_result())
    analysis_result()$result
  })

  current_log <- reactive({
    req(analysis_result())
    analysis_result()$log
  })

  interpretation_result <- eventReactive(input$run_interpretation, {
    req(current_result(), analysis_result())

    tryCatch(
      {
        plot_file <- tempfile(fileext = ".png")
        build_capability_plot(plot_file)
        imr_plot_file <- tempfile(fileext = ".png")
        build_imr_plot(imr_plot_file)
        study_plot_file <- tempfile(fileext = ".png")
        build_study_plot(study_plot_file)

        withProgress(message = "Consultando OpenAI", value = 0.2, {
          incProgress(0.4)
          text <- openai_interpret_capability(
            capability_result = current_result(),
            plot_path = plot_file,
            imr_plot_path = imr_plot_file,
            study_plot_path = study_plot_file,
            language = "es",
            extra_instructions = paste(
              "Contexto del analisis:",
              analysis_result()$label,
              if (nzchar(or_default(input$interpretation_instructions, ""))) input$interpretation_instructions else ""
            )
          )
          incProgress(0.4)
          list(ok = TRUE, text = text)
        })
      },
      error = function(e) {
        list(ok = FALSE, text = conditionMessage(e))
      }
    )
  })

  output$data_status <- renderUI({
    if (is.null(input$file)) {
      return(p("Carga un archivo CSV o Excel para habilitar el analisis."))
    }

    df <- data_reactive()
    tagList(
      p(sprintf("Filas: %s", nrow(df))),
      p(sprintf("Columnas: %s", ncol(df)))
    )
  })

  output$data_preview <- renderTable({
    req(data_reactive())
    utils::head(data_reactive(), 12)
  }, rownames = TRUE)

  output$analysis_log <- renderText({
    req(analysis_result())
    log_text <- current_log()
    if (nzchar(log_text)) log_text else "Analisis ejecutado sin mensajes adicionales."
  })

  output$summary_table <- renderTable({
    req(current_result())
    current_result()$summary
  }, rownames = FALSE, digits = function() input$digits)

  output$capability_table <- renderTable({
    req(current_result())
    current_result()$capability
  }, rownames = FALSE, digits = function() input$digits)

  output$capability_plot <- renderImage({
    req(current_result())
    outfile <- tempfile(fileext = ".png")
    build_capability_plot(outfile)
    list(
      src = outfile,
      contentType = "image/png",
      width = capability_plot_dims$width,
      height = capability_plot_dims$height,
      alt = "Grafico de capacidad de proceso"
    )
  }, deleteFile = TRUE)

  output$imr_plot <- renderImage({
    req(current_result())
    outfile <- tempfile(fileext = ".png")
    build_imr_plot(outfile)
    list(
      src = outfile,
      contentType = "image/png",
      width = imr_plot_dims$width,
      height = imr_plot_dims$height,
      alt = "Grafico IMR"
    )
  }, deleteFile = TRUE)

  output$study_plot <- renderImage({
    req(current_result())
    outfile <- tempfile(fileext = ".png")
    build_study_plot(outfile)
    list(
      src = outfile,
      contentType = "image/png",
      width = study_plot_dims$width,
      height = study_plot_dims$height,
      alt = "Capability sixpack"
    )
  }, deleteFile = TRUE)

  output$download_excel <- downloadHandler(
    filename = function() {
      label <- gsub("[^A-Za-z0-9_-]+", "-", analysis_result()$label)
      sprintf("capability-%s-%s.xlsx", label, Sys.Date())
    },
    content = function(file) {
      req(current_result())

      plot_file <- tempfile(fileext = ".png")
      build_capability_plot(plot_file)
      imr_plot_file <- tempfile(fileext = ".png")
      build_imr_plot(imr_plot_file)
      study_plot_file <- tempfile(fileext = ".png")
      build_study_plot(study_plot_file)

      interpretation_text <- NULL
      if (isTruthy(input$run_interpretation) && input$run_interpretation > 0) {
        interpretation_text <- interpretation_result()$text
      }

      write_capability_export_workbook(
        path = file,
        capability_result = current_result(),
        plot_path = plot_file,
        imr_plot_path = imr_plot_file,
        interpretation_text = interpretation_text,
        study_plot_path = study_plot_file
      )
    }
  )

  output$interpretation_status <- renderText({
    req(current_result())

    if (input$run_interpretation < 1) {
      return("Sin consumir tokens. Pulsa 'Generar interpretacion' para consultar OpenAI.")
    }

    if (isTRUE(interpretation_result()$ok)) {
      "Interpretacion generada."
    } else {
      "No se pudo generar la interpretacion."
    }
  })

  output$interpretation_text <- renderUI({
    req(interpretation_result())
    card(
      card_body(
        tags$div(
          style = paste(
            "white-space: pre-wrap; line-height: 1.5;",
            if (isTRUE(interpretation_result()$ok)) "" else "color: #b91c1c;"
          ),
          interpretation_result()$text
        )
      )
    )
  })
}
