#' @title Helpers para interpretar resultados con OpenAI
#'
#' @description
#' Reune utilidades para sanitizar resultados de capacidad de proceso,
#' codificar imagenes y construir solicitudes a la API de OpenAI para generar
#' una interpretacion ejecutiva del analisis.
#'
#' @keywords internal
NULL

#' Simplifica la salida de capacidad para serializarla
#'
#' @param capability_result Resultado devuelto por `run_process_capability()`.
#'
#' @return Una lista con tablas convertidas a `data.frame` y metadatos simples.
sanitize_capability_result <- function(capability_result) {
  list(
    summary = if (!is.null(capability_result$summary)) as.data.frame(capability_result$summary) else NULL,
    capability = if (!is.null(capability_result$capability)) as.data.frame(capability_result$capability) else NULL,
    limits = capability_result$limits,
    n = length(capability_result$data)
  )
}

#' Convierte una imagen local a data URL en base64
#'
#' @param image_path Ruta del archivo de imagen.
#'
#' @return Una cadena en formato data URL lista para enviar a OpenAI.
encode_image_data_url <- function(image_path) {
  if (!requireNamespace("base64enc", quietly = TRUE)) {
    stop("Falta instalar 'base64enc'. Usa install.packages('base64enc').", call. = FALSE)
  }

  ext <- tolower(tools::file_ext(image_path))
  mime_type <- switch(
    ext,
    png = "image/png",
    jpg = "image/jpeg",
    jpeg = "image/jpeg",
    webp = "image/webp",
    stop("Formato de imagen no soportado para OpenAI.", call. = FALSE)
  )

  paste0(
    "data:",
    mime_type,
    ";base64,",
    base64enc::base64encode(image_path)
  )
}

#' Extrae el texto util de una respuesta de la API Responses
#'
#' @param body Cuerpo de respuesta parseado como lista.
#'
#' @return Una cadena con el texto generado o `NULL` si no existe.
extract_response_text <- function(body) {
  if (!is.null(body$output_text) && is.character(body$output_text) && nzchar(body$output_text)) {
    return(body$output_text)
  }

  if (!is.null(body$output) && length(body$output) > 0) {
    text_chunks <- unlist(
      lapply(body$output, function(item) {
        if (!is.list(item) || !identical(item$type, "message") || is.null(item$content)) {
          return(character())
        }

        unlist(
          lapply(item$content, function(content_item) {
            if (
              is.list(content_item) &&
              identical(content_item$type, "output_text") &&
              !is.null(content_item$text)
            ) {
              content_item$text
            } else {
              character()
            }
          }),
          use.names = FALSE
        )
      }),
      use.names = FALSE
    )

    text_chunks <- text_chunks[nzchar(text_chunks)]
    if (length(text_chunks) > 0) {
      return(paste(text_chunks, collapse = "\n\n"))
    }
  }

  NULL
}

#' Solicita a OpenAI una interpretacion del analisis de capacidad
#'
#' @param capability_result Resultado devuelto por `run_process_capability()`.
#' @param plot_path Ruta opcional a una imagen PNG del histograma.
#' @param imr_plot_path Ruta opcional a una imagen PNG del grafico IMR.
#' @param study_plot_path Ruta opcional a una imagen PNG del estudio grafico de capacidad.
#' @param language Idioma en el que se solicita la respuesta.
#' @param extra_instructions Instrucciones adicionales para refinar la salida.
#'
#' @return Texto interpretativo generado por OpenAI o el cuerpo JSON
#'   serializado cuando no se puede extraer texto directo.
#' @export
openai_interpret_capability <- function(capability_result, plot_path = NULL, imr_plot_path = NULL, study_plot_path = NULL, language = "es", extra_instructions = "") {
  if (!requireNamespace("httr2", quietly = TRUE)) {
    stop("Falta instalar 'httr2'. Usa install.packages('httr2').", call. = FALSE)
  }

  if (!requireNamespace("jsonlite", quietly = TRUE)) {
    stop("Falta instalar 'jsonlite'. Usa install.packages('jsonlite').", call. = FALSE)
  }

  api_key <- Sys.getenv("OPENAI_API_KEY", unset = "")
  model <- Sys.getenv("OPENAI_MODEL", unset = "gpt-5-mini")

  if (!nzchar(api_key)) {
    stop(
      "OPENAI_API_KEY no esta configurada. Copia .Renviron.example a .Renviron y agrega tu clave.",
      call. = FALSE
    )
  }

  prompt_text <- paste(
    "Interpreta el siguiente resultado de capacidad de proceso en", language, ".",
    "Devuelve solo lo necesario con este formato:",
    "1. Resumen ejecutivo",
    "2. Riesgo",
    "3. Recomendacion",
    "Cada seccion debe ser breve y orientada a decision.",
    "Prioriza Cp, Cpk, Z y la cercania a limites de especificacion.",
    if (nzchar(extra_instructions)) extra_instructions else "",
    "\n\nResultado:\n",
    jsonlite::toJSON(sanitize_capability_result(capability_result), auto_unbox = TRUE, pretty = TRUE, null = "null")
  )

  message_content <- list(
    list(
      type = "input_text",
      text = prompt_text
    )
  )

  if (!is.null(plot_path) && file.exists(plot_path)) {
    message_content[[length(message_content) + 1]] <- list(
      type = "input_image",
      image_url = encode_image_data_url(plot_path),
      detail = "high"
    )
  }

  if (!is.null(imr_plot_path) && file.exists(imr_plot_path)) {
    message_content[[length(message_content) + 1]] <- list(
      type = "input_image",
      image_url = encode_image_data_url(imr_plot_path),
      detail = "high"
    )
  }

  if (!is.null(study_plot_path) && file.exists(study_plot_path)) {
    message_content[[length(message_content) + 1]] <- list(
      type = "input_image",
      image_url = encode_image_data_url(study_plot_path),
      detail = "high"
    )
  }

  payload <- list(
    model = model,
    instructions = paste(
      "Actua como experto en capacidad de proceso y control estadistico.",
      "Usa el resultado numerico y todos los graficos adjuntos disponibles, incluyendo histograma, IMR y estudio grafico.",
      "Responde en espanol claro y conciso para un usuario de negocio."
    ),
    input = list(
      list(
        role = "user",
        content = message_content
      )
    ),
    text = list(
      verbosity = "low"
    )
  )

  response <- httr2::request("https://api.openai.com/v1/responses") |>
    httr2::req_headers(
      Authorization = paste("Bearer", api_key),
      "Content-Type" = "application/json"
    ) |>
    httr2::req_body_json(payload, auto_unbox = TRUE) |>
    httr2::req_perform()

  body <- httr2::resp_body_json(response, simplifyVector = FALSE)
  text <- extract_response_text(body)

  if (!is.null(text) && nzchar(text)) {
    return(text)
  }

  jsonlite::toJSON(body, auto_unbox = TRUE, pretty = TRUE, null = "null")
}
