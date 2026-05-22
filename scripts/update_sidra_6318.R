#!/usr/bin/env Rscript

suppressPackageStartupMessages({
  if (!requireNamespace("jsonlite", quietly = TRUE)) {
    stop("Pacote 'jsonlite' não encontrado. Instale com: install.packages('jsonlite').")
  }
  if (!requireNamespace("openxlsx", quietly = TRUE)) {
    stop("Pacote 'openxlsx' não encontrado. Instale com: install.packages('openxlsx').")
  }
})

best_match_column <- function(needles, candidates, data_frame) {
  needles <- unique(tolower(needles))
  best_column <- NULL
  best_score <- 0

  for (column in candidates) {
    values <- unique(na.omit(as.character(data_frame[[column]])))
    if (length(values) == 0) {
      next
    }
    score <- sum(tolower(values) %in% needles)
    if (score > best_score) {
      best_score <- score
      best_column <- column
    }
  }

  if (best_score == 0) {
    return(NULL)
  }
  best_column
}

config <- list(
  workbook_path = Sys.getenv("SIDRA_WORKBOOK", "Suporte_IE_PnadC_Mensal.xlsx"),
  sheet_name = Sys.getenv("SIDRA_SHEET", "Tabela"),
  region_name = Sys.getenv("SIDRA_REGION_NAME", "Brasil"),
  data_start_marker = Sys.getenv("SIDRA_DATA_START_MARKER", "Brasil"),
  descriptor_url = Sys.getenv("SIDRA_DESCRIPTOR_URL", "https://apisidra.ibge.gov.br/DescritoresTabela/t/6318"),
  data_url = Sys.getenv(
    "SIDRA_DATA_URL",
    "https://apisidra.ibge.gov.br/values/t/6318/n1/all/v/all/p/all?formato=json"
  ),
  output_csv = Sys.getenv("SIDRA_OUTPUT_CSV", "sidra_6318_output.csv")
)

if (!file.exists(config$workbook_path)) {
  stop("Arquivo não encontrado: ", config$workbook_path)
}

sidra <- jsonlite::fromJSON(config$data_url)
if (!is.data.frame(sidra)) {
  stop("Resposta da API inválida. Verifique a URL: ", config$data_url)
}

value_column <- intersect(c("V", "Valor", "valor", "Value", "value"), names(sidra))[1]
if (is.na(value_column)) {
  stop("Coluna de valor não encontrada na resposta da API.")
}

name_columns <- names(sidra)[grepl("N$", names(sidra))]
if (length(name_columns) == 0) {
  stop("Nenhuma coluna de dimensões (terminada em 'N') foi encontrada na API.")
}

workbook <- openxlsx::loadWorkbook(config$workbook_path)
sheet_data <- openxlsx::read.xlsx(workbook, sheet = config$sheet_name, colNames = FALSE)

data_start_row <- which(sheet_data[[1]] == config$data_start_marker)[1]
if (is.na(data_start_row)) {
  stop("Linha inicial dos dados não encontrada (", config$data_start_marker, ").")
}

source_row <- which(grepl("^Fonte:", sheet_data[[1]]))[1]
if (is.na(source_row)) {
  source_row <- nrow(sheet_data) + 1
}

candidate_rows <- seq_len(data_start_row - 1)
non_empty_rows <- candidate_rows[
  vapply(candidate_rows, function(row_index) {
    row_values <- as.character(unlist(sheet_data[row_index, 3:ncol(sheet_data)]))
    any(!is.na(row_values) & nzchar(row_values))
  }, logical(1))
]
category_row <- tail(non_empty_rows, 1)
if (length(category_row) == 0 || is.na(category_row)) {
  stop("Linha de categorias não encontrada antes da linha de dados.")
}
categories <- as.character(unlist(sheet_data[category_row, 3:ncol(sheet_data)]))
categories <- categories[!is.na(categories) & nzchar(categories)]
if (length(categories) == 0) {
  stop("Categorias não encontradas na planilha.")
}

time_values <- as.character(unlist(sheet_data[data_start_row:(source_row - 1), 2]))
time_values <- time_values[!is.na(time_values) & nzchar(time_values)]

region_column <- best_match_column(config$region_name, name_columns, sidra)
time_column <- best_match_column(time_values, name_columns, sidra)
category_column <- best_match_column(categories, name_columns, sidra)

if (is.null(time_column) || is.null(category_column)) {
  stop(
    "Não foi possível identificar as dimensões da API. Verifique os descritores em ",
    config$descriptor_url
  )
}

filtered <- sidra
if (!is.null(region_column)) {
  filtered <- filtered[filtered[[region_column]] == config$region_name, , drop = FALSE]
}

filtered[[value_column]] <- as.numeric(filtered[[value_column]])

long_data <- filtered[, c(time_column, category_column, value_column)]
names(long_data) <- c("trimestre", "categoria", "valor")

long_data$trimestre <- as.character(long_data$trimestre)
long_data$categoria <- as.character(long_data$categoria)

wide_matrix <- xtabs(valor ~ trimestre + categoria, data = long_data, drop.unused.levels = FALSE)
wide_data <- as.data.frame.matrix(wide_matrix)

missing_categories <- setdiff(categories, colnames(wide_data))
for (missing_category in missing_categories) {
  wide_data[[missing_category]] <- NA
}
wide_data <- wide_data[, categories, drop = FALSE]

time_code_column <- sub("N$", "C", time_column)
if (time_code_column %in% names(filtered)) {
  time_map <- unique(filtered[, c(time_column, time_code_column)])
  time_map$code <- suppressWarnings(as.numeric(time_map[[time_code_column]]))
  time_map <- time_map[order(time_map$code, na.last = TRUE), ]
  ordered_times <- time_map[[time_column]]
  ordered_times <- ordered_times[ordered_times %in% rownames(wide_data)]
  if (length(ordered_times) > 0) {
    wide_data <- wide_data[ordered_times, , drop = FALSE]
  }
} else if (length(time_values) > 0) {
  ordered_times <- c(intersect(time_values, rownames(wide_data)), setdiff(rownames(wide_data), time_values))
  wide_data <- wide_data[ordered_times, , drop = FALSE]
}

output <- data.frame(
  Unidade = c(config$region_name, rep(NA, max(nrow(wide_data) - 1, 0))),
  `Trimestre Móvel` = rownames(wide_data),
  wide_data,
  check.names = FALSE
)

utils::write.csv(output, config$output_csv, row.names = FALSE, na = "")

data_end_row <- source_row - 1
clear_rows <- data_start_row:data_end_row
clear_cols <- seq_len(ncol(output))
openxlsx::writeData(
  workbook,
  sheet = config$sheet_name,
  x = matrix(NA, nrow = length(clear_rows), ncol = length(clear_cols)),
  startRow = data_start_row,
  startCol = 1,
  colNames = FALSE
)

openxlsx::writeData(
  workbook,
  sheet = config$sheet_name,
  x = output,
  startRow = data_start_row,
  startCol = 1,
  colNames = FALSE
)

source_text <- sheet_data[source_row, 1]
if (!is.na(source_text) && nzchar(source_text)) {
  new_source_row <- data_start_row + nrow(output)
  openxlsx::writeData(
    workbook,
    sheet = config$sheet_name,
    x = source_text,
    startRow = new_source_row,
    startCol = 1,
    colNames = FALSE
  )
}

openxlsx::saveWorkbook(workbook, config$workbook_path, overwrite = TRUE)

message("Atualização concluída: ", config$output_csv)
