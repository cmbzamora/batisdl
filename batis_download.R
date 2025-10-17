# SPDX-License-Identifier: Apache-2.0

#' Download OECD BaTIS (DSD_BATIS@DF_BATIS) and write Excel
#'
#' Minimal, no-retry downloader that:
#' - uses your **exact** `key_from_site` (as shown on the OECD site, keep dots),
#' - swaps the first token with `reporter_code`,
#' - handles 404/429 cleanly,
#' - excludes self==partner if columns exist,
#' - writes `"[REPORTER]_BATIS_[start]_[end].xlsx"` with a single sheet.
#'
#' @param reporter_code 3-letter ISO (e.g., "FRA").
#' @param start_year integer.
#' @param end_year integer (defaults to start_year).
#' @param output_dir folder path for the xlsx.
#' @param key_from_site character; the long plus-separated key from the website.
#' @param format character; "csvfilewithlabels" (default) or "csvfile".
#' @return Invisibly, the full output path.
#' @export
batis_download <- function(
  reporter_code,
  start_year,
  end_year = start_year,
  output_dir = ".",
  key_from_site,
  format = "csvfilewithlabels"
) {
  stopifnot(is.character(reporter_code), nchar(reporter_code) == 3)
  stopifnot(is.numeric(start_year), is.numeric(end_year))
  dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)

  prefix  <- "https://sdmx.oecd.org/public/rest/data/OECD.SDD.TPS,DSD_BATIS@DF_BATIS,1.0/"
  key_use <- sub("^[A-Z_]+", toupper(reporter_code), key_from_site)

  url <- paste0(
    prefix, key_use,
    "?startPeriod=", start_year,
    "&endPeriod=",   end_year,
    "&dimensionAtObservation=AllDimensions",
    "&format=",      format
  )

  # download (aborts on 404/429/other)
  df <- safe_read_csv_abort(url)

  # guards
  if (nrow(df) == 0) stop("No rows returned; not writing any file.", call. = FALSE)
  if (!("TIME_PERIOD" %in% names(df))) stop("TIME_PERIOD column missing; aborting.", call. = FALSE)
  df$TIME_PERIOD <- suppressWarnings(as.integer(df$TIME_PERIOD))
  if (!any(df$TIME_PERIOD >= start_year)) {
    stop(sprintf("No observations for %s or later; not writing any file.", start_year), call. = FALSE)
  }

  # exclude self-partner if columns exist
  if (all(c("REF_AREA", "COUNTERPART") %in% names(df))) {
    df <- dplyr::filter(df, REF_AREA != COUNTERPART, COUNTERPART != toupper(reporter_code))
  }

  # numeric coercion
  if ("OBS_VALUE" %in% names(df)) {
    df$OBS_VALUE <- suppressWarnings(as.numeric(df$OBS_VALUE))
  }

  # save
  out_file <- sprintf("%s_BATIS_%s_%s.xlsx", toupper(reporter_code), start_year, end_year)
  out_path <- file.path(output_dir, out_file)

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, sprintf("%s-%s", start_year, end_year))
  openxlsx::writeData(wb, sprintf("%s-%s", start_year, end_year), df)
  openxlsx::saveWorkbook(wb, out_path, overwrite = TRUE)

  message("Saved: ", normalizePath(out_path))
  invisible(out_path)
}
