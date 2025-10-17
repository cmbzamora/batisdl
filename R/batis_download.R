#' Download OECD BaTIS and write Excel
#'
#' Minimal, no-retry downloader that aborts on 404/429, excludes self-as-partner,
#' and saves as "[REPORTER]_BATIS_[start]_[end].xlsx".
#'
#' @param reporter_code 3-letter ISO code (e.g., "FRA")
#' @param start_year Integer start year
#' @param end_year Integer end year (defaults to start_year)
#' @param output_dir Output folder for the xlsx
#' @param key_from_site Long key copied from the OECD site
#' @param format "csvfilewithlabels" (default) or "csvfile"
#' @return Invisibly, the output file path
#' @examples
#' \dontrun{
#' batis_download("FRA", 2023, 2023, tempdir(), key_from_site = "AUS.AUT+...+S.A.USD_EXC.B")
#' }
#' @export
batis_download <- function(...) {  # keep your real args here
  # existing function body…
}

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
  url <- paste0(prefix, key_use,
                "?startPeriod=", start_year,
                "&endPeriod=",   end_year,
                "&dimensionAtObservation=AllDimensions",
                "&format=",      format)

  df <- safe_read_csv_abort(url)

  if (nrow(df) == 0) stop("No rows returned; not writing any file.", call. = FALSE)
  if (!("TIME_PERIOD" %in% names(df))) stop("TIME_PERIOD column missing; aborting.", call. = FALSE)
  df$TIME_PERIOD <- suppressWarnings(as.integer(df$TIME_PERIOD))
  if (!any(df$TIME_PERIOD >= start_year)) {
    stop(sprintf("No observations for %s or later; not writing any file.", start_year), call. = FALSE)
  }

  if (all(c("REF_AREA", "COUNTERPART") %in% names(df))) {
    df <- dplyr::filter(df, REF_AREA != COUNTERPART, COUNTERPART != toupper(reporter_code))
  }
  if ("OBS_VALUE" %in% names(df)) df$OBS_VALUE <- suppressWarnings(as.numeric(df$OBS_VALUE))

  out_file <- sprintf("%s_BATIS_%s_%s.xlsx", toupper(reporter_code), start_year, end_year)
  out_path <- file.path(output_dir, out_file)

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, sprintf("%s-%s", start_year, end_year))
  openxlsx::writeData(wb, sprintf("%s-%s", start_year, end_year), df)
  openxlsx::saveWorkbook(wb, out_path, overwrite = TRUE)

  message("Saved: ", normalizePath(out_path))
  invisible(out_path)
}
