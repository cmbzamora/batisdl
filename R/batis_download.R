#' OECD BaTIS API Downloader
#'
#' Download OECD BaTIS and write Excel, excludes self-as-partner,
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
#' key <- "AUS.AUT+BEL+CAN+CHL+COL+CRI+CZE+DNK+EST+FIN+FRA+DEU+GRC+HUN+ISL+IRL+ISR+ITA+JPN+KOR+LVA+LTU+LUX+MEX+NLD+NZL+NOR+POL+PRT+SVK+SVN+ESP+SWE+CHE+TUR+GBR+USA+AFG+ALB+DZA+AGO+AIA+ATG+ARG+ARM+ABW+AZE+BHS+BHR+BGD+BRB+BLR+BLZ+BEN+BMU+BTN+BOL+BIH+BWA+BRA+BRN+BGR+BFA+BDI+CPV+KHM+CMR+CYM+CAF+TCD+CHN+COM+COG+CIV+HRV+CUB+CUW+CYP+PRK+COD+DJI+DMA+DOM+ECU+EGY+SLV+GNQ+ERI+SWZ+ETH+FRO+FJI+PYF+GAB+GMB+GEO+GHA+GRD+GTM+GIN+GNB+GUY+HTI+HND+HKG+IND+IDN+IRN+IRQ+JAM+JOR+KAZ+KEN+KIR+XKV+KWT+KGZ+LAO+LBN+LSO+LBR+LBY+MAC+MDG+MWI+MYS+MDV+MLI+MLT+MRT+MUS+MDA+MNG+MNE+MSR+MAR+MOZ+MMR+NAM+NPL+NCL+NIC+NER+NGA+MKD+OMN+PAK+PSE+PAN+PNG+PRY+PER+PHL+QAT+ROU+RUS+RWA+KNA+LCA+VCT+WSM+STP+SAU+SEN+SRB+SYC+SLE+SGP+SXM+SLB+SOM+ZAF+LKA+SDN+SUR+SYR+TWN+TJK+TZA+THA+TLS+TGO+TON+TTO+TUN+TKM+TCA+TUV+UGA+UKR+ARE+URY+UZB+VUT+VEN+VNM+YEM+ZMB+ZWE+ANT_F+SCG_F+W+WXD.M+X..SC1+SDA+SI1+SJ1+SK1+SC2+SDB+SI2+SJ2+SK2+SC3+SI3+SJ3+SC4+SA+SB+SC+SD+SE+SF+SG+SH+SI+SJ+SK+SL+SOX+SOX1+SPX1+SPX4+S.A.USD_EXC.B"
#' batis_download(
#' reporter_code = "JPN",         #CHANGE COUNTRY CODE
#' start_year    = 2018,          #CHANGE
#' end_year      = 2018,          #CHANGE
#' output_dir    = '/Users/cmbzamora/Library/CloudStorage/OneDrive-AsianDevelopmentBank/Personal/MRIO Series Revisions 2017-24/JPN/BATIS',   #PATH / DIRECTORY WHERE YOU WANT THE FILE TO BE SAVED
#' key_from_site = key            #<- DO NOT CHANGE
#' )
#' }
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
