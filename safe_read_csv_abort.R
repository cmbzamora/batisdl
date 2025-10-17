#' Abort-on-status CSV fetch (no retries)
#'
#' Makes a GET request (httr2), **does not** auto-error on 4xx/5xx, and
#' aborts with clear messages for 404, 429, others; returns a parsed tibble for 2xx.
#' @param url Character. Full request URL.
#' @param ... Passed to readr::read_csv.
#' @return A tibble (on 2xx) or `stop()` on error statuses.
#' @keywords internal
safe_read_csv_abort <- function(url, ...) {
  req <- httr2::request(url) |>
    httr2::req_error(is_error = function(resp) FALSE) |>
    httr2::req_headers(
      `User-Agent` = sprintf("ADB-BaTIS-R/%s (httr2)", getRversion()),
      `Accept-Encoding` = "gzip, deflate"
    )

  resp <- httr2::req_perform(req)
  status <- httr2::resp_status(resp)

  if (status == 429) {
    ra <- httr2::resp_header(resp, "retry-after")
    stop(
      paste0(
        "Rate limited by OECD (HTTP 429).",
        if (!is.null(ra) && nzchar(ra)) paste0(" Retry after ", ra, "s.") else "",
        " Try later or reduce request size."
      ),
      call. = FALSE
    )
  }
  if (status == 404) stop("No data available yet for this request (HTTP 404).", call. = FALSE)
  if (status >= 400) stop(paste0("HTTP ", status, " from OECD."), call. = FALSE)

  readr::read_csv(httr2::resp_body_string(resp), show_col_types = FALSE, ...)
}
