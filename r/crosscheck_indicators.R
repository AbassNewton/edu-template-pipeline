# Independent re-computation of the gross enrolment ratio in base R, from the
# consolidated CSV files, compared with the Python results (indicators.csv).
# Usage: Rscript r/crosscheck_indicators.R out/processed
args <- commandArgs(trailingOnly = TRUE)
dir <- if (length(args)) args[1] else "out/processed"

obs <- read.csv(file.path(dir, "observations.csv"), stringsAsFactors = FALSE)
par <- read.csv(file.path(dir, "parameters.csv"), stringsAsFactors = FALSE)
ind <- read.csv(file.path(dir, "indicators.csv"), stringsAsFactors = FALSE)

param <- function(country, key, level) {
  v <- par$value[par$country == country & par$code == paste0("PAR|", key, "|", level, "|")]
  if (length(v) == 0 || is.na(v) || v == "") NA_real_ else as.numeric(v)
}

ger <- ind[ind$indicator == "GER", ]
ger$r_value <- NA_real_
for (i in seq_len(nrow(ger))) {
  g <- ger[i, ]
  e <- param(g$country, "entry_age", g$level)
  d <- param(g$country, "duration", g$level)
  num <- obs$value[obs$country == g$country & obs$code == paste0("ENR|", g$level, "|", g$sex, "|") &
                   obs$year == g$year]
  if (is.na(e) || is.na(d) || length(num) != 1 || is.na(num)) next
  sexes <- if (g$sex == "T") c("F", "M") else g$sex
  p <- obs[obs$country == g$country & obs$series == "POP" & obs$year == g$year &
           obs$sex %in% sexes & obs$age >= e & obs$age < e + d, "value"]
  if (length(p) == length(sexes) * d && all(!is.na(p)) && sum(p) > 0) ger$r_value[i] <- num / sum(p) * 100
}

both_na <- is.na(ger$value) & is.na(ger$r_value)
same <- both_na | (!is.na(ger$value) & !is.na(ger$r_value) & abs(ger$value - ger$r_value) < 1e-9)
cat(sprintf("GER cells compared: %d, mismatches: %d\n", nrow(ger), sum(!same)))
if (any(!same)) print(head(ger[!same, c("country", "level", "sex", "year", "value", "r_value")]))
quit(status = as.integer(any(!same)))
