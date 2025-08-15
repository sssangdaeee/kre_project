rm(list=ls())
setwd("C:\\Users\\mckim\\Desktop\\Mortality_Improvement")

# ============================================================
# Mortality Improvement Smoothing (Final, 2020–2035)
# - 2D Gaussian (separable) + Tail 1D Gaussian at boundary
# - Pandemic adj: none
# - 2035 floor: 0%  (min 0.00)
# - Viz: connected diverging palette (neg→warm, 0→soft neutral, pos→cool)
# - Output percent rounded to 0.1% (e.g., 0.12% → 0.1%)
# - Deps: ggplot2 (readxl optional for .xlsx)
# ============================================================

library(tidyverse)
library(scales)
library(readxl)
library(ggplot2)

# ------------------- User settings -------------------
male_csv   <- read_csv("4. male_future.csv")    # <- 남성 CSV 경로 (열: Age, 2020..2035)
female_csv <- read_csv("4. female_future.csv")  # <- 여성 CSV 경로 (열: Age, 2020..2035)

# 2D Gaussian bandwidths
sigma_age  <- 2.5
sigma_year <- 1.2

# Tail smoothing (boundary continuity near 2035)
tail_years <- 4       # last L years, e.g., 2032–2035
sigma_tail <- 1.0

# Boundary caps at 2035
floor_val <- -0.01   # 하단 -1%
ceil_val  <-  0.01   # 상단 +1%

# Heatmap contrast (symmetric clipping by abs quantile)
sat_q      <- 0.90    # 0.85–0.95 권장


out_dir <- "C:\\Users\\mckim\\Desktop\\Mortality_Improvement"

# ------------------- Helpers -------------------
# percent/char -> numeric (fraction)
to_num <- function(x) {
  if (is.character(x)) {
    x <- gsub(",", ".", x, fixed = TRUE)
    as.numeric(gsub("%", "", x)) / 100
  } else as.numeric(x)
}

# Read CSV/Excel or data.frame -> numeric matrix (rows=Age, cols=Year)
read_impr <- function(path_or_df) {
  if (is.data.frame(path_or_df)) {
    raw <- as.data.frame(path_or_df)
  } else if (is.character(path_or_df) && length(path_or_df) == 1) {
    if (!file.exists(path_or_df)) stop("파일이 존재하지 않습니다: ", path_or_df)
    ext <- tolower(tools::file_ext(path_or_df))
    if (ext %in% c("csv","txt")) {
      raw <- utils::read.csv(path_or_df, check.names = FALSE, fileEncoding = "UTF-8-BOM")
    } else if (ext %in% c("xlsx","xls")) {
      if (!requireNamespace("readxl", quietly = TRUE)) {
        stop("엑셀(.xlsx) 파일을 읽으려면 'readxl' 패키지가 필요합니다. CSV로 저장하거나 readxl을 설치하세요.")
      }
      raw <- as.data.frame(readxl::read_excel(path_or_df))
    } else stop("지원하지 않는 파일 형식: ", ext)
  } else stop("파일 경로(문자열) 또는 data.frame을 전달해 주세요.")
  
  if (!"Age" %in% names(raw)) names(raw)[1] <- "Age"
  for (j in seq.int(2, ncol(raw))) raw[[j]] <- to_num(raw[[j]])
  
  mat <- as.matrix(raw[, -1, drop = FALSE])
  rownames(mat) <- raw$Age
  colnames(mat) <- as.character(colnames(mat))
  # === check: 2025..2035 존재 여부(권장) ===
  yrs <- as.integer(colnames(mat))
  if (min(yrs) != 2020 || max(yrs) != 2035) {
    warning("권장 연도 범위는 2020–2035 입니다. (현재: ", min(yrs), "–", max(yrs), ")")
  }
  mat
}

# Percent string table (rounded to 0.1%)
to_pct_tbl <- function(mat) {
  sprintf01 <- function(x) sprintf("%.1f%%", 100 * x)  # 0.1% 자리 반올림
  out <- cbind(Age = rownames(mat), apply(mat, 2, function(col) sprintf01(as.numeric(col))))
  out <- as.data.frame(out, stringsAsFactors = FALSE)
  out$Age <- as.numeric(out$Age)
  out
}

# Percent label for plots (0.1% ticks)
percent_lab <- function(x) sprintf("%.1f%%", 100 * x)

# Gaussian weights (len = 2r+1, r≈4*sigma)
gaussian_weights <- function(sigma) {
  r  <- max(1, round(4 * sigma))
  ix <- (-r):r
  w  <- exp(-0.5 * (ix / sigma)^2)
  w / sum(w)
}

# Edge-replicated centered 1D smoothing
smooth_1d <- function(v, w) {
  r <- (length(w) - 1) / 2
  if (abs(r - floor(r)) > 1e-8) stop("커널 길이는 홀수여야 합니다.")
  vpad <- c(rep(v[1], r), v, rep(v[length(v)], r))
  ypad <- stats::filter(vpad, w, sides = 2, circular = FALSE)
  as.numeric(ypad[(r + 1):(r + length(v))])
}

# Separable 2D Gaussian: age -> year
gaussian_2d <- function(mat, sigma_age, sigma_year) {
  wa <- gaussian_weights(sigma_age)
  wy <- gaussian_weights(sigma_year)
  tmp <- apply(mat, 2, smooth_1d, w = wa)        # by column (age direction)
  out <- t(apply(tmp, 1, smooth_1d, w = wy))     # by row (year direction)
  colnames(out) <- colnames(mat)
  rownames(out) <- rownames(mat)
  out
}

# Tail boundary smoothing with dual anchors (floor & ceiling)
boundary_tail_smooth_dual <- function(mat, tail_years = 4, sigma_tail = 1.0,
                                      floor_val = -0.01, ceil_val = 0.01, tol = 1e-12) {
  wy   <- gaussian_weights(sigma_tail)
  nY   <- ncol(mat); last <- nY
  if (tail_years < 2) return(mat)
  k0   <- max(1, last - tail_years + 1)
  
  # (A) floor-anchored rows
  idx_floor <- which(abs(as.numeric(mat[, last]) - floor_val) < tol)
  if (length(idx_floor)) {
    for (i in idx_floor) {
      v <- as.numeric(mat[i, ])
      v[last] <- floor_val
      vs <- smooth_1d(v, wy)
      mat[i, k0:last] <- vs[k0:last]
      mat[i, last]    <- floor_val
    }
  }
  
  # (B) ceiling-anchored rows
  idx_ceil <- which(abs(as.numeric(mat[, last]) - ceil_val) < tol)
  if (length(idx_ceil)) {
    for (i in idx_ceil) {
      v <- as.numeric(mat[i, ])
      v[last] <- ceil_val
      vs <- smooth_1d(v, wy)
      mat[i, k0:last] <- vs[k0:last]
      mat[i, last]    <- ceil_val
    }
  }
  mat
}

# Long dataframe for ggplot
mat_to_long <- function(mat) {
  yrs  <- as.integer(colnames(mat))
  ages <- as.numeric(rownames(mat))
  data.frame(
    Age = rep(ages, times = length(yrs)),
    Year = rep(yrs,  each  = length(ages)),
    Improvement = as.vector(mat),
    check.names = FALSE
  )
}

# oob squish (no 'scales' dependency)
oob_squish <- function(x, range) pmin(pmax(x, range[1]), range[2])

# === Connected diverging palette around zero ===
# softer neutral near 0 (not pure white) to reduce “boundary break” feeling
div_scale_connected <- function(lims = NULL) {
  scale_fill_gradientn(
    colours = c("#B40426", "#F08A7C", "#E6E1DF", "#8DB0E3", "#3B4CC0"),
    values  = c(0.00,      0.45,      0.50,      0.55,      1.00),
    limits  = lims,
    oob     = oob_squish,        # <- 여기! 익명함수 대신 그대로 전달
    labels  = percent_lab
  )
}

# Heatmap save
save_heatmap <- function(mat, title, path_png, lims = NULL) {
  df <- mat_to_long(mat)
  p <- ggplot(df, aes(Year, Age, fill = Improvement)) +
    geom_tile() +
    div_scale_connected(lims) +
    labs(title = title, x = "Year", y = "Age", fill = "Improvement") +
    theme_minimal(base_size = 13) +
    theme(plot.title = element_text(face = "bold"),
          panel.grid.minor = element_blank())
  ggsave(path_png, p, width = 7.6, height = 5.4, dpi = 200)
}

# Age-section line plots (e.g., 30/60/85)
save_age_lines <- function(orig, sm, title, path_png, ages = c(30, 60, 85)) {
  yrs <- as.integer(colnames(orig))
  yr_range <- paste0(min(yrs), "–", max(yrs))
  df <- do.call(rbind, lapply(ages, function(a) {
    rbind(
      data.frame(Age = a, Year = yrs, Improvement = as.numeric(orig[as.character(a), ]), Type = "Original"),
      data.frame(Age = a, Year = yrs, Improvement = as.numeric(sm[as.character(a), ]),   Type = "Smoothed")
    )
  }))
  p <- ggplot(df, aes(Year, Improvement, color = Type)) +
    geom_line(linewidth = 1) +
    facet_wrap(~ Age, nrow = 1, scales = "free_y") +
    scale_y_continuous(labels = percent_lab) +
    scale_color_manual(values = c("Original" = "#7A7A7A", "Smoothed" = "#1B73E8")) +
    labs(title = paste(title, "— Age profiles (", yr_range, ")", sep = ""),
         x = "Year", y = "Improvement", color = NULL) +
    theme_minimal(base_size = 13) +
    theme(plot.title = element_text(face = "bold"),
          legend.position = "top",
          panel.grid.minor = element_blank())
  ggsave(path_png, p, width = 9.6, height = 3.2, dpi = 200)
}

# ------------------- Pipeline -------------------
# 1) Read
male_mat   <- read_impr(male_csv)
female_mat <- read_impr(female_csv)

# 2) 2D Gaussian smoothing
male_sm0   <- gaussian_2d(male_mat,   sigma_age, sigma_year)
female_sm0 <- gaussian_2d(female_mat, sigma_age, sigma_year)

# 3) Clamp 2035 at min 0% (negatives -> 0)
col_2035 <- "2035"
if (!(col_2035 %in% colnames(male_sm0)) || !(col_2035 %in% colnames(female_sm0))) {
  stop("입력 데이터에 '2035' 컬럼이 필요합니다.")
}
male_sm0[,   col_2035] <- pmin(pmax(male_sm0[,   col_2035], floor_val), ceil_val)
female_sm0[, col_2035] <- pmin(pmax(female_sm0[, col_2035], floor_val), ceil_val)

# 4) Tail boundary smoothing with dual anchors
male_sm   <- boundary_tail_smooth_dual(male_sm0,   tail_years = tail_years, sigma_tail = sigma_tail,
                                       floor_val = floor_val, ceil_val = ceil_val)
female_sm <- boundary_tail_smooth_dual(female_sm0, tail_years = tail_years, sigma_tail = sigma_tail,
                                       floor_val = floor_val, ceil_val = ceil_val)

# 5) Export CSVs
write.csv(cbind(Age = rownames(male_sm),   round(male_sm,   6)), file.path(out_dir, "male_smoothed_numeric.csv"),   row.names = FALSE)
write.csv(cbind(Age = rownames(female_sm), round(female_sm, 6)), file.path(out_dir, "female_smoothed_numeric.csv"), row.names = FALSE)
# Percent CSV rounded to 0.1%
write.csv(to_pct_tbl(male_sm),   file.path(out_dir, "male_smoothed_percent.csv"),   row.names = FALSE)
write.csv(to_pct_tbl(female_sm), file.path(out_dir, "female_smoothed_percent.csv"), row.names = FALSE)

# 6) Heatmaps (with common symmetric limits per gender using abs quantile)
abs_m <- abs(c(male_mat, male_sm))
L_m   <- as.numeric(quantile(abs_m, sat_q, na.rm = TRUE))
male_lim <- c(-L_m, L_m)

abs_f <- abs(c(female_mat, female_sm))
L_f   <- as.numeric(quantile(abs_f, sat_q, na.rm = TRUE))
female_lim <- c(-L_f, L_f)

save_heatmap(male_mat,   "Male: Original (2020–2035)",           file.path(out_dir, "male_original_heatmap.png"),   male_lim)
save_heatmap(male_sm,    "Male: 2D Gaussian + Tail smoothing",   file.path(out_dir, "male_smoothed_heatmap.png"),   male_lim)
save_heatmap(female_mat, "Female: Original (2020–2035)",         file.path(out_dir, "female_original_heatmap.png"), female_lim)
save_heatmap(female_sm,  "Female: 2D Gaussian + Tail smoothing", file.path(out_dir, "female_smoothed_heatmap.png"), female_lim)
# 7) Age-line plots
save_age_lines(male_mat,   male_sm,   "Male",   file.path(out_dir, "male_age_lines.png"))
save_age_lines(female_mat, female_sm, "Female", file.path(out_dir, "female_age_lines.png"))

message("Done. Files written to: ", normalizePath(out_dir))
