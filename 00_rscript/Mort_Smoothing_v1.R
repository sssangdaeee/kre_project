rm(list=ls())
setwd("~/WS/250812_Mortality_Smoothing")

# ============================================================
# Mortality Improvement Smoothing (2D Gaussian + Tail Smoothing)
# - Pandemic adjustment: None
# - 2035 column: floor at 0 (min=0), then LAST-YEARS ONLY 1D year-smoothing
# - Input  : CSV with columns: Age, 2024, ..., 2035   (values allow %, e.g. "1.23%")
# - Output : Smoothed CSVs (numeric / percent) + heatmaps + age-line plots (PNG)
# - Deps   : ggplot2 (only)
# ============================================================

library(tidyverse)
library(scales)
library(readxl)
library(ggplot2)

# --------- 0) 사용자 설정 ----------
# 입력 파일 경로 (남/여 각각)
male_csv   <- read_csv("4. male_future.csv")
female_csv <- read_csv("4. female_future.csv")

# 2D Gaussian bandwidths (tune as needed)
sigma_age  <- 2.5    # 연령축 표준편차(≈5~6세 영향)
sigma_year <- 1.2    # 기간축 표준편차(≈2~3년 영향)

# Tail smoothing (2035=0 경계 부드럽게)
tail_years <- 6      # 마지막 몇 년만 재-스무딩 (예: 2032~2035 → 4)
sigma_tail <- 1.0    # 꼬리 스무딩 표준편차(0.8~1.2 권장)

out_dir <- "~/WS/250812_Mortality_Smoothing"
#if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

# ------------------- Helpers -------------------
# 문자/퍼센트 → 소수
to_num <- function(x) {
  if (is.character(x)) {
    x <- gsub(",", ".", x, fixed = TRUE)   # 콤마 소수점 → 점
    as.numeric(gsub("%", "", x)) / 100
  } else {
    as.numeric(x)
  }
}

# CSV/Excel 또는 data.frame → 행=Age, 열=Year 의 numeric matrix
read_impr <- function(path_or_df) {
  if (is.data.frame(path_or_df)) {
    raw <- as.data.frame(path_or_df)
  } else if (is.character(path_or_df) && length(path_or_df) == 1) {
    if (!file.exists(path_or_df)) stop("파일이 존재하지 않습니다: ", path_or_df)
    ext <- tolower(tools::file_ext(path_or_df))
    if (ext %in% c("csv", "txt")) {
      raw <- utils::read.csv(path_or_df, check.names = FALSE, fileEncoding = "UTF-8-BOM")
    } else if (ext %in% c("xlsx", "xls")) {
      if (!requireNamespace("readxl", quietly = TRUE)) {
        stop("엑셀 파일을 읽으려면 'readxl'이 필요합니다. CSV로 저장하거나 readxl을 설치해 주세요.")
      }
      raw <- as.data.frame(readxl::read_excel(path_or_df))
    } else {
      stop("지원하지 않는 파일 형식: ", ext)
    }
  } else {
    stop("파일 경로(문자열) 또는 data.frame을 전달해 주세요.")
  }
  
  if (!"Age" %in% names(raw)) names(raw)[1] <- "Age"
  for (j in seq.int(2, ncol(raw))) raw[[j]] <- to_num(raw[[j]])
  
  mat <- as.matrix(raw[, -1, drop = FALSE])
  rownames(mat) <- raw$Age
  colnames(mat) <- as.character(colnames(mat))
  mat
}

# 퍼센트 문자열 테이블
to_pct_tbl <- function(mat, accuracy = 0.01) {
  digs <- max(0, ceiling(-log10(accuracy)))
  out  <- cbind(Age = rownames(mat), apply(mat, 2, function(col)
    sprintf(paste0("%.", digs, "f%%"), 100 * as.numeric(col))))
  out <- as.data.frame(out, stringsAsFactors = FALSE)
  out$Age <- as.numeric(out$Age)
  out
}

# 퍼센트 라벨 포매터 (범례/축에 사용)
percent_lab <- function(x, accuracy = 0.1) {
  digs <- max(0, ceiling(-log10(accuracy)))
  sprintf(paste0("%.", digs, "f%%"), 100 * x)
}

# Gaussian kernel weights (길이 = 2r+1, r ≈ 4*sigma)
gaussian_weights <- function(sigma) {
  r  <- max(1, round(4 * sigma))
  ix <- (-r):r
  w  <- exp(-0.5 * (ix / sigma)^2)
  w / sum(w)
}

# Edge-replicated 1D smoothing (중심형 컨볼루션)
smooth_1d <- function(v, w) {
  r <- (length(w) - 1) / 2
  if (abs(r - floor(r)) > 1e-8) stop("커널 길이는 홀수여야 합니다.")
  vpad <- c(rep(v[1], r), v, rep(v[length(v)], r))
  ypad <- stats::filter(vpad, w, sides = 2, circular = FALSE)
  as.numeric(ypad[(r + 1):(r + length(v))])
}

# Separable 2D Gaussian smoothing: age → year
gaussian_2d <- function(mat, sigma_age, sigma_year) {
  wa <- gaussian_weights(sigma_age)
  wy <- gaussian_weights(sigma_year)
  tmp <- apply(mat, 2, smooth_1d, w = wa)         # age 방향(열별)
  out <- t(apply(tmp, 1, smooth_1d, w = wy))      # year 방향(행별)
  colnames(out) <- colnames(mat)
  rownames(out) <- rownames(mat)
  out
}

# Boundary-tail smoothing: 2035=0으로 고정한 뒤, 마지막 몇 년만 연도축 재-스무딩
boundary_tail_smooth <- function(mat, tail_years = 4, sigma_tail = 1.0) {
  wy   <- gaussian_weights(sigma_tail)
  nY   <- ncol(mat)
  last <- nY
  if (tail_years < 2) return(mat)
  k0   <- max(1, last - tail_years + 1)
  
  # 2035(마지막 열)가 0으로 클램프된 연령만 조정
  idx <- which(as.numeric(mat[, last]) <= 1e-15)
  if (length(idx) == 0) return(mat)
  
  for (i in idx) {
    v        <- as.numeric(mat[i, ])
    v[last]  <- 0                            # 마지막 값 0 하드 고정
    vs       <- smooth_1d(v, wy)             # 연도축 1D 스무딩
    mat[i, k0:last] <- vs[k0:last]           # 꼬리 구간만 업데이트
    mat[i, last]    <- 0                     # 마지막은 다시 0으로 확정
  }
  mat
}

# Long 데이터 생성 (ggplot용)
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

# 공통 컬러 스케일 (원본/스무딩 비교시 같은 limits)
div_scale <- function(lims = NULL) {
  scale_fill_gradient2(
    low = "#B2182B",           # negative -> RED
    mid = "#FBFBFB",           # near 0 -> very light
    high = "#2166AC",          # positive -> BLUE
    midpoint = 0,
    limits = lims,
    oob = scales::squish,      # limits 밖은 색상범위로 클램프
    labels = function(x) percent_lab(x, 0.1)
  )
}

# Heatmap 저장
save_heatmap <- function(mat, title, path_png, lims = NULL) {
  df <- mat_to_long(mat)
  p <- ggplot(df, aes(Year, Age, fill = Improvement)) +
    geom_tile() +
    div_scale(lims) +
    labs(title = title, x = "Year", y = "Age", fill = "Improvement") +
    theme_minimal(base_size = 13) +
    theme(plot.title = element_text(face = "bold"),
          panel.grid.minor = element_blank())
  ggsave(path_png, p, width = 7.6, height = 5.4, dpi = 200)
}

# 연령 단면 라인 플롯 저장 (예: 30/60/85세)
save_age_lines <- function(orig, sm, title, path_png, ages = c(30, 60, 85)) {
  yrs <- as.integer(colnames(orig))
  df <- do.call(rbind, lapply(ages, function(a) {
    rbind(
      data.frame(Age = a, Year = yrs, Improvement = as.numeric(orig[as.character(a), ]), Type = "Original"),
      data.frame(Age = a, Year = yrs, Improvement = as.numeric(sm[as.character(a), ]),   Type = "Smoothed")
    )
  }))
  p <- ggplot(df, aes(Year, Improvement, color = Type)) +
    geom_line(linewidth = 1) +
    facet_wrap(~ Age, nrow = 1, scales = "free_y") +
    scale_y_continuous(labels = function(x) percent_lab(x, 0.1)) +
    scale_color_manual(values = c("Original" = "#7a7a7a", "Smoothed" = "#1b73e8")) +
    labs(title = paste(title, "— Age profiles (2024–2035)"),
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

# 3) Clamp 2035 at min 0 (negatives → 0)
col_2035 <- "2035"
if (!(col_2035 %in% colnames(male_sm0)) || !(col_2035 %in% colnames(female_sm0))) {
  stop("입력 데이터에 '2035' 컬럼이 필요합니다.")
}
male_sm0[,   col_2035] <- pmax(male_sm0[,   col_2035], 0)
female_sm0[, col_2035] <- pmax(female_sm0[, col_2035], 0)

# 4) Tail boundary smoothing (last few years; only for rows clamped to 0 at 2035)
male_sm   <- boundary_tail_smooth(male_sm0,   tail_years = tail_years, sigma_tail = sigma_tail)
female_sm <- boundary_tail_smooth(female_sm0, tail_years = tail_years, sigma_tail = sigma_tail)

# 5) Save CSVs (numeric & percent strings)
write.csv(cbind(Age = rownames(male_sm),   round(male_sm,   6)), file.path(out_dir, "male_smoothed_numeric.csv"),   row.names = FALSE)
write.csv(cbind(Age = rownames(female_sm), round(female_sm, 6)), file.path(out_dir, "female_smoothed_numeric.csv"), row.names = FALSE)
write.csv(to_pct_tbl(male_sm,   0.01), file.path(out_dir, "male_smoothed_percent.csv"),   row.names = FALSE)
write.csv(to_pct_tbl(female_sm, 0.01), file.path(out_dir, "female_smoothed_percent.csv"), row.names = FALSE)

# 6) Plots (heatmaps with common limits per gender, plus age-line plots)
sat_q <- 0.90  # 0.85~0.95 사이로 조정해보세요 (값↓ → 대비↑)

abs_m <- abs(c(male_mat, male_sm))
L_m   <- as.numeric(quantile(abs_m, sat_q, na.rm = TRUE))
male_lim <- c(-L_m, L_m)

abs_f <- abs(c(female_mat, female_sm))
L_f   <- as.numeric(quantile(abs_f, sat_q, na.rm = TRUE))
female_lim <- c(-L_f, L_f)

save_heatmap(male_mat,   "Male: Original (2024–2035)",               file.path(out_dir, "male_original_heatmap.png"), male_lim)
save_heatmap(male_sm,    "Male: 2D Gaussian + Tail smoothing",       file.path(out_dir, "male_smoothed_heatmap.png"), male_lim)
save_heatmap(female_mat, "Female: Original (2024–2035)",             file.path(out_dir, "female_original_heatmap.png"), female_lim)
save_heatmap(female_sm,  "Female: 2D Gaussian + Tail smoothing",     file.path(out_dir, "female_smoothed_heatmap.png"), female_lim)

save_age_lines(male_mat,   male_sm,   "Male",   file.path(out_dir, "male_age_lines.png"))
save_age_lines(female_mat, female_sm, "Female", file.path(out_dir, "female_age_lines.png"))

message("Done. Files written to: ", normalizePath(out_dir))
