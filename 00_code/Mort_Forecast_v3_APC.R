rm(list=ls())

# ======================================================================
# APC + (κ, γ) 시뮬레이션 경로 예측 + 신뢰밴드(P50/P05/P95)
# - Excel: "response"(deaths Dxt), "dose"(exposure Ext), 행=연령(0..100), 열=연도
# - Fit: StMoMo::apc
# - Forecast: κ, γ를 ARIMA/ETS로 "시뮬레이션 경로" 생성 → rates 재계산
# - Output: P50을 베이스라인으로 사용, P05/P95 밴드 포함
# ======================================================================

# -------- 1) 패키지 -----------------------------------------------------
need <- c("readxl","dplyr","tidyr","StMoMo","openxlsx","forecast","zoo")
for(p in need){
  if(!requireNamespace(p, quietly = TRUE)) install.packages(p)
}
library(readxl); library(dplyr); library(tidyr)
library(StMoMo)
library(openxlsx)
library(forecast)  # auto.arima, ets, simulate
library(zoo)       # na.approx

# -------- 0) 옵션/경로 --------------------------------------------------
file_path         <- "C:\\Users\\mckim\\Desktop\\Mortality_Improvement\\0.rawdata\\Data_mortality_M_1995_2023_v1.xlsx"
out_xlsx          <- "C:\\Users\\mckim\\Desktop\\Mortality_Improvement\\1.result\\forecast_APC_simpaths_M.xlsx"

years_obs_target  <- 1995:2019
years_fut_target  <- 2020:2035
ages_use          <- 0:100

# κ/γ 예측 방법: "arima" 또는 "ets"
kt_method <- "arima"
gc_method <- "arima"

# 시뮬레이션 설정 (신뢰밴드용)
simulate_indices <- TRUE
nsim             <- 100     # 권장: 100 이상

set.seed(20250812)          # 재현성(옵션)

# -------- 2) 데이터 읽기 ------------------------------------------------
read_matrix <- function(file, sheet){
  df <- readxl::read_excel(file, sheet = sheet, col_names = TRUE)
  ages_chr <- as.character(df[[1]])
  m <- as.matrix(df[-1]); mode(m) <- "numeric"
  rownames(m) <- ages_chr
  colnames(m) <- names(df)[-1]
  m
}

Dxt_all <- read_matrix(file_path, "response")
Ext_all <- read_matrix(file_path, "dose")

# 교집합 정렬
ages_all   <- intersect(rownames(Dxt_all), rownames(Ext_all))
years_all  <- intersect(colnames(Dxt_all), colnames(Ext_all))
Dxt_all    <- Dxt_all[ages_all, years_all, drop=FALSE]
Ext_all    <- Ext_all[ages_all, years_all, drop=FALSE]

ages_all_n  <- as.integer(ages_all)
years_all_n <- as.integer(years_all)

obs_years <- years_obs_target[years_obs_target %in% years_all_n]
fut_years <- years_fut_target
stopifnot(length(obs_years) >= 3)

ages     <- ages_use[ages_use %in% ages_all_n]
ages_chr <- as.character(ages)

Dxt <- Dxt_all[ages_chr, as.character(obs_years), drop=FALSE]
Ext <- Ext_all[ages_chr, as.character(obs_years), drop=FALSE]

h <- length(fut_years)

# -------- 3) APC 적합 ---------------------------------------------------
wxt      <- StMoMo::genWeightMat(ages, obs_years, clip = 3)
APCmodel <- StMoMo::apc(link = "log")
APCfit   <- StMoMo::fit(APCmodel, Dxt = Dxt, Ext = Ext,
                        ages = ages, years = obs_years, wxt = wxt)

# -------- 4) κ, γ를 ARIMA/ETS "시뮬레이션 경로"로 예측 -------------------
# 한 경로(평균 or 시뮬) 생성
path_one <- function(x, h, method=c("arima","ets"), simulate=FALSE){
  method <- match.arg(method)
  x <- as.numeric(x)
  if(method=="arima"){
    fit <- forecast::auto.arima(x, stepwise=FALSE, approximation=FALSE)
    if(simulate){
      as.numeric(simulate(fit, nsim=h, future=TRUE))
    } else {
      as.numeric(forecast::forecast(fit, h=h)$mean)
    }
  } else {
    fit <- forecast::ets(x)
    if(simulate){
      as.numeric(simulate(fit, nsim=h, future=TRUE))
    } else {
      as.numeric(forecast::forecast(fit, h=h)$mean)
    }
  }
}

# nsim개 경로(열=시뮬) 생성 -> 행=h
paths_matrix <- function(x, h, method, simulate, nsim){
  if(!simulate || nsim<=1){
    mat <- matrix(path_one(x, h, method, simulate), nrow=h, ncol=1)
  } else {
    mat <- replicate(nsim, path_one(x, h, method, TRUE))
    if(is.null(dim(mat))) mat <- matrix(mat, nrow=h, ncol=1)
  }
  rownames(mat) <- NULL
  mat
}

# 관측 κ, γ / 코호트 인덱스
kt_hist <- as.numeric(APCfit$kt)
gc_hist <- as.numeric(APCfit$gc)
coh_hist<- as.integer(names(APCfit$gc))

# γ NA 보간(끝단 연장)
gc_hist_full <- zoo::na.approx(zoo::zoo(gc_hist, coh_hist), xout = coh_hist, rule = 2)
gc_hist_full <- as.numeric(gc_hist_full); names(gc_hist_full) <- coh_hist

# 미래에 필요한 cohort 범위
c_max_needed <- max(fut_years) - min(ages)
coh_max_obs  <- max(coh_hist, na.rm = TRUE)
h_coh        <- max(0, c_max_needed - coh_max_obs)

# κ_t / γ_c 경로: [h × nsim], [h_coh × nsim]
kt_fut_mat <- paths_matrix(kt_hist, h,    kt_method, simulate_indices, nsim)
if(h_coh > 0){
  gc_fut_mat <- paths_matrix(gc_hist_full, h_coh, gc_method, simulate_indices, nsim)
  rownames(gc_fut_mat) <- (coh_max_obs + 1):c_max_needed
} else {
  gc_fut_mat <- matrix(numeric(0), nrow=0, ncol=nsim)
}

# α_x
ax <- as.numeric(APCfit$ax[ages_chr]); names(ax) <- ages

# 시뮬별 rates 계산: 배열 [age × year × sim]
rates_arr <- array(NA_real_, dim=c(length(ages), h, ncol(kt_fut_mat)),
                   dimnames=list(ages_chr, fut_years, paste0("sim", seq_len(ncol(kt_fut_mat)))))

for(s in seq_len(ncol(kt_fut_mat))){
  kt_fut <- kt_fut_mat[, s]; names(kt_fut) <- fut_years
  
  gc_all <- c(gc_hist_full, if(nrow(gc_fut_mat)>0) as.numeric(gc_fut_mat[, s]) else numeric(0))
  names(gc_all) <- c(names(gc_hist_full), if(nrow(gc_fut_mat)>0) rownames(gc_fut_mat) else character(0))
  gc_all <- gc_all[as.character(sort(unique(as.integer(names(gc_all)))))]
  
  for(j in seq_along(fut_years)){
    yr <- fut_years[j]
    k  <- kt_fut[j]
    coh <- yr - ages
    gamma_vals <- gc_all[as.character(coh)]
    if(anyNA(gamma_vals)){   # 안전장치: 보간
      nz <- !is.na(gamma_vals)
      if(any(nz)){
        gamma_vals <- approx(x=which(nz), y=gamma_vals[nz],
                             xout=seq_along(gamma_vals), rule=2)$y
      } else gamma_vals[] <- 0
    }
    rates_arr[, j, s] <- exp(ax + k + gamma_vals)
  }
}

# 노출(Ext) 고정(보수적) → counts
ExtF_arr <- array(Ext[, ncol(Ext)], dim=c(length(ages), h, dim(rates_arr)[3]))
counts_arr <- rates_arr * ExtF_arr

# 요약 통계(P50/P05/P95). P50을 '베이스라인'으로 사용
qfun <- function(a, p){ apply(a, c(1,2), quantile, probs=p, na.rm=TRUE) }
rates_p50 <- qfun(rates_arr, 0.50); rates_p05 <- qfun(rates_arr, 0.05); rates_p95 <- qfun(rates_arr, 0.95)
counts_p50<- qfun(counts_arr,0.50); counts_p05<- qfun(counts_arr,0.05); counts_p95<- qfun(counts_arr,0.95)

# 베이스라인(central path) = P50
rates_base  <- rates_p50
counts_base <- counts_p50

# -------- 5) 관측구간 적합(참고) ----------------------------------------
fitted_rates  <- fitted(APCfit)
colnames(fitted_rates) <- as.character(obs_years)
fitted_counts <- fitted_rates * Ext

# -------- 6) 진단: stepwise ratio --------------------------------------
# A. 베이스라인 rates의 연속 연도 배율
if(ncol(rates_base) >= 2){
  g_rates <- rates_base[, -1, drop=FALSE] / rates_base[, -ncol(rates_base), drop=FALSE]
  colnames(g_rates) <- paste0(colnames(rates_base)[-1], "_over_", colnames(rates_base)[-length(colnames(rates_base))])
} else g_rates <- matrix(NA_real_, nrow=nrow(rates_base), ncol=0)

# B. period 공통 배율 exp(Δκ_t) (미래-미래만, 길이 h-1) — P05/P50/P95 제공
dkt_mat <- apply(kt_fut_mat, 2, diff)              # [h-1 × nsim]
if(is.null(dim(dkt_mat))) dkt_mat <- matrix(dkt_mat, nrow=h-1, ncol=1)
exp_dk  <- exp(dkt_mat)                            # [h-1 × nsim]
labels  <- paste0(fut_years[-1], "_over_", fut_years[-length(fut_years)])
qs <- apply(exp_dk, 1, quantile, probs=c(0.05,0.5,0.95), na.rm=TRUE)
period_step_ratio <- data.frame(
  year_step      = labels,
  exp_dkappa_p05 = as.numeric(qs[1,]),
  exp_dkappa_p50 = as.numeric(qs[2,]),
  exp_dkappa_p95 = as.numeric(qs[3,])
)

# -------- 7) Long 포맷 & 메타 -------------------------------------------
to_long <- function(mat, value_name){
  as.data.frame(mat) %>%
    mutate(age = as.integer(rownames(mat))) %>%
    pivot_longer(-age, names_to="year", values_to=value_name) %>%
    mutate(year = as.integer(year))
}

rates_P50_df  <- to_long(rates_p50,  "APC_rate_P50")
rates_P05_df  <- to_long(rates_p05,  "APC_rate_P05")
rates_P95_df  <- to_long(rates_p95,  "APC_rate_P95")
counts_P50_df <- to_long(counts_p50, "APC_count_P50")
counts_P05_df <- to_long(counts_p05, "APC_count_P05")
counts_P95_df <- to_long(counts_p95, "APC_count_P95")

g_rates_df    <- as.data.frame(g_rates)

kt_series_obs <- data.frame(year = obs_years, kappa_t = kt_hist)
# κ 미래 경로의 P05/P50/P95 (연도별)
kt_series_fut <- data.frame(
  year       = fut_years,
  kappa_t_P05= apply(kt_fut_mat, 1, quantile, probs=0.05),
  kappa_t_P50= apply(kt_fut_mat, 1, quantile, probs=0.50),
  kappa_t_P95= apply(kt_fut_mat, 1, quantile, probs=0.95)
)

# 새 코호트 γ 요약(P05/P50/P95)
if(nrow(gc_fut_mat) > 0){
  gc_series_fut <- data.frame(
    cohort      = as.integer(rownames(gc_fut_mat)),
    gamma_c_P05 = apply(gc_fut_mat, 1, quantile, probs=0.05),
    gamma_c_P50 = apply(gc_fut_mat, 1, quantile, probs=0.50),
    gamma_c_P95 = apply(gc_fut_mat, 1, quantile, probs=0.95)
  )
} else {
  gc_series_fut <- data.frame()
}
gc_series_obs <- data.frame(cohort = coh_hist, gamma_c = as.numeric(APCfit$gc))

pc_sd <- data.frame(Index=c("kt (obs)","gc (obs)"),
                    SD   =c(sd(kt_hist), sd(as.numeric(APCfit$gc), na.rm=TRUE)))

meta <- data.frame(
  item  = c("file_path","ages_used","years_obs","years_fut",
            "kt_method","gc_method","simulate_indices","nsim",
            "baseline_path","note"),
  value = c(file_path,
            paste0(min(ages),"..",max(ages)),
            paste0(min(obs_years),"..",max(obs_years)),
            paste0(min(fut_years),"..",max(fut_years)),
            kt_method, gc_method, simulate_indices, nsim,
            "P50 (median of simulated paths)",
            "rates = exp(ax + kappa + gamma); counts = rates * last-year Ext")
)

# -------- 8) 엑셀 저장 ---------------------------------------------------
wb <- createWorkbook()
addWorksheet(wb, "rates_P50_base")
addWorksheet(wb, "rates_P05")
addWorksheet(wb, "rates_P95")
addWorksheet(wb, "counts_P50_base")
addWorksheet(wb, "counts_P05")
addWorksheet(wb, "counts_P95")
addWorksheet(wb, "fitted_rates_obs")
addWorksheet(wb, "fitted_counts_obs")
addWorksheet(wb, "kt_series_obs")
addWorksheet(wb, "kt_series_fut")
addWorksheet(wb, "gc_series_obs")
addWorksheet(wb, "gc_series_fut")
addWorksheet(wb, "step_ratio_rates_base")
addWorksheet(wb, "step_ratio_period")
addWorksheet(wb, "meta")

writeData(wb, "rates_P50_base",   rates_P50_df)
writeData(wb, "rates_P05",        rates_P05_df)
writeData(wb, "rates_P95",        rates_P95_df)
writeData(wb, "counts_P50_base",  counts_P50_df)
writeData(wb, "counts_P05",       counts_P05_df)
writeData(wb, "counts_P95",       counts_P95_df)
writeData(wb, "fitted_rates_obs", to_long(fitted_rates, "fitted_rate"))
writeData(wb, "fitted_counts_obs",to_long(fitted_counts, "fitted_count"))
writeData(wb, "kt_series_obs",    kt_series_obs)
writeData(wb, "kt_series_fut",    kt_series_fut)
writeData(wb, "gc_series_obs",    gc_series_obs)
writeData(wb, "gc_series_fut",    gc_series_fut)
writeData(wb, "step_ratio_rates_base", g_rates_df, rowNames = TRUE)
writeData(wb, "step_ratio_period",     period_step_ratio)
writeData(wb, "meta",                  meta)

dir.create(dirname(out_xlsx), recursive = TRUE, showWarnings = FALSE)
saveWorkbook(wb, out_xlsx, overwrite = TRUE)
message("✅ Complete. Excel saved: ", out_xlsx)

# -------- 9) (선택) 시각화 ----------------------------------------------
plot(APCfit, param="kt"); plot(APCfit, param="gc")
