# ======================================================================
# apc_arima_glm_forecast_FULL.R
# - Data: Excel with sheets "response"(deaths Dxt), "dose"(exposure Ext)
# - Ages: 0..100 rows, Years as column names (e.g., 1995..2023)
# - Models:
#    * APC (StMoMo) -> forecast rates -> counts (Ext_last * rates)
#    * ARIMA (no xreg / +year / +year^2)
#    * GLM (Poisson with log(Ext) offset; 1차/2차)
# - Diagnostics:
#    * Period/Cohort index plots, SD(kt)/SD(gc), AIC(Lc vs APC vs RH)
# - Output: one Excel with multi-sheets
# ======================================================================

# -------- 1) 패키지 -----------------------------------------------------
need <- c("readxl","dplyr","tidyr","StMoMo","openxlsx")
for(p in need){
  if(!requireNamespace(p, quietly = TRUE)) install.packages(p)
}
library(readxl); library(dplyr); library(tidyr)
library(StMoMo)
library(openxlsx)

# -------- 0) 옵션/경로 --------------------------------------------------
# 데이터 파일명(남성 예시): 작업 폴더에 파일을 둬 주세요.
file_path   <- "C:/Users/sdppa/Desktop/제안논문/01_rawdata/Data_mortality_M_1995_2023_v1.xlsx"

# 관측(학습)·예측 기간
years_obs_target  <- 1995:2015
years_fut_target  <- 2016:2024

# 사용할 연령대 (극단연령 잡음 줄이려면 20:89 권장)
ages_use          <- 0:100

# 출력 파일명
out_xlsx          <- "C:/Users/sdppa/Desktop/제안논문/02_result/forecast_APC_only_M.xlsx"

# -------- 2) 데이터 읽기 ------------------------------------------------
read_matrix <- function(file, sheet){
  df <- readxl::read_excel(file, sheet = sheet, col_names = TRUE)
  ages_chr <- as.character(df[[1]])
  m <- as.matrix(df[-1]); mode(m) <- "numeric"
  rownames(m) <- ages_chr
  colnames(m) <- names(df)[-1]
  m
}

Dxt_all <- read_matrix(file_path, "response")  # deaths
Ext_all <- read_matrix(file_path, "dose")      # exposures

# 교집합 정렬
ages_all   <- intersect(rownames(Dxt_all), rownames(Ext_all))
years_all  <- intersect(colnames(Dxt_all), colnames(Ext_all))
Dxt_all    <- Dxt_all[ages_all, years_all, drop=FALSE]
Ext_all    <- Ext_all[ages_all, years_all, drop=FALSE]

ages_all_n <- as.integer(ages_all)
years_all_n<- as.integer(years_all)

# 학습·예측 기간 보정
obs_years  <- years_obs_target[years_obs_target %in% years_all_n]
fut_years  <- years_fut_target
stopifnot(length(obs_years) >= 3)  # 최소 3년 이상 권장

# 연령 필터
ages       <- ages_use[ages_use %in% ages_all_n]
ages_chr   <- as.character(ages)

# 슬라이스(학습 구간)
Dxt <- Dxt_all[ages_chr, as.character(obs_years), drop=FALSE]
Ext <- Ext_all[ages_chr, as.character(obs_years), drop=FALSE]

# 예측 기간 길이
h <- length(fut_years)

# -------- 3) APC 적합 ---------------------------------------------------
# 극단연령 가중치 완화(권장)
wxt <- StMoMo::genWeightMat(ages, obs_years, clip = 3)

APCmodel <- StMoMo::apc(link = "log")
APCfit   <- StMoMo::fit(APCmodel, Dxt = Dxt, Ext = Ext,
                        ages = ages, years = obs_years, wxt = wxt)

# -------- 4) 예측(사망률) & 사망자수 변환 -------------------------------
APCfor    <- forecast(APCfit, h = h)
rates_fc  <- APCfor$rates                     # matrix: ages x h (사망률)
colnames(rates_fc) <- fut_years
rownames(rates_fc) <- ages_chr

# 노출(Ext) 전망이 없으면, 마지막 관측년도 노출을 고정 사용(보수적)
ExtF <- matrix(Ext[, ncol(Ext)], nrow = length(ages), ncol = h,
               dimnames = list(ages_chr, fut_years))

Dhat_apc <- rates_fc * ExtF                  # 예측 사망자수(ages x years)

# (선택) 학습구간 적합 사망률/사망자수도 보고 싶으면:
fitted_rates <- fitted(APCfit)                # ages x |obs_years|
colnames(fitted_rates) <- as.character(obs_years)
fitted_counts <- fitted_rates * Ext           # 적합 사망자수

# -------- 5) Long 포맷 & 인덱스 내보내기 -------------------------------
# Forecasted rates
rates_df <- as.data.frame(rates_fc) %>%
  mutate(age = as.integer(rownames(rates_fc))) %>%
  pivot_longer(-age, names_to="year", values_to="APC_rate") %>%
  mutate(year = as.integer(year))

# Forecasted counts
counts_df <- as.data.frame(Dhat_apc) %>%
  mutate(age = as.integer(rownames(Dhat_apc))) %>%
  pivot_longer(-age, names_to="year", values_to="APC_count") %>%
  mutate(year = as.integer(year))

# (옵션) fitted values (관측구간)
fitted_rates_df <- as.data.frame(fitted_rates) %>%
  mutate(age = as.integer(rownames(fitted_rates))) %>%
  pivot_longer(-age, names_to="year", values_to="fitted_rate") %>%
  mutate(year = as.integer(year))

fitted_counts_df <- as.data.frame(fitted_counts) %>%
  mutate(age = as.integer(rownames(fitted_counts))) %>%
  pivot_longer(-age, names_to="year", values_to="fitted_count") %>%
  mutate(year = as.integer(year))

# APC 지표(Period/Cohort) 진단용
kt_series <- data.frame(year = obs_years, kappa_t = as.numeric(APCfit$kt))
gc_series <- data.frame(cohort = as.integer(names(APCfit$gc)),
                        gamma_c = as.numeric(APCfit$gc))
pc_sd     <- data.frame(Index = c("kt (Period)", "gc (Cohort)"),
                        SD    = c(sd(APCfit$kt), sd(APCfit$gc, na.rm = TRUE)))

# 메타
meta <- data.frame(
  item  = c("file_path","ages_used","years_obs","years_fut",
            "note_rates_to_counts"),
  value = c(file_path,
            paste0(min(ages),"..",max(ages)),
            paste0(min(obs_years),"..",max(obs_years)),
            paste0(min(fut_years),"..",max(fut_years)),
            "Dhat = forecasted rates * last observed exposures (ExtF)")
)

# -------- 6) 엑셀 저장 ---------------------------------------------------
wb <- createWorkbook()
addWorksheet(wb, "forecast_rates_APC")
addWorksheet(wb, "forecast_counts_APC")
addWorksheet(wb, "fitted_rates_obs")
addWorksheet(wb, "fitted_counts_obs")
addWorksheet(wb, "APC_kappa_t")
addWorksheet(wb, "APC_gamma_c")
addWorksheet(wb, "period_cohort_SD")
addWorksheet(wb, "meta")

writeData(wb, "forecast_rates_APC",  rates_df)
writeData(wb, "forecast_counts_APC", counts_df)
writeData(wb, "fitted_rates_obs",    fitted_rates_df)
writeData(wb, "fitted_counts_obs",   fitted_counts_df)
writeData(wb, "APC_kappa_t",         kt_series)
writeData(wb, "APC_gamma_c",         gc_series)
writeData(wb, "period_cohort_SD",    pc_sd)
writeData(wb, "meta",                meta)

saveWorkbook(wb, out_xlsx, overwrite = TRUE)
message("✅ APC forecast complete. Excel saved: ", out_xlsx)

# -------- 7) (선택) 시각화 ----------------------------------------------
# plot(APCfit, param="kt")  # Period index κ_t
# plot(APCfit, param="gc")  # Cohort index γ_c
# plot(APCfit)              # α_x, κ_t, γ_c 종합
