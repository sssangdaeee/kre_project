# apc_arima_glm_forecast.R
# 0–100세 1995–2015 관측 데이터로
#  • APC (StMoMo) → 2016–2024 사망자수
#  • ARIMA(0/1/2차 추세) & GLM(1/2차) → 2016–2024 사망자수
# 을 모두 계산하고, 결과를 한꺼번에 Excel 파일로 출력합니다.

# 0. 작업 디렉터리 설정 (RStudio 전용) ---------------------------------------
library(rstudioapi)
rstudioapi::getActiveDocumentContext()$path
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))
source("Model_Library_v02.R")  # 필요 시

# 1. 패키지 로드 -------------------------------------------------------------
library(readxl)
library(dplyr)
library(tidyr)
library(apc)
library(StMoMo)
library(forecast)
library(openxlsx)

# 2. 데이터 읽기 함수 정의 --------------------------------------------------
read_matrix <- function(file, sheet) {
  df   <- read_excel(file, sheet = sheet, col_names = TRUE,
                     .name_repair = "unique_quiet")
  ages <- df[[1]] %>% as.character()
  m    <- df[-1] %>% as.matrix() %>% apply(2, as.numeric)
  rownames(m) <- ages
  colnames(m) <- names(df)[-1]
  m
}

# 3. 관측 데이터 불러오기 ----------------------------------------------------
setwd("C:/Users/sdpark/Desktop/자료/제안논문")
file_path  <- "./0_rawdata/Data_mortality_F_2010_2024_v2_자사통계_그룹.xlsx"
Dxt_all    <- read_matrix(file_path, "response")  # 사망자수(ages×years)
Ext_all    <- read_matrix(file_path, "dose")      # 노출(ages×years)

years_use  <- 2011:2023
Dxt        <- Dxt_all[, as.character(years_use)]  # 사망자수(ages×years)
Ext        <- Ext_all[, as.character(years_use)]  # 노출(ages×years)

ages       <- as.integer(rownames(Dxt))          # 0:100
years_obs  <- as.integer(colnames(Dxt))          # 1995:2015
years_fut  <- 2024:2035                          # 예측연도
h          <- length(years_fut)

# # 4. APC 예측 (StMoMo) -------------------------------------------------------
# # 4-1) 모델 정의 및 적합
# APCmodel  <- apc(link = "log")
# APCfit    <- fit(
#   APCmodel,
#   Dxt    = Dxt,
#   Ext    = Ext,
#   ages   = ages,
#   years  = years_obs
# )
# # 4-2) 예측 (2016–2024)
# APCfor    <- forecast(APCfit, h = h)
# rates_fc  <- APCfor$rates      # 사망률 직사각형(ages×h)
# ExtF      <- matrix(Ext[,ncol(Ext)], nrow=length(ages), ncol=h)
# Dhat_apc  <- rates_fc * ExtF   # 사망자수 예측 매트릭스
# colnames(Dhat_apc) <- years_fut

# 4. APC 예측 (StMoMo) -------------------------------------------------------
# 4-0) 라벨과 인덱스를 분리
age_lbl   <- as.integer(rownames(Dxt))          # c(30, 45, 55, 65, 80)
year_lbl  <- as.integer(colnames(Dxt))          # 2010:2020 (현재 사용 구간)
ages_idx  <- seq_len(nrow(Dxt))                 # 1:5  ← 연속 정수!
years_idx <- seq_len(ncol(Dxt))                 # 1:11 (연속 정수)

# 4-1) 모델 정의 및 적합  (패키지 충돌 방지: apc는 StMoMo의 것을 사용)
APCmodel  <- StMoMo::apc(link = "log")
APCfit    <- StMoMo::fit(
  APCmodel,
  Dxt    = Dxt,
  Ext    = Ext,
  ages   = ages_idx,     # 연속 정수 인덱스로!
  years  = years_idx
)

# 4-2) 예측
APCfor    <- forecast(APCfit, h = h)

# 4-3) 출력 라벨을 원래 연령/연도로 복구
rates_fc  <- APCfor$rates
rownames(rates_fc) <- age_lbl
colnames(rates_fc) <- years_fut

# 4-4) 사망자수로 변환
ExtF      <- matrix(Ext[, ncol(Ext)], nrow = nrow(Ext), ncol = h)
Dhat_apc  <- rates_fc * ExtF

# 5. ARIMA & GLM 예측 --------------------------------------------------------
# 5-1) 빈 매트릭스 생성
preds <- list(
  arima0 = matrix(NA, length(ages), h, dimnames=list(age=ages, year=years_fut)),
  arima1 = matrix(NA, length(ages), h, dimnames=list(age=ages, year=years_fut)),
  arima2 = matrix(NA, length(ages), h, dimnames=list(age=ages, year=years_fut)),
  glm1   = matrix(NA, length(ages), h, dimnames=list(age=ages, year=years_fut)),
  glm2   = matrix(NA, length(ages), h, dimnames=list(age=ages, year=years_fut))
)

# 5-2) 나이별 모델링 & 예측 루프
for(i in seq_along(ages)) {
  age_str   <- as.character(ages[i])
  y_obs     <- Dxt[age_str, ]
  ts_full   <- ts(y_obs, start=min(years_obs), frequency=1)
  
  # ARIMA0: 기본 auto.arima
  fit0                <- auto.arima(ts_full)
  preds$arima0[i, ]   <- as.numeric(forecast(fit0, h=h)$mean)
  
  # ARIMA1: xreg = year
  fit1                <- auto.arima(ts_full, xreg = years_obs)
  preds$arima1[i, ]   <- as.numeric(forecast(fit1, h=h, xreg=years_fut)$mean)
  
  # ARIMA2: xreg = cbind(year, year^2)
  fit2                <- auto.arima(ts_full, xreg = cbind(years_obs, years_obs^2))
  preds$arima2[i, ]   <- as.numeric(forecast(fit2, h=h,
                                             xreg=cbind(years_fut, years_fut^2))$mean)
  
  # GLM(1차): Poisson log link
  df_age              <- tibble(year=years_obs, count=y_obs)
  fit_g1              <- glm(count ~ year, family=poisson(link="log"), data=df_age)
  preds$glm1[i, ]     <- predict(fit_g1, newdata=tibble(year=years_fut), type="response")
  
  # GLM(2차): Poisson log link
  fit_g2              <- glm(count ~ poly(year,2,raw=TRUE),
                             family=poisson(link="log"), data=df_age)
  preds$glm2[i, ]     <- predict(fit_g2, newdata=tibble(year=years_fut), type="response")
}

# 6. 결과 결합 & 엑셀 저장 ---------------------------------------------------
# 6-1) APC_df 만들기
apc_df <- as.data.frame(Dhat_apc) %>%
  mutate(age=ages) %>%
  pivot_longer(-age, names_to="year", values_to="APC") %>%
  mutate(year=as.integer(year))

# 6-2) ARIMA/GLM long 포맷으로
df_list <- lapply(names(preds), function(m) {
  as.data.frame(preds[[m]]) %>%
    mutate(age=ages) %>%
    pivot_longer(-age, names_to="year", values_to=m) %>%
    mutate(year=as.integer(year))
})
library(dplyr)
forecast_compare <- reduce(df_list, left_join, by=c("age","year")) %>%
  left_join(apc_df, by=c("age","year"))

# 6-3) 엑셀 쓰기
write.xlsx(forecast_compare,
           file="C:/Users/sdpark/Desktop/자료/제안논문/2_result/자사통계_그룹/forecast_all_models_자사통계_F_exc2010_2035추정.xlsx")

# 7. 마무리 ---------------------------------------------------------------
message("Forecasting complete: APC + ARIMA + GLM results saved.")
