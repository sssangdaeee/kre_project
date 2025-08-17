# -----------------------------
# Libraries
# -----------------------------
req_pkgs <- c("readxl", "dplyr", "tidyr", "purrr", "forecast", "Metrics", "stringr", "readr", "tibble")
for (p in req_pkgs) {
  if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(p, character.only = TRUE)
}

set.seed(42)

# -----------------------------
# 설정
# -----------------------------
excel_path <- "C:/Users/sdpark/Desktop/자료/제안논문/0_rawdata/ML_Input_v1_변수추정용.xlsx"

train_years <- 2010:2020
test_years  <- 2021:2024
fcst_years  <- 2025:2035

# -----------------------------
# 유틸
# -----------------------------
clamp01 <- function(x) pmin(pmax(x, 0), 1)

safe_rmse <- function(actual, pred) {
  keep <- !(is.na(actual) | is.na(pred))
  Metrics::rmse(actual[keep], pred[keep])
}
safe_mae <- function(actual, pred) {
  keep <- !(is.na(actual) | is.na(pred))
  Metrics::mae(actual[keep], pred[keep])
}

# -----------------------------
# 데이터 로드
# -----------------------------
sheets <- c("gender", "bloodtest", "uwclass")
data_list <- lapply(sheets, function(s) readxl::read_excel(excel_path, sheet = s))
names(data_list) <- sheets

# -----------------------------
# 단일 시계열 모델링 & 비교
# -----------------------------
fit_and_compare_models <- function(df, sheet_name, var_name) {
  df <- df %>% dplyr::select(year, !!var_name := dplyr::all_of(var_name)) %>% arrange(year)
  colnames(df) <- c("year", "y")
  
  train_df <- df %>% filter(year %in% train_years)
  test_df  <- df %>% filter(year %in% test_years)
  
  ts_train <- ts(train_df$y, start = min(train_df$year), frequency = 1)
  
  # 비율형 자동 감지 (0~1 범위)
  is_ratio <- (min(df$y, na.rm = TRUE) >= 0) && (max(df$y, na.rm = TRUE) <= 1)
  
  # 1) 선형회귀 (LM)
  lm_fit <- lm(y ~ year, data = train_df)
  lm_pred_test <- predict(lm_fit, newdata = test_df)
  if (is_ratio) lm_pred_test <- clamp01(lm_pred_test)
  
  # 2) ARIMA (auto.arima)
  arima_fit <- forecast::auto.arima(ts_train, stepwise = FALSE, approximation = FALSE)
  arima_fc_test <- forecast::forecast(arima_fit, h = length(test_years))
  arima_pred_test <- as.numeric(arima_fc_test$mean)
  if (is_ratio) arima_pred_test <- clamp01(arima_pred_test)
  
  # 3) GLM (비율: quasibinomial-logit, 그 외: gaussian)
  if (is_ratio) {
    eps <- 1e-6
    train_df$y_clamped <- pmin(pmax(train_df$y, eps), 1 - eps)
    glm_fit <- glm(y_clamped ~ year, data = train_df,
                   family = quasibinomial(link = "logit"),
                   weights = rep(1000, nrow(train_df)))
    glm_pred_test <- predict(glm_fit, newdata = test_df, type = "response")
    glm_pred_test <- clamp01(glm_pred_test)
  } else {
    glm_fit <- glm(y ~ year, data = train_df, family = gaussian(link = "identity"))
    glm_pred_test <- predict(glm_fit, newdata = test_df, type = "response")
  }
  
  # 성능지표 계산 (테스트 기준)
  lm_rmse <- safe_rmse(test_df$y, lm_pred_test)
  lm_mae  <- safe_mae(test_df$y, lm_pred_test)
  
  arima_rmse <- safe_rmse(test_df$y, arima_pred_test)
  arima_mae  <- safe_mae(test_df$y, arima_pred_test)
  
  glm_rmse <- safe_rmse(test_df$y, glm_pred_test)
  glm_mae  <- safe_mae(test_df$y, glm_pred_test)
  
  perf <- tibble::tibble(
    sheet    = sheet_name,
    variable = var_name,
    model    = c("LM", "ARIMA", "GLM"),
    RMSE     = c(lm_rmse, arima_rmse, glm_rmse),
    MAE      = c(lm_mae,  arima_mae,  glm_mae)
  )
  
  # 최적 모델 선택: RMSE ⬇, MAE ⬇
  perf <- perf %>% arrange(RMSE, MAE)
  best_model <- perf$model[1]
  
  # 2025~2035 예측 (훈련 기반 그대로)
  if (best_model == "LM") {
    newdf <- data.frame(year = fcst_years)
    pred  <- predict(lm_fit, newdata = newdf)
    if (is_ratio) pred <- clamp01(pred)
  } else if (best_model == "ARIMA") {
    arima_fc <- forecast::forecast(arima_fit, h = length(fcst_years))
    pred <- as.numeric(arima_fc$mean)
    if (is_ratio) pred <- clamp01(pred)
  } else { # GLM
    newdf <- data.frame(year = fcst_years)
    pred <- predict(glm_fit, newdata = newdf, type = if (is_ratio) "response" else "response")
    if (is_ratio) pred <- clamp01(pred)
  }
  
  fcst <- tibble::tibble(
    sheet    = sheet_name,
    variable = var_name,
    model    = best_model,
    year     = fcst_years,
    forecast = as.numeric(pred)
  )
  
  list(perf = perf, fcst = fcst)
}

# -----------------------------
# 전체 루프
# -----------------------------
all_perf <- list(); all_fcst <- list()

for (s in names(data_list)) {
  df <- data_list[[s]]
  stopifnot("year" %in% names(df))
  var_cols <- setdiff(names(df), "year")
  for (v in var_cols) {
    res <- fit_and_compare_models(df, s, v)
    all_perf[[length(all_perf) + 1]] <- res$perf
    all_fcst[[length(all_fcst) + 1]] <- res$fcst
  }
}

perf_tbl <- dplyr::bind_rows(all_perf)
fcst_tbl <- dplyr::bind_rows(all_fcst)

# -----------------------------
# 저장 & 요약
# -----------------------------
readr::write_csv(perf_tbl, "C:/Users/sdpark/Desktop/자료/제안논문/0_rawdata/model_performance.csv")
readr::write_csv(fcst_tbl, "C:/Users/sdpark/Desktop/자료/제안논문/0_rawdata/forecasts_2025_2035.csv")

cat("\n[각 시계열의 최상 모델 (기준: RMSE -> MAE)]\n")
best_summary <- perf_tbl %>%
  group_by(sheet, variable) %>%
  arrange(RMSE, MAE, .by_group = TRUE) %>%
  slice(1) %>%
  ungroup()
print(best_summary)

cat("\nCSV 파일로 저장됨:\n",
    "- model_performance.csv (RMSE/MAE 비교)\n",
    "- forecasts_2025_2035.csv (최적 모델 예측 2025–2035)\n")
