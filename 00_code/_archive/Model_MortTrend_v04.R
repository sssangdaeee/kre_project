# Set working directory to the source file
# Only works in RStudio
rstudioapi::getActiveDocumentContext()$path
setwd(dirname(rstudioapi::getActiveDocumentContext()$path))

source("Model_Library_v02.R")

library(readxl) # 이 패키지 로드가 잘 안되는 것 같아서, 여기서 로드 다시 함
library(apc) # APC 패키지 로드


###########################################################
# User functions
###########################################################

# Read function
readdata <- function(file, sheet="response"){ # file 읽을 엑셀 파일 경로, sheet 읽을 엑셀 시트 이름 기본값은 "response"
  rownames.tmp <- 
    read_excel(file, sheet = sheet, col_names = TRUE, .name_repair = "unique_quiet") %>% 
    select(1) %>%
    unlist()
  # 첫 번째 열이 행 이름(rownames.tmp) 으로 저장됨. read_excel 로 엑셀파일을 읽고, 지정된 sheet 에서 첫번재 열만 선택(select(1)) 을 하고, unlist() 를 사용해 벡터 형태로 변환-> 나중에 행 이름으로 사용
  
  data.name <- paste("data.cnc", ".", sheet, sep = "")
  # 문자열을 결합하여 새로운 변수명을 동적으로 생성. paste 함수는 여러 개의 문자열을 이어 붙이는 함수
  # 엑셀 시트별로 데이터를 구분하기 위해 해당 함수를 사용
  
  
  data.tmp <- 
    read_excel(file, sheet = sheet, col_names = TRUE, .name_repair = "unique_quiet") %>% 
    select(-1) %>% 
    as.matrix()
  # 엑셀을 읽어서 data.frame 으로 가져옴. 첫번째 열을 삭제하고, 숫자로 이루어진 행렬(as.matrix())로 변환 
  
  rownames(data.tmp) <- rownames.tmp
  colnames.tmp <- colnames(data.tmp)
  assign(data.name, data.tmp, envir = .GlobalEnv)
  
  colnames.name <- paste("colnames.cnc", ".", sheet, sep = "")
  rownames.name <- paste("rownames.cnc", ".", sheet, sep = "")
  assign(colnames.name, colnames.tmp, envir = .GlobalEnv)
  assign(rownames.name, rownames.tmp, envir = .GlobalEnv)
  
  return(list(data=get(data.name), colnames=get(colnames.name), rownames=get(rownames.name)))
}
# Read function 종료


# Excel 데이터 중 연령 구간의 대표 값을 가져오게 하는 함수, 예를 들어 10-15 => 10으로 저장
extract_left_first_any_symbol <- function(text, symbols = c("-", "+")) {
  pattern <- paste0("[", paste(symbols, collapse = ""), "]")  # 특수기호 패턴 생성
  parts <- sapply(text, function(s) strsplit(s, pattern, perl = TRUE)[[1]][1])
  return(as.numeric(parts))  # 첫 번째 부분 반환
}
extract_right_last_any_symbol <- function(text, symbols = c("-", "+")) {
  pattern <- paste0("[", paste(symbols, collapse = ""), "]")  # 특수기호 패턴 생성
  parts <- sapply(text, function(s) {
    parts <- strsplit(s, pattern, perl = TRUE)[[1]]
    if (length(parts) > 1) parts[length(parts)] else NA  # 마지막 부분 반환
  })
  return(as.numeric(parts))
}


###########################################################
# Data Read: Lung
###########################################################

# # Data: Responses of mortality cancer of women in 1-year ages and 1-year periods
readdata("../2_rawdata/Data_mortality_M_2008_2023_v1.xlsx", sheet="response")
readdata("../2_rawdata/Data_mortality_M_2008_2023_v1.xlsx", sheet="dose")
readdata("../2_rawdata/Data_mortality_M_2008_2023_v1.xlsx", sheet="rate")

# Data.cnc.response (발생건수 정수형으로 변환)(SD)
# data.cnc.response <- round(data.cnc.response)

###########################################################
# Calculate data parameters
###########################################################

# 5세 단위 작업하는 경우 필요한 코드

#연령/기간 구간의 시작 부분을 추출
agegrp.start <- extract_left_first_any_symbol(rownames(data.cnc.response), symbols=c("-","+"))
pergrp.start <- extract_left_first_any_symbol(colnames(data.cnc.response), symbols=c("-","+"))

#연령/기간 구간의 종료 부분을 추출, 1세/1년 단위의 경우 NA 가 반환됨
agegrp.end <- extract_right_last_any_symbol(rownames(data.cnc.response), symbols=c("-","+"))
pergrp.end <- extract_right_last_any_symbol(colnames(data.cnc.response), symbols=c("-","+"))

#연령/기간/코호트의 기간을 반환, 
unit.age <- agegrp.start[2]-agegrp.start[1]
unit.per <- pergrp.start[2]-pergrp.start[1]
unit.coh <- min(unit.age, unit.per)

# agegrp.end <- agegrp.start+unit.age
# pergrp.end <- pergrp.start+unit.per

max.age.I <- nrow(data.cnc.response) # 연령 개수
max.per.J <- ncol(data.cnc.response) # 기간 개수
max.coh.K <- max.age.I+max.per.J-1 # 코호트 개수

min.age <- agegrp.start[1] #최소연령
# max.age <- min.age + unit.age*(max.age.I-1)
# max.age <- agegrp.end[length(agegrp.end)-1]

max.age <- agegrp.start[length(agegrp.start)] #최대연령
# min.per <- as.numeric(substr(colnames.cnc.response,1,4)[1])

min.per <- pergrp.start[1] #기간 시작 연도
# max.per <- min.per + unit.per*(max.per.J-1)
# max.per <- pergrp.end[length(pergrp.end)-1]
max.per <- pergrp.start[length(pergrp.start)] #기간 종료 연도
min.coh <- min.per - max.age #최초 코호트 태어난 연도, 1999년의 100세는 1899년생
# max.coh <- min.coh + unit.coh*(max.coh.K-1)
max.coh <- max.per - min.age #최종 코호트 태어난 연도, 2022년의 0세는 2022년생


###########################################################
# Data transform from age-period to cohort-period
# cohort양식은 평행사변형으로 변형해주는 것을 의미함 (행의 의미는 ~년도생, 컬럼의 의미는 ~년에 발생자수)(SD)
###########################################################

data.NA <- matrix(NA,nrow=max.per.J-1,ncol=max.per.J)
data.tmp <- rbind(rbind(data.NA,data.cnc.response),data.NA)

# Data 를 Cohort 양식으로 바꾸어 주는 것, 연령/연도별 발생자수 => 코호트/연도별 발생자수로 정리
data.cnc.response.cp <- matrix(nrow=max.coh.K, ncol=max.per.J)
for (k in 1:max.coh.K){
  for (j in 1:max.per.J){
    # data.cnc.response.cp <- data.cnc.response[row(data.cnc.response) == (col(data.cnc.response)+nrow(data.cnc.response)-k)]
    data.cnc.response.cp[k,j] <- data.tmp[max.age.I+max.per.J-k+j-1,j]
  }
}
colnames(data.cnc.response.cp) <- colnames.cnc.response
rownames(data.cnc.response.cp) <- paste(seq(min.coh, max.coh, unit.coh), seq(min.coh, max.coh, unit.coh)+unit.coh-1, sep="-")


# Data format change
# apc.data.list 는 R에서 APC 모델을 실행할 데이터를 구성하는 함수
# 반올림 처리 추가(SD)
data.cnc <- apc.data.list(response= data.cnc.response,
                          dose=data.cnc.dose,
                          data.format="AP", # 데이터의 형식을 지정, AP 는 Age Period 모델을 의미
                          age1=min.age, per1= min.per, unit=unit.per) #분석기간이 시작되는 연령/기간 및 기간의 단위를 지정


###########################################################
# Fit
###########################################################

# dose-reponse model
# apc.fit.table(data.cnc,"poisson.dose.response")
# APC(Age Period Cohort)

# reponse model
# 연령,기간,코호트 모델을 각각 적합한 후 모델비교(APC, AP, AC 등등)를 위한 통계를 테이블 형식으로 출력. 
# 결과를 통해 어떤 모델이 데이터에 가장 적합한지 평가할 수 있음
apc.fit.table(data.cnc,"poisson.response")


# Fit result for Poisson dose-response model, 
# 1. The APC structure is the best
# 2. The AC structure is the second best

# fit.cnc.apc <- apc.fit.model(data.cnc,"poisson.dose.response","APC")
# fit.cnc.ac <- apc.fit.model(data.cnc,"poisson.dose.response","AC")

# Fit result for Poisson response model
# 1. The APC structure is the best
# 2. The AC structure is the second best

# 포아송 회귀 모델을 사용하여, APC/AC모델을 데이터에 적합 (Fit) 시킴

rm(fit.cnc.apc, fit.cnc.ac) # 기존 모델 삭제

fit.cnc.apc <- apc.fit.model(data.cnc,"poisson.response","APC")
fit.cnc.ac <- apc.fit.model(data.cnc,"poisson.response","AC")
# We will focus on the Poisson response model


# Basic plot of the fit
graphics.off()

# 모델이 실제 데이터를 얼마나 잘 설명하는지 그래프로 확인
# 연령, 기간, 코호트 별 변동성, 발생률 등등 확인 가능 (SD)
apc.plot.fit(fit.cnc.apc, main.outer="Figure: APC model of cancer")
apc.plot.fit(fit.cnc.ac, main.outer="Figure: AC model of cancer")

# Plot of probability transform of responses given fitted values
# Probability Transform (확률 변환)을 적용하여 모델 적합도를 확인
# 모델이 실제 값과 얼마나 잘 적합하는지 보여줌(검정색이면 거의 일치) (SD)
apc.plot.fit.pt(fit.cnc.apc, main="Probability transform of responses (APC)")
apc.plot.fit.pt(fit.cnc.ac, main="Probability transform of responses (AC)")


# All plots of fit
graphics.off()
apc.plot.fit.all(fit.cnc.apc)

graphics.off()
apc.plot.fit.all(fit.cnc.ac)

graphics.off()


# Parameter estimates
#APC 모델에서 추정된 정준 회귀 계수(β 값)**를 출력하는 코드입니다.
#연령(Age), 기간(Period), 출생 코호트(Cohort) 각각의 효과를 정량적으로 확인할 수 있습니다.
#이를 활용하면 특정 연령대, 특정 시기, 특정 세대에서 질병 발생률이 증가하거나 감소하는 패턴을 분석할 수 있습니다.

fit.cnc.apc$coefficients.canonical
fit.cnc.ac$coefficients.canonical


### 1차 차분 계산 및 시각화 코드 수정 ###
# 대상: APC 모델 추정 결과에서 연령, 기간, 코호트 효과에 대한 1차 차분 계산

# 추정된 canonical coefficients는 atomic vector로, 순서대로 age, period, cohort 효과가 저장됨
coefs <- fit.cnc.apc$coefficients.canonical

# 각 효과의 길이를 파악
n_age <- max.age.I
n_per <- max.per.J
n_coh <- length(coefs) - n_age - n_per

# 연령, 기간, 코호트 효과 분리
alpha <- coefs[1:n_age]
beta  <- coefs[(n_age+1):(n_age+n_per)]
gamma <- coefs[(n_age+n_per+1):(n_age+n_per+n_coh)]

# 1차 차분 계산
alpha.diff <- diff(alpha)
beta.diff  <- diff(beta)
gamma.diff <- diff(gamma)

# X축 레이블 생성 (원본 그룹 시작점 벡터 사용)
# agegrp.start, pergrp.start, cohort start 계산을 활용
age.label <- agegrp.start[-1]                  # 원본 연령 그룹 시작 벡터에서 첫번째 이후
per.label <- pergrp.start[-1]                  # 원본 기간 그룹 시작 벡터에서 첫번째 이후
coh.starts <- seq(min.coh, max.coh, by=unit.coh)  # 코호트 시작값 전체
coh.label <- coh.starts[-1]                     # 첫번째 이후

# 길이 일치 확인
length(alpha.diff) == length(age.label)
length(beta.diff)  == length(per.label)
length(gamma.diff) == length(coh.label)


###########################################################
# Forecast with the apc structure
###########################################################

# Forecast of Poisson response model with the apc structure
forecast.cnc.apc <- apc.forecast.apc(fit.cnc.apc)
forecast.cnc.apc$response.forecast.age
forecast.cnc.apc$response.forecast.per
forecast.cnc.apc$response.forecast.coh
forecast.cnc.apc$response.forecast.all


# Collect all forecast point estimates
Table.forecast.cnc.apc <- forecast.cnc.apc$response.forecast.age
Table.forecast.cnc.apc <- rbind(Table.forecast.cnc.apc, forecast.cnc.apc$response.forecast.per)
Table.forecast.cnc.apc <- rbind(Table.forecast.cnc.apc, forecast.cnc.apc$response.forecast.coh)
Table.forecast.cnc.apc <- rbind(Table.forecast.cnc.apc, forecast.cnc.apc$response.forecast.all)
Table.forecast.cnc.apc

#csv 파일로 내리기
write.csv(Table.forecast.cnc.apc, file = "forecast_cnc_apc_mortality.csv")

forecast.cnc.apc$trap.linear.predictors.forecast


#APC 모델에 비해 AC 모형의 실행 시간이 더 오래 걸림
#APC 모델은 기간(Period)을 이용하여 직접 예측할 수 있어 계산량이 비교적 일정함.
#AC 모델은 기간이 없기 때문에 출생 코호트(Cohort)별로 모든 연령을 다시 계산해야 해서 연산량이 기하급수적으로 증가함.
#미래 데이터를 예측할 때, APC 모델은 단순히 새로운 기간을 추가하면 되지만, AC 모델은 새로운 출생 코호트와 연령을 모두 고려해야 해서 계산이 더 어려움.
###########################################################
# Forecast with the ac structure
###########################################################
# Forecast of Poisson response model with the ac structure
forecast.cnc.ac <- apc.forecast.ac(fit.cnc.ac)
forecast.cnc.ac$response.forecast.age
forecast.cnc.ac$response.forecast.per
forecast.cnc.ac$response.forecast.coh
forecast.cnc.ac$response.forecast.all

# Collect all forecast point estimates
Table.forecast.cnc.ac <- forecast.cnc.ac$response.forecast.age
Table.forecast.cnc.ac <- rbind(Table.forecast.cnc.ac, forecast.cnc.ac$response.forecast.per)
Table.forecast.cnc.ac <- rbind(Table.forecast.cnc.ac, forecast.cnc.ac$response.forecast.coh)
Table.forecast.cnc.ac <- rbind(Table.forecast.cnc.ac, forecast.cnc.ac$response.forecast.all)
Table.forecast.cnc.ac

#csv 파일로 내리기
write.csv(Table.forecast.cnc.ac, file = "forecast_cnc_ac_mortality.csv")


forecast.cnc.ac$trap.linear.predictors.forecast


# 1. 특정 코호트에 대한 예측

# 2. 특정 연령에 대한 예측, AC 모델만 가능
# 1세단위
# 특정연령구간에 대한 추정, 40~59세를 대상으로 예측
# 대장암의 경우 50-69세를 대상으로 예측하는 것이 apc 모형의 장점을 잘 보여줄 수 있을 것 같음
# 그냥 3가지 암종 모두 50-69세로 하는게 더 깔끔할 것 같아서 그렇게 함
forecast.cnc.ac.age.sub1 <- apc.forecast.ac(fit.cnc.ac, sum.per.by.age=c(40,60)) #40-59세
# forecast.cnc.ac.age.sub1 <- apc.forecast.ac(fit.cnc.ac, sum.per.by.age=c(50,69)) #50-69세


#특정연령구간의 과거 연도별 response 합계
data.cnc.sum.per.by.age.sub1 <- colSums(data.cnc.response[c(40:60),]) # 40-59세
# data.cnc.sum.per.by.age.sub1 <- colSums(data.cnc.response[c(51:70),]) # 50-69세


#특정연령구간의 미래 추정 연도별 response 합계
Table.forecast.cnc.ac <- forecast.cnc.ac.age.sub1$response.forecast.per.by.age

#csv 파일로 내리기
write.csv(Table.forecast.cnc.ac, file = "forecast_cnc_ac_mortality_40-60.csv")


###########################################################
# All Data for Time Series
###########################################################

# Change of format to time series in long format ==> 이거로 하면 오류 발생함 밑에 수정된 코드로 실행하면 됨
# data.ts.cnc <- melt(data.cnc.response)
# data.ts.cnc <- melt(data.cnc.dose)
# data.ts.cnc <- melt(data.cnc.rate)

# melt 함수 오류가 나서 아래 코드로 실행함. 세로 연령/가로 연도 => 세로에 연령/연도 양식으로 변환함
library(reshape2)  # reshape2 패키지 로드
data.ts.cnc <- reshape2::melt(data.cnc.response) #response 를 기준으로 실행.

# 컬럼 이름을 아래와 같이 지정함
colnames(data.ts.cnc) <- c("age","period","value")

# 연도를 숫자형식으로 바꾸고, t 라는 컬럼에 저장
data.ts.cnc <- data.ts.cnc %>% mutate(t=as.numeric(substr(period, 1, 4)))


# 특정 연령 그룹 지정
group.age <- c("40","41","42","43","44","45","46","47","48","49","50","51","52","53","54","55","56","57","58","59","60") #40-60세만 추출
# group.age <- c("50","51","52","53","54","55","56","57","58","59","60","61","62","63","64","65","66","67","68","69") #50-69세만 추출

# group.age <- c("50")

# 특정 연령만 추출하여 sub1 데이터에 저장
data.ts.cnc.sub_age <- data.ts.cnc %>% filter(age %in% group.age)

# sqldf 패키지 install 및 불러오기
# install.packages("sqldf")
library(sqldf)

# 연령별 연도별로 합산. as age 부분은 추출 연령에 맞게 수정하는게 좋음
# 특정기간 이후로 프로젝션 하려고 하면 where 부분을 수정해줌
data.ts.cnc.sub1<-sqldf("SELECT
 60 as age
,period
,t
,sum(value) as value
from 'data.ts.cnc.sub_age'
GROUP BY
 period
,t")

#where t > 2010 Colon
#where t > 2008 Stomach
###########################################################
# Time series of data
###########################################################

# Time series data preparation
ts.cnc <- ts(data.ts.cnc.sub1$value, 
             start=data.ts.cnc.sub1$t[1],
             #deltat=(data.ts.cnc.sub1$t[2]-data.ts.cnc.sub1$t[1]), 오류가 생겨서 어차피 1년단위니깐 1로 셋팅함
             deltat=1,
             names=c("value"))
data.cnc.t <- data.ts.cnc.sub1$t


# start_check <- data.ts.cnc.sub1$t[1]

# 패키지 설치 및 로드를 해야지 실행이 됨
# install.packages("ggfortify")
library(ggplot2)
library(ggfortify)

# Plot of time series data
autoplot(ts.cnc)+xlab("Time")+ylab("Response")
# autoplot(ts.cnc)+xlab("Time")+ylab("Dose")
# autoplot(ts.cnc)+xlab("Time")+ylab("Rate")


poly_orders <- c(1, 2)

###########################################################
# Time series in Auto fit of ARIMA
###########################################################

for (poly.ord in poly_orders){
  
  # Order of polynomial in t
  # 다항식 조건
  # poly.ord <- 1
  # poly.ord <- 2
  # poly.ord <- 3
  
  # t_poly <- poly(data.ts.cnc.sub1$t, poly.ord, raw=TRUE)
  
  #poly.ord <- 2
  t_poly <- poly(data.ts.cnc.sub1$t, poly.ord, raw=FALSE)
  # poly.ord = 2차 다항식으로 변환하여 회귀분석을 수행 
  # raw = false 다항식 항목들이 서로 독립적으로 변형됨 -> 다중공선성 문제를 줄이는데 유리함
  
  
  library(forecast)
  
  fit.arima <- auto.arima(data.ts.cnc.sub1$value, xreg = t_poly)
  fit.arima
  checkresiduals(fit.arima)
  fit.arima$fitted
  
  ###########################################################
  # Forecast Time series in Auto fit of ARIMA
  ###########################################################
  
  # Forecast of time series
  h <- 50
  data.cnc.tfor <- seq(data.cnc.t[length(data.cnc.t)]+data.cnc.t[2]-data.cnc.t[1], data.cnc.t[length(data.cnc.t)]+h*(data.cnc.t[2]-data.cnc.t[1]),
                       length.out=h)
  
  #check_1 <- data.cnc.t[length(data.cnc.t)] # 2022
  #check_2 <- data.cnc.t[2] # 2012 시작년도 + 1
  #check_3 <- data.cnc.t[1] # 2011 시작년도
  #check_4 <- h*(data.cnc.t[2]-data.cnc.t[1]) # 50, 1년 단위이기 때문
  
  tfor_poly <- predict(t_poly, data.cnc.tfor)
  
  # 결과 그래프 확인
  # autoplot(forecast(fit.arima, xreg=tfor_poly))+xlab("Time")+ylab("Response")
  # autoplot(forecast(fit.arima, xreg=tfor_poly))+xlab("Time")+ylab("Dose")
  # autoplot(forecast(fit.arima, xreg=tfor_poly))+xlab("Time")+ylab("Rate")
  
  # 결과 값 출력
  arima_forecast_result <- forecast(fit.arima, xreg=tfor_poly)
  arima_forecast_table <- as.data.frame(arima_forecast_result)
  # write.csv(arima_forecast_table, file = "arima_forecast_table_mortality_F_40-60_01.csv")
  write.csv(arima_forecast_table, file = paste0("arima_forecast_table_mortality_40-60_", poly.ord, ".csv"))
  
}


###########################################################
# Time series in GLM
###########################################################

for (poly.ord in poly_orders){
  
  # Order of polynomial in t
  # poly.ord <- 1
  # poly.ord <- 2
  # poly.ord <- 3
  
  # t_poly <- poly(data.ts.cnc.sub1$t, poly.ord, raw=TRUE)
  t_poly <- poly(data.ts.cnc.sub1$t, poly.ord, raw=FALSE)
  
  
  # Poisson model with log link, 암발생 건수 예측
  fit.glm <- glm(value ~ poly(t, poly.ord, raw=FALSE), data=data.ts.cnc.sub1, family=poisson(link="log"))
  fit.glm
  
  # Poisson model with identity link
  # fit.glm <- glm(value ~ poly(t, poly.ord, raw=FALSE), data=data.ts.cnc.sub1, family=poisson(link="identity"))
  # fit.glm
  
  # Quasi-Poisson model with log link, 암발생 건수 예측이며, 과산포가 있는 경우 적절
  # fit.glm <- glm(value ~ poly(t, poly.ord, raw=FALSE), data=data.ts.cnc.sub1, family=quasipoisson(link="log"))
  # fit.glm
  
  # Gaussian model with identity link, 발생률을 예측할 때는 아래 모형이 적절함
  # fit.glm <- glm(value ~ poly(t, poly.ord, raw=FALSE), data=data.ts.cnc.sub1, family=gaussian(link="identity"))
  # fit.glm
  
  fit.glm
  summary(fit.glm)
  BIC(fit.glm)
  
  
  # 패키지 설치 및 로드
  # install.packages("pscl")
  library(pscl)
  
  # checkresiduals(fit.glm)
  fit.glm$fitted
  pR2(fit.glm)
  fit.glm$fitted.values
  
  
  ###########################################################
  # Forecast Time series in GLM
  ###########################################################
  
  h <- 50
  data.cnc.tfor <- seq(data.cnc.t[length(data.cnc.t)]+data.cnc.t[2]-data.cnc.t[1], data.cnc.t[length(data.cnc.t)]+h*(data.cnc.t[2]-data.cnc.t[1]),
                       length.out=h)
  tfor_polydata <- predict(t_poly, newdata=data.cnc.tfor)
  
  
  forecast.glm <- exp(predict(fit.glm, newdata=data.frame(t=data.cnc.tfor), type="link")) # 로그 변환 하였기 때문에 결과 값을 exp 변환 하여야 함
  forecast.glm
  
  
  autoplot(ts(forecast.glm))+xlab("Time")+ylab("Response")
  # autoplot(ts(forecast.glm))+xlab("Time")+ylab("Dose")
  # autoplot(ts(forecast.glm))+xlab("Time")+ylab("Rate")
  
  
  # 결과 값 출력
  glm_forecast_table <- as.data.frame(forecast.glm)
  # write.csv(glm_forecast_table, file = "glm_forecast_table_mortality_F_50-69_01.csv")
  # write.csv(glm_forecast_table, file = "glm_forecast_table_mortality_40-60_02.csv")
  write.csv(glm_forecast_table, file = paste0("glm_forecast_table_mortality_40-60_", poly.ord, ".csv"))
  
}




