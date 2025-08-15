rm(list=ls())
setwd("C:\\Users\\mckim\\Desktop\\Mortality_Improvement")

library(readr)
library(dplyr)
library(tidyr)
library(ggplot2)

# 공통 처리 함수
read_and_process <- function(path, sex, type) {
  df <- readr::read_csv(path, show_col_types = FALSE)
  age_col <- names(df)[1]
  df %>%
    rename(Age = all_of(age_col)) %>%
    pivot_longer(
      cols = -Age,
      names_to = "Year",
      values_to = "ImprovementRate"
    ) %>%
    mutate(
      Sex = sex,
      Type = type,
      Age = as.integer(Age),
      Year = as.integer(Year),
      ImprovementRate = readr::parse_number(ImprovementRate)
    )
}

# 데이터 읽기
male_current   <- read_and_process("4. male_current.csv",   "Male",   "Actual")
male_future    <- read_and_process("4. male_future.csv",    "Male",   "Future")
female_current <- read_and_process("4. female_current.csv", "Female", "Actual")
female_future  <- read_and_process("4. female_future.csv",  "Female", "Future")

# 합치기 + 필터
df <- bind_rows(male_current, male_future, female_current, female_future) %>%
  filter(Age %in% c(40, 60))

# 그래프 객체 생성
p <- ggplot(df, aes(x = Year, y = ImprovementRate, color = Type, linetype = Type)) +
  geom_line(size = 1) +
  facet_grid(Age ~ Sex) +
  scale_color_manual(values = c("Actual" = "#4B4B4B", "Future" = "#1A72E7")) +
  scale_linetype_manual(values = c("Actual" = "solid", "Future" = "solid")) +
  labs(
    title = "Actual vs Future Mortality Improvement Rates (Ages 40 & 60)",
    x = "Year",
    y = "Improvement Rate (%)",
    color = "Series",
    linetype = "Series"
  ) +
  theme_minimal() +
  theme(
    plot.title = element_text(hjust = 0.5),
    legend.position = "top"
  )

# PNG 파일로 저장
ggsave(
  filename = "mortality_improvement_rates.png",
  plot = p,
  width = 10, height = 6, dpi = 300
)
