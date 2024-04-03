df_ATP <- readRDS("C:/Users/Victoria/Desktop/Master Bioestadistica/TFM/Default/df_ATP.rds")
WOCAUSES <- read_excel("WOCAUSES.xlsx")
library(readxl)
library(openxlsx)
library(dplyr)
library(rlang) 
library(nortest)
library(SmartEDA)
library(ggplot2)
library(compareGroups)
library(epiR)

##Exploratory analysis.
##Analysis of the normality. 
qqnorm(df_ATP$winner_age, main = "QQ Plot of Winner Age")
qqline(df_ATP$winner_age)
lillie.test(x =  df_ATP$winner_age)

qqnorm(df_ATP$loser_age, main = "QQ Plot of Loser Age")
qqline(df_ATP$loser_age)
lillie.test(x =  df_ATP$loser_age)

qqnorm(df_ATP$dif_age, main = "QQ Plot of Age Difference")
qqline(df_ATP$dif_age)
lillie.test(x =  df_ATP$dif_age)

qqnorm(df_ATP$winner_rank, main = "QQ Plot of Winner Rank")
qqline(df_ATP$winner_rank)
lillie.test(x =  df_ATP$winner_rank)

qqnorm(df_ATP$loser_rank, main = "QQ Plot of Loser Rank")
qqline(df_ATP$loser_rank)
lillie.test(x =  df_ATP$loser_rank)

qqnorm(df_ATP$dif_rank, main = "QQ Plot of Rank Difference")
qqline(df_ATP$dif_rank)
lillie.test(x =  df_ATP$dif_rank)

qqnorm(df_ATP$games, main = "QQ Plot of Games")
qqline(df_ATP$games)
lillie.test(x =  df_ATP$games)

#Creation of an exploratory report.
ExpReport(df_ATP, Template = NULL, Target = NULL, label = NULL, theme = "WalkOver", op_file = "report.html", op_dir = getwd(), sc = NULL, sn = NULL, Rc = NULL)

Summary_Cat<- ExpCTable(df_ATP, Target = NULL, margin = 1, clim = 10, nlim = 10, round = 2, bin = 3, per = FALSE, weight = NULL)
write.xlsx(Summary_Cat, "C:/Users/Victoria/Desktop/TFM/Datos ATP/Summary_Cat.csv", rowNames = FALSE)

Summary_Num<-ExpNumStat(df_ATP,by="A",gp=NULL,Qnt=c(0.25,0.75),MesofShape=2,Outlier=TRUE,round=2)
selected_columns <- c("Vname", "mean", "median", "SD", "IQR","25%","75%")

Summary_numeric <- Summary_Num[, selected_columns]
write.xlsx(Summary_Num, "C:/Users/Victoria/Desktop/Master Bioestadistica/TFM/Datos ATP/Summary_Num.csv", rowNames = FALSE)

Summary_Cat_WOC<- ExpCTable(WOCAUSES, Target = NULL, margin = 1, clim = 10, nlim = 10, round = 2, bin = 3, per = FALSE, weight = NULL)
write.xlsx(Summary_Cat_WOC, "C:/Users/Victoria/Desktop/Master Bioestadistica/TFM/Datos ATP/Summary_CatWO.csv", rowNames = FALSE)


##Epidemiological analysis. 

#Calculation of PI and 95%CI
prop_walkover <- sum(df_ATP$WalkOver == "YES") / nrow(df_ATP)
ci_walkover <- prop.test(sum(df_ATP$WalkOver == "YES"), nrow(df_ATP))$conf.int
prop_per_1000_walkover <- prop_walkover * 1000
ci_per_1000_walkover <- prop.test(sum(df_ATP$WalkOver == "YES"), nrow(df_ATP), conf.level = 0.95)$conf.int * 1000
cat("Proporci?n de incidencia de WalkOver:", prop_walkover, "\n")
cat("Intervalo de confianza para WalkOver:", ci_walkover[1], "-", ci_walkover[2], "\n\n")
cat("Proporci?n de incidencia de WalkOver por cada 1000 partidos:", prop_per_1000_walkover, "%\n\n")
cat("Intervalo de confianza para WalkOver por cada 1000 partidos:", ci_per_1000_walkover[1], "-", ci_per_1000_walkover[2], "%\n\n")


#Analysis of the time trend of the WalkOver.
walkover_cases_by_year <- df_ATP %>%
  group_by(year) %>%
  summarize(
    Total_Matches = n(),
    WalkOver = sum(WalkOver == "YES", na.rm = TRUE)
  ) %>%
  ungroup() %>%
  mutate(Incidence_Proportion = (WalkOver / Total_Matches) * 1000)

ggplot(walkover_cases_by_year, aes(x = year, y = Incidence_Proportion)) +
  geom_line() +  
  geom_smooth(method = "loess", se = TRUE, color = "blue") +  
  theme_minimal() +
  scale_x_continuous(breaks = seq(min(walkover_cases_by_year$year), max(walkover_cases_by_year$year), by = 2)) +  
  scale_y_continuous(breaks = seq(0, max(walkover_cases_by_year$Incidence_Proportion), by = 1)) + 
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) +  
  labs(y = "Incidence Proportion per 1000 Matches",
       x = "Year")


##Bivariate analysis.

#Summary of covariates by groups
Summary_walkover<-compareGroups(WalkOver~tourney_level+surface+round+games+sets+winner_hand+loser_hand+winner_age+loser_age+dif_age+winner_rank+loser_rank+dif_rank, df_ATP, byrow=TRUE)
df_Summary_walkover<- createTable(Summary_walkover)
export2word(df_Summary_walkover, file='Summary_walkover.docx')


numeric_vars <- c("winner_age", "loser_age", "dif_age", "winner_rank", "loser_rank", "dif_rank", "games")
calculate_stats <- function(data, var_name) {
  data %>%
    group_by(WalkOver) %>%
    summarise(
      Median = median(!!sym(var_name), na.rm = TRUE),
      IQR = IQR(!!sym(var_name), na.rm = TRUE),
      Q1 = quantile(!!sym(var_name), 0.25, na.rm = TRUE),
      Q3 = quantile(!!sym(var_name), 0.75, na.rm = TRUE),
      .groups = 'drop' # Para evitar el mensaje de agrupaci?n no eliminada
    ) %>%
    mutate(Variable = var_name) %>%
    select(Variable, WalkOver, Median, IQR, Q1, Q3)}

all_stats <- lapply(numeric_vars, calculate_stats, data = df_ATP) %>%
  bind_rows()
print(all_stats)


#Epidemiological analysis by groups 
#Tourney level and WO
(tabla_tourney_walkover <- table(df_ATP$tourney_level, df_ATP$WalkOver))
valores_M_GS <- c( tabla_tourney_walkover["Masters", "YES"], tabla_tourney_walkover["Masters", "NO"], tabla_tourney_walkover["Grand Slams", "YES"], tabla_tourney_walkover["Grand Slams", "NO"])
epi.2by2(dat = valores_M_GS, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

valores_250.500_GS <- c( tabla_tourney_walkover["250 or 500", "YES"], tabla_tourney_walkover["250 or 500", "NO"], tabla_tourney_walkover["Grand Slams", "YES"], tabla_tourney_walkover["Grand Slams", "NO"])
epi.2by2(dat = valores_250.500_GS, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

valores_TF_GS<- c( tabla_tourney_walkover["Tour Finals", "YES"], tabla_tourney_walkover["Tour Finals", "NO"], tabla_tourney_walkover["Grand Slams", "YES"], tabla_tourney_walkover["Grand Slams", "NO"])
epi.2by2(dat = valores_TF_GS, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

#Surface and WO
(tabla_surface_walkover <- table(df_ATP$surface, df_ATP$WalkOver))

valores_G_C<- c(tabla_surface_walkover["Grass", "YES"], tabla_surface_walkover["Grass", "NO"],tabla_surface_walkover["Clay", "YES"], tabla_surface_walkover["Clay", "NO"])
epi.2by2(dat = valores_G_C, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

valores_H_C<- c(tabla_surface_walkover["Hard", "YES"], tabla_surface_walkover["Hard", "NO"], tabla_surface_walkover["Clay", "YES"], tabla_surface_walkover["Clay", "NO"])
epi.2by2(dat = valores_H_C, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

valores_C_C<- c(tabla_surface_walkover["Carpet", "YES"], tabla_surface_walkover["Carpet", "NO"], tabla_surface_walkover["Clay", "YES"], tabla_surface_walkover["Clay", "NO"])
epi.2by2(dat = valores_C_C, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

#Sets and WO
(tabla_sets_walkover <- table(df_ATP$sets, df_ATP$WalkOver))
valores_3_5<- c(tabla_sets_walkover["3", "YES"], tabla_sets_walkover["3", "NO"], tabla_sets_walkover["5", "YES"], tabla_sets_walkover["5", "NO"])
epi.2by2(dat = valores_3_5, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

#Round and WO
(tabla_round_walkover <- table(df_ATP$round, df_ATP$WalkOver))
valores_F_Q<- c(tabla_round_walkover["Final", "YES"], tabla_round_walkover["Final", "NO"], tabla_round_walkover["Qualifying", "YES"], tabla_round_walkover["Qualifying", "NO"])
epi.2by2(dat = valores_F_Q, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

valores_P_Q<- c(tabla_round_walkover["Preliminary", "YES"], tabla_round_walkover["Preliminary", "NO"], tabla_round_walkover["Qualifying", "YES"], tabla_round_walkover["Qualifying", "NO"])
epi.2by2(dat = valores_P_Q, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")


