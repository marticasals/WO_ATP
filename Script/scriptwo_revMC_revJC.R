################################################################################
# Load libraries
################################################################################
library(readxl)
library(openxlsx)
library(dplyr)
library(rlang) 
library(nortest)
library(SmartEDA)
library(ggplot2)
library(compareGroups)
library(epiR)
library(predtools)

################################################################################
# Load data
################################################################################
df_ATP   <- readRDS("df_ATP.rds")
WOCAUSES <- read_excel("WOCAUSES.xlsx")



################################################################################
#
# Descriptive statistics
#
################################################################################


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
ExpReport(df_ATP, Template = NULL, Target = NULL, label = NULL, 
          theme = "WalkOver", op_file = "report.html", 
          op_dir = 'figures_tables', sc = NULL, sn = NULL, Rc = NULL)

Summary_Cat <- ExpCTable(df_ATP, Target = NULL, margin = 1, clim = 10, 
                        nlim = 10, round = 2, bin = 3, per = FALSE, weight = NULL)
write.xlsx(Summary_Cat, "figures_tables/Summary_Cat.csv", rowNames = FALSE)

Summary_Num <- ExpNumStat(df_ATP,by="A",gp=NULL,Qnt=c(0.25,0.75),MesofShape=2,Outlier=TRUE,round=2)
selected_columns <- c("Vname", "mean", "median", "SD", "IQR","25%","75%")

Summary_numeric <- Summary_Num[, selected_columns]
write.xlsx(Summary_Num, "figures_tables/Summary_Num.csv", rowNames = FALSE)

Summary_Cat_WOC<- ExpCTable(WOCAUSES, Target = NULL, margin = 1, clim = 10, nlim = 10, round = 2, bin = 3, per = FALSE, weight = NULL)
write.xlsx(Summary_Cat_WOC, "figures_tables/Summary_CatWO.csv", rowNames = FALSE)


## Epidemiological analysis. 

# Calculation of PI and its 95%CI
prop_walkover          <- sum(df_ATP$WalkOver == "YES") / nrow(df_ATP)
ci_walkover            <- prop.test(sum(df_ATP$WalkOver == "YES"), nrow(df_ATP))$conf.int
prop_per_1000_walkover <- prop_walkover * 1000
ci_per_1000_walkover   <- prop.test(sum(df_ATP$WalkOver == "YES"), nrow(df_ATP), conf.level = 0.95)$conf.int * 1000
cat("Incidence Proportion WalkOver:", prop_walkover, "\n")
cat("Confidence interval WalkOver:",  ci_walkover[1], "-", ci_walkover[2], "\n\n")
cat("Incidence Proportion de WalkOver per 1000 matches:", prop_per_1000_walkover, "%\n\n")
cat("Confidence interval para WalkOver per 1000 matches:", ci_per_1000_walkover[1], "-", ci_per_1000_walkover[2], "%\n\n")


# Analysis of the time trend of the WalkOver.
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


## Bivariate analysis.

# Summary of covariates by groups
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
      .groups = 'drop' # Para evitar el mensaje de agrupacion no eliminada
    ) %>%
    mutate(Variable = var_name) %>%
    dplyr::select(Variable, WalkOver, Median, IQR, Q1, Q3)
  }

all_stats <- lapply(numeric_vars, calculate_stats, data = df_ATP) %>%
  bind_rows()
print(all_stats)


# Epidemiological analysis by groups 
# Tourney level and WO
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

# Surface and WO
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

# Sets and WO
(tabla_sets_walkover <- table(df_ATP$sets, df_ATP$WalkOver))
valores_3_5<- c(tabla_sets_walkover["3", "YES"], tabla_sets_walkover["3", "NO"], tabla_sets_walkover["5", "YES"], tabla_sets_walkover["5", "NO"])
epi.2by2(dat = valores_3_5, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

# Round and WO
(tabla_round_walkover <- table(df_ATP$round, df_ATP$WalkOver))
valores_F_Q<- c(tabla_round_walkover["Final", "YES"], tabla_round_walkover["Final", "NO"], tabla_round_walkover["Qualifying", "YES"], tabla_round_walkover["Qualifying", "NO"])
epi.2by2(dat = valores_F_Q, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")

valores_P_Q<- c(tabla_round_walkover["Preliminary", "YES"], tabla_round_walkover["Preliminary", "NO"], tabla_round_walkover["Qualifying", "YES"], tabla_round_walkover["Qualifying", "NO"])
epi.2by2(dat = valores_P_Q, method = "cohort.count", digits = 2,
         conf.level = 0.95, units = 1000, interpret = FALSE, outcome = "as.columns")



################################################################################ 
# GLM: Logistic regression
################################################################################
df_ATP$WO <- 2-as.numeric(df_ATP$WalkOver)
## Descriptive analysis logodds sample size ------------------------------------
unique_year <- sort(unique(df_ATP$year))
n_year <- length(unique_year)
logodd <- c()
for(i in 1:n_year){
  numer <- with(df_ATP, sum(WO==1 & year ==  unique_year[i]))
  denom <- with(df_ATP, sum(WO==0 & year ==  unique_year[i]))
  logodd[i] <- log(numer/denom)
}
##-- Plot logodds
plot(logodd ~ unique_year, pch=19, xlab='Year',col=rgb(0,0,1,0.5))
lines(lowess(logodd~unique_year,f = 0.4),col='darkblue',lwd=2)




################################################################################
### Multivariable analysis 
##-- Full model ----------------------------------------------------------------
model_full <- glm(WO ~ tourney_level + surface + round + 
                    dif_age + dif_rank + poly(year,2), 
                  data = df_ATP, family = binomial)
summary(model_full)

# BIC
model_step <- step(model_full, k = log(nrow(df_ATP)))

##-- model without dif_rank
model_1 <- glm(WO ~ tourney_level + surface + round + 
                 dif_age + poly(year,2), 
                  data = df_ATP, family = binomial)
summary(model_1)

# BIC
model_step <- step(model_1, k = log(nrow(df_ATP)))


##-- model Sense dif_age
model_2 <- glm(WO ~ tourney_level + surface + round + 
                 poly(year,2), 
               data = df_ATP, family = binomial)
summary(model_2)

# BIC
model_step <- step(model_2,k = log(nrow(df_ATP))) # trial design is the less relevant

##-- model Sense surface
model_3 <- glm(WO ~ tourney_level + round + poly(year,2), 
               data = df_ATP, family = binomial)
summary(model_3)

# BIC
model_step <- step(model_3,k = log(nrow(df_ATP)))

model_def <- model_3


# Validation
vif(model_def)
residualPlot(model_def)
library(predtools)
d <- data.frame(y    = df_ATP$WO,
                pred = predict(model_def, type = 'response'))
calibration_plot(data  = d, 
                 obs   = "y", 
                 pred  = "pred", 
                 title = "Calibration plot", 
                 y_lim = c(0, 0.01), 
                 x_lim = c(0, 0.01))

################################################################################
##-- Interactions --> Improve calibration but less interpretability
################################################################################
model_4 <- glm(WO ~ (tourney_level + round + poly(year,2))^2, 
               data = df_ATP, family = binomial)
summary(model_4)

##-- Dejar solo round con year?
model_5 <- glm(WO ~ tourney_level + round * poly(year,2), 
               data = df_ATP, family = binomial)
summary(model_5)

# Validation
d <-  data.frame(y     = df_ATP$WO,
                 pred  = predict(model_5, type = 'response'))
calibration_plot(data  = d, 
                 obs   = "y", 
                 pred  = "pred", 
                 title = "Calibration plot", 
                 y_lim = c(0, 0.01), 
                 x_lim = c(0, 0.01))

################################################################################
##-- Interpretation
################################################################################
##-- ORS -----------------------------------------------------------------------
sel_rm <- c(1,7,8)
confint_res <- confint(model_def)[-sel_rm,]
OR1 <- data.frame(OR=round(exp(coef(model_def)[-sel_rm]),3),
                  LI=round(exp(confint_res),3)[,1],
                  LS=round(exp(confint_res),3)[,2])

##-- OR Years --------------------------------------------------------- 
em <- emmeans(model_def, specs = ~ year, at =list(year=c(2020,2000,1980)) , type='response')
em_years <- summary(pairs(em))
row_names <- c('year 2000 vs. 1980','year 2020 vs. 1980')
OR_years <- em_years$odds.ratio[c(3,2)]
SE_years <- em_years$SE[c(3,2)]
OR2 <- data.frame(
  OR   = OR_years,
  LI  = exp(log(OR) - qnorm(0.975) * SE_years),
  LS  = exp(log(OR) + qnorm(0.975) * SE_years),
  row.names = row_names
)
OR2

OR <- rbind(OR1,OR2)

##-- Forest_plot ---------------------------------------------------------------
LABELS_VAR <- c('Tourney level',rep(NA,3),
                'Round',rep(NA,2),
                'Year',rep(NA,2))
LABELS_CAT <- c(NA, 'Grand Slams vs. 250 or 500','Masters vs. 250 or 500','Tour Finals vs. 250 or 500',
                NA,'Preliminary vs. Final','Qualifying vs. Final',
                NA,'2000 vs. 1980','2020 vs. 1980')
OR_plot <- rbind(rep(NA,3),OR[1:3,],
                 rep(NA,3),OR[4:5,],
                 rep(NA,3),OR[6:7,])
par(mar=c(4,10,1,5), cex.lab=0.8, font.lab= 2)
h <- nrow(OR_plot)
plot(NA, xlim= c(1/5,5), ylim=c(0.5,h+0.5), log='x',
     xlab='Odds Ratio',ylab='',
     xaxt='n',yaxt='n',bty='n')
axis(1,at=c(1/5,1/2,1,2,5),cex=0.7)
# Background
par(xpd=NA)
rect(0.000001,h+0.5,1000,h-3.5,  col='lightblue',border = NA)
rect(0.000001,h-6.5,1000,h-9.5,col='lightblue',border = NA)
par(xpd=FALSE)

abline(v=1,lty=1)
abline(v=c(1/5,1/2,2,5),lty=2, col='grey')
mtext('OR [95%CI]',side = 4, at = h +1, line=2.5, las=1, adj=0.5, 
      cex=0.8, font=2)

for(i in 1:h){
  points(OR_plot[i,'OR'],h,pch=15, col='darkblue', cex=1.2)
  segments(OR_plot[i,'LI'],h,OR_plot[i,'LS'],h, col='darkblue', lwd=1.5)
  
  lab_var <- ifelse(is.na(LABELS_VAR[i]),"",LABELS_VAR[i])
  lab_cat <- ifelse(is.na(LABELS_CAT[i]),"",LABELS_CAT[i])
  lab_ci  <- ifelse(is.na(OR_plot[i,'OR']),"",
                    paste0(formatC(OR_plot[i,'OR'],digits = 2, format = 'f'),' ',
                           '[', formatC(OR_plot[i,'LI'],digits = 2, format = 'f'),' , ',
                           formatC(OR_plot[i,'LS'],digits = 2, format = 'f'),']'))
  mtext(lab_var,side = 2, at = h, line=10, las=1, adj=0,   cex=0.7, font=2)
  mtext(lab_cat,side = 2, at = h, line=9, las=1, adj=0,   cex=0.7, font=1)
  mtext(lab_ci, side = 4, at = h, line=2.5,  las=1, adj=0.5, cex=0.7, font=1)
  
  h <- h - 1
}



