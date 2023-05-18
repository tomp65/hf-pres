install.packages("tidyverse")
install.packages("stringr")
install.packages("readxl")
install.packages("writexl")
install.packages("openxlsx")

library(tidyverse)
library(stringr)
library(readxl)
library(writexl)
library(openxlsx)


###PARAMETERS###

yr1 <- '2019-20' #fill in financial year 1 in the YYYY-YY format

yr2 <- '2020-21' #fill in financial year 2 in the YYYY-YY format

HES_url_yr1 <- "https://files.digital.nhs.uk/95/484BF0/hosp-epis-stat-outp-trea-spec-2019-20-tab.xlsx"
  
HES_url_yr2 <- "https://files.digital.nhs.uk/E8/C1C42B/hosp-epis-stat-outp-trea-spec-2020-21-tab.xlsx"

NCC_url <- "https://www.england.nhs.uk/wp-content/uploads/2021/06/National_Schedule_of_NHS_Costs_FY1920.xlsx"

HES_urls <- list(HES_url_yr1, HES_url_yr2)


########################################################################
#                               ACTIVITY                               #
########################################################################

#Load in HES activity data

HES_xlsx <- list("HES_yr1.xlsx", "HES_yr2.xlsx")

rawdata <- lapply(c(1:2), function(i){
  
  download.file(HES_urls[[i]], destfile = HES_xlsx[[i]])
  
  read_excel(HES_xlsx[[i]], skip = 3) %>%
    select(1:8) %>%
    rename(Service_code = 1, Service_description = 2, All_attendances = 3, Attended_FA = 4, Attended_FTC = 5, Attended_SA = 6, Attended_STC = 7, FS_unknown = 8)
})


#Sum attendance types and join years into a single activity dataset

sumdata <- lapply(rawdata, function(i){ 
  
  df <- i %>%
  mutate(FA = (Attended_FA + Attended_FTC + ((FS_unknown*(Attended_FA + Attended_FTC))/(Attended_FA + Attended_FTC +
                                                                                          Attended_SA + Attended_STC)))) %>%
    mutate(SA = (Attended_SA + Attended_STC + ((FS_unknown*(Attended_SA + Attended_STC))/(Attended_FA + Attended_FTC +
                                                                                            Attended_SA + Attended_STC)))) %>%
    slice(2:157)
  
  df[df == 'NaN'] <- 0
  
  return(df)
  
  })

HES_Activity <- full_join(sumdata[[1]], sumdata[[2]], by = "Service_code") %>%
  select(1,2,9,10,18,19) %>%
  rename("FA_yr1" = FA.x, "SA_yr1" = SA.x, "FA_yr2" = FA.y, "SA_yr2" = SA.y)


########################################################################
#                            UNIT COSTS                                #
########################################################################

download.file(NCC_url, destfile = "NCC.xlsx")

CL <- read_excel("NCC.xlsx", sheet = 'CL', skip = 4) %>%
  unite(MatchCode, 3, 1, sep = '_', remove = FALSE) %>%
  rename(Attendances_CL = 6, Total_cost_CL = 8)

NCL <- read_excel("NCC.xlsx", sheet = 'NCL', skip = 4) %>%
  unite(MatchCode, 3, 1, sep = '_', remove = FALSE) %>%
  rename(Attendances_NCL = 6, Total_cost_NCL = 8)

NCL_CL_join <- full_join(CL, NCL, by = 'MatchCode') %>%
  select(1, 6, 7, 8, 14, 15 , 16) %>%
  mutate_at(vars("Attendances_CL", "Attendances_NCL",
                 "Total_cost_CL", "Total_cost_NCL"), ~replace_na(.,0)) %>%
  mutate(Full_attendances = Attendances_CL + Attendances_NCL) %>% 
  mutate(Full_expenditure = Total_cost_CL + Total_cost_NCL) %>%
  separate(MatchCode, c('Currency_code', 'Service_code'), sep = '_', remove = FALSE)


## Aggregate AC and BD currency codes together, calculate unit costs and volumes
unitcost_aggregated <- NCL_CL_join %>%
  mutate(Aggregated_code = case_when(Currency_code %in% c('WF01A', 'WF01C', 'WF02A', 'WF02C') ~ 'WF_AC',
                                     Currency_code %in% c('WF01B', 'WF01D', 'WF02B', 'WF02D') ~ 'WF_BD', 
                                     TRUE ~ 'NA')) %>%
  group_by(Aggregated_code, Service_code) %>%
  summarise(Attendances = sum(Full_attendances), Expenditure = sum(Full_expenditure)) %>%
  mutate(UnitCost = Expenditure/Attendances)

FA_unitcost<- subset(unitcost_aggregated, Aggregated_code %in% "WF_BD")

SA_unitcost <- subset(unitcost_aggregated, Aggregated_code %in% "WF_AC")

########################################################################
#                              VOLUMES                                 #
########################################################################

FA_volume <- full_join(HES_Activity, FA_unitcost, by = 'Service_code') %>%
  select(-c(4,6,7,8,9)) %>%
  rename('UnitCost_FA' = UnitCost) %>%
  mutate(FA_volume_yr1 = (FA_yr1*UnitCost_FA)) %>%
  mutate(FA_volume_yr2 = (FA_yr2*UnitCost_FA)) 

FA_matched <- FA_volume %>%
  filter(!(is.na(FA_volume_yr1)|is.na(FA_volume_yr2))) %>%
  mutate(growth_FA = FA_volume_yr2/FA_volume_yr1-1) %>%
  mutate(contribution_FA = growth_FA*FA_volume_yr1/sum(FA_volume_yr1))

FA_unmatched <- FA_volume %>%
  filter((is.na(FA_volume_yr1)|is.na(FA_volume_yr2)))

SA_volume <- full_join(HES_Activity, SA_unitcost, by = 'Service_code') %>%
  select(-c(3,5,7,8,9)) %>%
  rename('UnitCost_SA' = UnitCost) %>%
  mutate(SA_volume_yr1 = (SA_yr1*UnitCost_SA)) %>%
  mutate(SA_volume_yr2 = (SA_yr2*UnitCost_SA)) 

SA_matched <- SA_volume %>%
  filter(!(is.na(SA_volume_yr1)|is.na(SA_volume_yr2))) %>%
  mutate(growth_SA = SA_volume_yr2/SA_volume_yr1-1) %>%
  mutate(contribution_SA = growth_SA*SA_volume_yr1/sum(SA_volume_yr1))

SA_unmatched <- SA_volume %>%
  filter((is.na(SA_volume_yr1)|is.na(SA_volume_yr2)))

volumes_matched_all <- full_join(FA_matched, SA_matched, by = 'Service_code')

all_unmatched <- full_join(FA_unmatched, SA_unmatched, by = 'Service_code')


########################################################################
#                                SUMMARIES                              #
########################################################################
volume_summaries <-  data.frame(FA_yr1_activity = sum(FA_matched$FA_yr1), FA_yr2_activity = sum(FA_matched$FA_yr2),
                                FA_yr1_expenditure = sum(FA_matched$FA_volume_yr1), FA_yr2_volume = sum(FA_matched$FA_volume_yr2),
                                SA_yr1_activity = sum(SA_matched$SA_yr1), SA_yr2_activity = sum(SA_matched$SA_yr2),
                                SA_yr1_expenditure = sum(SA_matched$SA_volume_yr1), SA_yr2_volume = sum(SA_matched$SA_volume_yr2)) %>%
                                mutate(FA_growth = FA_yr2_volume/FA_yr1_expenditure-1) %>%
                                mutate(SA_growth = SA_yr2_volume/SA_yr1_expenditure-1)

write_xlsx(list(Summary = volume_summaries, Matched = volumes_matched_all, unmatched = all_unmatched),
           paste0('Outpatients_processing_', yr2, '.xlsx'))
