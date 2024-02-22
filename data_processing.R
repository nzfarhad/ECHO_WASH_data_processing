library(readxl)
library(dplyr)
library(tidyr)
library(atRfunctions)

# Read the data set
data <- read_xlsx_sheets("input/Data from Activity info_merged.xlsx")

# Fix the Zone name for Daykundi Province
data$`Access to Sanitation Facilities` <- data$`Access to Sanitation Facilities` %>% 
  mutate(
    `Zone Name` = case_when(
      grepl("Daykundi", `Province Name`) ~ "Central Zone\\مرکزی زون",
      TRUE ~ `Zone Name`
    )
  )

data$`WASH Supplies Distribution` <- data$`WASH Supplies Distribution` %>% 
  mutate(
    `Zone Name` = case_when(
      grepl("Daykundi", `Province Name`) ~ "Central Zone\\مرکزی زون",
      TRUE ~ `Zone Name`
    )
  )

data$`WASH in Health Care Facilities` <- data$`WASH in Health Care Facilities` %>% 
  mutate(
    `Zone Name` = case_when(
      grepl("Daykundi", `Province Name`) ~ "Central Zone\\مرکزی زون",
      TRUE ~ `Zone Name`
    )
  )

## WASH IN SCHOOL --------------------------------------------------------------
wash_schools_anlysis_village <- data$`WASH in Schools` %>% 
  group_by(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`) %>%
  summarise(
    `# of Schools` = n(),
    `Male Students` = sum(`Male Students`, na.rm = T),
    `Female Students` = sum(`Female Students`, na.rm = T),
  ) %>%
  ungroup() %>% 
  rowwise() %>% 
  mutate(
    `Total Number of Students for WASH in schools interventions` = sum(c_across(7:8), na.rm = T),
    key = paste0(`District Name`, `Village Name`, `Community/ Area/ Camp Name`)
  ) %>% 
  arrange(key) %>% 
  mutate(
    component = "wash_school"
  )

## WASH IN HEALTH CARE FACILITIES ----------------------------------------------
wash_hf_village <- data$`WASH in Health Care Facilities` %>% 
  filter(`Catchment Population` > 0) %>% # Exclude 0 population reported
  mutate(key = paste0(`District Name`, `Village Name`, `Community/ Area/ Camp Name`)) %>% 
  group_by(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`) %>% 
  summarise(
    `# of HFs` = n(),
    `Catchment Population on HF` = sum(`Catchment Population`, na.rm = T)
  ) %>%
  mutate(key = paste0(`District Name`, `Village Name`, `Community/ Area/ Camp Name`)) %>% 
  ungroup() %>% 
  mutate(
    component = "wash_hf"
  )


## ACCESS TO SANITATION FACILITIES ---------------------------------------------
sanitation_hygiene <- data$`Access to Sanitation Facilities` %>% 
  mutate(
    key_plus_activity = paste0(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En`),
  ) 

sanitation_hygiene_dup <- sanitation_hygiene %>% 
  filter(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F)) %>% 
  filter(!is.na(Families) & !is.na(Beneficiaries)) %>% 
  group_by(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En`, key_plus_activity) %>% 
  summarise(
    Families = max(Families, na.rm = T),
    Beneficiaries = max(Beneficiaries, na.rm = T)
    )

# PROVINCE
sanitation_hygiene_province <- bind_rows(
    sanitation_hygiene %>% 
      filter(!(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F))),
    sanitation_hygiene_dup
  ) %>% 
  filter(!is.na(Families)) %>% # Exclude NA reported # of families
  group_by(`Zone Name`, `Province Name`, `Activity Name En` ) %>% 
  summarise(
    Families = sum(Families, na.rm = T),
    Beneficiaries = sum(Beneficiaries, na.rm = T),
  ) %>% 
  select(-c(Beneficiaries)) %>% 
  pivot_wider(names_from = `Activity Name En`, values_from = c(Families))

# VILLAGE
sanitation_hygiene_village <- bind_rows(
    sanitation_hygiene %>%
      filter(!(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F))),
    sanitation_hygiene_dup
  ) %>%
  filter(!is.na(Families)) %>% # Exclude NA reported # of families
  group_by(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En` ) %>% 
  summarise(
    Families = sum(Families, na.rm = T),
    Beneficiaries = sum(Beneficiaries, na.rm = T),
  ) %>% 
  select(-c(Beneficiaries)) %>%
  pivot_wider(names_from = `Activity Name En`, values_from = c(Families)) %>% 
  ungroup() %>% 
  rowwise() %>% 
  mutate(
    `Number of Families for all Access to Sanitation interventions` = max(c_across(6:11), na.rm = T),
    key = paste0(`District Name`, `Village Name`, `Community/ Area/ Camp Name`)
  ) %>% mutate(
    component = "sanitation_hygiene"
  )

## ACCESS TO WATER -------------------------------------------------------------
access_to_water <- data$`Access to Water` %>% 
  mutate(
    key_plus_activity = paste0(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En`)
  )

access_to_water_dup <- access_to_water %>% 
  filter(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F)) %>% 
  filter(!is.na(Families)) %>% 
  group_by(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En`, key_plus_activity) %>% 
  summarise(
    Families = max(Families, na.rm = T)
  )

# PROVINCE
access_to_water_province <- bind_rows(
    access_to_water %>% 
      filter(!(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F))),
    access_to_water_dup
  ) %>% 
  group_by(`Zone Name`, `Province Name`, `Activity Name En` ) %>% 
  summarise(
    Families = sum(Families, na.rm = T),
  )  %>% 
  pivot_wider(names_from = `Activity Name En`, values_from = c(Families))

# VILLAGE
access_to_water_village <- bind_rows(
    access_to_water %>% 
      filter(!(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F))),
    access_to_water_dup
  ) %>%  
  group_by(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En`) %>% 
  summarise(
    Families = sum(Families, na.rm = T),
  )  %>% 
  pivot_wider(names_from = `Activity Name En`, values_from = c(Families)) %>% 
  ungroup() %>% 
  rowwise() %>% 
  mutate(
    `Number of Families for all Access to water interventions` = max(c_across(6:22), na.rm = T),
    key = paste0(`District Name`, `Village Name`, `Community/ Area/ Camp Name`)
  ) %>% mutate(
    component = "access_to_water"
  )


## WASH SUPPLIES DISTRIBUTION
wash_supp <- data$`WASH Supplies Distribution` %>% 
  mutate(
    key_plus_activity = paste0(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En`)
  )

wash_supp_dup <- wash_supp %>% 
  filter(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F)) %>% 
  filter(!is.na(Families)) %>% 
  group_by(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En`, key_plus_activity) %>% 
  summarise(
    Families = max(Families, na.rm = T)
  )

# PROVINCE
wash_supp_province <- bind_rows(
    wash_supp %>% 
      filter(!(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F))),
    wash_supp_dup
  ) %>% 
  filter(Families > 0) %>% 
  group_by(`Zone Name`, `Province Name`, `Activity Name En` ) %>% 
  summarise(
    Families = sum(Families, na.rm = T),
  ) %>% 
  pivot_wider(names_from = `Activity Name En`, values_from = c(Families))

# VILLAGE
wash_supp_village <- bind_rows(
    wash_supp %>% 
      filter(!(duplicated(key_plus_activity, fromLast = T) | duplicated(key_plus_activity, fromLast = F))),
    wash_supp_dup
  ) %>% 
  filter(Families > 0) %>% 
  group_by(`Zone Name`, `Province Name`, `District Name`, `Village Name`, `Community/ Area/ Camp Name`, `Activity Name En`  ) %>% 
  summarise(
    Families = sum(Families, na.rm = T),
  ) %>%
  pivot_wider(names_from = `Activity Name En`, values_from = c(Families)) %>% 
  ungroup() %>% 
  rowwise() %>% 
  mutate(
    `Number of Families for all WASH supply distribution interventions` = max(c_across(6:11), na.rm = T),
    key = paste0(`District Name`, `Village Name`, `Community/ Area/ Camp Name`)
  ) %>% mutate(
    component = "wash_supp"
  )


output_list <- list(
  sanitation_hygiene_province = sanitation_hygiene_province,
  sanitation_hygiene_village = sanitation_hygiene_village,
  access_to_water_province = access_to_water_province,
  access_to_water_village = access_to_water_village,
  wash_supp_province = wash_supp_province,
  wash_supp_village = wash_supp_village,
  wash_schools_anlysis_village = wash_schools_anlysis_village,
  wash_hf_village = wash_hf_village
)


merged <- sanitation_hygiene_village %>% 
  full_join(wash_supp_village, by = "key") %>% 
  full_join(access_to_water_village, by = "key") %>% 
  full_join(wash_schools_anlysis_village, by = "key") %>%
  full_join(wash_hf_village, by = "key") %>% 
  mutate(
    key = tolower(key)
  )

# Cleaning merged output
merged <- merged %>% 
  mutate(
    Zone_name = case_when(
      !is.na(component.x) ~ `Zone Name.x`,
      !is.na(component.y) ~ `Zone Name.y`,
      !is.na(component.x.x) ~ `Zone Name.x.x`,
      !is.na(component.y.y) ~ `Zone Name.y.y`,
      !is.na(component) ~ `Zone Name`
    ),
    Province_name = case_when(
      !is.na(component.x) ~ `Province Name.x`,
      !is.na(component.y) ~ `Province Name.y`,
      !is.na(component.x.x) ~ `Province Name.x.x`,
      !is.na(component.y.y) ~ `Province Name.y.y`,
      !is.na(component) ~ `Province Name`
    ),
    District_name = case_when(
      !is.na(component.x) ~ `District Name.x`,
      !is.na(component.y) ~ `District Name.y`,
      !is.na(component.x.x) ~ `District Name.x.x`,
      !is.na(component.y.y) ~ `District Name.y.y`,
      !is.na(component) ~ `District Name`
    ),
    Village_name = case_when(
      !is.na(component.x) ~ `Village Name.x`,
      !is.na(component.y) ~ `Village Name.y`,
      !is.na(component.x.x) ~ `Village Name.x.x`,
      !is.na(component.y.y) ~ `Village Name.y.y`,
      !is.na(component) ~ `Village Name`
    ),
    Community_name = case_when(
      !is.na(component.x) ~ `Community/ Area/ Camp Name.x`,
      !is.na(component.y) ~ `Community/ Area/ Camp Name.y`,
      !is.na(component.x.x) ~ `Community/ Area/ Camp Name.x.x`,
      !is.na(component.y.y) ~ `Community/ Area/ Camp Name.y.y`,
      !is.na(component) ~ `Community/ Area/ Camp Name`
    )
  ) %>% 
  select(KEY = key,Zone_name, Province_name, District_name, Village_name, Community_name,
         Component_sanitation_hygiene = component.x, Component_wash_supp = component.y,
         Component_access_to_water = component.x.x, Component_wash_school = component.y.y,
         Component_wash_hf = component, everything())

merged <- merged[, !grepl("\\.x$|\\.y$|\\.x.x$|\\.y.y$", colnames(merged))] %>% 
  select(-c(`Zone Name`, `District Name`, `Province Name`, `Village Name`, `Community/ Area/ Camp Name`))

openxlsx::write.xlsx(merged, "output/New_Merge_updated.xlsx")
openxlsx::write.xlsx(output_list, "output/All_outputs.xlsx")

## Analysis -------------------------------------------------------------
merged_analysis_all_province <- merged %>%
  ungroup() %>%
  group_by(Zone_name, Province_name) %>% 
  summarize(
    `# of Districts` = length(unique(District_name)),
    `# of Villages - schools interventions` = sum(!is.na(`Total Number of Students for WASH in schools interventions`)),
    `# of Villages -  health facility interventions` = sum(!is.na(`Catchment Population on HF`)),
    `# of Villages -  Sanitation interventions` = sum(!is.na(`Number of Families for all Access to Sanitation interventions`)),
    `# of Villages -  Access to water interventions` = sum(!is.na(`Number of Families for all Access to water interventions`)),
    `# of Villages -  WASH supply distribution interventions` = sum(!is.na(`Number of Families for all WASH supply distribution interventions`)),
  ) %>% 
  mutate(Remark = "All villages") %>% 
  arrange(Province_name)

openxlsx::write.xlsx(merged_analysis_all_province, "output/Interventions_by_province.xlsx")

merged_analysis_all_province_interventions <- merged %>%
  ungroup() %>%
  group_by(Zone_name, Province_name) %>% 
  summarize(
    `# of Districts` = length(unique(District_name)),
    
    `# of Villages - schools interventions - Male Students` = sum(!is.na(`Male Students`) & `Male Students` != 0),
    `# of Villages - schools interventions - Female Students` = sum(!is.na(`Female Students`) & `Female Students` != 0),
    `# of Villages - schools interventions` = sum(!is.na(`Total Number of Students for WASH in schools interventions`)),
    
    `# of Villages -  health facility interventions` = sum(!is.na(`Catchment Population on HF`)),
    
    `# of Villages -  Sanitation interventions - Reach affected people with hygiene promotion` = sum(!is.na(`Reach affected people with hygiene promotion`)),
    `# of Villages -  Sanitation interventions - Construct emergency household latrines` = sum(!is.na(`Construct emergency household latrines`)),
    `# of Villages -  Sanitation interventions - Conduct clean up campaigns in emergency` = sum(!is.na(`Conduct clean up campaigns in emergency`)),
    `# of Villages -  Sanitation interventions - Distribute IEC materials` = sum(!is.na(`Distribute IEC materials`)),
    `# of Villages -  Sanitation interventions - Construct emergency and/or urban communal latrines and/or with handwashing facilities` = sum(!is.na(`Construct emergency and/or urban communal latrines and/or with handwashing facilities`)),
    `# of Villages -  Sanitation interventions - Rehabilitate emergency and/or urban communal latrines and/or with handwashing facilities` = sum(!is.na(`Rehabilitate emergency and/or urban communal latrines and/or with handwashing facilities`)),
    `# of Villages -  Sanitation interventions` = sum(!is.na(`Number of Families for all Access to Sanitation interventions`)),
    
    `# of Villages -  Access to water interventions - Drill new boreholes and equip them with hand pumps` = sum(!is.na(`Drill new boreholes and equip them with hand pumps`)),
    `# of Villages -  Access to water interventions - Construct deep boreholes with any other type of pumps - household connection` = sum(!is.na(`Construct deep boreholes with any other type of pumps - household connection`)),
    `# of Villages -  Access to water interventions - Rehabilitate boreholes and equip them with hand pumps` = sum(!is.na(`Rehabilitate boreholes and equip them with hand pumps`)),
    `# of Villages -  Access to water interventions - Construct deep boreholes with solar powered piped system - household connection` = sum(!is.na(`Construct deep boreholes with solar powered piped system - household connection`)),
    `# of Villages -  Access to water interventions - Rehabilitate Gravity-fed/spring piped system - public taps` = sum(!is.na(`Rehabilitate Gravity-fed/spring piped system - public taps`)),
    `# of Villages -  Access to water interventions - Rehabilitate dug wells and equip them with hand pumps` = sum(!is.na(`Rehabilitate dug wells and equip them with hand pumps`)),
    `# of Villages -  Access to water interventions - Support people with safe drinking water through water trucking` = sum(!is.na(`Support people with safe drinking water through water trucking`)),
    `# of Villages -  Access to water interventions - Rehabilitate deep boreholes with solar powered piped system - public taps` = sum(!is.na(`Rehabilitate deep boreholes with solar powered piped system - public taps`)),
    `# of Villages -  Access to water interventions - Treat water systems with chlorine` = sum(!is.na(`Treat water systems with chlorine`)),
    `# of Villages -  Access to water interventions - Rehabilitate deep boreholes with solar powered piped system - household connection` = sum(!is.na(`Rehabilitate deep boreholes with solar powered piped system - household connection`)),
    `# of Villages -  Access to water interventions - Rehabilitate deep boreholes with any other type of pumps - public taps` = sum(!is.na(`Rehabilitate deep boreholes with any other type of pumps - public taps`)),
    `# of Villages -  Access to water interventions - Construct deep boreholes with solar powered piped system - public taps` = sum(!is.na(`Construct deep boreholes with solar powered piped system - public taps`)),
    `# of Villages -  Access to water interventions - Construct Gravity-fed/spring piped system - household connection` = sum(!is.na(`Construct Gravity-fed/spring piped system - household connection`)),
    `# of Villages -  Access to water interventions - Extend water systems` = sum(!is.na(`Extend water systems`)),
    `# of Villages -  Access to water interventions - Construct Gravity-fed/spring piped system - public taps` = sum(!is.na(`Construct Gravity-fed/spring piped system - public taps`)),
    `# of Villages -  Access to water interventions - Conduct water quality testing (includes physical, microbial and chemical parameters)` = sum(!is.na(`Conduct water quality testing (includes physical, microbial and chemical parameters)`)),
    `# of Villages -  Access to water interventions - Rehabilitate deep boreholes with any other type of pumps - household connection` = sum(!is.na(`Rehabilitate deep boreholes with any other type of pumps - household connection`)),
    `# of Villages -  Access to water interventions` = sum(!is.na(`Number of Families for all Access to water interventions`)),
    
    `# of Villages -  WASH supply distribution interventions - Provide people with WASH supplies including soap or jerrycans, and or other WASH supplies` = sum(!is.na(`Provide people with WASH supplies including soap or jerrycans, and or other WASH supplies`)),
    `# of Villages -  WASH supply distribution interventions - Provide people with WASH supplies including soap or jerrycans` = sum(!is.na(`Provide people with WASH supplies including soap or jerrycans`)),
    `# of Villages -  WASH supply distribution interventions - Provide people with hygiene kits` = sum(!is.na(`Provide people with hygiene kits`)),
    `# of Villages -  WASH supply distribution interventions - Provide people with consumable kits` = sum(!is.na(`Provide people with consumable kits`)),
    `# of Villages -  WASH supply distribution interventions - Provide people with household water treatment products including aqua-tabs, PUR, or ceramic filters` = sum(!is.na(`Provide people with household water treatment products including aqua-tabs, PUR, or ceramic filters`)),
    `# of Villages -  WASH supply distribution interventions - Provide people with water kits` = sum(!is.na(`Provide people with water kits`)),
    `# of Villages -  WASH supply distribution interventions` = sum(!is.na(`Number of Families for all WASH supply distribution interventions`))
  ) %>% 
  mutate(Remark = "All villages") %>% 
  arrange(Province_name)

openxlsx::write.xlsx(merged_analysis_all_province_interventions, "output/Interventions_by_province_activities.xlsx")
