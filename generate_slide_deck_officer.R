library(officer)
library(rvg)
library(readxl)
library(tidyr)
library(dplyr)
library(mschart)
library(ggplot2)
library(ggforce)
library(lubridate)

#################
## Data input
#################
excel_fn <- "officer_demo_cat_data.xlsx"
template_fn <- "slide-deck-officer-template.pptx"
report_params <- read_excel(excel_fn, sheet = "Report_Params")

report_type <- report_params$Report_Name
report_type_short <- report_params$Report_Name_Short

start_date <- paste(day(report_params$Start_Date), 
                    month(report_params$Start_Date, label = TRUE), 
                    year(report_params$Start_Date))
end_date <- paste(day(report_params$End_Date), 
                  month(report_params$End_Date, label = TRUE), 
                  year(report_params$End_Date))

#################
## ppt template
#################

ppt_out <- read_pptx(template_fn) 

# color palettes
oi_colors <- as.vector(palette.colors()[2:9])
ppt_colors <- c("#4F81BD", "#C0514D", "#9BBB59", "#8064A2", "#4bABC6")
gbr_colors <- c("#9BBB59","#4F81BD", "#C0514D", "#8064A2", oi_colors)

#################
## SLIDE 1 bullet points
#################
## S1 TEXT
slide1_notes <- "Slide 1 features bulleted text and an image. 
unordered_list() can be used to create bulletpoints with different levels. 
external_img() allows you to insert image files into a picture location in your pptx template."

slide1_txt_title <- "Data driven decision making"
slide1_txt_bodytxt <- "We are making better use of the data and information collected during Cat Encounters (CEs) to inform decision-making and appropriate actions to help increase positive cat experiences"

CE_txt <- read_excel(excel_fn, sheet = "Open CE graphic text", col_names = TRUE)
slide1_bullets <- unordered_list(
  level_list = c(1, 2, 2, 1, 2),
  str_list = c(CE_txt$Slide_Text))

image_file <- external_img("mr_jones.png")

## S1 format
ppt_out <- add_slide(ppt_out, layout = "slide1", master = "Theme1") %>%
  ph_with(value = slide1_txt_title, location = ph_location_label(ph_label = "title")) %>%
  ph_with(value = slide1_txt_bodytxt,location = ph_location_label(ph_label = "bodytext")) %>%
  ph_with(value = slide1_bullets, location = ph_location_label(ph_label = "bodybullet")) %>%
  ph_with(value = image_file, location = ph_location_label(ph_label = "image")) %>%
  set_notes(value = slide1_notes, location = notes_location_type("body"))



# #################
# ## SLIDE 1 alt graphic text
# #################
# 
# ppt_out <- add_slide(ppt_out, layout = "slide1_alt", master = "Theme1") %>%
#   ph_with(value = "Cat Encounters", location = ph_location_label(ph_label = "title")) %>%
#   ph_with(value = report_type,location = ph_location_label(ph_label = "subtitle1")) %>%
#   ph_with(value = paste(start_date, "to", end_date),location = ph_location_label(ph_label = "subtitle2")) %>%
#   ph_with(value = paste0(CE_txt$Slide_Text[1], start_date, " to ", end_date), 
#           location = ph_location_label(ph_label = "textbox0")) %>%
#   ph_with(value = CE_txt$Slide_Text[2], location = ph_location_label(ph_label = "textbox1")) %>%
#   ph_with(value = CE_txt$Slide_Text[3], location = ph_location_label(ph_label = "textbox2")) %>%
#   ph_with(value = CE_txt$Slide_Text[4], location = ph_location_label(ph_label = "textbox3")) %>%
#   ph_with(value = CE_txt$Slide_Text[5], location = ph_location_label(ph_label = "textbox4")) %>%
#   set_notes(value = slide1_notes, location = notes_location_type("body"))

# #################
# ## SLIDE 2 ggplot + mschart
# #################

slide2_notes <- "Slide 2 features editable plots made with ggplot2 and mschart.
Wrapping the ggplot in dml() turn the ggplot into graphic vectors in which design and text can be edited (left plot). 
ms_barchart() creates a native microsoft chart in which both chart design, text, and underlying data can be edited (right plot)."


CE_month_type <- read_excel(excel_fn, sheet = "CEs by month type") %>%
  replace(is.na(.), 0) %>%
  mutate(FY = factor(FY, ordered = TRUE)) %>%
  mutate(Quarter = factor(Quarter, ordered = TRUE)) %>%
  mutate(Month = factor(Month,  c("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"), ordered = TRUE))

CE_month_type_long <- CE_month_type %>%
  pivot_longer(cols = 4:(dim(CE_month_type)[2]), names_to = "type", values_to = "count") %>%
  unite("time", FY:Month, remove = FALSE)

# CE_quarter_type_long <- CE_month_type_long %>%
#   group_by(Quarter, type) %>%
#   summarize(`CE Count` = sum(count)) %>% ungroup()
# 
# slide3_ms_bars <- ms_barchart(data = CE_quarter_type_long,
#                               x = "Quarter", y = "CE Count", group = "type")
# slide3_ms_bars <- chart_settings( x = slide3_ms_bars,
#                                   dir="vertical", grouping="clustered", gap_width = 150 )
# 
# ppt_out <- add_slide(ppt_out, layout = "slide2_alt", master = "Theme1") %>%
#   ph_with(value = slide3_ms_bars, location = ph_location_label(ph_label = "graph1"))

#########################
### s2 CHART 1 STACKED BARS X TIME

## https://stackoverflow.com/questions/20571306/multi-row-x-axis-labels-in-ggplot-line-chart
## https://github.com/tidyverse/ggplot2/issues/2933
gg_vbar_month_by_type <- ggplot(CE_month_type_long, aes(y = count, x = Month, fill = reorder(type, count, sum))) + 
  geom_col(position="stack", color = "white") +
  facet_row(vars(FY, Quarter),  scales = "free_x", space = "free", strip.position = "bottom") +
  labs(x = "", y = "", fill = "") +
  theme_minimal() +
  scale_fill_manual(values = rev(ppt_colors)) +
  theme(panel.spacing.x = unit(0,"line"), strip.placement = 'outside', strip.background.x = element_blank(), 
        panel.grid.minor.x = element_blank(), panel.grid.major.x = element_blank(),
        axis.title = element_text(size=14, face="bold"),
        strip.text = element_text(size=12, face = "bold", vjust = 3),
        legend.text = element_text(size=12), legend.title = element_text(size=12), legend.position = "right",
        #legend.key.spacing.y = unit(1.75, "cm"), legend.margin = margin(1.75, 0, 0, 0, "cm"), #for ggplot2 v3.5+
        legend.spacing.y = unit(1.75, "cm"), #for ggplot2 < v3.5
        plot.title = element_blank(),
        axis.text.x = element_text(angle=45, vjust = 1.75, hjust = 1.25, size = 12),
        rect = element_rect(fill = "transparent"), 
        panel.background = element_rect(fill = "transparent", color = NA), plot.background = element_rect(fill = "transparent", color = NA),
        legend.background = element_rect(fill = "transparent", color = NA), legend.box.background = element_rect(fill = "transparent", color = NA), legend.key = element_rect(fill = "transparent", color = NA)) +
  guides(fill = guide_legend(byrow = TRUE))

gg_vbar_month_by_type <- dml(ggobj = gg_vbar_month_by_type, bg = "transparent")

#########################
### s2 CHART 2 horiz bars mschart

CE_grade <- read_excel(excel_fn, sheet = "CE Grade", col_names = TRUE) %>%  mutate(Type = factor(Type, levels = Type))

CE_grade_long <- CE_grade %>% pivot_longer(cols = 2:dim(CE_grade)[2], names_to = "grade", values_to = "count")

# slide4_gg_hbars <- ggplot(CE_grade_long, aes(x = count, y = reorder(Type, -count, sum, na.rm=TRUE), fill = grade)) + 
#   geom_bar(position = position_fill(reverse = TRUE), stat="identity", color = "white", width = .45) +
#   geom_text(aes(label = grade, fontface = "bold"),  position = position_fill(reverse = TRUE, vjust = .5), color = "white", size = 4.75) +
#   labs(x = "", y = "", fill = "", title = "CE grade") +
#   theme_minimal() +
#   scale_fill_manual(values = oi_colors[1:6], drop = FALSE) +
#   theme(panel.grid.minor = element_blank(), panel.grid.major = element_blank(),
#         axis.text = element_blank(), axis.ticks = element_blank(),
#         legend.position = "none", 
#         plot.title = element_blank(),
#         panel.background = element_rect(fill = "transparent", color = NA), plot.background = element_rect(fill = "transparent", color = NA))
# 
# slide4_gg_hbars <- dml(ggobj = slide4_gg_hbars, bg = "transparent")

ms_hbars_grade <- ms_barchart(data = CE_grade_long, x = "Type", y = "count", group = "grade") %>%
  chart_settings(., dir="horizontal", grouping="percentStacked", gap_width = 150, overlap = 100 ) %>%
  chart_labels(., title = NULL, xlab = NULL, ylab = NULL) %>%
  chart_data_labels(., position = "ctr", show_serie_name = TRUE) %>%
  chart_labels_text(., values = list(A = fp_text(color = "white", bold = TRUE, font.size=20), B = fp_text(color = "white", bold = TRUE, font.size=20),
                                     C = fp_text(color = "white", bold = TRUE, font.size=20), D = fp_text(color = "white", bold = TRUE, font.size=20))) %>%
  chart_data_fill(., values = c(A = oi_colors[1], B = oi_colors[2], C = oi_colors[3], D = oi_colors[4])) %>%
  chart_data_stroke(., values = c(A = oi_colors[1], B = oi_colors[2], C = oi_colors[3], D = oi_colors[4])) %>%
  chart_ax_x(.,major_tick_mark = "none", minor_tick_mark = "none", display = FALSE, tick_label_pos = "none") %>%
  chart_ax_y(.,major_tick_mark = "none", minor_tick_mark = "none", display = FALSE, tick_label_pos = "none") %>%
  chart_theme(., legend_position = "n", grid_major_line = fp_border(style = "none"), grid_major_line_y = fp_border(style = "none"))

#print(slide4_ms_bars, preview = TRUE)

## S2 ppt out
ppt_out <- add_slide(ppt_out, layout = "slide2", master = "Theme1") %>%
  ph_with(value = paste(report_type_short, "CEs by Month and Type"), location = ph_location_label(ph_label = "title")) %>%
  ph_with(value = paste(start_date, "and", end_date), location = ph_location_label(ph_label = "subtitle")) %>%
  ph_with(value = "CEs by Month and Type", location = ph_location_label(ph_label = "titlegraph1")) %>%
  ph_with(value = gg_vbar_month_by_type, location = ph_location_label(ph_label = "graph1")) %>%
  ph_with(value = "CE Grade", location = ph_location_label(ph_label = "titlegraph2")) %>%
  ph_with(value = ms_hbars_grade, location = ph_location_label(ph_label = "graph2")) %>%
  set_notes(value = slide2_notes, location = notes_location_type("body"))

#################
## SLIDE 3 ggplot
#################

slide3_notes <- "Slide 2 features static plots made with ggplot2.
Without wrapping the ggplots in dml(), the plots are saved as static image files in the final powerpoint."

#########################
### s3 CHART 2 horiz bar - outcome
CE_outcomes <- read_excel(excel_fn, sheet = "CE Outcome", col_names = TRUE) 

gg_hbar_outcome <- ggplot(CE_outcomes, aes(x = `Outcome Count`, y = "", fill = reorder(`Outcome Label`, -`Outcome Count`))) + geom_bar(position = position_fill(reverse = TRUE), stat="identity", color = "white") +
  labs(x = "", y = "", fill = "", title = "") +
  theme_minimal() +
  scale_fill_manual(values = gbr_colors, drop = FALSE) +
  theme(panel.grid.minor.y = element_blank(), panel.grid.major.y = element_blank(),
        axis.text = element_blank(), axis.ticks = element_blank(),
        plot.title = element_blank(),
        legend.text = element_text(size=12), legend.position = "bottom", legend.margin=margin(t=-25), legend.key.size = unit(.75, "line"),
        rect = element_rect(fill = "transparent")) + 
  guides(fill=guide_legend(nrow=3, byrow = TRUE))

#gg_hbar_outcome <- dml(ggobj = gg_hbar_outcome, bg = "transparent")

#########################
### s3 CHART 2 table - actions
CE_actions <- read_excel(excel_fn, sheet = "Actions", col_names = TRUE) 

CE_actions_long <- CE_actions %>% pivot_longer(cols = 2:dim(CE_actions)[2], names_to = "type", values_to = "count") %>%
  mutate(type = ordered(type, levels = c("Positive Experience", "Neutral Experience", "Negative Experience"))) 

gg_table_actions <- ggplot(CE_actions_long, aes(x = count, y = reorder(`Action Count`, count, sum, na.rm = TRUE), fill = type)) + 
  geom_bar(stat = "identity") + 
  facet_grid(~ type, drop = FALSE) + 
  geom_text(aes(x = (max(count, na.rm = TRUE)/2), label = count, fontface = "bold"),  color = "black", size = 4) +
  labs(x = "", y = "", fill = "") + 
  theme_bw() +
  scale_fill_manual(values = gbr_colors) +
  theme(panel.grid.minor.x = element_blank(), panel.grid.major.x = element_blank(),
        axis.text.x = element_blank(), axis.ticks = element_blank(),axis.text.y = element_text(size = 12),
        legend.position = "none", 
        plot.title = element_blank(),
        panel.background = element_rect(fill = "transparent", color = NA), plot.background = element_rect(fill = "transparent", color = NA),
        strip.background = element_blank(), panel.spacing = unit(0, "mm"),
        strip.text.x = element_text(size = 16, face = "bold"), panel.border = element_rect(color = "darkgray"),
        rect = element_rect(fill = "transparent"))

ppt_out <- add_slide(ppt_out, layout = "slide3", master = "Theme1") %>%
  ph_with(value = "Cat Encounter Outcomes and Actions", location = ph_location_label(ph_label = "title")) %>%
  ph_with(value = paste(start_date, "and", end_date, "-", sum(CE_outcomes$`Outcome Count`, na.rm = TRUE), "CEs"), location = ph_location_label(ph_label = "subtitle")) %>%
  ph_with(value =  "Cat Encounter Outcomes", location = ph_location_label(ph_label = "titlegraph1")) %>%
  ph_with(value = gg_hbar_outcome, location = ph_location_label(ph_label = "graph1")) %>%
  ph_with(value = "Actions associated with Cat Encounter Outcomes", location = ph_location_label(ph_label = "titlegraph2")) %>%
  ph_with(value = gg_table_actions, location = ph_location_label(ph_label = "graph2")) %>%
  set_notes(value = slide3_notes, location = notes_location_type("body"))

#############
## Save ppt
##############

ppt_out_fn <- paste0("Cat Encounters_", report_type_short, " Numbers_", format(report_params$Start_Date, '%d-%b-%Y'), " to ", format(report_params$End_Date, '%d-%b-%Y'), ".pptx")
print(ppt_out, target = ppt_out_fn)
