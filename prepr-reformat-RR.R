# Load required libraries
library(readxl)
library(dplyr)
library(hms)
library(lubridate)
library(svDialogs)

# Function to select a file using a GUI
select_file <- function() {
  dlg_open(title = "Select Excel File", filters = matrix(c("Excel Files", "*.xlsx"), ncol = 2))$res
}

# Function to select a folder using a GUI
select_folder <- function() {
  dlg_dir(title = "Select Output Folder")$res
}

# Select the input file
file_path <- select_file()

# Check if a file was selected
if (length(file_path) == 0) {
  stop("No file selected. Exiting...")
}

# Read the "Time Series" sheet
time_series_data <- read_excel(file_path, sheet = "Time Series")

# Extract the "Real time (hh:mm:ss)" column
time_column <- time_series_data %>% select(`Real time (hh:mm:ss)`)

# Read the "R-R" sheet
rr_data <- read_excel(file_path, sheet = "R-R")

# Extract the "RR" column
rr_column <- rr_data %>% select(RR)

# Convert the initial time to POSIXct
initial_time <- as.POSIXct(time_column$`Real time (hh:mm:ss)`[1], format = "%H:%M:%S")

# Calculate the cumulative time vector using the RR intervals, starting from the initial time
cumulative_time <- initial_time + cumsum(rr_column$RR) / 1000

# Function to format time with milliseconds accurately
format_time_with_ms <- function(time) {
  time_str <- format(time, "%Y.%m.%d %H:%M:%S")
  ms <- round((time - trunc(time)) * 1000)
  sprintf("%s.%03d", time_str, ms)
}

# Apply the formatting function to the cumulative time vector
formatted_time <- sapply(cumulative_time, format_time_with_ms)

# Create the final matrix
initial_time_formatted <- format_time_with_ms(initial_time)
final_matrix <- data.frame(
  Time = c(initial_time_formatted, formatted_time),
  Peak_to_Peak_Distance = c(NA, rr_column$RR)
)

# Remove the first row (initial NA) to ensure no NA values
final_matrix <- final_matrix[-1, ]

# Select the output folder
output_folder <- select_folder()

# Check if a folder was selected
if (length(output_folder) == 0) {
  stop("No folder selected. Exiting...")
}

# Extract the original file name without extension
original_file_name <- tools::file_path_sans_ext(basename(file_path))

# Define the output file path with "PREPR" added
output_file_path <- file.path(output_folder, paste0(original_file_name, "_PREPR.txt"))

# Write the final matrix to a text file
write.table(final_matrix, file = output_file_path, row.names = FALSE, col.names = TRUE, sep = "\t", quote = FALSE)

# Inform the user that the file has been saved
cat("The output file has been saved to:", output_file_path, "\n")
