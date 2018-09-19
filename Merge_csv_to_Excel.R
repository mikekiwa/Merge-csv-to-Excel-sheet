library(data.table)  # fread() function
library(XLConnect)   # Excel and csv files
library(svDialogs)   # Dialog boxes

# Ask user for path to csv files
folder <- dlgInput(title = "Merge csv", "Enter path to csv files (use '/' instead of '\\': ", Sys.info()["user"])$res

# Set the working directory to user path
setwd(folder)

# Create and load Excel file
wb <- loadWorkbook("Output.xlsx", create=TRUE)

# Get list of csv files in directory
pattern.ext <- "\\.csv$"
files <- dir(folder, full=TRUE, pattern=pattern.ext)

# Use substring of file names for sheet names (Excel limit) and remove extension 
files.nms <- substr(basename(files),1,31)
files.nms <- gsub(pattern.ext, "", files.nms)

# Set the names 
names(files) <- files.nms

# Iterate over each csv and output to Excel sheet
for (nm in files.nms) {
  
  # Ingest csv file 
  df <- fread(files[nm])
  
  # Create the sheet and name as substr of file name  
  createSheet(object = wb, name = nm)
  
  # Output the contents of the csv 
  writeWorksheet(object = wb, data = df, sheet = nm, header = TRUE, rownames = NULL)
  
  # Create a custom anonymous cell style
  cs <- createCellStyle(wb)
  
  # Specify to wrap the text
  setWrapText(object = cs, wrap = TRUE)
  
  # Set column width
  setColumnWidth(object = wb, sheet = nm, column = 1:50, width = -1)
}

saveWorkbook(wb)

# Check to see if Excel file exists
if (file.exists("Output.xlsx") & file.size("Output.xlsx") > 8731) {
  dlg_message("Your Excel file has been created.")$res
} else {
  dlg_message("Error: Your file may not have been created or compelted properly. Please verify and try again if necessary.")$res
}