# This script uses arbitrary commission rates to avoid disclosing sensitive business info.

library(readxl)
library(writexl)

INPUT_FILE = "C:\\Users\\bk.kwon\\Downloads\\FY26 Billable Report - 29 Jan 2026.xlsx"
OUTPUT_FILE = "C:\\Users\\bk.kwon\\Downloads\\FY26 Billable Report - 29 Jan 2026 - Incorrect Commission.xlsx"

# Read the billable report in and trim the last row
billable <- read_excel(INPUT_FILE)
billable <- billable[-nrow(billable), ]

# Calculate the correct commission based on ESTNUM
commission <- function (est) {
  est <- as.numeric(est)
  if ((200 <= est & est <= 299) | (900 <= est & est <= 999)) {
    return (0.000)
  } else if (600 <= est & est <= 799) {
    return (0.05)           # actual commission rates are different
  } else {
    return (0.1)            # these commission rates are arbitrary
  }
}

# Round commissions to 3 decimal places
round_to_3dp <- function (x) {
  return (round(x, digits = 3))
}

billable$Commission <- sapply(as.numeric(billable$Commission), round_to_3dp)

billable$"Correct Commission" <- sapply(billable$ESTNUM, commission)

# Move Correct Commission column next to ESTNUM
new_index <- c(1:6, 44, 7:43)
billable <- billable[, new_index]

# Check Correct Commission against Commission column
# and select rows with incorrect commission
incorrect_billable <- billable[billable$"Correct Commission" != billable$Commission, ]

# Write in OUTPUT_FILE
write_xlsx(incorrect_billable, OUTPUT_FILE)