
#https://cran.r-project.org/web/packages/xlsx/index.html

#https://cran.r-project.org/web/packages/openxlsx/index.html


# Fetches last version from CRAN
install.packages('openxlsx')

# Fetches a specific version as of a specific date from Posit
#install.packages('openxlsx', repos = 'https://packagemanager.posit.co/cran/2023-06-30')

library(openxlsx)

packageVersion("openxlsx")
