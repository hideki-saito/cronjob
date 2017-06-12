source_directory = "E:\_Python\etc\cronjob\sample" # directory that contains csv files

interval_minutes = 15    # for cronjob
cronJob_user = "username" # username of OS

template_excel = "template.xlsx" # the parent of the output excel file
csv_sheet = {'EECensus': "Employees", "DEPCensus": "Spouse & Dependents", "DEPBenefits": "Current Benefits", "EEBenefits": "Current Benefits"} # pairs of the csvfilename and sheet of excel.