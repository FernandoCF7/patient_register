#import pandas as pd
#orig_url = 'https://github.com/FernandoCF7/denatbioRegistroPacientes/blob/main/'
#filePath_listExam = ("{0}"+"listadoDeExamenes/listExam.csv?raw=true").format(orig_url)
#pd_listExam = (pd.read_csv(filePath_listExam, usecols=["COD INT", "EXAMEN"]))

import settings

#-----------------------------------------------------------------------------#
#set parameters to make excel files
day = '070123'

#list to generate the excel file by enterprise; Note _allarrived_ makes for all registered enterprises
exel_enterprises = ['/WONKA CHOCOLATE/']

#set subsidiary
subsidiary = '01'#hermita

#inline excell files
inlineEF = False
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#set parameters to settings.py
settings.set_daily_parameters(day, exel_enterprises, subsidiary, inlineEF)
#-----------------------------------------------------------------------------#

#----------------------------------------------------------------------------#
#make excel's (for day)

#antygen
# settings.antigen_excel()

#antybody
# settings.antybody_excel()

#laboratory all
# settings.laboratory_excel()

#laboratory no-covids
settings.laboratoryNoCovid_excel()

# #cobranza
# settings.cobranza_excel()

#enterprises
# settings.enterprises_excel()
#-----------------------------------------------------------------------------#

# #-----------------------------------------------------------------------------#
# #make excel's, (for month)

# dummy_counter = 0
# for day_tmp in range( 1, int(day[0:2])+1 ):
        
#     settings.set_daily_parameters( "{}{}".format(str(day_tmp).zfill(2),day[2:]), exel_enterprises, subsidiary, inlineEF )
#     settings.join_month_parameters(dummy_counter)
#     dummy_counter += max( [max(settings.idx_enterprise), max(settings.idx_patients)] ) + 1

# # #laboratory all
# settings.laboratory_excel_m()

# #cobranza
# settings.cobranza_excel_m()

# #antygen
# settings.antigen_excel_m()

# #antybody
# settings.antybody_excel_m()

# #laboratory no-covids
# settings.laboratoryNoCovid_excel_m()

# #enterprises
# settings.enterprises_excel_m()
#-----------------------------------------------------------------------------#

