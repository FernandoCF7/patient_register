 import settings
import gime

#-----------------------------------------------------------------------------#
day = '080123'

#list to generate the excel file by enterprise
exel_enterprises = ['/particular/']#_allarrived_

#set subsidiary
subsidiary = '01'#hermita

#inline excell files
inlineEF = True
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

