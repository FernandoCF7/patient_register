from pandas import read_csv as pd_read_csv, concat as pd_concat, isnull as pd_isnull, Index as pd_Index, DataFrame as pd_DataFrame, ExcelWriter as pd_ExcelWriter, Series as pd_Series


from os import path as os_path
from numpy import any as np_any, prod as np_prod, array as np_array, NaN as np_NaN, argwhere as np_argwhere, where as np_where
from re import compile as re_compile, IGNORECASE as re_IGNORECASE
from time import strftime as time_strftime
from unicodedata import normalize as unicodedata_normalize
from sys import exit as sys_exit
from copy import deepcopy as copy_deepcopy
from datetime import datetime


#inline excell files
orig_url = 'https://github.com/FernandoCF7/denatbioRegistroPacientes/blob/main/'

color_by_day = {
    0: "AMARILLO",#monday
    1: "ROSA",#tuesday
    2: "ANARANJADO",#wednesday
    3: "GRIS",#thursday
    4: "VERDE",#friday
    5: "BLANCO",#saturday
    6: "AZUL"#sunday
}

#-----------------------------------------------------------------------------#
#setting variables from settings
def set_projectmodule_parameters(currentPath, inlineEF):

    #read pd_listExam file
    set_pd_listExam(currentPath, inlineEF)

    #read pd_listSurrogate file
    set_pd_listSurrogate(currentPath, inlineEF)

    #read clavesNombresEmpresa file
    set_df_enterpriseNames(currentPath, inlineEF)

    #read codeEnterprise file
    set_codeEnterpriseFile(currentPath, inlineEF)
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#read registro file
def get_csvFile(currentPath, yymmddPath):
    
    filePath_registro = os_path.join("{0}", "input_src", "{1}.txt").format(currentPath, yymmddPath)

    csvFile = (pd_read_csv(filePath_registro,sep='*', dtype={2: str}))

    csvFile.columns = ["firstName", "secondName", 'thirdName']

    #set as upper
    csvFile["firstName"] = csvFile["firstName"].str.upper()
    csvFile["secondName"] = csvFile["secondName"].str.upper()

    #remove spaces at the begining and end of the string
    csvFile["firstName"] = csvFile["firstName"].str.strip()
    csvFile["secondName"] = csvFile["secondName"].str.strip()
    csvFile["thirdName"] = csvFile["thirdName"].str.strip()
    
    #mask Ñ's with @@'s
    for idx, _ in enumerate(csvFile["firstName"]):
        csvFile["firstName"][idx] = csvFile["firstName"][idx].replace("Ñ", "@@")
        csvFile["secondName"][idx] = csvFile["secondName"][idx].replace("Ñ", "@@")
    
    #remove acents
    csvFile["firstName"] = csvFile["firstName"].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
    csvFile["secondName"] = csvFile["secondName"].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

    #restore @@'s with Ñ's
    for idx, _ in enumerate(csvFile["firstName"]):
        csvFile["firstName"][idx] = csvFile["firstName"][idx].replace("@@","Ñ")
        csvFile["secondName"][idx] = csvFile["secondName"][idx].replace("@@","Ñ")
    
    return csvFile
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#read pd_listExam file
def set_pd_listExam(currentPath, inlineEF):
    
    global pd_listExam

    if inlineEF:
        filePath_listExam = ("{0}"+"src_lists/exams/listExam.csv?raw=true").format(orig_url)
    else:
        filePath_listExam = os_path.join("{0}", "src_lists", "exams", "listExam.csv").format(currentPath)

    pd_listExam = (pd_read_csv(filePath_listExam, usecols=["NUMERIC_CODE", "SHORT_EXAMN_NAME"]))

    #set index of pd_listExam as NUMERIC_CODE column
    pd_listExam.set_index("NUMERIC_CODE", inplace=True)

    #read listExam locally file
    filePath_listExam_tmp = os_path.join("{}", "src_lists", "dischargeWhenOffline", "listExam.csv").format(currentPath)
    listExam_locally = pd_read_csv(filePath_listExam_tmp, usecols=["NUMERIC_CODE", "SHORT_EXAMN_NAME"])

    listExam_locally.set_index("NUMERIC_CODE", inplace=True)

    # append listExam_locally to pd_listExam
    for idx, row in listExam_locally.iterrows():
        if idx in pd_listExam.index:#update the Exam
            pd_listExam.EXAMEN[idx] = row["SHORT_EXAMN_NAME"]
        # else:#append examn
        #     pd_listExam = pd_concat([pd_listExam, listExam_locally.loc[idx]], axis=0)
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#read clavesNombresEmpresa file
def set_df_enterpriseNames(currentPath, inlineEF):

    global df_enterpriseNames

    if inlineEF:
        filePath_clavesNombresEmpresa = ("{0}"+"src_lists/enterprise/claveEnterpriseName.csv?raw=true").format(orig_url)
    else:
        filePath_clavesNombresEmpresa = os_path.join("{0}", "src_lists", "enterprise", "claveEnterpriseName.csv").format(currentPath)

    df_enterpriseNames = pd_read_csv(filePath_clavesNombresEmpresa, keep_default_na=False)

    #set index of enterpriseNames as clave column
    df_enterpriseNames.set_index("clave", inplace=True)

    #read clavesNombresEmpresa locally file
    filePath_clavesNombresEmpresa_tmp = os_path.join("{}", "src_lists", "dischargeWhenOffline", "claveEnterpriseName.csv").format(currentPath)
    clavesNombresEmpresa_locally = pd_read_csv(filePath_clavesNombresEmpresa_tmp, encoding='latin-1', keep_default_na=False)

    clavesNombresEmpresa_locally.set_index("clave", inplace=True)

    #Concatenated df_enterpriseNames and clavesNombresEmpresa_locally
    df_enterpriseNames = pd_concat([df_enterpriseNames, clavesNombresEmpresa_locally], axis=0)
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#read pd_listSurrogate file
def set_pd_listSurrogate(currentPath, inlineEF):
    
    global pd_listSurrogate

    if inlineEF:
        filePath_list = ("{0}"+"src_lists/surrogate/surrogateList.csv?raw=true").format(orig_url)
    else:
        filePath_list = os_path.join("{0}", "src_lists", "surrogate", "surrogateList.csv").format(currentPath)

    pd_listSurrogate = (pd_read_csv(filePath_list, usecols=["CODIGO", "NAME"]))

    #set index of pd_listSurrogate as CODE column
    pd_listSurrogate.set_index("CODIGO", inplace=True)

    #read pd_listSurrogate locally file
    filePath_list_tmp = os_path.join("{}", "src_lists", "dischargeWhenOffline", "surrogateList.csv").format(currentPath)
    list_locally = pd_read_csv(filePath_list_tmp, usecols=["CODIGO", "NAME"])

    list_locally.set_index("CODIGO", inplace=True)
    
    #check no repeated surrogate of list_locally
    for idx, row in list_locally.iterrows():
        
        #repeated code 
        if idx in pd_listSurrogate.index:
            sys_exit("ERROR: EL código {} del subrrogante {} (archivo: src_lists/dischargeWhenOffline/surrogate/surrogateList.csv) ya es utilizado en el listado de subrrogantes del sistema; asocie un código distinto".format(idx, row.NAME))
        
        #repeated NAME
        if row.NAME in list(pd_listSurrogate.NAME):
            sys_exit("ERROR: EL subrrogante {} con código {} (archivo: src_lists/dischargeWhenOffline/surrogate/surrogateList.csv) ya está en el listado de subrrogantes del sistema".format(row.NAME, idx))

    #append
    pd_listSurrogate = pd_concat([pd_listSurrogate, list_locally])
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#search for enterprise; extract index of the enterprises
def get_idx_enterprise(csvFile_firstName):

    idx_enterprise = []
    patern = re_compile(r'/.*?/')

    for idx, val in enumerate(csvFile_firstName):#for each field
        
        #search for "ENTERPRISE" word
        if val.find("ENTERPRISE") != -1:
            idx_enterprise.append(idx)

    return idx_enterprise
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#validate that all patients values are filled
def checkFilledfiels(columnIdx, csvFile, idx_patients, day):
    
    if np_any(pd_isnull(csvFile.iloc[idx_patients, columnIdx])) or np_any(csvFile.iloc[idx_patients, columnIdx]==""):
        
        tmp = np_where(pd_isnull(csvFile.iloc[idx_patients, columnIdx]))
        
        if not(np_any(tmp)):
            tmp = np_where(csvFile.iloc[idx_patients,columnIdx]=="")

        infoPatients = csvFile.iloc[idx_patients[tmp],:]
        sys_exit("""Registro no valido (fecha: {}) para el (los) paciente(s):\n {}""".format(day, infoPatients))
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get enterprise names and codecs
def get_enterpriseNames(OSR, csvFile, idx_enterprise, day):
    #idx_enterprise --> array, index of enterprice at CSV file
    
    patern = re_compile(r'/.*?/')

    enterpriseNames = []
    enterpriseCodecs = []
    
    for val in idx_enterprise:#for each enterprise
        
        #get the enterprise name
        enterprise_name = patern.search(csvFile["firstName"][val]).group()
        
        #search enterprise_name in codeEnterpriseFile
        logicTMP = enterprise_name == codeEnterpriseFile['enterprise'].str.upper()
        
        #not defined enterprise mesage error
        if not(any(logicTMP)):
            print('''OPERACION FALLIDA (fecha: {})\nEmpresa no definida: {} en el archivo \
    codeEnterprise.csv; folio OSR: {}'''.format(day, enterprise_name,OSR[val]))
            sys_exit("")
        
        #get enterprise code of clavesNombresEmpresa
        try:
            enterprise_code = codeEnterpriseFile.clave[logicTMP==True].item()
        except ValueError:
            print("""OPERACION FALLIDA (fecha: {})\nLa empresa {} se encuentra definida más \
    de una vez de la misma manera en el archivo codeEnterprise.csv""".format(day, enterprise_name))
            sys_exit()
            
        #append enterpriseCodecs
        enterpriseCodecs.append(enterprise_code)
        
        #append enterprise_code
        enterpriseNames.append(df_enterpriseNames.loc[enterprise_code].item())

    return enterpriseNames, enterpriseCodecs
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get the enterprise names and codecs forExclusiveExcel
def get_enterpriseNames_exclusiveExcel(exel_enterprises, day):
    
    enterpriseNames_forExclusiveExcel = []
    enterpriseCodecs_forExclusiveExcel = []
    
    for enterprise_name_ in exel_enterprises:
            
        #set as upper
        enterprise_name = enterprise_name_.upper()
    
        #mask Ñ's with @@'s
        enterprise_name = enterprise_name.replace("Ñ", "@@")
        
        #remove acents
        enterprise_name = unicodedata_normalize('NFKD', enterprise_name)
        enterprise_name = enterprise_name.encode('ascii', errors='ignore').decode('utf-8')

        #restore @@'s with Ñ's
        enterprise_name = enterprise_name.replace("@@", "Ñ")
        
        #search enterprise_name in codeEnterpriseFile
        logicTMP = enterprise_name == codeEnterpriseFile['enterprise'].str.upper()
        
        #not defined enterprise mesage error
        if not(any(logicTMP)):
            print('''OPERACION FALLIDA (fecha: {})\n La empresa {} del listado list_enterprise_forExclusiveExcel no está definida en el archivo codeEnterprise.csv'''.format(day, enterprise_name))
            sys_exit("")
        
        #get enterprise code of clavesNombresEmpresa
        try:
            enterprise_code = codeEnterpriseFile.clave[logicTMP==True].item()
        except ValueError:
            print("""OPERACION FALLIDA(fecha: {})\nLa empresa {} se encuentra definida más \
    de una vez de la misma manera en el archivo codeEnterprise.csv""".format(day, enterprise_name))
            sys_exit()
            
        #append enterpriseCodecs_forExclusiveExcel
        enterpriseCodecs_forExclusiveExcel.append(enterprise_code)
        
        #append enterpriseNames_forExclusiveExcel
        enterpriseNames_forExclusiveExcel.append(df_enterpriseNames.loc[enterprise_code].item())

    return enterpriseNames_forExclusiveExcel, enterpriseCodecs_forExclusiveExcel
    #-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#read codeEnterprise file
def set_codeEnterpriseFile(currentPath, inlineEF):
    
    global codeEnterpriseFile

    #codeEnterpriseFile --> pd data frame, /empresa/ CODE

    if inlineEF:
        filePath_codeEnterprise = ("{0}"+"src_lists/enterprise/codeEnterprise.csv?raw=true").format(orig_url)
    else:
        filePath_codeEnterprise = os_path.join("{0}", "src_lists", "enterprise", "codeEnterprise.csv").format(currentPath)

    codeEnterpriseFile = pd_read_csv(filePath_codeEnterprise, encoding='latin-1', keep_default_na=False)

    #set as upper
    codeEnterpriseFile["enterprise"] = codeEnterpriseFile["enterprise"].str.upper()

    #remove acents
    codeEnterpriseFile["enterprise"] = codeEnterpriseFile["enterprise"].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

    #read codeEnterprise locally file
    filePath_codeEnterprise_tmp = os_path.join("{}", "src_lists", "dischargeWhenOffline", "codeEnterprise.csv").format(currentPath)

    codeEnterpriseFile_locally = pd_read_csv(filePath_codeEnterprise_tmp, encoding='latin-1', keep_default_na=False)

    #set as upper
    codeEnterpriseFile_locally["enterprise"] = codeEnterpriseFile_locally["enterprise"].str.upper()

    #remove acents
    codeEnterpriseFile_locally["enterprise"]=codeEnterpriseFile_locally["enterprise"].str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

    #Concatenated codeEnterpriseFile and codeEnterpriseFile_locally
    codeEnterpriseFile = pd_concat([codeEnterpriseFile, codeEnterpriseFile_locally], axis=0)
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#set the enterprise name by patient
def get_listEnterpriseNameByPatient(idx_enterprise, enterpriseNames, idx_patients, len_csvFile):
    
    listEnterpriseNameByPatient = []
    
    for idx, val in enumerate(idx_enterprise[:-1]):
        for _ in range(val, idx_enterprise[idx+1]):
            listEnterpriseNameByPatient.append(enterpriseNames[idx])
    else:
        if len(idx_enterprise)==1: idx=-1
        for _ in range(idx_enterprise[idx+1], len_csvFile):
            listEnterpriseNameByPatient.append(enterpriseNames[idx+1])       

    #listEnterpriseNameByPatient --> convert list to dict
    tmp = dict()
    for val in idx_patients:
        tmp[val] = listEnterpriseNameByPatient[val]

    return tmp
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#set the enterprise code by patient
def get_listEnterpriseCodeByPatient(idx_enterprise, enterpriseCodecs, idx_patients, len_csvFile):
    
    listEnterpriseCodeByPatient = []
    
    for idx, val in enumerate(idx_enterprise[:-1]):
        for _ in range(val, idx_enterprise[idx+1]):
            listEnterpriseCodeByPatient.append(enterpriseCodecs[idx])
    else:
        try:
            for _ in range(idx_enterprise[idx+1], len_csvFile):
                listEnterpriseCodeByPatient.append(enterpriseCodecs[idx+1])
        except:
            if idx_enterprise:
                for _ in range(0, len(idx_patients)+1):
                    listEnterpriseCodeByPatient.append(enterpriseCodecs[0])

    #listEnterpriseNameByPatient --> convert list to dict
    tmp = dict()
    for val in idx_patients:
        tmp[val] = listEnterpriseCodeByPatient[val]

    return tmp
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get shift (turno)
def get_shift(idx_enterprise, csvFile, OSR, day):

    shift = []

    #for each enterprise
    for val in idx_enterprise:
        
        #get shift
        tmp=""
        
        if csvFile["firstName"][val].find("VESPERTINO")!=-1:
            tmp="V"
        elif csvFile["firstName"][val].find("MATUTINO")!=-1:
            tmp="M"
        
        #not assigned shift
        if not tmp: sys_exit(NAS.format(day, OSR[val]))
        
        #append in shift dict
        shift.append(tmp)
    
    return shift
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#set the enterprise code by patient
def get_listShiftByPatient(idx_enterprise, idx_patients, shift, len_csvFile):

    listShiftByPatient = []
    
    for idx, val in enumerate(idx_enterprise[:-1]):
        for _ in range(val, idx_enterprise[idx+1]):
            listShiftByPatient.append(shift[idx])
    else:

        try:
            for _ in range(idx_enterprise[idx+1], len_csvFile):
                listShiftByPatient.append(shift[idx+1])
        except:
            if idx_enterprise:
                for _ in range(0, len(idx_patients)+1):
                    listShiftByPatient.append(shift[0])

    #listEnterpriseNameByPatient --> convert list to dict
    tmp = dict()
    for val in idx_patients:
        tmp[val] = listShiftByPatient[val]

    return tmp
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#search for 'express' examns
def get_idx_express(idx_patients, csvFile):

    patern_express = re_compile(r'EXPRESS', flags=re_IGNORECASE)#'EXPRESS' pattern
    
    idx_express = []

    for val in idx_patients:#search for each patient
        
        #csvFile.firstName --> patient name
        patientName = csvFile.firstName[val]

        if patientName.find("EXPRESS")!=-1:#finded 'EXPRESS' inside the patient name
            idx_express.append(val)
    
    return idx_express
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#search for vuelo examns
def get_idx_vuelo(idx_patients, csvFile):

    patern_vuelo = re_compile(r'FLIGHT', flags=re_IGNORECASE)#'FLIGHT' pattern
    
    idx_vuelo = []

    for val in idx_patients:#search for each patient
        
        #csvFile.firstName --> patient name
        patientName = csvFile.firstName[val]
        if patientName.find("FLIGHT")!=-1:#finded 'FLIGHT' in the patient name
            idx_vuelo.append(val)

    return idx_vuelo
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#csvFile.firstName --> cut the 'FLIGHT' part
def update_csvFile_firstName_vuelo(idx_vuelo, csvFile):

    patern_vuelo = re_compile(r'FLIGHT', flags=re_IGNORECASE)#'FLIGHT' pattern
    
    for val in idx_vuelo:
        
        #csvFile.firstName --> patient name
        patientName = csvFile.firstName[val]

        tmp = patern_vuelo.search(patientName)
        csvFile.firstName[val] = patientName[0:tmp.start()]+patientName[tmp.end():]
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get the exam code by patient as dictionary
def get_ECBP(idx_patients, csvFile):

    ECBP = dict()#ECBP-->ExamCodeByPatien
    listSurrogate_code = list(pd_listSurrogate.index)

    for val in idx_patients:#for each patien
        
        #get the string examCodecs
        exams = csvFile.thirdName[val]
        
        #Split the string considering " "
        exams = exams.split()

        #asociate each exam by their surrogant
        tmp_ = {}
        for exam in exams:
            tmp = exam.split("_")
            exam_code = int(tmp[0])
            try:
                surrogate_code = tmp[1]
            except:
                surrogate_code = 1

            if type(surrogate_code) == str:
                try:
                    surrogate_code = int(surrogate_code)
                except:
                    print("ERROR: el código de subrrogante '{}', del examen con código {} del paciente {} {} debe ser un número".format(surrogate_code, exam_code, csvFile.firstName[val], csvFile.secondName[val]))
                    sys_exit()

            #validate if surrogate_code is registered
            if surrogate_code != 1:
               
                if surrogate_code not in listSurrogate_code:
                    
                    print("ERROR: el código de subrrogante '{}', del examen con código {} del paciente {} {} no existe".format(surrogate_code, exam_code, csvFile.firstName[val], csvFile.secondName[val]))
                    sys_exit()
            
            subrrogate_name = "" if surrogate_code == 1 else pd_listSurrogate.loc[surrogate_code]["NOMBRE"].upper()
            tmp_[exam_code] = subrrogate_name
        
        ECBP[val] = tmp_
    
    return ECBP
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#update csvFile.firstName cuting the 'EXPRESS' part
def update_csvFile_firstName_express(idx_express, csvFile):

    patern_express = re_compile(r'EXPRESS', flags=re_IGNORECASE)#'EXPRESS' pattern
    
    for val in idx_express:
        
        patientName = csvFile.firstName[val]

        tmp = patern_express.search(patientName)
        csvFile.firstName[val] = patientName[0:tmp.start()]+patientName[tmp.end():]
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#set ECBP as list of str´s
def get_ECBP_str(ECBP):
    
    ECBP_str = dict()
    
    for key, value in ECBP.items():

        tmp = list(map(str, value.keys()))
        
        ECBP_str[key]="\n".join(tmp)
    
    return ECBP_str
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#examn product code by patient
def get_EPCBP(ECBP):
 
    EPCBP = dict()
 
    for key, value in ECBP.items():
        
        EPCBP[key] = np_prod(list(value.keys()))
    
    return EPCBP
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#set each study kind with a color
def get_color_as_study(ECBP):

    #ECBNC --> Exam color by no PCR covits
    ECBNC = dict()
    for idx, val in ECBP.items():
        if any(np_array(list(val.keys()))!=2):
            ECBNC[idx] = True

    #ECBAC --> Exam color by antigen covit
    ECBAC = dict()
    for idx, val in ECBP.items():
        if any(np_array(list(val.keys())) == 487):
            ECBAC[idx] = True

    #ECBABC --> Exam color by anti body covit
    ECBABC = dict()
    for idx, val in ECBP.items():
        if any(np_array(list(val.keys())) == 491):
            ECBABC[idx] = True

    #ECBCABC --> Exam color by cuantitative anti body covit
    ECBCABC = dict()
    for idx, val in ECBP.items():
        if any(np_array(list(val.keys())) == 569):
            ECBCABC[idx] = True

    #ECBSP --> Exam color by sars plus
    ECBSP = dict()
    for idx, val in ECBP.items():
        if any(np_array(list(val.keys())) == 1009):
            ECBSP[idx] = True
        
    return ECBNC, ECBAC, ECBABC, ECBCABC, ECBSP
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#set the exams name 
def get_examNameList(idx_patients, csvFile, ECBP, format_, day):

    examNameList = dict()

    for val in idx_patients:#for each patien with spetial prices
        
        #ensure exams code are recored
        try:
            examsName = pd_listExam.SHORT_EXAMN_NAME[list((ECBP[val]).keys())].tolist()
        except KeyError:
            print(CEND.format(day, csvFile.firstName[val], csvFile.secondName.iloc[val]))        
            sys_exit()
        
        #ensure exams name are recored
        for tmp in examsName:
            if type(tmp) == float:#(nan is type float); exam name is empty at the excel file
                print(CEND.format(day, csvFile.firstName[val], csvFile.secondName.iloc[val]))        
                sys_exit()

        if format_ == "as_str":
            tmp = "\n"
            examNameList[val] = tmp.join(examsName)
        elif format_ == "as_list":
            examNameList[val] = examsName
    
    return examNameList
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#make internal code
def get_codeInt_Lab_Cob(idx_patients, idx_express, day, subsidiary, ECBP, listEnterpriseCodeByPatient, listShiftByPatient, EPCBP):

    codeIntLab = dict()
    codeIntCob = dict()
    
    for idx, val in enumerate(idx_patients):

        express = "U" if val in idx_express else "N"
        
        codeIntCob[val] = day+'-'+subsidiary+'-'+str(idx+1).zfill(3)
        
        #Check if the exam is covid type
        examCovid = "C"
        if not [i for i in [2, 487, 491, 492, 569, 1009] if  i in list((ECBP[val]).keys())]:
            examCovid = "O"
        
        codeIntLab[val] = day+'-'+subsidiary+'-'+str(idx+1).zfill(3)+express+examCovid+listEnterpriseCodeByPatient[val]+listShiftByPatient[val]+"P"+str(EPCBP[val]).zfill(5)
        
    return codeIntLab, codeIntCob
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#dict --> asociate each patient (key) with its corresponding enterprise (value)
def get_dict_pattient_enterprise(idx_patients, idx_enterprise):

    counter = 0
    dict_pattient_enterprise = {}

    for idx, val in enumerate(idx_patients):
        
        if counter+1 < len(idx_enterprise):
            if idx_enterprise[counter+1] < val:
                counter += 1

        dict_pattient_enterprise[val] =  idx_enterprise[counter]
    
    return dict_pattient_enterprise
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get the index's of no covit patients
def get_idx_noCovits(idx_patients, ECBP, dict_pattient_enterprise):

    #covid_exam_codecs = [2, 487, 491, 492, 569, 1009]
    covid_exam_codecs = []#this part has been set to include all studies (eaven covid) at no covid list
    tmp = []

    for val in idx_patients:
        
        for i_ in list((ECBP[val]).keys()):
            if i_ not in covid_exam_codecs:
                tmp.append(val)
                break

    idx_patients_noCovits = pd_Index(data=tmp)#convert list into pd index

    #get the enterprise idx associated with idx_patients_noCovits
    idx_enterprise_patients_noCovits = list(set([dict_pattient_enterprise[tmp] for tmp in idx_patients_noCovits]))

    return idx_patients_noCovits, idx_enterprise_patients_noCovits
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get index's for antigen patients
def get_idx_antigenCovit(idx_patients, ECBP, dict_pattient_enterprise):
    
    tmp = []

    for val in idx_patients:

        if [i for i in [487] if  i in list((ECBP[val]).keys())]:
            tmp.append(val)

    idx_patients_antigenCovit = pd_Index(data=tmp)#convert list into pd index

    #get the enterprise idx associated with idx_patients_antigenCovit
    idx_enterprise_patients_antigenCovit = list(set([dict_pattient_enterprise[tmp] for tmp in idx_patients_antigenCovit]))

    return idx_patients_antigenCovit, idx_enterprise_patients_antigenCovit
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get index's for qualitative antybody covit patients
def get_idx_antibodyCovit(idx_patients, ECBP, dict_pattient_enterprise):
    
    tmp = []
    
    for val in idx_patients:

        if [i for i in [491] if  i in list((ECBP[val]).keys())]:
            tmp.append(val)

    idx_patients_antibodyCovit = pd_Index(data=tmp)#convert list into pd index

    #get the enterprise idx associated with idx_patients_antibodyCovit
    idx_enterprise_patients_antibodyCovit = list(set([dict_pattient_enterprise[tmp] for tmp in idx_patients_antibodyCovit]))

    return idx_patients_antibodyCovit, idx_enterprise_patients_antibodyCovit
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get the index's of list_enterprise_forExclusiveExcel patients
def get_idx_enterpriseExclusive(enterpriseCodecs_forExclusiveExcel, idx_patients, listEnterpriseCodeByPatient, dict_pattient_enterprise):

    idx_patients_enterprise_forExclusiveExcel_asDict = {}
    idx_enterprise_enterprise_forExclusiveExcel_asDict = {}

    for codeEnterprise_ in enterpriseCodecs_forExclusiveExcel:
        
        tmp = []
        for val in idx_patients:

            if [i for i in [codeEnterprise_] if  i in listEnterpriseCodeByPatient[val]]:
                tmp.append(val)

        idx_patients_enterprise_forExclusiveExcel_asDict[codeEnterprise_] = pd_Index(data=tmp)

        #get the enterprise idx associated with codeEnterprise_
        idx_enterprise_enterprise_forExclusiveExcel_asDict [codeEnterprise_] = list(set([dict_pattient_enterprise[tmp] for tmp in idx_patients_enterprise_forExclusiveExcel_asDict[codeEnterprise_]]))
    
    return idx_patients_enterprise_forExclusiveExcel_asDict, idx_enterprise_enterprise_forExclusiveExcel_asDict
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
def delete_multiple_element(list_object, indices):
    indices = sorted(indices, reverse=True)
    for idx in indices:
        if idx < len(list_object):
            list_object.pop(idx)
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
def make_laboratory_excel(idx_patients_, idx_enterprise_, codeIntLab, csvFile, ECBP_str, currentPath, yymmddPath, day, idx_express, idx_vuelo, examNameList, ECBNC, ECBAC, ECBABC, ECBCABC, ECBSP, enterpriseNames_asDict, path_=""):
    #idx_patients_ --> pandas index, the index (in the CSV file) of patients to show
    #idx_enterprise --> list, the index (in the CSV file) of enterprises to show

    #----------------------------------------------------------------------------#
    #merge, as list, idx_patients_ and idx_enterprise_    
    idx = idx_patients_.tolist() + idx_enterprise_
    
    idx.sort()
    #----------------------------------------------------------------------------#

    #----------------------------------------------------------------------------#
    #Export to excel-->laboratori
    df_toExcel = pd_DataFrame(
        {
            'OSR': np_NaN,
            'COD INT': {x:codeIntLab[x] for x in idx_patients_},
            'NOMBRE': csvFile['firstName'][idx].str.strip(), 
            'APELLIDO': csvFile['secondName'][idx].str.strip(),
            'EXAMEN': {x:examNameList[x] for x in idx_patients_},
            'COD': {x:ECBP_str[x] for x in idx_patients_},
            'ESTATUS': np_NaN,
            'RESULTADO': np_NaN,
            'ENVIO': np_NaN,
            'REVISO': np_NaN,
            'HORA ENVIO': np_NaN
        }
    )
    #----------------------------------------------------------------------------#

    #----------------------------------------------------------------------------#
    #set the OSR code in df_toExcel
    df_toExcel.loc[idx_enterprise_, "OSR"] = csvFile["secondName"][idx_enterprise_].str.strie()
    #----------------------------------------------------------------------------#

    #----------------------------------------------------------------------------#
    pathTosave = os_path.join("{0}","output_src","{1}", "{2}{3}.xlsx").format(currentPath, yymmddPath,day,path_)

    with pd_ExcelWriter(pathTosave, engine='xlsxwriter') as writer:
                                                           
        #Convert the dataframe to an XlsxWriter Excel object.
        df_toExcel.to_excel(writer, sheet_name=day, index=False)
        
        #Get the xlsxwriter workbook and worksheet objects.
        workbook  = writer.book
        worksheet = writer.sheets[day]
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #set formats
        
        #Wrap 'EXAMEN' and 'PRECIO' column
        widthColumn = workbook.add_format({'text_wrap': True})
        worksheet.set_column('E:E', 40, widthColumn)
        worksheet.set_column('F:F', 6, widthColumn)
        worksheet.set_column('H:H', 18, widthColumn)
        
        border_format = workbook.add_format({'border': 1})
        #-----------------------------------------------------------------------------#
    
        #-----------------------------------------------------------------------------#
        #Set express format
        
        #express
        expressFormat = workbook.add_format(
            {
                'align': 'left',
                'valign': 'vcenter',
                'bold': True,
                'font_color': 'black',
                'bg_color': 'orange'
            }
        )
    
        for tmp in list(set(idx_express) & set(idx_patients_.tolist())):
            tmp_ = idx.index(tmp)
            worksheet.write_string('G'+str(tmp_+2)+':G'+str(tmp_+2), "EXPRESS", expressFormat)
        #-----------------------------------------------------------------------------#
    
        #-----------------------------------------------------------------------------#
        #Set vuelo format
        
        #express
        vueloFormat = workbook.add_format(
            {
                'align': 'left',
                'valign': 'vcenter',
                'bold': True,
                'font_color': 'black',
                'bg_color': '#0080FF'
            }
        )
        
        for tmp in  list(set(idx_vuelo) & set(idx_patients_.tolist())):
            tmp_ = idx.index(tmp)
            worksheet.write_string('H'+str(tmp_+2)+':H'+str(tmp_+2), "FLIGHT", vueloFormat)
        #-----------------------------------------------------------------------------#
    
        #-----------------------------------------------------------------------------#
        #Add color to the cell according to the exam
        
        #colors:
        cell_format_mostaza = workbook.add_format({'bg_color': '#FF9933'})
        cell_format_mostaza.set_text_wrap()
        cell_format_magenta = workbook.add_format({'bg_color': 'magenta'})
        cell_format_magenta.set_text_wrap()
        cell_format_yellow = workbook.add_format({'bg_color': 'yellow'})
        cell_format_yellow.set_text_wrap()
        cell_format_green = workbook.add_format({'bg_color': 'green'})
        cell_format_green.set_text_wrap()
        cell_format_lime = workbook.add_format({'bg_color': 'lime'})
        cell_format_lime.set_text_wrap()

        #ECBNC --> Exam color by no covits; Mostaza
        tmp = ECBNC.keys()
        tmp_ = [x for x in idx_patients_ if x in tmp]
        ECBNC_tmp = {x:ECBNC[x] for x in tmp_}
        for key in ECBNC_tmp:
            key_ = idx.index(key)
            worksheet.write_string('E'+str(key_+2)+':E'+str(key_+2), examNameList[key], cell_format_mostaza)
        
        #ECBAC --> Exam color by antigen covit
        tmp = ECBAC.keys()
        tmp_ = [x for x in idx_patients_ if x in tmp]
        ECBAC_tmp = {x:ECBAC[x] for x in tmp_}
        for key in ECBAC_tmp:
            key_ = idx.index(key)
            worksheet.write_string('E'+str(key_+2)+':E'+str(key_+2), examNameList[key], cell_format_yellow)
        
        #ECBABC --> Exam color by anti body covit
        tmp = ECBABC.keys()
        tmp_ = [x for x in idx_patients_ if x in tmp]
        ECBABC_tmp = {x:ECBABC[x] for x in tmp_}
        for key in ECBABC_tmp:
            key_ = idx.index(key)
            worksheet.write_string('E'+str(key_+2)+':E'+str(key_+2), examNameList[key], cell_format_magenta)
        
        #ECBCABC --> Exam color by cuantitative anti body covit
        tmp = ECBCABC.keys()
        tmp_ = [x for x in idx_patients_ if x in tmp]
        ECBCABC_tmp = {x:ECBCABC[x] for x in tmp_}
        for key in ECBCABC_tmp:
            key_ = idx.index(key)
            worksheet.write_string('E'+str(key_+2)+':E'+str(key_+2), examNameList[key], cell_format_magenta)
        
        #ECBSP --> Exam color by sars plus
        tmp = ECBSP.keys()
        tmp_ = [x for x in idx_patients_ if x in tmp]
        ECBSP_tmp = {x:ECBSP[x] for x in tmp_}
        for key in ECBSP_tmp:
            key_ = idx.index(key)
            worksheet.write_string('E'+str(key_+2)+':E'+str(key_+2), examNameList[key], cell_format_lime)
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #Add border
        numRows=len(df_toExcel)
        
        worksheet.conditional_format('A1:K'+str(numRows+1), {'type':'no_blanks', 'format':border_format})
        
        worksheet.conditional_format('A1:K'+str(numRows+1), {'type':'blanks', 'format':border_format})
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #Set format to enterprise
        merge_format = workbook.add_format(
            {
                'align': 'center',
                'valign': 'vcenter',
                'bold': True,
                'font_color': 'white',
                'bg_color': 'black'
            }
        )
        
        for indx_, val, in enumerate(idx_enterprise_):
            val_ = idx.index(val)
            worksheet.merge_range('B'+str(val_+2)+':K'+str(val_+2), enterpriseNames_asDict[val], merge_format)
        #-----------------------------------------------------------------------------#
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
def make_no_covid_excel(
            idx_patients_,
            idx_enterprise_,
            codeIntLab,
            csvFile,
            currentPath,
            yymmddPath,
            day,
            idx_express,
            idx_vuelo,
            examNameList_nested,
            ECBP,
            enterpriseNames_asDict,
            path_=""
        ):

    #idx_patients_ --> pandas index, the index (in the CSV file) of patients to show
    #idx_enterprise --> list, the index (in the CSV file) of enterprises to show

    idx_patients_ = idx_patients_.tolist()
    idx_patients_.sort()
    idx_enterprise_.sort()

    #----------------------------------------------------------------------------#
    #merge idx_patients_ and idx_enterprise_    
    idx_patients_and_enterprise = idx_patients_ + idx_enterprise_
    
    idx_patients_and_enterprise.sort()
    #----------------------------------------------------------------------------#

    #----------------------------------------------------------------------------#
    #set correlation pd index and idx_patients_and_enterprise
    excelIdx_pdIndx = {val:idx_ for idx_, val in enumerate(idx_patients_and_enterprise)}
    
    #update excelIdx_pdIndx considering there patients that have more than one exam
    examNameList_nested_without_covid = {x:y for x,y in examNameList_nested.items() if x in idx_patients_}
    ECBP_without_covid = copy_deepcopy(ECBP)
    
    #quit the covid part (if have) to examNameList_nested_without_covid
    for key in examNameList_nested_without_covid:
        idx_to_quit = []
        key_to_quit = []

        for idx, tmp in enumerate(list((ECBP[key]).keys())):    
            # if tmp in [2, 487, 491, 492, 569, 1009]: idx_to_quit.append(idx)#this part has been set to include all studies (eaven covid) at no covid list
            if tmp in []:
                idx_to_quit.append(idx)
                key_to_quit.append(tmp)
        
        if idx_to_quit:
            tmp = examNameList_nested_without_covid[key]
            delete_multiple_element(tmp, idx_to_quit)
            examNameList_nested_without_covid[key] = tmp

            tmp = ECBP_without_covid[key]
            for key_ in key_to_quit:
                del tmp[key_]
            
            #delete_multiple_element(tmp, idx_to_quit)
            ECBP_without_covid[key] = tmp
    
    excelIdxExams_pdIndx = {}
    for key, value in examNameList_nested_without_covid.items():
        
        if len(value) > 1:# --> more than one exam
            
            for key_, val_ in excelIdx_pdIndx.items():
                if key_ > key: excelIdx_pdIndx[key_] += (len(value)-1)

            excelIdxExams_pdIndx[key] = [excelIdx_pdIndx[key]+index_ for index_, value_ in enumerate(value)]
        else:
            excelIdxExams_pdIndx[key] = [excelIdx_pdIndx[key]]
    #----------------------------------------------------------------------------#
    
    #----------------------------------------------------------------------------#
    #Export to excel-->laboratori
    end_index = max( excelIdxExams_pdIndx.get(idx_patients_[-1]) )+1+4 if idx_patients_and_enterprise else 0
    df_toExcel = pd_DataFrame(
        {
            'OSR':np_NaN,#A --> 0
            'COD INT':np_NaN,#B --> 1
            'NOMBRE':np_NaN,#C --> 2
            'APELLIDO':np_NaN,#D --> 3
            'EXAMEN\nNOMBRE':np_NaN,#E --> 4
            'EXAMEN\nCOD':np_NaN,#F --> 5
            'ESTATUS':np_NaN,#G --> 6
            'RESULTADO\nSARS CoV2':np_NaN,#H --> 7
            'FECHA: RECEPCIÓN\nRESULTADO':np_NaN,#I --> 8
            'ENTREGA\nRESULTADO':np_NaN,#J --> 9
            'RECIBE\nRESULTADO':np_NaN,#K --> 10
            'ENVIÓ':np_NaN,#L --> 11
            'REVISÓ':np_NaN,#M --> 12
            'FECHA DE ENVÍO':np_NaN,# --> 13
            'HORA DE ENVÍO':np_NaN# --> 14
        }
        , index=range(0, end_index)
    )
    #----------------------------------------------------------------------------#
    
    #set valus at df_toExcel
    #OSR
    df_toExcel.loc[[excelIdx_pdIndx[tmp]+4 for tmp in idx_enterprise_], ["OSR"]] = [csvFile["secondName"][[tmp]].str.strip() for tmp in idx_enterprise_]

    #COD INT
    df_toExcel.loc[[excelIdx_pdIndx[tmp]+4 for tmp in idx_patients_],"COD INT"] = [codeIntLab[tmp] for tmp in idx_patients_]
    
    #NOMBRE
    df_toExcel.loc[[excelIdx_pdIndx[tmp]+4 for tmp in idx_patients_], ["NOMBRE"]] = [csvFile["firstName"][[tmp]].str.strip() for tmp in idx_patients_]

    #APELLIDO
    df_toExcel.loc[[excelIdx_pdIndx[tmp]+4 for tmp in idx_patients_], ["APELLIDO"]] = [csvFile["secondName"][[tmp]].str.strip() for tmp in idx_patients_]

    #EXAMEN NOMBRE
    tmp_ = []
    for tmp in idx_patients_:
        tmp_.extend(examNameList_nested_without_covid[tmp])
    
    tmp_0 = []
    for tmp in idx_patients_:
        tmp_0.extend([tmp_1+4 for tmp_1 in excelIdxExams_pdIndx[tmp]])
    
    df_toExcel.loc[tmp_0, "EXAMEN\nNOMBRE"] = tmp_

    #EXAMEN COD
    tmp_ = []
    for tmp in idx_patients_:
        tmp_.extend(list((ECBP_without_covid[tmp]).keys()))
    
    tmp_0 = []
    for tmp in idx_patients_:
        tmp_0.extend([tmp_1+4 for tmp_1 in excelIdxExams_pdIndx[tmp]])
    
    df_toExcel.loc[tmp_0, "EXAMEN\nCOD"] = tmp_

    #processed by (subrrogate)
    tmp_ = []
    for tmp in idx_patients_:
        tmp_.extend(list((ECBP_without_covid[tmp]).values()))
    
    df_toExcel.loc[tmp_0, "ENTREGA\nRESULTADO"] = tmp_
    #----------------------------------------------------------------------------#

    #----------------------------------------------------------------------------#
    pathTosave = os_path.join("{0}", "output_src", "{1}", "{2}{3}.xlsx").format(currentPath, yymmddPath, day, path_)

    with pd_ExcelWriter(pathTosave, engine='xlsxwriter') as writer:
                                                           
        #Convert the dataframe to an XlsxWriter Excel object.
        df_toExcel.to_excel(writer, sheet_name=day, index=False, startrow=1, header=False)
        
        #Get the xlsxwriter workbook and worksheet objects
        workbook  = writer.book
        worksheet = writer.sheets[day]
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #set pre-amble
        merge_format = workbook.add_format({'align':'left', 'valign':'vcenter'})
        worksheet.merge_range(0, 0, 0, 2, "RELACIÓN DE ÓRDENES DE SERVICIO", merge_format)
        worksheet.merge_range(1, 0, 1, 2, "POE-34-A", merge_format)
        worksheet.merge_range(2, 0, 2, 2, "COLOR: {}".format(color_by_day[datetime.strptime(day, "%d%m%y").weekday()]), merge_format)
        worksheet.merge_range(3, 0, 3, 2, "FECHA: {}".format(datetime.strptime(day, "%d%m%y")), merge_format)
        #-----------------------------------------------------------------------------#

        #-----------------------------------------------------------------------------#
        #set header
        
        #header format
        header_format = workbook.add_format(
            {
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1,
                'align':'center'
            }
        )

        #write the column headers with the defined format
        for idx_, value_ in enumerate(df_toExcel.columns.values):
            worksheet.write(4, idx_, value_, header_format)
        #-----------------------------------------------------------------------------#

        #-----------------------------------------------------------------------------#
        #set formats
        
        #Wrap EXAMEN and PRECIO column
        widthColumn = workbook.add_format({'text_wrap': True})
        worksheet.set_column(0,0, 12, widthColumn)# --> A
        worksheet.set_column(1,1, 30, widthColumn)# --> B
        worksheet.set_column(2,2, 30, widthColumn)# --> C
        worksheet.set_column(3,3, 30, widthColumn)# --> D
        worksheet.set_column(4,4, 40, widthColumn)# --> E
        worksheet.set_column(5,5, 10, widthColumn)# --> F
        worksheet.set_column(6,6, 15, widthColumn)# --> G
        worksheet.set_column(7,7, 12, widthColumn)# --> H
        worksheet.set_column(8,8, 24, widthColumn)# --> I
        
        worksheet.set_column(9,9, 18, widthColumn)# --> J
        worksheet.set_column(10,10, 18, widthColumn)# --> K
        worksheet.set_column(11,11, 18, widthColumn)# --> L
        worksheet.set_column(12,12, 18, widthColumn)# --> M
        worksheet.set_column(13,13, 16, widthColumn)# --> 
        worksheet.set_column(14,14, 16, widthColumn)# --> 
        
        border_format = workbook.add_format({'border': 1})
        #-----------------------------------------------------------------------------#
    
        #-----------------------------------------------------------------------------#
        #Set express and vuelo format
        
        #express
        expressFormat = workbook.add_format(
            {
                'align': 'left',
                'valign': 'vcenter',
                'bold': True,
                'font_color': 'black',
                'bg_color': 'orange'
            }
        )
    
        for tmp in list(set(idx_express) & set(idx_patients_)):
            tmp_ = excelIdx_pdIndx[tmp] 
            worksheet.write_string('G'+str(tmp_+2+4)+':G'+str(tmp_+2+4), "EXPRESS", expressFormat)
        
        #vuelo
        vueloFormat = workbook.add_format(
            {
                'align': 'left',
                'valign': 'vcenter',
                'bold': True,
                'font_color': 'black',
                'bg_color': 'blue'
            }
        )
        
        for tmp in list(set(idx_vuelo) & set(idx_patients_)):
            tmp_ = excelIdx_pdIndx[tmp] 
            worksheet.write_string('H'+str(tmp_+2+4)+':H'+str(tmp_+2+4), "FLIGHT", vueloFormat)
        #-----------------------------------------------------------------------------#
    
        #-----------------------------------------------------------------------------#
        #Add border
        numRows = len(df_toExcel)
        
        #15 is the "P" column
        worksheet.conditional_format(4, 0, numRows, 14, {'type':'no_blanks', 'format':border_format})
        
        worksheet.conditional_format(4, 0, numRows, 14, {'type':'blanks', 'format':border_format})
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #Set format to enterprise
        merge_format = workbook.add_format(
            {
                'align': 'center',
                'valign': 'vcenter',
                'bold': True,
                'font_color': 'white',
                'bg_color': 'black'
            }
        )
        
        for val in idx_enterprise_:
            #14 is the "O" column
            worksheet.merge_range(excelIdx_pdIndx[val]+1+4, 1, excelIdx_pdIndx[val]+1+4, 14, enterpriseNames_asDict[val], merge_format)
        #-----------------------------------------------------------------------------#
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
def make_excel_antigen_antibody(idx_patients_, resultadoColumn, exam, path_, day, day_to_save, codeIntCob, yymmddPath, currentPath):

    #----------------------------------------------------------------------------#
    #Export to excel-->antigen and antybody
    dictForDF = {
        'FECHA':day,
        'FOLIO': {x:codeIntCob[x] for x in idx_patients_},#codeIntLab
        #'PACIENTE':csvFile['firstName'][idx_patients_].str.strip()+' '+csvFile['secondName'][idx_patients_].str.strip(),
        'EXAMEN': exam#{x:examNameList[x] for x in idx_patients_},
    }
    
    for key, value in resultadoColumn.items():
        dictForDF[key] = value
    
    dictForDF['VALIDO'] = np_NaN
    dictForDF['RECIBE RESULTADOS'] = np_NaN
    
    
    df_toExcel=pd_DataFrame(dictForDF)
    #----------------------------------------------------------------------------#

    #----------------------------------------------------------------------------#
    pathTosave = os_path.join("{0}", "..", "..", "output_src", "{1}", "byExamCategory", "{2}_{3}.xlsx").format(currentPath, yymmddPath,day_to_save,path_)

    with pd_ExcelWriter(pathTosave, engine='xlsxwriter') as writer:

        #Convert the dataframe to an XlsxWriter Excel object
        df_toExcel.to_excel(writer, sheet_name=day_to_save, index=False)
        
        #Get the xlsxwriter workbook and worksheet objects
        workbook  = writer.book
        worksheet = writer.sheets[day_to_save]
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #set formats
        
        #Wrap EXAMEN
        widthColumn = workbook.add_format({'text_wrap': True})
        worksheet.set_column('B:B', 20, widthColumn)
        worksheet.set_column('C:C', 22, widthColumn)
        
        tmp = ["D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
        for x in range(0,len(resultadoColumn)):
            worksheet.set_column('{}:{}'.format(tmp[x],tmp[x]), 10, widthColumn)

        worksheet.set_column('{}:{}'.format(tmp[x+1],tmp[x+1]), 25, widthColumn)
        worksheet.set_column('{}:{}'.format(tmp[x+2],tmp[x+2]), 25, widthColumn)
        
        border_format = workbook.add_format({'border': 1})
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #Add border
        numRows=len(df_toExcel)
        
        dictt = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N"]
        
        
        worksheet.conditional_format('A1:{}'.format(dictt[len(dictForDF)-1])+str(numRows+1), {'type':'no_blanks', 'format':border_format})
        
        worksheet.conditional_format('A1:{}'.format(dictt[len(dictForDF)-1])+str(numRows+1), {'type':'blanks', 'format':border_format})
        #-----------------------------------------------------------------------------#
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
def make_excel_enterprise_forExclusiveExcel(idx_patients_, path_, day, day_to_save, codeIntLab, csvFile, examNameList, currentPath, yymmddPath, idx_express):

    #----------------------------------------------------------------------------#
    #Export to excel-->antigen and antybody
    dictForDF = {
        'FECHA':day,
        'FOLIO': {x:codeIntLab[x] for x in idx_patients_},
        'PACIENTE':csvFile['firstName'][idx_patients_].str.strip()+' '+csvFile['secondName'][idx_patients_].str.strip(),
        'EXAMEN': {x:examNameList[x] for x in idx_patients_},
        'ESTATUS':np_NaN
    }
    
    df_toExcel = pd_DataFrame(dictForDF)
    #----------------------------------------------------------------------------#

    #----------------------------------------------------------------------------#
    pathTosave = os_path.join("{0}", "..", "..", "output_src", "{1}","byEnterprise", "{2}{3}.xlsx").format(currentPath, yymmddPath,day_to_save,path_)

    with pd_ExcelWriter(pathTosave, engine='xlsxwriter') as writer:

        #Convert the dataframe to an XlsxWriter Excel object
        df_toExcel.to_excel(writer, sheet_name=day_to_save, index=False)
        
        #Get the xlsxwriter workbook and worksheet objects.
        workbook  = writer.book
        worksheet = writer.sheets[day_to_save]
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #set formats
        
        #Wrap EXAMEN
        widthColumn = workbook.add_format({'text_wrap': True})
        worksheet.set_column('B:B', 28, widthColumn)
        worksheet.set_column('C:C', 40, widthColumn)
        worksheet.set_column('D:D', 28, widthColumn)
        worksheet.set_column('E:E', 10, widthColumn)
        
        border_format = workbook.add_format({'border': 1})
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #Set express format
        
        #express
        expressFormat = workbook.add_format(
            {
                'align': 'left',
                'valign': 'vcenter',
                'bold': True,
                'font_color': 'black',
                'bg_color': 'orange'
            }
        )
    
        idx = idx_patients_.tolist()
        idx.sort()
        for tmp in list(set(idx_express) & set(idx_patients_.tolist())):
            tmp_ = idx.index(tmp)
            worksheet.write_string('E'+str(tmp_+2)+':E'+str(tmp_+2), "EXPRESS", expressFormat)
        #-----------------------------------------------------------------------------#

        #-----------------------------------------------------------------------------#
        #Add border
        numRows=len(df_toExcel)

        worksheet.conditional_format('A1:E'+str(numRows+1), {'type':'no_blanks', 'format':border_format})
        worksheet.conditional_format('A1:E'+str(numRows+1),{'type':'blanks', 'format':border_format})
        #-----------------------------------------------------------------------------#
#-----------------------------------------------------------------------------#

#----------------------------------------------------------------------------#
#excel-->cobranza
def make_excel_cobranza(idx_patients_, codeIntCob, day, csvFile, examNameList, ECBP_str, listEnterpriseNameByPatient, day_to_save, currentPath, yymmddPath, idx_express):
    
    df_toExcel = pd_DataFrame(
        {
            'COD INT': codeIntCob,
            'FECHA': day,
            'NOMBRE': csvFile['firstName'][idx_patients_].str.strip()+' '+csvFile['secondName'][idx_patients_].str.strip(),
            'EXAMEN': examNameList,
            'COD': ECBP_str,
            # 'PRECIO':pricesList_str,
            'ESTATUS': np_NaN,
            ' ': np_NaN,
            'EMPRESA': listEnterpriseNameByPatient
        }
    )
    #----------------------------------------------------------------------------#


    #----------------------------------------------------------------------------#
    pathTosave = os_path.join("{0}", "..", "..", "output_src", "{1}", "{2}_cobranza.xlsx").format(currentPath, yymmddPath,day_to_save)

    with pd_ExcelWriter(pathTosave, engine='xlsxwriter') as writer:

        #Convert the dataframe to an XlsxWriter Excel object
        df_toExcel.to_excel(writer, sheet_name=day_to_save, index=False)
        
        #Get the xlsxwriter workbook and worksheet objects.
        workbook  = writer.book
        worksheet = writer.sheets[day_to_save]
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #set formats
        
        #express
        merge_expressFormat = workbook.add_format(
            {
                'align': 'left',
                'valign': 'vcenter',
                'bold': True,
                'font_color': 'black',
                'bg_color': 'red'
                }
            )
        
        #Wrap 'EXAMEN' and 'PRECIO' column
        widthColumn = workbook.add_format({'text_wrap': True})
        worksheet.set_column('D:D', 40, widthColumn)
        worksheet.set_column('E:E', 6, widthColumn)
        
        border_format = workbook.add_format({'border': 1})
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #express_index
        express_index=[]
        for val in idx_express:
            express_index.append(np_argwhere(idx_patients_==val).item())
        
        for tmp in express_index:
            
            worksheet.merge_range('F'+str(tmp+2)+':G'+str(tmp+2), "EXPRESS", merge_expressFormat)
        #-----------------------------------------------------------------------------#
        
        #-----------------------------------------------------------------------------#
        #Add border
        numRows=len(df_toExcel)
        
        worksheet.conditional_format('A1:I'+str(numRows+1),{'type':'no_blanks', 'format':border_format})
        
        worksheet.conditional_format('A1:I'+str(numRows+1),{'type':'blanks', 'format':border_format})
    #-----------------------------------------------------------------------------#
#----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#define error mesages

#Precio no definido archivo express (listadoPreciosExpress.csv)
PNDAU="""OPERACION FALLIDA\nPrecio no definido en archivo: \
listadoPreciosExpress.csv para la empresa \"{0}\" con el examen \"{1}\"; \
paciente: \"{2}\", folio de paciente: \"{3}\""""

#Precio no definido archivo standar (listadoPreciosEtandar.csv)
PNDAS="""OPERACION FALLIDA\nPrecio no definido en archivo: \
listadoPreciosEtandar.csv para la empresa \"{0}\" con el examen \"{1}\"; \
paciente: \"{2}\", folio de paciente: \"{3}\" """

#Not autorized for assign spetial price
NAFASP="""OPERACION FALLIDA\nLa empresa \"{0}\" no está autorizada para \
ofertar precios especiales; retire el precio especial del archivo de ingreso\ 
de pacientes (.txt), o modifique el permiso asignado a dicha empresa en el \ 
archivo listadoPermisosCostosEspecialesEmpresas.csv"""

#Code of exam not defined
CEND='''OPERACION FALLIDA (fecha: {}):\nCódigo de examen no definido; paciente: {} {}'''

#Not assigned shift
NAS="""OPERACION FALLIDA (fecha: {})\nTurno no asignado a la OSR {}; asigne turno \
MATUTINO/VESPERTINO"""
#-----------------------------------------------------------------------------#
#-----------------------------------------------------------------------------#


#-----------------------------------------------------------------------------#
#local variables
currentPath = os_path.dirname(os_path.abspath(__file__))
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
#get parameters from make_excel_file.py
def set_daily_parameters(day, exel_enterprises, subsidiary, inlineEF):
    
    global idx_patients, idx_enterprise, idx_patients_antigenCovit, idx_patients_antibodyCovit, codeIntCob, yymmddPath, codeIntLab, csvFile, ECBP, ECBP_str, idx_express, idx_vuelo, examNameList, examNameList_nested, day_
    global ECBNC, ECBAC, ECBABC, ECBCABC, ECBSP, enterpriseNames_asDict, idx_patients_noCovits, idx_enterprise_patients_noCovits, listEnterpriseNameByPatient, idx_patients_enterprise_forExclusiveExcel_asDict, idx_enterprise_enterprise_forExclusiveExcel_asDict
    global enterpriseNames_forExclusiveExcel, enterpriseCodecs_forExclusiveExcel

    day_ = day

    #set parameters to projectmodule
    set_projectmodule_parameters(currentPath, inlineEF)

    #day, month and year
    dayy = day[0:2]
    month = day[2:4]
    year = day[4:6]
    
    #date as save file format
    yymmddPath = os_path.join('____'+year, '__'+month+year, day)

    #read csv patients file
    csvFile = get_csvFile(currentPath, yymmddPath)

    #search for enterprise; extract index of the enterprises
    idx_enterprise = get_idx_enterprise(csvFile["firstName"])

    #search for patients; extract index of the patients
    idx_patients = csvFile.index[~csvFile.index.isin(idx_enterprise)]

    #get the OSR (orden de servicio de referencia) code
    OSR = dict(csvFile["secondName"][idx_enterprise].str.strip())

    #-----------------------------------------------------------------------------#
    #check fill values in all patients entries
    checkFilledfiels(0, csvFile, idx_patients, day)
    checkFilledfiels(1, csvFile, idx_patients, day)
    checkFilledfiels(2, csvFile, idx_patients, day)
    #-----------------------------------------------------------------------------#

    #get enterprise names and codecs
    enterpriseNames, enterpriseCodecs = get_enterpriseNames(OSR, csvFile, idx_enterprise, day)

    #get the enterprise names and codecs forExclusiveExcel
    enterpriseNames_forExclusiveExcel, enterpriseCodecs_forExclusiveExcel = get_enterpriseNames_exclusiveExcel(exel_enterprises, day)

    #set idx_enterprise and enterpriseNames as dict
    enterpriseNames_asDict = dict(zip(idx_enterprise, enterpriseNames))

    #set the enterprise name by patient
    listEnterpriseNameByPatient = get_listEnterpriseNameByPatient(idx_enterprise, enterpriseNames, idx_patients, len(csvFile))

    #set the enterprise code by patient
    listEnterpriseCodeByPatient = get_listEnterpriseCodeByPatient(idx_enterprise, enterpriseCodecs, idx_patients, len(csvFile))

    #get shift (turno)
    shift = get_shift(idx_enterprise, csvFile, OSR, day)

    #set the enterprise code by patient
    listShiftByPatient = get_listShiftByPatient(idx_enterprise, idx_patients, shift, len(csvFile))

    #search for express examns
    idx_express = get_idx_express(idx_patients, csvFile)

    #update csvFile.firstName cuting the "express" part
    update_csvFile_firstName_express(idx_express, csvFile)

    #search for vuelo examns
    idx_vuelo = get_idx_vuelo(idx_patients, csvFile)

    #update csvFile.firstName cuting the "vuelo" part
    update_csvFile_firstName_vuelo(idx_vuelo, csvFile)

    #get the exam code by patient as dictionary
    ECBP = get_ECBP(idx_patients, csvFile)

    #set ECBP as list of str´s
    ECBP_str = get_ECBP_str(ECBP)

    #examn product code by patient
    EPCBP = get_EPCBP(ECBP)

    #-----------------------------------------------------------------------------#
    #set color as each study
    #ECBNC --> Exam color by no PCR covits
    #ECBAC --> Exam color by antigen covit
    #ECBABC --> Exam color by anti body covit
    #ECBCABC --> Exam color by cuantitative anti body covit
    #ECBSP --> Exam color by sars plus
    ECBNC, ECBAC, ECBABC, ECBCABC, ECBSP = get_color_as_study(ECBP)
    #-----------------------------------------------------------------------------#

    #set the exams name
    examNameList = get_examNameList(idx_patients, csvFile, ECBP, "as_str", day)

    #set the exams name nested list
    examNameList_nested = get_examNameList(idx_patients, csvFile, ECBP, "as_list", day)

    #make inter code
    codeIntLab, codeIntCob = get_codeInt_Lab_Cob(idx_patients, idx_express, day, subsidiary, ECBP, listEnterpriseCodeByPatient, listShiftByPatient, EPCBP)
    
    #asociate, as dict, each patient with its corresponding enterprise
    dict_pattient_enterprise = get_dict_pattient_enterprise(idx_patients, idx_enterprise)
    
    #get the index's of no covit patients
    idx_patients_noCovits, idx_enterprise_patients_noCovits = get_idx_noCovits(idx_patients, ECBP, dict_pattient_enterprise)
    
    #index's for antigen patients
    idx_patients_antigenCovit, idx_enterprise_patients_antigenCovit = get_idx_antigenCovit(idx_patients, ECBP, dict_pattient_enterprise)
    
    #index's for antibody patients
    idx_patients_antibodyCovit, idx_enterprise_patients_antibodyCovit = get_idx_antibodyCovit(idx_patients, ECBP, dict_pattient_enterprise)
    
    #get the index's of list_enterprise_forExclusiveExcel patients
    idx_patients_enterprise_forExclusiveExcel_asDict, idx_enterprise_enterprise_forExclusiveExcel_asDict = get_idx_enterpriseExclusive(enterpriseCodecs_forExclusiveExcel, idx_patients, listEnterpriseCodeByPatient, dict_pattient_enterprise)
#-----------------------------------------------------------------------------#

#-----------------------------------------------------------------------------#
def join_month_parameters(dummy_counter):

    if dummy_counter == 0:

        #_m --> mounth
        global idx_patients_m, idx_enterprise_m, idx_enterprise_patients_noCovits_m
        global codeIntLab_m, csvFile_m, ECBP_str_m, idx_express_m, idx_vuelo_m
        global examNameList_m, examNameList_nested_m, ECBP_m, enterpriseNames_asDict_m
        global ECBNC_m, ECBAC_m, ECBABC_m, ECBCABC_m, ECBSP_m
        global codeIntCob_m, listEnterpriseNameByPatient_m
        global day_list_m, day_list_antigenCovit_m, day_list_antibodyCovit_m
        global idx_patients_antigenCovit_m, idx_patients_antibodyCovit_m, idx_patients_noCovits_m
        global idx_patients_enterprise_forExclusiveExcel_asDict_m, idx_enterprise_enterprise_forExclusiveExcel_asDict_m, day_list_enterprises_excel_m

        idx_patients_m = pd_Index(data=[])
        idx_patients_antigenCovit_m = pd_Index(data=[])
        idx_patients_antibodyCovit_m = pd_Index(data=[])
        idx_patients_noCovits_m = pd_Index(data=[])
        idx_enterprise_m = []
        idx_enterprise_patients_noCovits_m = []
        codeIntLab_m = {}
        csvFile_m = pd_DataFrame()
        ECBP_str_m = {}
        idx_express_m = []
        idx_vuelo_m = []
        examNameList_m = {}
        examNameList_nested_m = {}
        ECBP_m = {}
        ECBNC_m = {}
        ECBAC_m = {}
        ECBABC_m = {}
        ECBCABC_m = {}
        ECBSP_m = {}
        enterpriseNames_asDict_m = {}
        codeIntCob_m = {}
        listEnterpriseNameByPatient_m = {}
        day_list_m = {}
        day_list_antigenCovit_m = {}
        day_list_antibodyCovit_m = {}
        idx_patients_enterprise_forExclusiveExcel_asDict_m, idx_enterprise_enterprise_forExclusiveExcel_asDict_m = get_idx_enterpriseExclusive(enterpriseCodecs_forExclusiveExcel, pd_Index(data=[]), {}, {})
        day_list_enterprises_excel_m = {x:{} for x in enterpriseCodecs_forExclusiveExcel}
    
    for key, value in idx_enterprise_enterprise_forExclusiveExcel_asDict.items():
        idx_enterprise_enterprise_forExclusiveExcel_asDict_m[key] += [tmp+dummy_counter for tmp in value]
    
    for key, value in idx_patients_enterprise_forExclusiveExcel_asDict.items():
        idx_patients_enterprise_forExclusiveExcel_asDict_m[key] = (idx_patients_enterprise_forExclusiveExcel_asDict_m[key]).union(value + dummy_counter) 
    
    idx_patients_m = idx_patients_m.union(idx_patients+dummy_counter)
    idx_patients_antigenCovit_m = idx_patients_antigenCovit_m.union(idx_patients_antigenCovit+dummy_counter)
    idx_patients_antibodyCovit_m = idx_patients_antibodyCovit_m.union(idx_patients_antibodyCovit+dummy_counter)
    idx_patients_noCovits_m = idx_patients_noCovits_m.union(idx_patients_noCovits+dummy_counter)
    idx_enterprise_m.extend([tmp+dummy_counter for tmp in idx_enterprise])
    idx_enterprise_patients_noCovits_m.extend([tmp+dummy_counter for tmp in idx_enterprise_patients_noCovits])
    idx_express_m.extend([tmp+dummy_counter for tmp in idx_express])
    idx_vuelo_m.extend([tmp+dummy_counter for tmp in idx_vuelo])

    tmp = {x:day_[0:2]+'/'+day_[2:4]+'/'+day_[4:6] for x in idx_patients}
    for key, value in tmp.items():
        day_list_m[key+dummy_counter] = value

    tmp = {x:day_[0:2]+'/'+day_[2:4]+'/'+day_[4:6] for x in idx_patients_antigenCovit}
    for key, value in tmp.items():
        day_list_antigenCovit_m[key+dummy_counter] = value
    
    tmp = {x:day_[0:2]+'/'+day_[2:4]+'/'+day_[4:6] for x in idx_patients_antibodyCovit}
    for key, value in tmp.items():
        day_list_antibodyCovit_m[key+dummy_counter] = value

    for key, value in idx_patients_enterprise_forExclusiveExcel_asDict.items():

        tmp = {x:day_[0:2]+'/'+day_[2:4]+'/'+day_[4:6] for x in value}
        for key1, value1 in tmp.items():
            day_list_enterprises_excel_m[key][key1+dummy_counter] = value1

    for key, value in codeIntLab.items():
        codeIntLab_m[key+dummy_counter] = value

    for key, value in ECBP_str.items():
        ECBP_str_m[key+dummy_counter] = value
    
    for key, value in examNameList.items():
        examNameList_m[key+dummy_counter] = value

    for key, value in examNameList_nested.items():
        examNameList_nested_m[key+dummy_counter] = value

    for key, value in ECBP.items():
        ECBP_m[key+dummy_counter] = value

    for key, value in ECBNC.items():
        ECBNC_m[key+dummy_counter] = value
    
    for key, value in ECBAC.items():
        ECBAC_m[key+dummy_counter] = value
    
    for key, value in ECBABC.items():
        ECBABC_m[key+dummy_counter] = value
    
    for key, value in ECBCABC.items():
        ECBCABC_m[key+dummy_counter] = value
    
    for key, value in ECBSP.items():
        ECBSP_m[key+dummy_counter] = value
    
    for key, value in enterpriseNames_asDict.items():
        enterpriseNames_asDict_m[key+dummy_counter] = value

    for key, value in codeIntCob.items():
        codeIntCob_m[key+dummy_counter] = value

    for key, value in listEnterpriseNameByPatient.items():
        listEnterpriseNameByPatient_m[key+dummy_counter] = value
    
    csvFile_m = pd_concat([ csvFile_m, csvFile.set_index(csvFile.index+dummy_counter) ])
#-----------------------------------------------------------------------------#

def antigen_excel():
    make_excel_antigen_antibody(idx_patients_antigenCovit, {'RESULTADO':np_NaN}, "Antígeno SARS CoV-2", "_antigenSARS_COV2", day_[0:2]+'/'+day_[2:4]+'/'+day_[4:6], day_, codeIntCob, yymmddPath, currentPath)

def antybody_excel():
    make_excel_antigen_antibody(idx_patients_antibodyCovit, {'IgG':np_NaN, 'IgM':np_NaN}, "IgG IgM SARS CoV-2", "_antibodySARS_COV2", day_[0:2]+'/'+day_[2:4]+'/'+day_[4:6], day_, codeIntCob, yymmddPath, currentPath)

def laboratory_excel():
    make_laboratory_excel(idx_patients, idx_enterprise, codeIntLab, csvFile, ECBP_str, currentPath, yymmddPath, day_, idx_express, idx_vuelo, examNameList, ECBNC, ECBAC, ECBABC, ECBCABC, ECBSP, enterpriseNames_asDict, "")

def laboratoryNoCovid_excel():
    make_no_covid_excel(idx_patients_noCovits, idx_enterprise_patients_noCovits, codeIntLab, csvFile, currentPath, os_path.join(yymmddPath,"byExamCategory"), day_, idx_express, idx_vuelo, examNameList_nested, ECBP, enterpriseNames_asDict, "")

def cobranza_excel():
    make_excel_cobranza(idx_patients, codeIntCob, day_[0:2]+'/'+day_[2:4]+'/'+day_[4:6], csvFile, examNameList, ECBP_str, listEnterpriseNameByPatient, day_, currentPath,  yymmddPath, idx_express)

def enterprises_excel():

    for codeEnterprise_ in idx_patients_enterprise_forExclusiveExcel_asDict:
        
        make_excel_enterprise_forExclusiveExcel(
            idx_patients_enterprise_forExclusiveExcel_asDict[codeEnterprise_], "_{}".format(codeEnterprise_),
            day_[0:2]+'/'+day_[2:4]+'/'+day_[4:6], day_, codeIntLab, csvFile, examNameList, currentPath, yymmddPath, idx_express
            )
    
def laboratory_excel_m():
    make_laboratory_excel(idx_patients_m, idx_enterprise_m, codeIntLab_m, csvFile_m, ECBP_str_m, currentPath, os_path.join(yymmddPath[:-6],"byMonth"), yymmddPath[7:-7], idx_express_m, idx_vuelo_m, examNameList_m, ECBNC_m, ECBAC_m, ECBABC_m, ECBCABC_m, ECBSP_m, enterpriseNames_asDict_m, "")

def cobranza_excel_m():
    make_excel_cobranza(idx_patients_m, codeIntCob_m, day_list_m, csvFile_m, examNameList_m, ECBP_str_m, listEnterpriseNameByPatient_m, yymmddPath[7:-7], currentPath, os_path.join(yymmddPath[:-6],"byMonth"), idx_express_m)
        
def antigen_excel_m():
    make_excel_antigen_antibody(idx_patients_antigenCovit_m, {'RESULTADO':np_NaN}, "Antígeno SARS CoV-2", "_antigenSARS_COV2", day_list_antigenCovit_m, yymmddPath[7:-7], codeIntCob_m, os_path.join(yymmddPath[:-6],"byMonth"), currentPath)

def antybody_excel_m():
    make_excel_antigen_antibody(idx_patients_antibodyCovit_m, {'IgG':np_NaN, 'IgM':np_NaN}, "IgG IgM SARS CoV-2", "_antibodySARS_COV2", day_list_antibodyCovit_m, yymmddPath[7:-7], codeIntCob_m, os_path.join(yymmddPath[:-6],"byMonth"), currentPath)

def laboratoryNoCovid_excel_m():
    make_no_covid_excel(idx_patients_noCovits_m, idx_enterprise_patients_noCovits_m, codeIntLab_m, csvFile_m, currentPath, os_path.join(yymmddPath[:-6],"byMonth","byExamCategory"), yymmddPath[7:-7], idx_express_m, idx_vuelo_m, examNameList_nested_m, ECBP_m, enterpriseNames_asDict_m, "_NoCovid")

def enterprises_excel_m():

    for codeEnterprise_ in idx_patients_enterprise_forExclusiveExcel_asDict_m:
        
        make_excel_enterprise_forExclusiveExcel(
            idx_patients_enterprise_forExclusiveExcel_asDict_m[codeEnterprise_], "_{}".format(codeEnterprise_),
            day_list_enterprises_excel_m[codeEnterprise_],
            yymmddPath[7:-7],
            codeIntLab_m,
            csvFile_m,
            examNameList_m,
            currentPath,
            os_path.join(yymmddPath[:-6],"byMonth"),
            idx_express_m
        )
