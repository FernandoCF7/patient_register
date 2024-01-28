# patient_register

> ## **This project contains a python software to load and generate excel list (reports) of the patient laboratory register process**
> - #### Patient laboratory register
> - #### Python

The software is designed to efficiently manage and made the process of patient register and laboratory tasks. The images below shows how the software works.

This image shows how to register the patient data in a txt file. The data are: name, required laboratory examns, client (enterprise or claimant), shift, and other features of the request as express and/or flight process, etc.

![](/aux_src/at_txt.png)

Once data are registered just run the /code/make_excel_file.py python program to generate the 'laboratory process excel file' (image below).

![](/aux_src/at_excel.png)

## **Setting your system**

### The enterprises

Suppse that periodically you process patient examns of the same place (enterprise, or other laboratories); example, today you received in a request sheet $10$ patients of the 'enterprise A', $6$ patient of the 'enterprise B' (in other request sheet), and other $4$ patients of the 'enterprise A' (in other request sheet). This workflow is very common in the laboratory process, then it is helpful to have these enterprises stored in some data base. Note: you could manage the common person requests as 'person' enterprise (or any other name that makes sense). To load your enterprises, you would have to put a list of them inside the /code/src_list/enterprise/codeEnterprise.csv and /code/src_list/enterprise/claveEnterpriseName.csv. To explain these files it will be registered two enterprises: 'Wonka chocolate' and 'Acme enterprise'

1. The codeEnterprise.csv file

![](/aux_src/codeEnterprise.png)

This excel file (sheet) has two columns: 'enterprise' and 'clave' (never change the header ('column names') of this file; or any other header of the other load DB files). The enterprise have to be registered inside of '/' character /WONKA CHOCOLATE/ as example. The choice of the 'clave' is custom (but unique for each enterprise), here it is suggested to use a set of [AAA, AAB,  . . ., ABA, ABB, . . . ZZZ] since it generates a $17576$ registers ($2^3$). Note that the 'WONKA' enterprise appears three times (excel image above, rows 2, 3 adn 4), also notice that them have the same clave (AAA), it means that you can register the 'WONKA' enterprise in the 'patient register text file' as:

* /WONKA CHOCOLATE/
* /WILLY WONKA CHOCOLATE/

    or

* /WK CHOCOLATE/

This is helpful when the capture person uses similar terminology for the same enterprise. Note: It is not necesary register the same enterprise more than one time, as is the case of the 'ACME ENTERPRISE'. 

2. The claveEnterpriseName.csv file

![](/aux_src/claveEnterpriseName.png)

This excel file has the same header of the last one ('enterprise' and 'clave'). Following the 'Wonka' enterprise registration, it will have to be register in this file too ( using the same clave AAA). In this file, the Wonka enterprise has to be register just once time: 'WILLY WONKA CHOCOLATE' (now without the '/' charter). This file is dessigned to set the name of the enterprise which will be deployed in the system generated excel report (figure 2), then, it is recommended no so large names. The enterprise name ('ENTERPRISE' column) should not necessarily be the same as those choosen in the last file ('WONKA CHOCOLATE', 'WILLY WONKA CHOCOLATE' or 'WK CHOCOLATE'); remember, just the 'clave' have to be the same; anyway, in this example it has been set as 'WILLY WONKA CHOCOLATE'. In both files the system manage the enterprise name as case insensitive; the excel report always will show the enterprise name as upercase, and regardless of the upper, or lower, case in the register text file the system will be recognize the enterprise registered in the codeEnterprise.csv file.


### The exam list

![](/aux_src/examList.png)

This excel file has three columns: 'NUMERIC_CODE', 'SHORT_EXAMN_NAME' and 'EXAMN NAME'. You could register any exam, the condition will be the unique 'NUMERIC_CODE' value by register (an integeer but not necessary sequentially). The system uses the 'SHORT_EXAMN_NAME' to show it in the excel report file. The 'EXAMN NAME' is ignored by the system and it is used just as a more complete description of the exam.

### The load patient file (.txt)

The load patient file will be made by day. Let's take as an exampe the load patient stack ot the day January 7, 2023. Open your favorite text editor program and made a text file /code/input_src/____23/070123.txt. Supose we have to load a total of 5 patients with the next features:

1. The first three patients belong to the 'WONKA CHOCOLATE' enterprise
2. The fourth patient belongs to the 'ACME ENTERPRISE' enterprise
3. The last two patients belong to the 'WONKA CHOCOLATE' enterprise
4. The names and lastnames of the patients will be 'patient1 name' 'patient1 lastname', 'patient2 name' 'patient2 lastname', etc.
5. THe firs patient must have associated 'VIH load' exame. The other patients could have any exams, heaven some of them with more than one examn.
6. The second and last patient must be process as 'express'
7. The thirth patient must be process as 'flight' (denotes special laboratory treatment and send results according to the 'SARS CoV-2 flight standards)
8. The 'ACME ENTERPRISE' could be register as 'acme enterprise' (lower case); this will show the case insensitive of the system
9. The last two patients could be registered as 'WK CHOCOLATE' (not as 'WONKA CHOCOLATE'), this will demonstrate the multi register ways of one enterprise

According to the last list, the file should be created as:
```text
firstName*SecondName*thirdName
enterprise/WONKA CHOCOLATE/matutino*0015*
patient1 name* patient1 lastname* 003
express patient2 name* patient2 lastname* 003
patient3 name* patient3 lastname* 003
enterprise/acme enterprise/matutino*0015*
patient4 name* patient4 lastname* 003
enterprise/WK CHOCOLATE/matutino*0015*
patient5 name* patient5 lastname* 003

```

**Notes:**

    - The first row always has to be 'firstName*SecondName*thirdName'
    - Uses the asterisk symbol '*' to separate the info at three columns
    - The second line begins with 'enterprise', it delimites a patient group requirement
    - The last number of the 'enterprise' lines is used to identify the patient group requirement sheet; it can be used any number
    - You can set the exam code as simple integer (example: 1) or as a padding zero number (example: 001)
    - The blank space into '*'s are ignored; i.e. is the same: 'name1*lastname1','name1 * lastname1', 'name1* lastname1', 'name1 *lastname1', etc.

[//]: # "4) saca listados por dia, por mes y por empresa"
