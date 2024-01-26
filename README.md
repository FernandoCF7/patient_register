# patient_register

> ## **This project contains a python sofware to load and generate excel list (reports) of the patient laboratory register proccess**
> - #### Patient laboratory register
> - #### Python

The software is designed to manage and efficiently the process of patient register and laboratory tasks. The image below shows how the software works.

This image show how to register the patient data in a txt file. These data are: name, laboratory test, the client (enterprise or claimant), shift, and if the test is express.

![](/aux_src/at_txt.png)

Onces the data are register just run the /code/make_excel_file.py to generate the laboratory process excel file with the registered data (image below)

![](/aux_src/at_excel.png)

## **Setting your sistem**

### The enterprises

To setting your sistem, you would have to put a list of enterprises (or laboratories) as your clients inside of the /code/src_list/enterprise/codeEnterprise.csv and /code/src_list/enterprise/claveEnterpriseName.csv To explaint these files it will be registered two enterprises: 'Wonka chocolate' and 'Acme enterprise'

1. The codeEnterprise.csv file

![](/aux_src/codeEnterprise.png)

This excel file (sheet) has two columns: 'enterprise' and 'clave'. The enterprise have to be register inside of '/' chacter /WONKA CHOCOLATE/ as example. The choice of the 'clave' is custom (but unique for each enterprise), here it is suggested to use a set of [AAA, AAB,  . . ., ABA, ABB, . . . ZZZ] since it generates a 17576 registers ($2^3$). Note that the 'WONKA' enterprise appears three times (rows 2, 3 adn 4), also notice that them have the same clave (AAA), it means that you can register the 'WONKA' enterprise in the 'patient register text file' (figure 1) as:

* /WONKA CHOCOLATE/
* /WILLY WONKA CHOCOLATE/

    or

* /WK CHOCOLATE/

This is helpulf when the capture person uses simmilar terminology for the same enterprise. It is not neccesary register the same enterprise more than one time, as is the case of the 'ACME ENTERPRISE'. 

2. The claveEnterpriseName.csv file

![](/aux_src/claveEnterpriseName.png)

This excel file (sheet) has the same two columns of the last one ('enterprise' and 'clave'). Following the 'Wonka' enterprise registration, it will have to be register in this file too using the same clave (AAA). In this file, the Wonka enterprise has to be register just once time: 'WILLY WONKA CHOCOLATE' (now without the '/' charter). This file is dessigned to set the name of the enterprise which will be deployed in the system generated excel report (figure 2). The enterprise name ('ENTERPRISE' column) should not necessarily be the same as those choosen in the last file ('WONKA CHOCOLATE', 'WILLY WONKA CHOCOLATE' or 'WK CHOCOLATE'); remember, just the 'clave' have to be the same; anyway, in this example it has been set as 'WILLY WONKA CHOCOLATE'. In both files the system manage the enterprise name as case insensitive; the excel report always will show the enterprise name as upercase, and regardless of the upper, or lower, case in the register text file the system will be recognize the enterprise registered in the codeEnterprise.csv file.


### The exam list

![](/aux_src/examList.png)

This excel file has three columns: 'NUMERIC CODE', 'SHORT EXAMN NAME' and 'EXAMN NAME'. You could register any exam, the unique condition will be the unique 'NUMERIC CODE' value by register (not necessary sequentially). The 'SHORT EXAMN NAME' is used by the system to show it in the excel file report. The 'EXAMN NAME' is ignored by the system and it is used just as a more complex description of the exam.


### The load patient file (.txt)

The load patient file will be made by day. Let's take as exampe the load patient stack ot the day January 7, 2023. Open your favorite text editor program and made a text file /code/input_src/____23/070123.txt Supose we have to load 5 patient in total with the next features:

1. The first three patients allows to the 'WONKA CHOCOLATE' enterprise 
2. The fourth patient allows to the 'ACME ENTERPRISE' enterprise
3. The last two patients allows to the 'WONKA CHOCOLATE' enterprise
4. The names and lastnames of the patients will be 'patient1 name' 'patient1 lastname', 'patient2 name' 'patient2 lastname', etc.
5. THe firs patient must have associated 'VIH load' exame. The other patients could have any examns, eaven some of them with more than one examn.
6. The second and last patient must be process as 'express'
7. The thirth patient must be process as 'flight' (denotates special laboratory treatment and send results according to the SARS CoV-2 'flight' standars)
8. The 'ACME ENTERPRISE' could be register as 'acme enterprise' (lower case); this will show the case insensitive of the system
9. The last two patients could be registered as 'WK CHOCOLATE' (not as 'WONKA CHOCOLATE'), this will demostrate the multiregister ways of one enterprise

The file will be made as:
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

se puede cargar como 3 o 003
los espacios los ignora el sistema (por claridad yo los pongo como patient1 name* patient1 lastname* 003)
the * separates columns
4) saca listados por dia, por mes y por empresa 
