How to use the program?

1. At first you need "serviceAccountKey.json" placed in the program run directory. It's already given but you can always generate new one from your firebase console. To do this go to the firebase console, then "Project settings", then tab "Service accounts", then press button "Generate new private key". Then you will get "*.json" file and then you have to rename it to "serviceAccountKey.json" and move it to program run directory.

2. Prepare source data file "*.xlsx". Use proposed "upload sheet.xlsx" file as template. Do not change header names and order of sheets cause work of the  program will be broken.

3. Run the program. By the default program will use as source "upload sheet.xlsx" that put in the run directory. You can change in the command line the path or name of the source file.
Examples: firestoreUpload.exe "project1.xlsx", firestoreUpload.exe "C:\MyFolder\project3.xlsx". 

4. If an error occurs while the program is running, you will see a message on the screen, as well as in the "log_errors.txt" file for further examination.
---


What the program do?
1. If there are rows in the Users sheet the program will check accounts with identifiers and will add new accounts (if they are not exist) as well as add related records in the "Users" collection.

2.The program will check rows in the next 3 sheets and add records to the firestore collection "Project" and its subcollections (measurements, designations, measurement-refs, contacts)
2.1. The measurements, designations, measurement-refs data fills from Manipulate sheet.