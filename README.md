# py_email_application

This application allows a user to import a list of email contacts from a CSV file and will automatically create a blank email in outlook for the user selected contact that fills in the To, CC and subject line. Email signatures are retained in emails which are created. It uses tkinter for the front end interface and interacts with outlook via the pywin32 module. 
While running the script from my IDE I did recevie an error which seemed to be due to an issue with the gen_py data located in the users temp AppData folder. Navigating to C:\Users\<your username>\AppData\Local\Temp\gen_py and removing folders/files that have a naming convention like "00020813-0000-0000-C000-000000000046x0x1x9 resolved that error and allowed the application to execute successfully. 
However, I have not received any errors when running the pacakaged application I created via pysintaller nor has the end user which I built the application for.



