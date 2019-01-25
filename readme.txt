Workflow script is intended to read from a "WorkFlow.xlsx" file on the desktop.
When downloaded, you will initally need to edit the code to add your own email, password, workflow email and smtp server configuration.

Python script that uses endless loops to check dates and read an excel file that is used to put completed tasks in. The script reads that file and prints it into an email body and sends it to a workflow email(removed for obvious reasons) and end of day, 530 monday through thursday and 12 on friday. Checks if it is the weekend and puts in an endless loop until monday and restarts the script. Parts of code were used as reference and some was implemented throughout the code.
