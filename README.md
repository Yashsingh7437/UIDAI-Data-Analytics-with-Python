In this repository there are many files but main file is final_file.py
In final_file.py the code logs into your email and then search for the specific sender's mail, if mail from that sender is found then it proceeds further else it will print "No emails found with the specified subject and sender."
If the mail is found from that specific sender then it will check for specific subject, if subject is found it will proceed further else it will print "No emails found with the specified subject and sender."
If the subject is found it will look for the attachment and if that attachment is a excel sheet it will proceed further else it will return with mesaage "Attachment is there but it's not an Excel fileAttachment is there but it's not an Excel file"
After all this it will extract the row with specific data (in my code it is RO Lucknow) and copy that to the new excel sheet (if sheet exist else will create a new with that name) along with the date of email received.
All the above processes will be repeated for all the mail from specific sender.
This code will not make the copy of any specific data twice. It will remove the identical rows.
Once the new sheet is create by fetching the data from the mail, it will make a heat map of some columns(in my case those are column E+F, column G, Column H)
On y-axis it will plot the columns and on x-axis it will plot the rows.

This is a Data Automation Code with Python
