## Invoice Extract Script
These scripts were created because it was noticed that the invoices we thought we sent (stored in our ERP system) was not reflected in the actual invoices sent to our producers. This script had a pre-requiste of a Microsoft Flow that looked into a specific account, checked if it got an email from the DL that had the invoices sent to, and pulled the invoice number and put it into a csv.

The powershell script first downloads the data gathered from the Flow, then pulls all the invoices from a database (by invoice number, time of the invoice, and who it was sent to), and export it into a csv file. It then invokes the python script pulls
two csv files, one from the aforementionned script and the other from a Microsoft Flow program that collects all the actual invoices sent and gets the differences between them.
It then compiles the differences and attaches it to an email to send automatically to multiple people.