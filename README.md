# Data Retrieval using XPath, Visual Basic and XML
# Description: 
This folder contains data retrieval process implemented with XPath Language version 3.1 (XML). The whole process can be pictured or generally divided into the following subprocesses:

1. Use OAuth 2.0 (which is frequently used by big companies such as Facebook, Google and Microsoft) to securely access database with a given API key.
2. Do subsequent API calls (defined with XPath and XML) to retrieve data from the database and parse data into spreadsheet.
3. Use the data to generate reports (business, contacts, membership renewal, membership lapsed etc.) 

To be specific, the application uses OAuth authentication to gain access of the admin view of the database supported by Wild Apricot (https://www.wildapricot.com/). It retrieves contact information, auditLog items information and event registrations information from the database and utilizes the data to plot subsequent pivot charts. To guarantee the efficiency of the User Interface (UI), the program is also supported with a progress bar to indicate current progress.   

# Purpose:
The main purpose for this project is to provide a means to securely access the database and retrieve data for business analysis purposes.

# Application:
Business Analysis (such as creating charts or pivot charts to store business information so as to reckon and analyse market trend)

# Reference:
Wild Apricot Help (syntax and API calls references):
https://gethelp.wildapricot.com/en

XPath References (Syntax and examples):
https://msdn.microsoft.com/en-us/library/ms256086(v=vs.110).aspx

// Last Edited by Edmond Tsoi, 5-27-2018.
