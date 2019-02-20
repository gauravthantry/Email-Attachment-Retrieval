# Email-Attachment-Retrieval

                                                             EAR 1.0 (Email Attachment retrieval)                                                                                                
1.This C# console application traverses the outlook account to fetch the clearstream reports that are recieved from Tripartyrepo.cs@clearstream.com on every working day (Monday - Friday)
2.The application stores the clearstream report in the path mentioned in the app.config file. Any changes to the path or the file name in the future must be done in the app.config file.
3.The application checks for the file(path mentioned in the app.config file) for two consecutive days.
4.The errorLog for the file is populated for every scheduled run if the file is not yet recieved. Once the file is recieved, the successlog is populated once and the errorLog is cleared off completely for the particular date.
5.The application checks the mail account for the report for two days. For example, if the report for the current day is not yet recieved, it continues the check for it on the next date, along with the report of the current date. 
6.General Exceptions / Run Time errors are also logged using the generalExceptions() method defined in the program below.
7.If the files are not yet recieved, they are logged using the errorLogging() method as defined below.
8.The success log are logged using the successLogging() method as defined below.
