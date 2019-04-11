# Auto_Service_Check
This VBScript will check all the Automatic Services on a windows machine and restart them if they are not up and running. The script has the ability to check the local machine or a remote machine. This is task that we normally run on Task Scheduler using the AutoServe.bat this is included in the project. Since the script runs on it’s own interval we normally set the Task Scheduler Task to run on system start. If you look in the Batch Script we have the vbscript pipe the output to a log file, with in the batch script there is a function to check the log size, if it is over 700KB then delete before running the script, the way we have task scheduler setup, a new log will be created on the next restart, only if it is over 700KB.

Additional Documentation for this application can be found at:

https://github.com/burnsoftnet/Documentation/blob/master/Auto_Service_Check%20-%20VBScript.pdf


[![Donate](https://www.paypalobjects.com/en_US/i/btn/btn_donateCC_LG.gif)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=JSW8XEMQVH4BE)]
