# pywintasks
zac fawcett

Windows Python task menu for junior admins

the basic concept of the script will be as follows; 

a command line menu that flashes a numbered options list. Entering one of the numbers on the menu will trigger the beginning of a task or will populate certain info on the screen. 

There will be error checking on the numbered options, wherein if the user enters an option not listed it will tell them that and reset the menu. 

Each of the options will have option specific error checking as well, always going back to the main menu if a problem arises so the user can try again if desired. 

There will also be a program terminating trigger from the main menu so the user can exit when they see fit. Certain actions on the menu will also trigger a logging feature.

Upon termination of the program a log file will be written to that contains the actions and corresponding time to that action so management can always know the environment changing actions the junior administrators have taken.

This code only has full functionality if run from a command line with administrator elevation.
The command line needs to be reading from the same directory that the file is housed in.
So if the file is at C:\\ then the command line needs to change directories to C:\\

If ran from the python editor certain features will not work:

   a. The output file will not be created as its path requires admin priveleges
   
   b. Groups cannot be created
   
   c. Users cannot be added to groups
   
   d. The log file will not be written to
   
   

If not in elevated cmd line all other features will still work




   Menu Options

**************************************************

1. Display Printer settings
2. Display Disk and Machine attributes
3. Display User attributes and Domain name
4. Display Local Users
5. Create a group
6. Add a user to a group
7. Display groups a user belongs to
8. List all local groups
9. Find Encryption status of a file
10. List System Envrironment String Paths

