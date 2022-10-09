"""
Project 2
October 19 2021
by Zac Fawcett

v1 October 19 2021
-testing functions 
-Pseudo code
-Menu and splash screen

v2 October 24 2021
-finalize functions to be used from Pywin32
-create all menu options
-create user made functions to implement win32 functions

v3 October 31 2021
-error test all possibilites
-add log file write out feature
-test on different virtual machine
"""
"""

Preamble
This code only has full functionality if run from a command line with administrator elevation.
The command line needs to be reading from the same directory that the file is housed in.
So if the file is at C:\\ then the command line needs to change directories to C:\\
If ran from the python editor certain features will not work:
    1. The output file will not be created as its path requires admin priveleges
    2. Groups cannot be created
    3. Users cannot be added to groups
    4. The log file will not be written to
If not in elevated cmd line all other features will still work


Pseudo Code

OUTPUTNAME = "Path of desired log file output.txt"
open a file with OUTPUTNAME as path, "w"

define user functions

For getfunctions
def getfunction
    list = win32module.function(arguments)
    for items in list
        print(items)
    (alternatively ask for input to define arguments that customize the list)
    (alternatively create variables from list indecies and concatenate a string with desired output
    

For add or create functions
def add/create function
    list = win32module.function(arguments)
    variable = input(user or group)
    error check input against list to check for redundancy
    if non redundant:
        execute function
    
for log function
    outputfile.write(time/date + output of a create or add command)

print splash screen on first load with title



menu
Menu loops until a proper command is entered or exit character is entered
while true:
    list options
    choice = input
    if input == option
        execute function within option
    if input != any options
        print(error message)
    if input = "Exit character"
        break

signal that program is ending
signal output file has been written/path of file
close output file
sys exit
"""

# import needed modules from win32 
import win32api
import win32net
import win32print
import win32con
import win32file
import win32profile
import sys

# Constant string used for easy changing of file output name and also to signal to
# user what the file name is at end of code
OUTPUTNAME = ("C:\\Users\\Junior Task Log.txt")

# Function that will create a sound, used for when an incorrect input is entered
def Beep():
    win32api.MessageBeep(win32con.MB_ICONWARNING)

# Create a file that will log certain actions including date and time
try:
    taskLogFile = open(OUTPUTNAME, "w")
    
except:
    Beep()
    print("Failed to create log file at that location")

# Function for attractively listing dictionaries with custom separators
def listDict(dict, separator):
     for key, value in dict.items():
         print (key, separator, str(value))

# Function for logging time and output of user actions
def LogApp(x:str):
    taskLogFile.write(DateTime() + " - " + x)

# Function for printing date and time for session log
def DateTime():
     dateTuple = win32api.GetLocalTime()
     year = str(dateTuple[0])
     month = str(dateTuple[1])
     day = str(dateTuple[3])
     hour = str(dateTuple[4])
     minute = str(dateTuple[5])
     second = str(dateTuple[6])
     if (hour == "1" or hour == "2" or hour == "3" or hour == "4" or hour == "5" or hour == "6"
         or hour == "7" or hour == "8" or hour == "9"):
         hour = "0" + hour
     if (second == "1" or second == "2" or second == "3" or second == "4" or second == "5" or second == "6"
         or second == "7" or second == "8" or second == "9"):
         second = "0" + second
     if (minute == "1" or minute == "2" or minute == "3" or minute == "4" or minute == "5" or minute == "6"
         or minute == "7" or minute == "8" or minute == "9"):
         minute = "0" + minute
     out = (str(month + "-" + day + "-" + year + " at " + hour + ":" + minute + ":" + second))
     return out
    
# Function that allows for local users to be added to local groups by name
# After adding to group, this function lists the groups user belongs to
def LocalGroupAdd():
    while True:
        # Create a dictionary of the users list and cast to string for parsing
        enumUsers = win32net.NetUserEnum(None, 0, 0, 0, 512)
        usersList = (str(enumUsers))
        name = input("Enter desired user: ")
        # error checking for if the user entered doesn't exist
        if name not in usersList:
             print("No such user exists")
             print("\n")
             break
        # Find domain name
        localMach = win32api.GetDomainName()
        userGroupsList = win32net.NetUserGetLocalGroups(None, name, 0)
        # create list of groups on computer
        enumGroups = win32net.NetLocalGroupEnum(None, 0, 0)
        # cast groups list to a string
        groupsList = (str(enumGroups))
        groupAdd = input("Enter group to add user to: ")
        # error check for if that group does not exist
        if groupAdd not in groupsList:
            print("No such group exists" + "\n")
            print("The user " + name + " belongs to these groups: ")
            for groups in userGroupsList:
                print (groups)
            print("\n")
            break
        # create dictionary of domain name + username for add members function argument
        addDict = [ {"domainandname" : localMach+"\\" + name} ]
        # error check for if the user already belongs to that group
        if groupAdd in userGroupsList:
            print("User is already in that group" + "\n")
            print("The user " + name + " belongs to these groups: ")
            for groups in userGroupsList:
                print (groups)
                print("\n")
                break
        # add user to group function
        try:
            win32net.NetLocalGroupAddMembers(localMach, groupAdd, 3, addDict)
            print("User " + name + " added to group: " + groupAdd)
            # Write output to log file
            LogApp("User " + name + " added to group: " + groupAdd)
            print("\n")
        except:
            Beep()
            print("Failed to add user " + name + " to group: " + groupAdd)
            break
        userGroupsList = win32net.NetUserGetLocalGroups(None, name, 0)
        groupsList = (str(enumGroups))
        print("The user " + name + " belongs to these groups: ")
        for groups in userGroupsList:
                print (groups)
        print("\n")
        break
        
# Function to create a group with a description
def CreateGroup():
    while True:
        enumGroups = win32net.NetLocalGroupEnum(None, 0, 0)
        groupsList = (str(enumGroups))
        groupsTuple = (tuple(enumGroups))
        group = input("Enter group name: ")
        for groups in groupsList:
            if group in groupsList:
                Beep()
                print("That name is too similar to a current group name, choose a different name" + "\n")
                break
        comment = input("Enter a group description: ")
        group_data = {'name': group, 'comment': comment}
        try:
            win32net.NetLocalGroupAdd(None, 1, group_data)
            print("Group named: " + group + " was successfully created" + "\n")
            # Write output to log file
            LogApp(group + " successfully created" + "\n")
        except:
            Beep()
            print("Failed to create group")
            print("\n")
            break
        finally:
            break

# Function to list groups a particular user belongs to
def UserGroups():
    while True:
        try:
            enumUsers = win32net.NetUserEnum(None, 0, 0, 0, 512)
            usersList = (str(enumUsers))
            name = input("Enter name of user: ")
            if name not in usersList:
                 print("No such user exists" + "\n")
                 break
            userGroupsList = win32net.NetUserGetLocalGroups(None, name, 0)
            print("The user " + name + " belongs to these groups: ")
            for groups in userGroupsList:
                print (groups)
                break
            print("\n")
            break
        except:
            Beep()
            print("Failed to find the groups list" + "\n")
            break

# Function that lists machines version of windows with build number
def PrintWinVer():
     vers = win32api.GetVersionEx(1)
     major = vers[0]
     minor = vers[1]
     build = vers[2]
     print("This system is running Windows " + str(major) + "." + str(minor) + " Build number " + str(build))

# Function to list users
def ListUsers():
    userList = win32net.NetUserEnum(None, 0, 0, 0, 512)
    print("List of local users: ")
    for j in userList[0]:
         print(j)
    print("\n")
     
# Function to list local groups
def ListGroups():
    try:
        enumGroups = win32net.NetLocalGroupEnum(None, 0, 0)
        groupsList = (str(enumGroups))
        groupsTuple = (tuple(enumGroups))
        print("Current List of local groups: ")
        for groups in groupsTuple[0]:
             print(groups)    
    except:
        Beep()
        print("Could not parse group data")

# Function for returning info about system environment variables
def EnvirOp():
     default = win32profile.GetDefaultUserProfileDirectory()
     print ("Default User Profile Directory is: + " + default)
     profiles = win32profile.GetProfilesDirectory()
     print ("Default Profiles Directory is: + " + profiles)
     print("The system enrionment variables are: ")
     sysStrs = win32profile.GetEnvironmentStrings()
     for key in sysStrs.keys():
          print(key)
          
     print("\n")  
     while True:
          while True:
               print("Enter item from list to expand upon ")
               print("Item must be entered exactly as shown, or press enter to exit")
               choice = input("Entry: " + "\n")
               if choice == "":
                    break
               if choice not in sysStrs:
                   print("That string was not listed" + "\n")
               if choice in sysStrs:
                    break
          choiceQ = ("%"+choice+"%")
               
          sysVar = win32profile.ExpandEnvironmentStringsForUser(None, choiceQ)
          if choice == "":
               break
          print ("Setting: " + sysVar)
          print("\n")

# Function that lists info about specified user
def UserAttr():
    try:
        name = input("Enter a user name: ")
        print("\n")
        userInfo = win32net.NetUserGetInfo(None, name, 2)
        print("Attributes for user " + name + ":")
        listDict(userInfo, ':')
        print("\n")
        currentUser = win32api.GetUserName()
        print("Currently logged in as: " + currentUser)
        currentDomain = win32api.GetDomainName()
        print("Current domain is: " + currentDomain)
        print("\n")
    except:
        Beep()
        print("Can not retrieve user info" + "\n")

# Function that prints a file's encrpytion status
def CryptStat():
    try:                    
        filePath = input("Enter the absolute path of a file without quotations: " + "\n")
        cryptInfo = (win32file.FileEncryptionStatus(filePath))
        if cryptInfo == 0:
            print("This file is not encrypted" + "\n")
        else:
            print("This file is encrypted" + "\n")
    except:
        Beep()
        print("This is not a valid absolute file path")
        print("Please enter a proper absolute file path")
        print("\n")

# Function that returns disk space info about a specified Volume as well as current system memory usage
def FileInfo():
    try:
        rootPath = input("Enter drive volume path (in the format: C:\\) : ")
        print("\n")
        freeSpace = (win32file.GetDiskFreeSpace(rootPath))
        print("Current Disk Free Space for drive " + rootPath)
        print("Sectors per cluster: " + str(freeSpace[0]))
        print("Bytes per sector: " + str(freeSpace[1]))
        print("Number of free clusters: " + str(freeSpace[2]))
        print("Total number of clusters: " + str(freeSpace[3]) + "\n")
        tempPath = (win32api.GetTempPath())
        dnsName = win32api.GetComputerNameEx(1)
        PrintWinVer()
        print("Current machine name: " + dnsName)
        print("Current Drive temp path: " + tempPath)
        print("Windows Directory path: " + win32api.GetWindowsDirectory())
        print("\n")
        memoryDict = win32api.GlobalMemoryStatusEx()
        print("Current memory usage: ")
        listDict(memoryDict, '=')
        print("\n")
            
    except:
        Beep()
        print("Cannot find that drive. Please enter volume path (in the format: C:\\)" + "\n")

# Function that displays system printer information
def PrintInfo():
    try:
        print("\n")
        dePrint = win32print.GetDefaultPrinter()
        printerHandle = win32print.OpenPrinter(dePrint, None)
        printerInfo = win32print.GetPrinter(printerHandle, 2)
        print("Printer Info: " + "\n")
        dePrint = win32print.GetDefaultPrinter()
        print("Default Printer: " + dePrint)
        printerDirectory = win32print.GetPrinterDriverDirectory(None, None)
        print("Printer Driver directory: " + printerDirectory)
        print("\n")
        print("Current Printer attributes: ")
        listDict(printerInfo, '=')
        print("\n")
    except:
        Beep()
        print("Not able to retrieve printer attributes")
        print("\n")
            
# Function for Splash Screen for first app opening
def Splash():
    print("**************************************************" + "\n")
    print("Welcome to the Junior Admin Task Menu" + "\n")
    print("**************************************************" + "\n")
    print("Please make a selection or press 0 to exit" + "\n")
    print("**************************************************" + "\n")
    print("\n" + "\n")

# Function for Main Menu
def Menu():
    print("Menu Options" + "\n")
    print("**************************************************" + "\n")
    print("1. Display Printer settings")
    print("2. Display Disk and Machine attributes")
    print("3. Display User attributes and Domain name")
    print("4. Display Local Users")
    print("5. Create a group")
    print("6. Add a user to a group")
    print("7. Display groups a user belongs to")
    print("8. List all local groups")
    print("9. Find Encryption status of a file")
    print("10. List System Envrironment String Paths" + "\n")
    
# Splash screen on script open
Splash()

while True:
    
    Menu()
    
    choiceEntered = input("Enter choice or 0 to quit: ")
    print("\n")
    # Error checking for if an incorrect input is entered. Will display message and keep looping menu
    if (choiceEntered != "0" and choiceEntered != "1" and choiceEntered !="2" and choiceEntered != "3" and
        choiceEntered != "4" and choiceEntered != "5" and choiceEntered != "6" and choiceEntered != "7" and
        choiceEntered != "8" and choiceEntered != "9" and choiceEntered != "10"):
        Beep()
        print("ERROR**ERROR**ERROR")
        print("Enter a digit between 1 and 9 or enter 0 to exit program" + "\n")
      
    if choiceEntered == "1":
        PrintInfo()
        
    if choiceEntered == "2":
        FileInfo()

    if choiceEntered == "3":
        UserAttr()
            
    if choiceEntered == "4":
        ListUsers()
        
    if choiceEntered == "5":
        CreateGroup()
  
    if choiceEntered == "6":
        LocalGroupAdd()
                 
    if choiceEntered == "7":    
        UserGroups()
        print("\n")

    if choiceEntered == "8":
        ListGroups()
        print("\n")
        
    if choiceEntered == "9":    
        CryptStat()

    if choiceEntered == "10":
        EnvirOp()
# Exit program option. This option will terminate the program with a message saying so
# also gives the log file output path
    if  choiceEntered == "0":
        print("Exiting Program ---- Goodbye!")
        print("Log file located at C:\\Users\\Junior Task Log.txt")
        break
# closes the log file so it can write
taskLogFile.close()
# Terminates program
sys.exit()     
