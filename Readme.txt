Yet another FTP Class...
Copyright Graham Daly December 2001 
Any questions or queries, 
feel free to drop me a mail at g.daly@iol.ie
===================================================

Public Properties
-----------------
-----------------

Directory:
+++++++++++
Reading the Directory property (Get) will return the current host directory (string) for the active FTP session. If there is no active connection, this property will return "<Unknown>". Setting a new Directory value (Let) will attempt to change the host directory for the active FTP session to the new specified directory.

Host:
+++++++++++
Reading the Host property (Get) will return the current host address (string) for the active FTP session. Setting a new Host address value (Let) will set the host address for an FTP session which is not yet active. An error will be returned if an attempt is made to change the host address while there is an active connection for the current clsFTP object. 

IsConnected:
+++++++++++
This is a read-only property (Boolean) to determine whether the current clsFTP object has an active connection. It does this by attempting to read the current FTP directory. If we can read the directory, then there is an active connection. The IsConnected property will return True. If we cannot read the directory, then there is no active connection and the property will return False.

Password:
+++++++++++
Reading the password property (Get) will return the current host password (string) for the active FTP session. Setting a new host password value (Let) will set the host password for an FTP session which is not yet active. An error will be returned if an attempt is made to change the host password while there is an active connection for the current clsFTP object. 

Port:
+++++++++++
Reading the port property (Get) will return the current host port (long integer) for the active FTP session. Setting a new host port value (Let) will set the host port for an FTP session which is not yet active. An error will be returned if an attempt is made to change the host password while there is an active connection for the current clsFTP object. 

User:
+++++++++++
Reading the User property (Get) will return the current host username (string) for the active FTP session. Setting a new host User value (Let) will set the host username for an FTP session which is not yet active. An error will be returned if an attempt is made to change the host username while there is an active connection for the current clsFTP object. 

Usage of these properties is shown in the examples below.


Public Methods
--------------
--------------

Connect:
+++++++++++
This function establishes a connection with the FTP host server. 4 optional parameters may be specified:

Host - if specified, this argument over-rides the clsFTP.Host property
Port - if specified, this argument over-rides the clsFTP.Port property
User - if specified, this argument over-rides the clsFTP.User property
Password - if specified, this argument over-rides the clsFTP.Password property

The function itself returns a value (long integer) of 1 if the connection attempt is successful. If it is unsuccessful an error will be raised, and the function will return the error number.

Example 1:

Dim Host$, Port&, User$, Pwd$
Dim FTPSession as New clsFTP

Host$ = "ifftp03"
Port& = 21
User$ = "anonymous"
Pwd$ = vbNullString

ReturnVal& = FTPSession.Connect(Host$, Port&, User$, Pwd$)

Example 2:

Dim FTPSession as New clsFTP

FTPSession.Host = "ifftp03"
FTPSession.Port = 21
FTPSession.User = "anonymous"
FTPSession.Password = vbNullString

ReturnVal& = FTPSession.Connect


DeleteFile:
+++++++++++
This function deletes a file in the current folder for an active clsFTP object. It takes one argument: FileToDelete. This is a string value which specifies the name of the file to delete. Also, since FTP allows you to reference files in a sub-directory contained within the current directory, you can delete a file in a sub-directory by assigning the FileToDelete parameter to contain the sub-directory path and the filename. 

The function itself returns a value (long integer) of 1 if the deletion attempt is successful. If it is unsuccessful an error will be raised, and the function will return the error number.

Example 1: Delete a file within the current directory

Dim Host$, Port&, User$, Pwd$
Dim FTPSession as New clsFTP
Host$ = "ifftp03"
Port& = 21
User$ = "anonymous"
Pwd$ = vbNullString
ReturnVal& = FTPSession.Connect(Host$, Port&, User$, Pwd$)
FTPSession.Directory = "/users/lang/efm/ird/enq/in"
ReturnVal& = FTPSession.DeleteFile("pnl00472.txt")


Example 2: Delete a file within a sub-directory

Dim Host$, Port&, User$, Pwd$
Dim FTPSession as New clsFTP
Host$ = "ifftp03"
Port& = 21
User$ = "anonymous"
Pwd$ = vbNullString
ReturnVal& = FTPSession.Connect(Host$, Port&, User$, Pwd$)
FTPSession.Directory = "/users/lang/"
ReturnVal& = FTPSession.DeleteFile("efm/ird/enq/in/pnl00472.txt")


Disconnect:
+++++++++++
This simple function terminates the connection for the active clsFTP object.


GetDirListing:
+++++++++++
This function returns the filenames and filesizes in the current directory for an active clsFTP object. It takes 3 arguments: a string array called FileNames(), a long integer array called FileSizes() and an optional string value called SubDir$. Since FTP allows you to reference files in a sub-directory contained within the current directory, you can get a list of files in a sub-directory by assigning a value to the SubDir.

FileNames() - this is a ByRef string array which returns the name of all files in 
the current FTP directory.
FileSizes() - this is a ByRef string array which returns the byte size of all files in  
the current FTP directory.
SubDir - if specified, this argument will specify a sub-directory from which to 
search for the files.

The function itself returns a value (long integer) of 1 if the listing attempt is successful. If it is unsuccessful an error will be raised, and the function will return the error number.


Example 1: Get file listing within the current directory

Dim Host$, Port&, User$, Pwd$
Dim FTPSession as New clsFTP
Dim GetFileNames$(), GetFilSizes&()

Host$ = "ifftp03"
Port& = 21
User$ = "anonymous"
Pwd$ = vbNullString

ReturnVal& = FTPSession.Connect(Host$, Port&, User$, Pwd$)
FTPSession.Directory = "/users/lang/efm/ird/enq/in"
ReturnVal& = FTPSession.GetDirListing(GetFileNames$(), GetFileSizes&())


Example 2: : Get file listing within a sub-directory

Dim Host$, Port&, User$, Pwd$
Dim FTPSession as New clsFTP
Dim GetFileNames$(), GetFilSizes&()

Host$ = "ifftp03"
Port& = 21
User$ = "anonymous"
Pwd$ = vbNullString

ReturnVal& = FTPSession.Connect(Host$, Port&, User$, Pwd$)
FTPSession.Directory = "/users/lang/efm"
ReturnVal& = FTPSession.GetDirListing(GetFileNames$(), _
GetFileSizes&(), "ird/enq/in")


GetFile:
+++++++++++
This function retrieves a file from the current directory in an active FTP session and copies it to the local machine. It takes 3 arguments: a string value called HostFile, a string value called ToLocalFile and a user-defined value called tt. Since FTP allows you to reference files in a sub-directory contained within the current directory, you can also retrieve a file in a sub-directory by specifying the sub-directory as part of the HostFile variable.

HostFile - this specifies the name of the file to get in the current FTP directory.
ToLocalFile - this specifies the file path and name to use when saving the 
file to the local machine.
tt - if specified, this argument specifies the FTP file transfer type to use. For text 
files, the fttAscii setting should be used. For binary files (e.g. graphics), the fttBinary setting should be used. If unsure of the file type, the fttUnknown setting can be used as a catch-all.

The function itself returns a value (long integer) of 1 if the retrieval attempt is successful. If it is unsuccessful an error will be raised, and the function will return the error number.


Example 1: Get a file from a sub-directory

Dim Host$, Port&, User$, Pwd$
Dim FTPSession as New clsFTP
Dim LocalFile$, HostFile$

Host$ = "ifftp03"
Port& = 21
User$ = "anonymous"
Pwd$ = vbNullString

HostFile$ =  "enq/in/pnl00001.txt"
LocalFile$ = "c:\temp\pnl00001.txt"

ReturnVal& = FTPSession.Connect(Host$, Port&, User$, Pwd$)
FTPSession.Directory = "/users/lang/efm/ird"
ReturnVal& = FTPSession.GetFile(HostFile$, LocalFile$, fttAscii)



PutFile:
+++++++++++
This function uploads a file from the local machine to the current directory in an active FTP session. It takes 3 arguments: a string value called LocalFile, a string value called ToHostFile and a user-defined value called tt. Since FTP allows you to reference files in a sub-directory contained within the current directory, you can also copy a file to a sub-directory by specifiying the sub-directory as part of the ToHostFile variable.

LocalFile - this specifies the name of the local file to be copied.
ToHostFile - this specifies the name to use when saving the file to the current 
directory of the active FTP session.
tt - if specified, this argument specifies the FTP file transfer type to use. For text 
files, the fttAscii setting should be used. For binary files (e.g. graphics), the fttBinary setting should be used. If unsure of the file type, the fttUnknown setting can be used as a catch-all.

The function itself returns a value (long integer) of 1 if the upload attempt is successful. If it is unsuccessful an error will be raised, and the function will return the error number.

Logic an application is same as GetFile, just in reverse - see example above.



RenameFile:
+++++++++++
This function renames a file from in the current directory in an active FTP session. It takes 3 arguments: a string value called LocalFile, a string value called ToHostFile and a user-defined value called tt. Since FTP allows you to reference files in a sub-directory contained within the current directory, you can also effectively move a file from one sub-directory to another by specifiying the sub-directory as part of the ToHostFile variable.

FileNameOld - this specifies the name of the host file to be renamed
FileNameNew - this specifies the new name of the file

The function itself returns a value (long integer) of 1 if the upload attempt is successful. If it is unsuccessful an error will be raised, and the function will return the error number.


Example 1: Renaming a file in current host directory

Dim Host$, Port&, User$, Pwd$
Dim FTPSession as New clsFTP
Dim OldFileName$, NewFileName$

Host$ = "ifftp03"
Port& = 21
User$ = "anonymous"
Pwd$ = vbNullString

OldFileName$ = "pnl00001.txt"
NewFileName$ = "pnl00001.arc"
ReturnVal& = FTPSession.Connect(Host$, Port&, User$, Pwd$)
FTPSession.Directory = "/users/lang/efm/ird/enq/out"
ReturnVal& = FTPSession.RenameFile(OldFileName$, NewFileName$)


Example 2: Using RenameFile to move a file from one sub-directory to 
another

Dim Host$, Port&, User$, Pwd$
Dim FTPSession as New clsFTP
Dim OldFileName$, NewFileName$

Host$ = "ifftp03"
Port& = 21
User$ = "anonymous"
Pwd$ = vbNullString

OldFileName$ = "/ird/enq/out/pnl00001.txt"
NewFileName$ = "/ird/enq/out/pnl00001.arc"
ReturnVal& = FTPSession.Connect(Host$, Port&, User$, Pwd$)
FTPSession.Directory = "/users/lang/efm"
ReturnVal& = FTPSession.RenameFile(OldFileName$, NewFileName$)
