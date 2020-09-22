After quite a long time searching for a decent replacement for the FileSystemObject - which did not require the programmer to include a bloated DLL dependency with the application, making use of a tool that, although useful is rather slow and more often than not, does not yield the desired results one intends it to achieve - I decided to create my own file handling class, for no examples online helped me perform the actions I needed.

So, after many attempts and having several times recoded the class from scratch, tweaking and optimizing it for a better performance, and adding functions I haven't seen anywhere else, or at least improving on the most common ones, I think this version is good enough to be released. 

After all the hard work of compiling a class that is both simple yet featuring advanced functions and thoroughly commenting everything, so it is clear and easy to understand, I have decided to release what I originally intended to keep for private use, so that other people may benefit from this and are not forced to 'reinvent the wheel' every time, when it comes to file, folder and drives management. 

clsFileAnalyzer contains over 50 different functions for file, folder and disk handling, whereas most of them were created by myself, while a few have been heavily modified from other versions I have come accross online. I have tried to review each one carefully, but if you manage to find a bug or think you could make it faster or more reliable, please do not hesitate to contact me. 

You will find all my software on www.kaotix-crew.net, where many of my other open source applications have been released, so check back if you're interested.

I hope you enjoy this class,

- NightWolf

----------------------------
List of Functions
----------------------------

-> General Functions

Build Path - Returns a valid, available path to create a file of the specified name
Execute - Executes the specified file/folder/url with its assigned application
Exists - Determines wether the specified path is valid
FixPath - Appends a backslash to the specified path if required
GetAbsolutePath - Converts the specified path from relative to absolute
GetAttributeNames - Returns a string containing the specified attributes' names
GetBaseName - Returns the name portion of the specified path, discarding any extension
GetLongPath - Retrieves the long path of the specified short path
GetParentFolder - Retrieves the folder which holds the specified path
GetPathAttributes - Retrieves the specified path's attributes
GetPathDate - Retrieves information about a path's created/accessed/modified dates
GetPathName - Returns the File Name + Extension portion of the specified Path
GetPathType - Retrieves the description associated with the specified file's type
GetRelativePath - Returns a relative path from a specified path
GetShortPath - Gets the MS-DOS 6 characters ID from a specified path
MoveToRecycleBin - Moves the specified file/folder to the Recycle Bin
ParseSize - Returns a string containing the specified size, formatted accordingly
Rename - Renames the specified path to the determined name
Search - Searches the specified path for files matching the given criteria
SetPathAttributes - Sets the specified path's attributes to those passed to the function
SetPathDate - Sets the Created/Accessed/Modified date of the specified path

-> File Functions

CopyFile - Copies a file to the specified location
CreateFile - Creates a file and optionally places the specified content inside of it
DeleteFile - Deletes the specified file
GetFileExtension - Returns the extension portion of the specified path, if available
GetFileExecutable - Returns the executable associated with the specified file
GetFileSize - Returns the specified file's size
GetFileContents - Retrieves the contents of the specified file
GetTempFile - Returns an available file name to open within the system's temporary folder
IsFileInUse - Determines if the specified file is currently in use by another process

-> Folder Functions

CopyFolder - Copies an entire directory to the specified location
CreateFolder - Creates the specified folder, if it doesn't exist
DeleteFolder - Deletes the specified folder and its entire contents
GetFolderContents - Retrieves all files/folders inside the specified path
GetFolderSize - Retrieves the specified folder's size
GetSpecialFolder - Retrieves the paths of various system folders
IsFolder - Detects wether the specified path is a directory
IsEmpty - Detects wether the specified directory is empty
PathHasSubFolders - Checks for the existance of subfolders within the specified path

-> Disk/Drive Functions

DriveExists - Checks wether the specified drive exists
GetDrives - Retrieves all available drives
GetDriveType - Retrieves an integer specifying the passed Drive's type
GetDriveTypeName - Returns a string with the corresponding Drive's type association
GetDriveInfo - Retrieves the specified drive's information
GetDriveNumber - Retrieves the system order of the specified drive
GetDriveSpace - Retrieves information about the specified drive's size
GetPathDrive - Strips the Drive portion of a path
GetSystemDrive - Returns the drive in which the Operating System is running
SetDriveLabel - Sets the label of the specified drive