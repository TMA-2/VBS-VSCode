Option Explicit

''' <summary>An intrinsic global object that can send output to a script debugger, such as the Microsoft Script Debugger.</summary>
Class Debug

	''' <summary>Sends a string to the debugger.</summary>
	''' <param name="str">String to send to the debugger.</param>
	Sub Write(str)
	End Sub

	''' <summary>Sends a string followed by a newline character to the debugger.</summary>
	''' <param name="str">String to send to the debugger.</param>
	Sub WriteLine(str)
	End Sub
End Class


''' <summary>Object that stores data key, item pairs.</summary>
Class Dictionary

	''' <summary>Adds a key and item pair to a Dictionary object.</summary>
	''' <param name="key">The key associated with the item being added.</param>
	''' <param name="value">The item associated with the key being added.</param>
	Sub Add(key, value)
	End Sub

	''' <summary>Sets or returns the comparison mode for comparing string keys in a Dictionary object.</summary>
	Property Get CompareMode ' As Long
	End Property
	''' <summary>Sets or returns the comparison mode for comparing string keys in a Dictionary object.</summary>
	''' <param name="Mode">Comparison mode.</param>
	Property Let CompareMode(Mode)
	End Property

	''' <summary>Returns the number of items in a collection or Dictionary object.</summary>
	Property Get Count ' As Long
	End Property

	''' <summary>Returns true if a specified key exists in the Dictionary object, false if it does not.</summary>
	''' <param name="key">Key value being searched for in the Dictionary object.</param>
	Function Exists(key) ' as Boolean
	End Function

	''' <summary>Returns the hash value for a specified key in a Dictionary object.</summary>
	''' <param name="key">Key associated with the item for which the hash value is to be returned.</param>
	Property Get HashVal(key) ' As Long
	End Property

	''' <summary>Sets or returns an item for a specified key in a Dictionary object.</summary>
	''' <param name="key">Key associated with the item being retrieved or added.</param>
	Public Default Property Get Item(key)
	End Property

	''' <summary>Returns an array containing all the items in a Dictionary object.</summary>
	''' <param name="key">Key associated with the item being retrieved.</param>
	Function Items(key) ' As Variant
	End Function

	''' <summary>Sets a key in a Dictionary object.</summary>
	''' <param name="key">Key being changed.</param>
	Property Get Key(key)
	End Property

	''' <summary>Returns an array containing all existing keys in a Dictionary object.</summary>
	Function Keys() ' As Variant
	End Function

	''' <summary>Removes a key, item pair from a Dictionary object.</summary>
	''' <param name="key">Key associated with the key, item pair you want to remove from the Dictionary object.</param>
	Sub Remove(key)
	End Sub

	''' <summary>Removes all key, item pairs from a Dictionary object.</summary>
	Sub RemoveAll()
	End Sub

End Class


''' <summary>Contains information about run-time errors. Accepts the Raise and Clear methods for generating and clearing run-time errors.</summary>
Class Err

	''' <summary>Clears all property settings of the Err object.</summary>
	Sub Clear()
	End Sub

	''' <summary>Returns or sets a string expression containing a descriptive string associated with an error.</summary>
	Property Get Description
	End Property

	''' <summary>Returns or sets an integer containing the context ID for a topic in a Help file.</summary>
	Property Get HelpContext
	End Property

	''' <summary>Returns or sets a string expression containing the fully qualified path to a Help file.</summary>
	Property Get HelpFile
	End Property

	''' <summary>Returns or sets a numeric value specifying an error.</summary>
	Property Get Number
	End Property

	''' <summary>Generates a run-time error.</summary>
	''' <param name="number">Long integer that identifies the nature of the error.</param>
	Sub Raise(number)
	End Sub

	''' <summary>Generates a run-time error.</summary>
	''' <param name="number">Long integer that identifies the nature of the error.</param>
	''' <param name="source">String expression naming the object or application that generated the error.</param>
	''' <param name="description">String expression describing the error.</param>
	''' <param name="helpfile">The fully qualified path to the Help file in which help on this error can be found.</param>
	''' <param name="helpcontext">The context ID identifying a topic within helpfile that provides help for the error.</param>
	Sub Raise(number, source, description, helpfile, helpcontext)
	End Sub

	''' <summary>Returns or sets a string expression specifying the name of the object or application that originally generated the error.</summary>
	Property Get Source
	End Property

End Class


''' <summary>Provides access to all the properties of a file.</summary>
Class File

	''' <summary>Sets or returns the attributes of files or folders.</summary>
	Property Get Attributes ' as Long
	End Property

	''' <summary>Copies a specified file from one location to another.</summary>
	''' <param name="Destination">Destination where the file is to be copied.</param>
	Sub Copy(Destination)
	End Sub
	''' <summary>Copies a specified file from one location to another.</summary>
	''' <param name="Destination">Destination where the file is to be copied.</param>
	''' <param name="OverWriteFiles">Boolean value that indicates if existing files are to be overwritten.</param>
	Sub Copy(Destination, OverWriteFiles)
	End Sub

	''' <summary>Returns the date and time that the specified file or folder was created.</summary>
	Property Get DateCreated ' as Date
	End Property

	''' <summary>Returns the date and time that the specified file or folder was last accessed.</summary>
	Property Get DateLastAccessed ' as Date
	End Property

	''' <summary>Returns the date and time that the specified file or folder was last modified.</summary>
	Property Get DateLastModified ' as Date
	End Property

	''' <summary>Deletes a specified file or folder.</summary>
	Sub Delete()
	End Sub
	''' <summary>Deletes a specified file or folder.</summary>
	''' <param name="Force">Boolean value that is True if files or folders with the read-only attribute should be deleted; False if they should not.</param>
	Sub Delete(Force)
	End Sub

	''' <summary>Returns a Drive object corresponding to the drive on which the specified file or folder resides.</summary>
	Property Get Drive ' as Drive
	End Property

	''' <summary>Moves a specified file or folder from one location to another.</summary>
	''' <param name="Destination">Destination where the file or folder is to be moved.</param>
	Sub Move(Destination)
	End Sub

	''' <summary>Sets or returns the name of a specified file or folder.</summary>
	Property Get Name ' As String
	End Property

	''' <summary>Opens a specified file and returns a TextStream object that can be used to read from the file.</summary>
	Function OpenAsTextStream() ' As TextStream
	End Function
	''' <summary>Opens a specified file and returns a TextStream object that can be used to read from the file.</summary>
	''' <param name="IOMode">Indicates input/output mode.</param>
	Function OpenAsTextStream(IOMode) ' As TextStream
	End Function
	''' <summary>Opens a specified file and returns a TextStream object that can be used to read from the file.</summary>
	''' <param name="IOMode">Indicates input/output mode.</param>
	''' <param name="Format">One of three Tristate values used to indicate the format of the opened file.</param>
	Function OpenAsTextStream(IOMode, Format) ' As TextStream
	End Function

	''' <summary>Returns the folder object for the parent of the specified file or folder.</summary>
	Property Get ParentFolder ' As Folder
	End Property

	''' <summary>Returns the path for a specified file, folder, or drive.</summary>
	Property Get Path ' As String
	End Property

	''' <summary>Returns the short name used by programs that require the earlier 8.3 naming convention.</summary>
	Property Get ShortName ' As String
	End Property

	''' <summary>Returns the short path used by programs that require the earlier 8.3 file naming convention.</summary>
	Property Get ShortPath ' As String
	End Property

	''' <summary>For files, returns the size, in bytes, of the specified file. For folders, returns the size, in bytes, of all files and subfolders contained in the folder.</summary>
	Property Get Size ' as Long
	End Property

	''' <summary>Returns information about the type of a file or folder.</summary>
	Property Get Type ' As String
	End Property

End Class


''' <summary>Provides access to all the properties of a folder.</summary>
Class Folder

	''' <summary>Sets or returns the attributes of files or folders.</summary>
	Property Get Attributes ' as Long
	End Property

	''' <summary>Copies a specified folder from one location to another.</summary>
	''' <param name="Destination">Destination where the folder is to be copied.</param>
	Sub Copy(Destination)
	End Sub
	''' <summary>Copies a specified folder from one location to another.</summary>
	''' <param name="Destination">Destination where the folder is to be copied.</param>
	''' <param name="OverWriteFiles">Boolean value that indicates if existing files are to be overwritten.</param>
	Sub Copy(Destination, OverWriteFiles)
	End Sub

	''' <summary>Returns the date and time that the specified folder was created.</summary>
	Property Get DateCreated ' as Date
	End Property

	''' <summary>Returns the date and time that the specified folder was last accessed.</summary>
	Property Get DateLastAccessed ' as Date
	End Property

	''' <summary>Returns the date and time that the specified folder was last modified.</summary>
	Property Get DateLastModified ' as Date
	End Property

	''' <summary>Deletes a specified folder.</summary>
	Sub Delete()
	End Sub
	''' <summary>Deletes a specified folder.</summary>
	''' <param name="Force">Boolean value that is True if folders with the read-only attribute should be deleted; False if they should not.</param>
	Sub Delete(Force)
	End Sub

	''' <summary>Returns a Drive object corresponding to the drive on which the specified folder resides.</summary>
	Property Get Drive ' as Drive
	End Property

	''' <summary>Returns a Files collection consisting of all File objects contained in the specified folder, including those with hidden and system file attributes set.</summary>
	Property Get Files ' as FileCollection
	End Property

	''' <summary>Returns True if the specified folder is the root folder; False if it is not.</summary>
	Property Get IsRootFolder ' as Boolean
	End Property

	''' <summary>Moves a specified folder from one location to another.</summary>
	''' <param name="Destination">Destination where the folder is to be moved.</param>
	Sub Move(Destination)
	End Sub

	''' <summary>Sets or returns the name of a specified folder.</summary>
	Property Get Name ' As String
	End Property

	''' <summary>Creates a file as a TextStream object.</summary>
	''' <param name="FileName">String expression that identifies the file to create.</param>
	Function CreateTextFile(FileName) ' As TextStream
	End Function
	''' <summary>Creates a file as a TextStream object.</summary>
	''' <param name="FileName">String expression that identifies the file to create.</param>
	''' <param name="Overwrite">Boolean value that indicates whether you can overwrite an existing file.</param>
	Function CreateTextFile(FileName, Overwrite) ' As TextStream
	End Function
	''' <summary>Creates a file as a TextStream object.</summary>
	''' <param name="FileName">String expression that identifies the file to create.</param>
	''' <param name="Overwrite">Boolean value that indicates whether you can overwrite an existing file.</param>
	''' <param name="Unicode">Boolean value that indicates whether the file is created as a Unicode or ASCII file.</param>
	Function CreateTextFile(FileName, Overwrite, Unicode) ' As TextStream
	End Function

	''' <summary>Returns the folder object for the parent of the specified folder.</summary>
	Property Get ParentFolder ' As Folder
	End Property

	''' <summary>Returns the path for a specified folder.</summary>
	Property Get Path ' As String
	End Property

	''' <summary>Returns the short name used by programs that require the earlier 8.3 naming convention.</summary>
	Property Get ShortName ' As String
	End Property

	''' <summary>Returns the short path used by programs that require the earlier 8.3 file naming convention.</summary>
	Property Get ShortPath ' As String
	End Property

	''' <summary>For folders, returns the size, in bytes, of all files and subfolders contained in the folder.</summary>
	Property Get Size ' as Long
	End Property

	''' <summary>Returns a Folders collection consisting of all folders contained in a specified folder.</summary>
	Property Get SubFolders ' as FolderCollection
	End Property

	''' <summary>Returns information about the type of a folder.</summary>
	Property Get Type ' As String
	End Property

End Class


''' <summary>Provides access to a computer's file system.</summary>
Class FileSystemObject

	''' <summary>Appends a name to an existing path.</summary>
	''' <param name="path">Existing path to which name is appended.</param>
	''' <param name="name">Name being appended to the existing path.</param>
	Function BuildPath(path, name) ' As String
	End Function

	''' <summary>Copies one or more files from one location to another.</summary>
	''' <param name="source">Character string file specification, which can include wildcard characters, for one or more files to be copied.</param>
	''' <param name="destination">Character string destination where the file or files from source are to be copied.</param>
	Sub CopyFile(source, destination)
	End Sub
	''' <summary>Copies one or more files from one location to another.</summary>
	''' <param name="source">Character string file specification, which can include wildcard characters, for one or more files to be copied.</param>
	''' <param name="destination">Character string destination where the file or files from source are to be copied.</param>
	''' <param name="overwrite">Boolean value that indicates if existing files are to be overwritten.</param>
	Sub CopyFile(source, destination, overwrite)
	End Sub

	''' <summary>Recursively copies a folder from one location to another.</summary>
	''' <param name="source">Character string folder specification, which can include wildcard characters, for one or more folders to be copied.</param>
	''' <param name="destination">Character string destination where the folder and subfolders from source are to be copied.</param>
	Sub CopyFolder(source, destination)
	End Sub
	''' <summary>Recursively copies a folder from one location to another.</summary>
	''' <param name="source">Character string folder specification, which can include wildcard characters, for one or more folders to be copied.</param>
	''' <param name="destination">Character string destination where the folder and subfolders from source are to be copied.</param>
	''' <param name="overwrite">Boolean value that indicates if existing folders are to be overwritten.</param>
	Sub CopyFolder(source, destination, overwrite)
	End Sub

	''' <summary>Creates a folder.</summary>
	''' <param name="foldername">String expression that identifies the folder to create.</param>
	Function CreateFolder(foldername) ' As Folder
	End Function

	''' <summary>Creates a file as a TextStream object.</summary>
	''' <param name="filename">String expression that identifies the file to create.</param>
	Function CreateTextFile(filename) ' As TextStream
	End Function
	''' <summary>Creates a file as a TextStream object.</summary>
	''' <param name="filename">String expression that identifies the file to create.</param>
	''' <param name="overwrite">Boolean value that indicates whether you can overwrite an existing file.</param>
	Function CreateTextFile(filename, overwrite) ' As TextStream
	End Function
	''' <summary>Creates a file as a TextStream object.</summary>
	''' <param name="filename">String expression that identifies the file to create.</param>
	''' <param name="overwrite">Boolean value that indicates whether you can overwrite an existing file.</param>
	''' <param name="unicode">Boolean value that indicates whether the file is created as a Unicode or ASCII file.</param>
	Function CreateTextFile(filename, overwrite, unicode) ' As TextStream
	End Function

	''' <summary>Deletes a specified file.</summary>
	''' <param name="filename">String expression that specifies one or more files to delete.</param>
	Sub DeleteFile(filename)
	End Sub
	''' <summary>Deletes a specified file.</summary>
	''' <param name="filename">String expression that specifies one or more files to delete.</param>
	''' <param name="force">Boolean value that is True if files with the read-only attribute should be deleted; False if they should not.</param>
	Sub DeleteFile(filename, force)
	End Sub

	''' <summary>Deletes a specified folder and its contents.</summary>
	''' <param name="filename">String expression that specifies one or more folders to delete.</param>
	Sub DeleteFolder(filename)
	End Sub
	''' <summary>Deletes a specified folder and its contents.</summary>
	''' <param name="filename">String expression that specifies one or more folders to delete.</param>
	''' <param name="force">Boolean value that is True if folders with the read-only attribute should be deleted; False if they should not.</param>
	Sub DeleteFolder(filename, force)
	End Sub

	''' <summary>Returns a Drives collection consisting of all Drive objects available on the local machine.</summary>
	Property Get Drives ' As DriveCollection
	End Property

	''' <summary>Returns True if the specified drive exists; False if it does not.</summary>
	''' <param name="drive">Drive letter or complete path specification.</param>
	Function DriveExists(drive) ' As Boolean
	End Function

	''' <summary>Returns True if a specified file exists; False if it does not.</summary>
	''' <param name="filename">String expression that specifies a file name.</param>
	Function FileExists(filename) ' As Boolean
	End Function

	''' <summary>Returns True if a specified folder exists; False if it does not.</summary>
	''' <param name="foldername">String expression that specifies a folder name.</param>
	Function FolderExists(foldername) ' As Boolean
	End Function

	''' <summary>Returns a complete and unambiguous path from a provided path specification.</summary>
	''' <param name="path">Path specification to change to a complete and unambiguous path.</param>
	Function GetAbsolutePathName(path) ' As String
	End Function

	''' <summary>Returns a string containing the base name of the last component, less any file extension, in a path.</summary>
	''' <param name="path">Path specification for the component whose base name is to be returned.</param>
	Function GetBaseName(path) ' As String
	End Function

	''' <summary>Returns a Drive object corresponding to the drive in a specified path.</summary>
	''' <param name="drive">Path specification whose drive is to be returned.</param>
	Function GetDrive(drive) ' As Drive
	End Function

	''' <summary>Returns a string containing the name of the drive for a specified path.</summary>
	''' <param name="drive">Path specification for the component whose drive name is to be returned.</param>
	Function GetDriveName(drive) ' As String
	End Function

	''' <summary>Returns a string containing the extension name for the last component in a path.</summary>
	''' <param name="path">Path specification for the component whose extension name is to be returned.</param>
	Function GetExtensionName(path) ' As String
	End Function

	''' <summary>Returns a File object corresponding to the file in a specified path.</summary>
	''' <param name="filename">Path specification for the file that you want to get a File object for.</param>
	Function GetFile(filename) ' As File
	End Function

	''' <summary>Returns a string containing the name of the last component, less any file extension, in a path.</summary>
	''' <param name="filename">Path specification for the component whose file name is to be returned.</param>
	Function GetFileName(filename) ' As String
	End Function

	''' <summary>Returns version information for the specified file.</summary>
	''' <param name="filename">String expression that specifies a file name.</param>
	Function GetFileVersion(filename) ' As String
	End Function

	''' <summary>Returns a Folder object corresponding to the folder in a specified path.</summary>
	''' <param name="foldername">Path specification for the folder that you want to get a Folder object for.</param>
	Function GetFolder(foldername) ' As Folder
	End Function

	''' <summary>Returns a string containing the name of the parent folder of the last component in a specified path.</summary>
	''' <param name="foldername">Path specification for the component whose parent folder name is to be returned.</param>
	Function GetParentFolderName(foldername) ' As String
	End Function

	''' <summary>Returns the special folder specified.</summary>
	''' <param name="folderspec">The name of the special folder to be returned.</param>
	Function GetSpecialFolder(folderspec) ' As Folder
	End Function

	''' <summary>Returns a TextStream object corresponding to the standard input, output, or error stream.</summary>
	''' <param name="StandardStreamType">Indicates which standard stream to return.</param>
	Function GetStandardStream(StandardStreamType) ' As TextStream
	End Function
	''' <summary>Returns a TextStream object corresponding to the standard input, output, or error stream.</summary>
	''' <param name="StandardStreamType">Indicates which standard stream to return.</param>
	''' <param name="Unicode">Boolean value indicating whether the stream is Unicode or ASCII.</param>
	Function GetStandardStream(StandardStreamType, Unicode) ' As TextStream
	End Function

	''' <summary>Returns a randomly generated temporary file or folder name that is useful for performing operations that require a temporary file or folder.</summary>
	Function GetTempName() ' As String
	End Function

	''' <summary>Moves one or more files from one location to another.</summary>
	''' <param name="source">Path to the file or files to be moved.</param>
	''' <param name="destination">Path where the file or files are to be moved.</param>
	Sub MoveFile(source, destination)
	End Sub

	''' <summary>Moves one or more folders from one location to another.</summary>
	''' <param name="source">Path to the folder or folders to be moved.</param>
	''' <param name="destination">Path where the folder or folders are to be moved.</param>
	Sub MoveFolder(source, destination)
	End Sub

	''' <summary>Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.</summary>
	''' <param name="filename">String expression that identifies the file to open.</param>
	Function OpenTextFile(filename) ' As TextStream
	End Function
	''' <summary>Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.</summary>
	''' <param name="filename">String expression that identifies the file to open.</param>
	''' <param name="iomode">Indicates input/output mode.</param>
	Function OpenTextFile(filename, iomode) ' As TextStream
	End Function
	''' <summary>Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.</summary>
	''' <param name="filename">String expression that identifies the file to open.</param>
	''' <param name="iomode">Indicates input/output mode.</param>
	''' <param name="create">Boolean value that indicates whether a new file can be created if the specified filename doesn't exist.</param>
	Function OpenTextFile(filename, iomode, create) ' As TextStream
	End Function
	''' <summary>Opens a specified file and returns a TextStream object that can be used to read from, write to, or append to the file.</summary>
	''' <param name="filename">String expression that identifies the file to open.</param>
	''' <param name="iomode">Indicates input/output mode.</param>
	''' <param name="create">Boolean value that indicates whether a new file can be created if the specified filename doesn't exist.</param>
	''' <param name="format">One of three Tristate values used to indicate the format of the opened file.</param>
	Function OpenTextFile(filename, iomode, create, format) ' As TextStream
	End Function

End Class


''' <summary>Provides access to the properties of a particular disk drive or network share.</summary>
Class Drive
	''' <summary>Returns the amount of space available to a user on the specified drive or network share.</summary>
	Property Get AvailableSpace ' As Double
	End Property

	''' <summary>Returns the drive letter of a physical local drive or a network share.</summary>
	Property Get DriveLetter ' As String
	End Property

	''' <summary>Returns a value indicating the type of a specified drive.</summary>
	Property Get DriveType ' As Long
	End Property

	''' <summary>Returns the type of file system in use for the specified drive.</summary>
	Property Get FileSystem ' As String
	End Property

	''' <summary>Returns the amount of free space available to a user on the specified drive or network share.</summary>
	Property Get FreeSpace ' As Double
	End Property

	''' <summary>Returns True if the specified drive is ready; False if it is not.</summary>
	Property Get IsReady ' As Boolean
	End Property

	''' <summary>Returns the path for a specified drive.</summary>
	Property Get Path ' As String
	End Property

	''' <summary>Returns a Folder object representing the root folder of a specified drive.</summary>
	Property Get RootFolder ' As Folder
	End Property

	''' <summary>Returns the decimal serial number used to uniquely identify a disk volume.</summary>
	Property Get SerialNumber ' As Long
	End Property

	''' <summary>Returns the network share name for a specified drive.</summary>
	Property Get ShareName ' As String
	End Property

	''' <summary>Returns the total space, in bytes, of a drive or network share.</summary>
	Property Get TotalSize ' As Double
	End Property

	''' <summary>Sets or returns the volume name of the specified drive.</summary>
	Property Get VolumeName ' As String
	End Property

End Class


''' <summary>Provides access to the read-only properties of a regular expression match.</summary>
Class Match

	''' <summary>Gets the position in the original string where the first character of the match is found.</summary>
	''' <returns>A Long value representing the zero-based index of the first character of the match.</returns>
	Property Get FirstIndex ' As Long
	End Property

	''' <summary>Gets the length of the match in the original string.</summary>
	''' <returns>A Long value representing the number of characters in the match.</returns>
	Property Get Length ' As Long
	End Property

	''' <summary>Gets a collection containing all the submatches (captured groups) found during the regular expression search.</summary>
	''' <returns>A String array containing the submatches. Each element corresponds to a captured group in the regular expression pattern.</returns>
	Property Get SubMatches ' As String()
	End Property

	''' <summary>Gets the actual text of the match found in the original string.</summary>
	''' <returns>A String containing the matched text.</returns>
	Property Get Value ' As String
	End Property

End Class


''' <summary>Provides simple regular expression support.</summary>
Class RegExp

	''' <summary>Executes a regular expression search against a specified string.</summary>
	''' <param name="str">The text string upon which the regular expression is executed.</param>
	Function Execute(str) ' as Object
	End Function

	''' <summary>Sets or returns a Boolean value that indicates if a pattern should match all occurrences in an entire search string or just the first one.</summary>
	Property Get Global ' As Boolean
	End Property
	''' <summary>Sets or returns a Boolean value that indicates if a pattern should match all occurrences in an entire search string or just the first one.</summary>
	''' <param name="b">Boolean value.</param>
	Property Let Global(b)
	End Property

	''' <summary>Sets or returns a Boolean value that indicates if a pattern search is case-sensitive or not.</summary>
	Property Get IgnoreCase ' As Boolean
	End Property
	''' <summary>Sets or returns a Boolean value that indicates if a pattern search is case-sensitive or not.</summary>
	''' <param name="b">Boolean value.</param>
	Property Let IgnoreCase(b)
	End Property

	''' <summary>Sets or returns the regular expression pattern being searched for.</summary>
	Property Get Pattern ' As String
	End Property
	''' <summary>Sets or returns the regular expression pattern being searched for.</summary>
	''' <param name="s">Regular expression pattern.</param>
	Property Let Pattern(s)
	End Property

	''' <summary>Replaces text found in a regular expression search.</summary>
	''' <param name="string1">The text string in which the text replacement is to occur.</param>
	''' <param name="string2">The replacement text string.</param>
	Function Replace(string1, string2) ' As String
	End Function

	''' <summary>Executes a regular expression search against a specified string and returns a Boolean value that indicates if a pattern match was found.</summary>
	''' <param name="str">The text string upon which the regular expression is executed.</param>
	Function Test(str) ' As Boolean
	End Function

End Class


''' <summary>Facilitates sequential access to file.</summary>
Class TextStream

	''' <summary>Returns True if the file pointer is positioned immediately before the end-of-line marker in a TextStream file; False if it is not.</summary>
	Property Get AtEndOfLine ' As Boolean
	End Property

	''' <summary>Returns True if the file pointer is at the end of a TextStream file; False if it is not.</summary>
	Property Get AtEndOfStream ' As Boolean
	End Property

	''' <summary>Closes an open TextStream file.</summary>
	Sub Close()
	End Sub

	''' <summary>Returns the column number of the current character position in a TextStream file.</summary>
	Property Get Column ' As Long
	End Property

	''' <summary>Returns the current line number in a TextStream file.</summary>
	Property Get Line ' As Long
	End Property

	''' <summary>Reads a specified number of characters from a TextStream file and returns the resulting string.</summary>
	''' <param name="Characters">Number of characters you want to read from the file.</param>
	Function Read(Characters) ' As String
	End Function

	''' <summary>Reads an entire TextStream file and returns the resulting string.</summary>
	Function ReadAll() ' As String
	End Function

	''' <summary>Reads an entire line from a TextStream file and returns the resulting string.</summary>
	Function ReadLine() ' As String
	End Function

	''' <summary>Skips a specified number of characters when reading a TextStream file.</summary>
	''' <param name="Characters">Number of characters to skip when reading a file.</param>
	Sub Skip(Characters)
	End Sub

	''' <summary>Skips the next line when reading a TextStream file.</summary>
	Sub SkipLine()
	End Sub

	''' <summary>Writes a specified string to a TextStream file.</summary>
	''' <param name="Text">The text you want to write to the file.</param>
	Sub Write(Text)
	End Sub

	''' <summary>Writes a specified number of newline characters to a TextStream file.</summary>
	''' <param name="Lines">Number of newline characters you want to write to the file.</param>
	Sub WriteBlankLines(Lines)
	End Sub

	''' <summary>Writes a newline character to a TextStream file.</summary>
	Sub WriteLine()
	End Sub
	''' <summary>Writes a specified string and newline character to a TextStream file.</summary>
	''' <param name="text">The text you want to write to the file.</param>
	Sub WriteLine(text)
	End Sub

End Class


''' <summary>Provides access to the root object of the Windows Script Host object model.</summary>
Class WScript

	''' <summary>Returns the WScript object, which is the root object of the Windows Script Host object model.</summary>
	Property Get Application
	End Property

	''' <summary>Returns the WshArguments object (a collection of arguments).</summary>
	Property Get Arguments
	End Property

	''' <summary>Returns the build version of Windows Script Host.</summary>
	Property Get BuildVersion ' As String
	End Property

	''' <summary>Connects the object's event sources to functions with a given prefix.</summary>
	''' <param name="objEventSource">Object to be connected.</param>
	''' <param name="strPrefix">String value indicating the function prefix.</param>
	Sub ConnectObject(objEventSource, strPrefix)
	End Sub

	''' <summary>Creates an object.</summary>
	''' <param name="strProgID">String value indicating the programmatic identifier (ProgID) of the object you want to create.</param>
	Function CreateObject(strProgID) ' As Object
	End Function

	''' <summary>Creates an object.</summary>
	''' <param name="strProgID">String value indicating the programmatic identifier (ProgID) of the object you want to create.</param>
	''' <param name="strPrefix">String value indicating the function prefix.</param>
	Function CreateObject(strProgID, strPrefix) ' As Object
	End Function

	''' <summary>Disconnects a connected object's event sources.</summary>
	''' <param name="obj">Object to be disconnected.</param>
	Sub DisconnectObject(obj)
	End Sub

	''' <summary>Outputs text to either a message box or the command console window.</summary>
	''' <param name="args">List of items to be displayed.</param>
	Sub Echo(args)
	End Sub

	''' <summary>Returns the fully qualified path of the host executable.</summary>
	Property Get FullName ' As String
	End Property

	''' <summary>Retrieves an existing object with the specified ProgID from memory, or creates a new one from a file.</summary>
	''' <param name="strPathname">Fully qualified path to the file containing the object persisted to disk.</param>
	Function GetObject(strPathname) ' As Object
	End Function

	''' <summary>Retrieves an existing object with the specified ProgID from memory, or creates a new one from a file.</summary>
	''' <param name="strPathname">Fully qualified path to the file containing the object persisted to disk.</param>
	''' <param name="strProgID">Program identifier.</param>
	Function GetObject(strPathname, strProgID) ' As Object
	End Function

	''' <summary>Retrieves an existing object with the specified ProgID from memory, or creates a new one from a file.</summary>
	''' <param name="strPathname">Fully qualified path to the file containing the object persisted to disk.</param>
	''' <param name="strProgID">Program identifier.</param>
	''' <param name="strPrefix">String value indicating the function prefix.</param>
	Function GetObject(strPathname, strProgID, strPrefix) ' As Object
	End Function

	''' <summary>Sets the script mode, or identifies the script mode.</summary>
	Property Get Interactive ' As Boolean
	End Property

	''' <summary>Returns the name of the WScript object (the host executable file).</summary>
	Property Get Name ' As String
	End Property

	''' <summary>Returns the name of the directory containing the host executable.</summary>
	Property Get Path ' As String
	End Property

	''' <summary>Forces the script to stop execution at any time.</summary>
	Sub Quit()
	End Sub

	''' <summary>Forces the script to stop execution at any time.</summary>
	''' <param name="ErrorCode">Numeric value returned as the process exit code.</param>
	Sub Quit(ErrorCode)
	End Sub

	''' <summary>Returns the full path of the currently running script.</summary>
	Property Get ScriptFullName ' As String
	End Property

	''' <summary>Returns the file name of the currently running script.</summary>
	Property Get ScriptName ' As String
	End Property

	''' <summary>Suspends script execution for a specified length of time, then continues execution.</summary>
	''' <param name="ms">Numeric value indicating the interval (in milliseconds) you want the script to be inactive.</param>
	Sub Sleep(ms)
	End Sub

	''' <summary>Exposes the write-only error output stream of the current script.</summary>
	Property Get StdErr ' As TextStream
	End Property

	''' <summary>Exposes the read-only input stream of the current script.</summary>
	Property Get StdIn ' As TextStream
	End Property

	''' <summary>Exposes the write-only output stream of the current script.</summary>
	Property Get StdOut ' As TextStream
	End Property

	''' <summary>Returns the timeout setting in seconds for the WScript object.</summary>
	Property Get TimeOut ' As Integer
	End Property

	''' <summary>Returns the version of Windows Script Host.</summary>
	Property Get Version ' As String
	End Property

End Class

''' <summary>Provides access to Windows shell functionality.</summary>
Class Shell

	''' <summary>Activates an application window.</summary>
	''' <param name="App">Specifies which application to activate.</param>
	''' <param name="Wait">Optional. Boolean value indicating whether the script should wait for the application to become active.</param>
	Function AppActivate(App, Wait) ' As Boolean
	End Function

	''' <summary>Creates a shortcut, or assigns a value to an environment variable.</summary>
	''' <param name="PathLink">String value indicating the pathname of the shortcut to create.</param>
	Function CreateShortcut(PathLink)
	End Function

	''' <summary>Runs an application in a child command-shell, providing access to the StdIn/StdOut/StdErr streams.</summary>
	''' <param name="Command">String value indicating the command line used to run the script.</param>
	Function Exec(Command)
	End Function

	''' <summary>Returns an environment variable's expanded value.</summary>
	''' <param name="Src">String value indicating the name of the environment variable you want to expand.</param>
	Function ExpandEnvironmentStrings(Src) ' As String
	End Function

	''' <summary>Adds an event entry to a log file.</summary>
	''' <param name="Type">Numeric value indicating the type of entry.</param>
	''' <param name="Message">String value containing the log entry text.</param>
	''' <param name="Target">Optional. String value indicating the name of the computer system where the event log is stored.</param>
	Function LogEvent(Type, Message, Target) ' As Boolean
	End Function

	''' <summary>Displays text in a pop-up message box.</summary>
	''' <param name="Text">String value that contains the text you want to appear in the pop-up message box.</param>
	''' <param name="SecondsToWait">Optional. Numeric value indicating the maximum length of time (in seconds) you want the pop-up message box displayed.</param>
	''' <param name="Title">Optional. String value that contains the text you want to appear as the title of the pop-up message box.</param>
	''' <param name="Type">Optional. Numeric value indicating the type of buttons and icons you want in the pop-up message box.</param>
	Function Popup(Text, SecondsToWait, Title, Type) ' As Integer
	End Function

	''' <summary>Deletes a key or one of its values from the registry.</summary>
	''' <param name="Name">String value indicating the name of the registry key or key value you want to delete.</param>
	Sub RegDelete(Name)
	End Sub

	''' <summary>Returns the value of a key or value-name from the registry.</summary>
	''' <param name="Name">String value indicating the key or value-name whose value you want.</param>
	Function RegRead(Name)
	End Function

	''' <summary>Creates a new key, adds another value-name to an existing key (and assigns it a value), or changes the value of an existing value-name.</summary>
	''' <param name="Name">String value indicating the key-name, value-name, or value you want to create, add, or change.</param>
	''' <param name="Value">The name you want to assign to the value.</param>
	''' <param name="Type">Optional. String value indicating the value's data type.</param>
	Sub RegWrite(Name, Value, Type)
	End Sub

	''' <summary>Runs a program in a new process.</summary>
	''' <param name="Command">String value indicating the command line you want to run.</param>
	''' <param name="WindowStyle">Optional. Integer value indicating the appearance of the program's window.</param>
	''' <param name="WaitOnReturn">Optional. Boolean value indicating whether the script waits for the program to finish executing before continuing.</param>
	Function Run(Command, WindowStyle, WaitOnReturn) ' As Integer
	End Function

	''' <summary>Sends one or more keystrokes to the active window.</summary>
	''' <param name="Keys">String value indicating the keystroke(s) you want to send.</param>
	''' <param name="Wait">Optional. Boolean value indicating whether or not to wait for the keys to be processed before returning control to your script.</param>
	Sub SendKeys(Keys, Wait)
	End Sub

	''' <summary>Returns the WshEnvironment object (a collection of environment variables).</summary>
	''' <param name="Type">Optional. String value indicating the location of the environment variable.</param>
	Public Default Property Get Environment(Type)
	End Property

	''' <summary>Returns or sets the current active directory.</summary>
	Property Get CurrentDirectory
	End Property

	''' <summary>Returns or sets the current active directory.</summary>
	Property Let CurrentDirectory
	End Property

	''' <summary>Returns a WshSpecialFolders object (a collection of special folders).</summary>
	Property Get SpecialFolders
	End Property

End Class

''' <summary>Provides properties and methods for working with picture objects.</summary>
Private Class Picture

	''' <summary>Returns a handle to the picture.</summary>
	Property Get Handle ' As Long
	End Property

	''' <summary>Renders the picture to a specified device context.</summary>
	''' <param name="hdc">Handle to the device context where the picture is to be rendered.</param>
	''' <param name="x">Horizontal coordinate where the picture is placed.</param>
	''' <param name="y">Vertical coordinate where the picture is placed.</param>
	''' <param name="cx">Horizontal size of the destination rectangle.</param>
	''' <param name="cy">Vertical size of the destination rectangle.</param>
	''' <param name="xSrc">Horizontal offset in the source picture.</param>
	''' <param name="ySrc">Vertical offset in the source picture.</param>
	''' <param name="cxSrc">Horizontal extent of the source rectangle.</param>
	''' <param name="cySrc">Vertical extent of the source rectangle.</param>
	''' <param name="prcWBounds">Pointer to the bounding rectangle.</param>
	Sub Render(hdc, x, y, cx, cy, xSrc, ySrc, cxSrc, cySrc, prcWBounds)
	End Sub

	''' <summary>Returns the height of the picture.</summary>
	Property Get Height ' As Long
	End Property

	''' <summary>Returns the palette handle for the picture.</summary>
	Property Get hPal ' As Long
	End Property

	''' <summary>Returns the type of the picture.</summary>
	Property Get Type ' As Integer
	End Property

	''' <summary>Returns the width of the picture.</summary>
	Property Get Width ' As Long
	End Property

End Class