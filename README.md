<div align="center">

## Directory and file enumeration


</div>

### Description

ListSubDirs:

Lists names of all directories found under a given path. (Does not search recursively)

ListFiles:

Lists names of all files found under a given path. (Does not search recursively)
 
### More Info
 
Path = path to directory under which you would like to scan.

These are standalone functions.

Array of directories, or files found under path argument.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alex Wolfe](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alex-wolfe.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alex-wolfe-directory-and-file-enumeration__1-6233/archive/master.zip)





### Source Code

```
Function ListSubDirs(ByVal Path As String) As Variant
  'returns an array of directory names
  On Error Resume Next
  Dim Count, Dirs(), i, DirName ' Declare variables.
  DirName = Dir(Path, vbDirectory) ' Get first directory name.
  Count = 0
  Do While Not DirName = ""
    ' A file or directory name was returned
    If Not DirName = "." And Not DirName = ".." Then
      ' Not a parent or current directory entry so process it
      If GetAttr(Path & DirName) And vbDirectory Then
        ' This is a directory
        ' Increase the size of the array by one element
        ReDim Preserve Dirs(Count + 1)
        Dirs(Count) = DirName ' Add directory name to array
        Count = Count + 1 ' Increment counter.
      End If
    End If
    DirName = Dir ' Get another directory name.
  Loop
  ReDim Preserve Dirs(Count - 1) 'remove the last empty element
  ListSubDirs = Dirs()
End Function
Function ListFiles(ByVal Path As String) As Variant
  'returns an array of file names
  On Error Resume Next
  Dim Count, Files(), i, FileName ' Declare variables.
  Count = 0
  FileName = Dir(Path, 6) ' Get first file name.
  Do While Not FileName = ""
    If Not FileName = "." And Not FileName = ".." Then
      'Not a parent or current directory entry so process it
      If Not GetAttr(Path & FileName) And vbDirectory Then
        'This is a file
        'Increase the size of the array by one element
        ReDim Preserve Files(Count + 1)
        Files(Count) = FileName 'Add Filename to array.
        Count = Count + 1 'Increment counter
      End If
    End If
    FileName = Dir ' Get another file name.
  Loop
  ReDim Preserve Files(Count - 1) 'remove the last empty element
  ListFiles = Files()
End Function
```

