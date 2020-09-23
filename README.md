<div align="center">

## KillFiles


</div>

### Description

I received a request from someone on help with a problem in deleting

temporary files. It seems that they needed to delete all temporary files

except for those with the current date. This subroutine was the result, and I

though it would be good for those of you struggling with how to use the Dir and GetAttr

and SetAttr functions in VB
 
### More Info
 
Full path to the target directory including the drive letter and the extension type

to be deleted

Create a project with a single form and a command button and paste this code

into it.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jack Rizzo](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jack-rizzo.md)
**Level**          |Unknown
**User Rating**    |4.2 (165 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jack-rizzo-killfiles__1-3434/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
KillFiles "C:\windows\temp", ".tmp"
End Sub
Public Sub KillFiles(FilePath As String, Extension As String)
Dim curfile As String
Dim mydate As String
Dim tgtdate As String
Dim tgtpath As String
Dim oldpath As String
Dim indx As Integer
Dim attr As Integer
On Error GoTo TrapError
oldpath = CurDir      'Save Current Path and drive'
mydate = Format(Day(Now), "##00") 'Force current date to 2 digits
ChDrive FilePath         'make sure we change drive
ChDir FilePath          'and path to correct place
'
'Build full target path variable
'
If Right(FilePath, 1) = "\" Then
  tgtpath = FilePath & "*" & Extension
Else
  tgtpath = FilePath & "\*" & Extension
End If
'
' Get first target extension file in directory
'
curfile = Dir(tgtpath, vbNormal)
'
' Loop through directory of all extension files
'
While curfile <> ""
  tgtdate = FileDateTime(curfile)  'get file date
  indx = InStr(1, tgtdate, "/")   'find first date slash
  tgtdate = Mid(tgtdate, indx + 1) 'move in data
  indx = InStr(1, tgtdate, "/")   'find second slash
  tgtdate = Format(Left(tgtdate, indx - 1), "##00") 'form 2 digit date
  '
  ' Check to see if the dates are the same
  ' if not, delete the file
  '
  If tgtdate <> mydate Then
    '
    ' check attributes for readonly, system and hidden files
    '
    attr = GetAttr(curfile) And 31 ' and out unwanted bits
    If attr <> 0 Then 'file is special
     resp = MsgBox(curfile & " Is protected ... Delete?", vbYesNo)
     If resp = vbYes Then
       SetAttr curfile, vbNormal 'reset attributes so u can delete
       Kill curfile   ' delete the file
     End If
    Else
     Kill curfile ' file is normal file .. delete it
    End If
  End If
  curfile = Dir() ' get next file
Wend
ChDrive oldpath 'restore original drive
ChDir oldpath  'restore original path
Exit Sub
TrapError:
  MsgBox Error(Err) & " on " & curfile
  Resume Next
End Sub
```

