<div align="center">

## Get IE4 History URLs history folder


</div>

### Description

This code will open a DAT file in the c:\windows\history folder and pull out all sites visited
 
### More Info
 
For the Index.dat file the displacement is set to 119 for other files I have

set the displacement to 15.

For the Index.dat file the delimiter, or search string, is "URL "

For other files I have used "Visited: "

This example uses the Index.dat file, but you can easily modify it to get the others by making the documented changes to the code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Wells](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-wells.md)
**Level**          |Unknown
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-wells-get-ie4-history-urls-history-folder__1-1104/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
  Dim iDisplacement As Integer
  Dim iURLCount As Integer
  Dim sDelimiter As String
  Dim sData As String
  Dim sURLs(1 To 1000) As String
  Dim IEHistoryFile As String
  Dim i As Long
  Dim j As Long
  Dim x As Integer
  'For the Index.dat file the displacement is set to 119 for other files I  'have set the displacement to 15.
  iDisplacement = 119  'Index.dat = 119
  sDelimiter = "URL " '"Visited: "
  IEHistoryFile = "index.dat" 'Could also me an MM DAT file in History folder
  'For the Index.dat file the delimiter, or search string, is "URL "
  'For other files I have used "Visited: "
  'This is the History DAT file. I use Index.dat for this example, but there are MM files
  Open "c:\windows\history\" & IEHistoryFile For Binary As #1
  sData = Space$(LOF(1)) 'Data Buffer
  Get #1, , sData  'Places all data from file into buffer , sData
  Close #1  'Closes file
  i = InStr(i + 1, sData, sDelimiter) 'Looks for sdelimiter in sdata
  iURLCount = 0 'Sets URLCount to 0 because this is the beginning for the file
  While i < Len(sData)
   iURLCount = iURLCount + 1  'Keeps a count of how manu URLs are in the file
   If i > 0 Then
    j = InStr(i + iDisplacement - 1, sData, Chr$(0))
    'Place URL in an array
    sURLs(iURLCount) = Mid$(sData, i + iDisplacement, j - (i + iDisplacement))
   End If
   i = InStr(i + 1, sData, sDelimiter) 'Index = URL
   If i = 0 Then GoTo EndURLs 'If there are no more URLs then stop looping
  Wend
EndURLs:
  'This prints all URLs in Array in the debug window so you can see them
  For x = 1 To iURLCount
    Debug.Print sURLs(x)
  Next x
End Sub
```

