<div align="center">

## Creating Program Manager Groups & Icons


</div>

### Description

Create program manager groups and icons from your code!
 
### More Info
 
It requires 3 arguments to be passed to it. They are:

1.The form that contains Label1 (x) 2.A string variable containing the group's name (GroupName$) 3.A string variable containing the path to the group (*.GRP) file (GroupPath$)


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Tips and Source Code](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-tips-and-source-code.md)
**Level**          |Unknown
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-tips-and-source-code-creating-program-manager-groups-icons__1-163/archive/master.zip)





### Source Code

```
The first step is to add a label to a form. This example assumes you are using a label named "Label1". This label will be used in the DDE conversation between Program Manager and your proram. This example contains two SUBs. Both are placed into a BAS module. The first SUB creates the Program Manager Group, and the second SUB creates an icon within that group. These SUBs are called independantly (to allow for flexibility and clarity of illustration).
The following SUB creates the Program Manager group. It requires 3 arguments to be passed to it. They are:
1.The form that contains Label1 (x) 2.A string variable containing the group's name (GroupName$) 3.A string variable containing the path to the group (*.GRP) file (GroupPath$)
Sub CreateProgManGroup (x As Form, GroupName$, GroupPath$)
  Dim i%, z%        'Declare required working variables
  Screen.MousePointer = 11 'hourglass mousepointer while working
  On Error Resume Next   'Not good to have program crash :-)
  ' Set LinkTopic & LinkMode parameters
  x.Label1.LinkTopic = "ProgMan|Progman"
  x.Label1.LinkMode = 2
  For i% = 1 To 10     ' Give the DDE process time to take place
   z% = DoEvents()
  Next
  x.Label1.LinkTimeout = 100
  ' Actually create the group now
  x.Label1.LinkExecute "[CreateGroup(" + GroupName$ + Chr$(44) + GroupPath$ + ")]"
  ' Reset label properties and mousepointer
  x.Label1.LinkTimeout = 50
  x.Label1.LinkMode = 0
  Screen.MousePointer = 0
End Sub
The following SUB creates the Program Manager icon. It requires 3 arguments to be passed to it. They are:
1.The form that contains Label1 (x) 2.A string variable containing the icon's Command Line (CmdLine$) 3.A string variable containing the icon's Caption (IconTitle$)
Sub CreateProgManItem (x As Form, CmdLine$, IconTitle$)
  Dim i%, z%        'Declare required working variables
  Screen.MousePointer = 11 'hourglass mousepointer while working
  On Error Resume Next   'Not good to have program crash :-)
  ' Set LinkTopic & LinkMode parameters
  x.Label1.LinkTopic = "ProgMan|Progman"
  x.Label1.LinkMode = 2
  For i% = 1 To 10     ' Give the DDE process time to take place
   z% = DoEvents()
  Next
  x.Label1.LinkTimeout = 100
  x.Label1.LinkExecute "[AddItem(" + CmdLine$ + Chr$(44) + IconTitle$ + Chr$(44) + ",,)]"
  ' Reset label properties and mousepointer
  x.Label1.LinkTimeout = 50
  x.Label1.LinkMode = 0
  Screen.MousePointer = 0
End Sub
Finally, the last thing you need is for an event procedure (or any other form level routine) to call the 2 SUBs and provide the necessary information. In this example, I am creating a group window called VB Library and am placing it into the Windows directory. Then, I am creating an icon called "VB Library" within the group. This example creates an icon for the currently running program which happens to be Library.EXE.
' Refer to Tips 23 and 24 for obtaining the Windows Directory
CreateProgManGroup Me, "VB Library", "c:\windows"
CreateProgManItem Me, app.Path + "\library", "VB Library"
A little side note here. Thanks to Microsoft making Windows 95 backward-compatible, this routine runs fine within it. The group file will appear as an entry in the Start Menu's Programs section and the icon will be a sub-menu of that entry.
```

