<div align="center">

## Copying List Items to the Clipboard


</div>

### Description

This code is a demonstration of a quick and easy way to copy an entire list (without needing to select them all) or selected list items to the clipboard for pasting into other applications. It simply copies the items to a hidden text box and then uses the clipboard method. Follow the instructions below to see how it works
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David J Jenkins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-j-jenkins.md)
**Level**          |Unknown
**User Rating**    |4.2 (164 globes from 39 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-j-jenkins-copying-list-items-to-the-clipboard__1-1197/archive/master.zip)

### API Declarations

none


### Source Code

```
'*******************************
'demonstration on how to copy
'an entire list or selected
'list items to the clipboard
'for use in other apps.
'to see how it works, do the
'following:
'1. open a new project
'2. put a listbox on the form,
'  name it lstList and set its
'  MultiSelect value to 2
'3. put a command button on
'  the form and call it
'  cmdCopyList
'4. put another command button
'  on the form and call it
'  cmdCopyListItems.
'5. put a textbox on the form,
'  call it txtHidden, and set
'  its visible property to false.
'6. paste the code into the
'  code window, run, and test.
'  Be sure to select some items
'  before you choose
'  copy list items.
'
'******************************
Private Sub Form_Load()
   'add rainbow colors to list box
   lstList.AddItem "Red"
   lstList.AddItem "Orange"
   lstList.AddItem "Yellow"
   lstList.AddItem "Green"
   lstList.AddItem "Blue"
   lstList.AddItem "Indigo"
   lstList.AddItem "Violet"
End Sub
Private Sub cmdCopyList_Click()
'this procedure loops thru the list
'and copies each item to a textbox
   Dim I As Integer
   For I = 0 To lstList.ListCount - 1
     txtHidden.Text = txtHidden.Text & lstList.List(I) & vbCrLf
   Next I
   Call CopyText
End Sub
Private Sub cmdCopyListItems_Click()
'copy list item to textbox
   Dim I As Integer
   For I = 0 To lstList.ListCount - 1
     If lstList.Selected(I) Then
        txtHidden.Text = txtHidden.Text & lstList.List(I) & vbCrLf
     End If
   Next I
   Call CopyText
End Sub
Public Sub CopyText()
'select list and copy
'to clipboard
   txtHidden.SelLength = Len(txtHidden.Text)
   Clipboard.Clear
   Clipboard.SetText txtHidden.SelText
End Sub
```

