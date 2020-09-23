VERSION 5.00
Begin VB.Form AddEdit 
   BorderStyle     =   0  'None
   Caption         =   "  Editor"
   ClientHeight    =   2670
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4305
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command7 
      BackColor       =   &H8000000A&
      Caption         =   "Change"
      Height          =   375
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2040
      Width           =   915
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H8000000A&
      Height          =   525
      Left            =   3420
      Picture         =   "AddEdit.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1410
      Width           =   465
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H8000000A&
      Height          =   525
      Left            =   3420
      Picture         =   "AddEdit.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   465
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000A&
      Height          =   405
      Left            =   3690
      Picture         =   "AddEdit.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   300
      Width           =   525
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000A&
      Height          =   405
      Left            =   3060
      Picture         =   "AddEdit.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   300
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   330
      Width           =   2625
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "Delete"
      Height          =   375
      Left            =   1650
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "Add/Insert"
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   915
   End
   Begin VB.Label Label8 
      Caption         =   "00"
      Height          =   225
      Left            =   1470
      TabIndex        =   14
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "of"
      Height          =   225
      Left            =   1200
      TabIndex        =   13
      Top             =   1200
      Width           =   225
   End
   Begin VB.Label Label6 
      Caption         =   " 1"
      Height          =   225
      Left            =   870
      TabIndex        =   12
      Top             =   1200
      Width           =   285
   End
   Begin VB.Label Label5 
      Caption         =   "Row"
      Height          =   225
      Left            =   270
      TabIndex        =   11
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label Label4 
      Caption         =   "00"
      Height          =   225
      Left            =   1470
      TabIndex        =   6
      Top             =   870
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "of"
      Height          =   225
      Left            =   1200
      TabIndex        =   5
      Top             =   870
      Width           =   225
   End
   Begin VB.Label Label2 
      Caption         =   " 1"
      Height          =   225
      Left            =   870
      TabIndex        =   4
      Top             =   870
      Width           =   285
   End
   Begin VB.Label Label1 
      Caption         =   "Column"
      Height          =   225
      Left            =   270
      TabIndex        =   3
      Top             =   870
      Width           =   555
   End
   Begin VB.Menu ExitMnu 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "AddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rowcount%
Dim t$(100)
Private Sub Command1_Click()
'                                 The ADD/Insert Button
If colpos% = 1 Then
Form1.ListView1.ListItems.Add(rowpos%) = Text1
rowcount% = rowcount% + 1
Label8 = Str(rowcount%)
Else
tmpp$ = Form1.ListView1.ListItems(rowpos%).Text ' First column text
Form1.ListView1.ListItems(rowpos%).ListSubItems.Add(colpos% - 1, , Text1) = Text1
End If
Label4 = Str(Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1) '# of columns
Form1.ListView1.Refresh

End Sub
Private Sub Command2_Click()
' The Delete Command Button
If rowpos% > rowcount% Then Beep: Exit Sub
If colpos% > Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 Then Beep: Exit Sub 'columns in this row
If colpos% = 1 Then
Form1.ListView1.ListItems.Remove (rowpos%) 'remove the row
rowcount% = rowcount% - 1
Label8 = Str(rowcount%)
Else ' we have to delete it and then add it again without the selected column
 y% = Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 ' save the current columns
 For x% = 1 To y% ' store the data in the current columns
   If x% = 1 Then
      t$(x%) = Form1.ListView1.ListItems(rowpos%).Text ' First column text
   Else
      t$(x%) = Form1.ListView1.ListItems(rowpos%).ListSubItems(x% - 1).Text
   End If
Next x%
Form1.ListView1.ListItems.Remove (rowpos%) 'remove the entire row now anyway
For x% = 1 To y% ' we now add the row excluding the selected column to be deleted
 If x% = 1 Then
  Form1.ListView1.ListItems.Add(rowpos%) = t$(x%)
 Else
  ' (we enter the mod with q%=0)
  If x% <> colpos% Then q% = q% + 1: Form1.ListView1.ListItems(rowpos%).ListSubItems.Add(q%, , t$(x%)) = t$(x%)
 End If
Next x%
End If

If rowpos% > rowcount% Then Text1 = "": Text2 = "": Exit Sub
tp% = Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1: Label4 = Str(tp%) 'cur cols
If colpos% > tp% Then
colpos% = tp%
Label6 = Str(tp%): Label8 = Label6
End If
Label2 = Str(colpos%)
If colpos% = 1 Then
Text1 = Form1.ListView1.ListItems(rowpos%).Text ' First column text
Else
Text1 = Form1.ListView1.ListItems(rowpos%).ListSubItems(colpos% - 1).Text
End If

End Sub
Private Sub Command3_Click()

If colpos% > 1 Then
colpos% = colpos% - 1
Label2 = Str(colpos%)
  If colpos% = 1 Then
     Text1 = Form1.ListView1.ListItems(rowpos%).Text ' First column text
  Else
     Text1 = Form1.ListView1.ListItems(rowpos%).ListSubItems(colpos% - 1).Text
  End If
End If

End Sub
Private Sub Command4_Click()

If Form1.ListView1.ListItems.Count = 0 Then Beep: Exit Sub
If rowpos% > Form1.ListView1.ListItems.Count Then Beep: Exit Sub
If colpos% <= (Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1) Then
 colpos% = colpos% + 1
 Label2 = Str(colpos%)
    If colpos% > Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 Then
       Text1 = ""
       Exit Sub
    End If
   If colpos% = 1 Then
    Text1 = Form1.ListView1.ListItems(rowpos%).Text ' First column text
   Else
    Text1 = Form1.ListView1.ListItems(rowpos%).ListSubItems(colpos% - 1).Text
   End If
   
End If

End Sub
Private Sub Command5_Click()

If rowpos% > 1 Then
rowpos% = rowpos% - 1
 If colpos% > Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 Then
    colpos% = Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 'columns in this row
 End If
Label2 = Str(colpos%) 'column were currently on
Label4 = Str(Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1)
Label6 = Str(rowpos%)
If colpos% = 1 Then
   Text1 = Form1.ListView1.ListItems(rowpos%).Text ' First column text
Else
         If colpos% > Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 Then
            colpos% = Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1
            Label2 = Str(Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1)
         End If
   Text1 = Form1.ListView1.ListItems(rowpos%).ListSubItems(colpos% - 1).Text
End If

End If

End Sub
Private Sub Command6_Click()

If rowpos% <= rowcount% Then
rowpos% = rowpos% + 1
Label6 = Str(rowpos%) ' display the row position
If rowpos% > Form1.ListView1.ListItems.Count Then
colpos% = 1
Label2 = "1" 'column were currently on
Label4 = "0" 'currently there are no columns at this row position
Text1 = "" ' no text to display either
Exit Sub ' no row to display
End If
If rowpos% > Form1.ListView1.ListItems.Count Then
rowpos% = 1: Label6 = "1"
End If
 If colpos% > Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 Then
   colpos% = Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 'columns in this row
 End If
Label4 = Str(Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1) 'current columns

 If colpos% = 1 Then
    Text1 = Form1.ListView1.ListItems(rowpos%).Text ' First column text
 Else
          If colpos% > Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1 Then
             colpos% = Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1
             Label2 = Str(Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1)
          End If
    Text1 = Form1.ListView1.ListItems(rowpos%).ListSubItems(colpos% - 1).Text
 End If

End If

End Sub
Private Sub Command7_Click()
' Change Button
If colpos% = 1 Then
If Form1.ListView1.ListItems.Count < rowpos% Then Beep: Exit Sub
Form1.ListView1.ListItems(rowpos%).Text = Text1.Text
Else
If Form1.ListView1.ListItems(rowpos%).ListSubItems.Count < colpos% - 1 Then Beep: Exit Sub
Form1.ListView1.ListItems(rowpos%).ListSubItems(colpos% - 1).Text = Text1.Text
End If

End Sub
Private Sub ExitMnu_Click()

Unload Me

End Sub
Private Sub Form_Load()

'If colpos% = 0 Then colpos% = 1
'If rowpos% = 0 Then rowpos% = 1
colpos% = 1: rowpos% = 1

'Form1.ListView1.ListItems.Clear
'Form1.ListView1.ColumnHeaders.Add 1, , "Last Name"
'Form1.ListView1.ColumnHeaders.Add 2, , "First Name"
'Form1.ListView1.ColumnHeaders.Add 3, , "Middle Initial"
'Form1.ListView1.ColumnHeaders.Add 4, , "4th Column"

'Set lvi = Form1.ListView1.ListItems.Add(, "Aitken", "Aitken")
'lvi.SubItems(1) = "Peter"
'lvi.SubItems(2) = "G."

'Form1.ListView1.ListItems.Add , "Clinton", "Clinton"
'Set lvi = Form1.ListView1.ListItems("Clinton")
'lvi.SubItems(1) = "William"
'lvi.SubItems(2) = "J."
'Form1.ListView1.ListItems.Add , "Edison", "Edison"
'Form1.ListView1.ListItems("Edison").SubItems(1) = "Thomas"
'Form1.ListView1.ListItems("Edison").SubItems(2) = "A."

'Form1.ListView1.ListItems.Add , "He", "He"
'Form1.ListView1.ListItems("He").SubItems(1) = "Bobby"
'Form1.ListView1.ListItems("He").SubItems(2) = "Q."
'Form1.ListView1.ListItems("He").SubItems(3) = "LastOne"

On Error Resume Next
' Columns in selected row .... X
Label4 = Str(Form1.ListView1.ListItems(rowpos%).ListSubItems.Count + 1) ' first item + sub-items
'           Rows in list ..... Y
rowcount% = Form1.ListView1.ListItems.Count
Label8 = Str(rowcount%)

Text1 = Form1.ListView1.ListItems(rowpos%).Text
On Error GoTo 0

End Sub
