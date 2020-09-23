VERSION 5.00
Begin VB.Form OptForm2 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Column Heading Editor"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   1650
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Update Width"
      Height          =   315
      Left            =   2340
      TabIndex        =   12
      Top             =   1260
      Width           =   1365
   End
   Begin VB.TextBox Text2 
      Height          =   255
      Left            =   1350
      TabIndex        =   11
      Top             =   1290
      Width           =   825
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1350
      TabIndex        =   10
      Top             =   750
      Width           =   825
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Insert"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3030
      TabIndex        =   9
      Top             =   750
      Width           =   825
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2190
      TabIndex        =   6
      Top             =   750
      Width           =   825
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   3990
      Picture         =   "OptForm2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3990
      Picture         =   "OptForm2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   465
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1350
      TabIndex        =   0
      Top             =   390
      Width           =   2475
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Current Width :"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Select"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3900
      TabIndex        =   8
      Top             =   1380
      Width           =   615
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Columns"
      Height          =   225
      Left            =   420
      TabIndex        =   7
      Top             =   810
      Width           =   645
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "0"
      Height          =   225
      Left            =   30
      TabIndex        =   5
      Top             =   810
      Width           =   315
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Column"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3900
      TabIndex        =   4
      Top             =   1170
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Column 1 Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   1275
   End
   Begin VB.Menu ExitMnu 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "OptForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim col As Integer
Dim curcol As Integer
Private Sub Command1_Click()

If col < curcol + 1 Then ' Selection is valid or past the end by 1
col = col + 1
Label1 = "Column" + Str(col) + " Text $"
If curcol < col Then Text1 = "": Text2 = "": Exit Sub
Text1 = Form1.ListView1.ColumnHeaders.Item(col) ' display header text
Text2 = Str(Form1.ListView1.ColumnHeaders.Item(col).Width)
End If

End Sub
Private Sub Command2_Click()

If col > 1 Then
col = col - 1
Label1 = "Column" + Str(col) + " Text $"
Text1 = Form1.ListView1.ColumnHeaders.Item(col) ' display header text
Text2 = Str(Form1.ListView1.ColumnHeaders.Item(col).Width)
End If

End Sub
Private Sub Command3_Click()

If col > curcol Then Beep: Exit Sub
Form1.ListView1.ColumnHeaders.Remove (col) 'remove the column
curcol = curcol - 1
Label5 = Str(curcol)

If curcol > 0 And curcol >= col Then
Text1 = Form1.ListView1.ColumnHeaders.Item(col) ' display header text
Else
Text1 = "": Text2 = ""
End If


End Sub
Private Sub Command4_Click()

Form1.ListView1.ColumnHeaders.Add(col).Text = Text1
curcol = curcol + 1
Label5 = Str(curcol)
Text2 = Str(Form1.ListView1.ColumnHeaders.Item(col).Width)

End Sub
Private Sub Command5_Click()

If curcol < col Then Beep: Exit Sub
Form1.ListView1.ColumnHeaders.Remove (col) 'remove the column
Set LVHeader = Form1.ListView1.ColumnHeaders.Add(col)
If Val(Text2) = 0 Then Text2 = Str(Form1.ListView1.ColumnHeaders.Item(col).Width)
LVHeader.Text = Text1 ' restore the text
LVHeader.Width = Val(Text2) ' and set the selected width to item

End Sub
Private Sub Command6_Click()

Form1.ListView1.ColumnHeaders.Remove (col) 'remove the column
Set LVHeader = Form1.ListView1.ColumnHeaders.Add(col)
LVHeader.Text = Text1 ' restore the text
LVHeader.Width = Val(Text2) ' and set the selected width to item
Text2 = Str(Form1.ListView1.ColumnHeaders.Item(col).Width)

End Sub
Private Sub ExitMnu_Click()

Unload OptForm2

End Sub
Private Sub Form_Load()

col = 1 ' start at the first heading
curcol = Form1.ListView1.ColumnHeaders.Count
If curcol > 0 Then
Text1 = Form1.ListView1.ColumnHeaders.Item(1) 'display header text
Text2 = Str(Form1.ListView1.ColumnHeaders.Item(1).Width)
Else
' No header information is in the listview at the current time
End If
Label5 = Str(curcol)

End Sub
Private Sub Form_Unload(Cancel As Integer)

headcount% = Form1.ListView1.ColumnHeaders.Count

End Sub
