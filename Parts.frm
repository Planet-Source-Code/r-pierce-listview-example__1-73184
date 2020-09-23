VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "     ListView example/ turtorial in Report Mode by   Rp"
   ClientHeight    =   4935
   ClientLeft      =   150
   ClientTop       =   315
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9660
      Top             =   -270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4395
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   7752
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu FileMnu 
      Caption         =   "File"
      Begin VB.Menu OpenMnu 
         Caption         =   "Open"
      End
      Begin VB.Menu SaveMnu 
         Caption         =   "Save"
      End
      Begin VB.Menu ExitMnu 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu SettingsMnu 
      Caption         =   "Editors"
      Begin VB.Menu AddEditMnu 
         Caption         =   "Add/Edit"
      End
      Begin VB.Menu HeadingsMnu 
         Caption         =   "Headings"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is just another ListView example I did a few years ago
' while messing around with the control.
' Note that you can click on the default field once you have one
' to edit the field if you want to.
' Step thru the fields, delete some and see how the control works.
' The retarded editor provided is for an example or tutorial !
' Rp ....   Feel free to distribute this anywhere.  It's public domain now!
Private Sub AddEditMnu_Click()

AddEdit.Show 1

End Sub
Private Sub ExitMnu_Click()

Unload Me

End Sub
Private Sub Form_Load()

ListView1.ColumnHeaders.Clear ' Initialize the Headers to Nothing!

End Sub
Private Sub HeadingsMnu_Click()
OptForm2.Show 1
End Sub
Private Sub OpenMnu_Click()

On Error GoTo Cancel
CommonDialog1.InitDir = App.Path
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.Filter = "Specialty Coin Mfg files  (*.scm)|*.scm"
CommonDialog1.ShowOpen
ListView1.ColumnHeaders.Clear
ListView1.ListItems.Clear
filenum% = FreeFile ' Get a file number not being used
Open CommonDialog1.FileName For Input As #filenum%
Input #filenum%, headcount%
For x% = 1 To headcount% ' Save heading names
Set LVHeader = ListView1.ColumnHeaders.Add(x%)
Input #filenum%, Txt$
LVHeader.Text = Txt$
Input #filenum%, tmp%
LVHeader.Width = tmp%
Next x% '  Headings have been loaded
Input #filenum%, rows%
If rows% = 0 Then ' Only headers to load
Close #filenum%
Exit Sub
End If
For x% = 1 To rows%
Input #filenum%, cols%
For y% = 1 To cols%
If y% = 1 Then
Input #filenum%, tmpp$
ListView1.ListItems.Add(x%) = tmpp$
Else
Input #filenum%, Txt$
ListView1.ListItems(x%).ListSubItems.Add(y% - 1, , Txt$) = Txt$
End If
Next y%
Next x% ' ListView has been loaded via disk file
Close #filenum%

Cancel:

End Sub
Private Sub SaveMnu_Click()
' Save the listview control to disk
If ListView1.ColumnHeaders.Count = 0 Then Exit Sub
CommonDialog1.CancelError = True
On Error GoTo Cancel
CommonDialog1.Flags = cdlOFNHideReadOnly
CommonDialog1.InitDir = App.Path
CommonDialog1.Filter = "Specialty Coin Mfg files (*.scm)"
CommonDialog1.DefaultExt = "scm"
CommonDialog1.ShowSave
headcount% = ListView1.ColumnHeaders.Count ' Save the number of column headers
filenum% = FreeFile '  Get a file number not being used
Open CommonDialog1.FileName For Output As #filenum%
' Save the Heading number and names of each column in our ListView
Write #filenum%, headcount%
For x% = 1 To headcount% ' Save heading names and column widths
Write #filenum%, ListView1.ColumnHeaders.Item(x%) '  save heading
Write #filenum%, ListView1.ColumnHeaders.Item(x%).Width ' and width
Next x% '                  Headings have been saved
rows% = ListView1.ListItems.Count ' # or rows
Write #filenum%, rows% ' Save the number of rows in listview control
If rows% = 0 Then ' Only headers to save
Close #filenum%
Exit Sub
End If
For x% = 1 To rows%
cols% = ListView1.ListItems(x%).ListSubItems.Count + 1 '# of columns
Write #filenum%, cols% ' Save the number of columns in the current row
For y% = 1 To cols%
If y% = 1 Then
Write #filenum%, ListView1.ListItems(x%).Text ' First column
Else
Write #filenum%, ListView1.ListItems(x%).ListSubItems(y% - 1).Text ' > 1st column
End If
Next y%
Next x% ' Listview contents have been saved
Close #filenum%

Cancel:
End Sub
