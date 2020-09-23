VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Sample 2"
      Height          =   630
      Left            =   1230
      TabIndex        =   1
      Top             =   1215
      Width           =   2340
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sample 1"
      Height          =   660
      Left            =   1230
      TabIndex        =   0
      Top             =   435
      Width           =   2355
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------------------------------
' Code By Sudu_Sudu@hotmail.com
' Name : Muhamad Hanafiah Yahya
' Require
'            Microsoft Active X DataObject 2.7
'--------------------------------------------------------------------------
Dim MyCon As New ADODB.Connection
Dim MyRs As New ADODB.Recordset

Private Sub Command1_Click()
'------------------------------------------------------------------------------------------
'This sample to display Report direct from database without any additional
' operation like add data, combine and etc
'------------------------------------------------------------------------------------------

'To set Textbox Datafield
With DataReport1.Sections("Section1").Controls 'section1 mean that section you create in datareport
    .Item("txtFamily").DataField = MyRs("HouseholdName").Name
    .Item("txtAddress").DataField = MyRs("Address").Name
End With

'To Set Label caption
With DataReport1.Sections("Section2").Controls
    .Item("lblFamily").Caption = "Family Name"
    .Item("lblAddress").Caption = "Address"
End With

With DataReport1.Sections("Section4").Controls
    .Item("lblTitle").Caption = "My Family Name - Sample 1"
End With

'to set datasource for datareport
Set DataReport1.DataSource = MyRs

'show datareport
DataReport1.Show
End Sub

Private Sub Command2_Click()
'------------------------------------------------------------------------------------------
'This sample to display Report from database with any additional
' operation like add data, combine and etc
'------------------------------------------------------------------------------------------
' Create adodb record set
Dim intCount As Integer
Dim strAddress As String
Dim TempRS As ADODB.Recordset

Set TempRS = New ADODB.Recordset
TempRS.Fields.Append "tmpFamily", adVarChar, 30
TempRS.Fields.Append "tmpAddress", adVarChar, 100
TempRS.Open
MyRs.MoveFirst 'set to first record

For intCount = 1 To MyRs.RecordCount
    strAddress = MyRs("Address") & " , " & MyRs("city") & ", " & MyRs("StateOrProvince")
    TempRS.AddNew Array("tmpFamily", "tmpAddress"), Array(intCount & " " & MyRs("HouseholdName"), strAddress)
    MyRs.MoveNext
Next intCount

'To set Textbox Datafield
With DataReport1.Sections("Section1").Controls 'section1 mean that section you create in datareport
    .Item("txtFamily").DataField = TempRS("tmpFamily").Name
    .Item("txtAddress").DataField = TempRS("tmpAddress").Name
End With

'To Set Label caption
With DataReport1.Sections("Section2").Controls
    .Item("lblFamily").Caption = "Family Name"
    .Item("lblAddress").Caption = "Address"
End With

'set report title
With DataReport1.Sections("Section4").Controls
    .Item("lblTitle").Caption = "My Family Name - Sample 2"
End With

'to set datasource for datareport
Set DataReport1.DataSource = TempRS

'show datareport
DataReport1.Show
End Sub

Private Sub Form_Load()
Dim strPath As String

strPath = App.Path & "\database\ADDRBOOK.mdb"
 
'Set connection ke database ( strpath )
MyCon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                         "Data Source=" & strPath
MyCon.Open

'Open The The Recordset
MyRs.ActiveConnection = MyCon

'Open sebagai Keyset, dan LockOptimistic
MyRs.Open "HouseHold", MyCon, adOpenKeyset, adLockOptimistic, adCmdTableDirect
End Sub

Private Sub Form_Unload(Cancel As Integer)
MyRs.Close
Set MyCon = Nothing
End Sub
