VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   8370
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvDB 
      Height          =   4815
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8493
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Default         =   -1  'True
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "fMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'© 2000-2007 SCINER
'Created by SCINER: lenar2003@mail.ru
'03/11/2006 20:41
'Custom DataBase Engine v1.0

Dim db As New cMobileDB

Function RCP(ByVal P As String) As String
  RCP = P & IIf(VBA.Right$(P, 1) = "\", vbNullString, "\")
End Function

Private Sub cmdTest_Click()

  Const szTableName As String = "TestTable"
  Const szFiledNum As String = "Number"
  Const szFiledRandom As String = "Random number"
  Const szFiledComment As String = "Comment"

  Dim dwTable As Long
  Dim dwFiledNum As String
  Dim dwFiledRandom As String
  Dim dwFiledComment As String

  Dim lRow As Long
  Dim lField As Long
  Dim v
  Dim i As Long
  Dim z As Long
  Dim capcount As Long
  Dim lItem As ListItem

  'set database file
  db.Filename = RCP(App.Path) & "db.sbs"

'-------------------------------------------------------
'write to database table
'-------------------------------------------------------

  'add table
  dwTable = db.addTable(szTableName, mNormal)
  'get table offset
  dwTable = db.getTableOffset(szTableName)

  'if table exists
  If dwTable > 0 Then
    'add fields
    dwFiledNum = db.addField(dwTable, szFiledNum, fldLong)
    dwFiledRandom = db.addField(dwTable, szFiledRandom, fldDouble)
    dwFiledComment = db.addField(dwTable, szFiledComment, fldString)
  End If

  'get fields offset
  dwFiledNum = db.getFieldOffset(dwTable, szFiledNum)
  dwFiledRandom = db.getFieldOffset(dwTable, szFiledRandom)
  dwFiledComment = db.getFieldOffset(dwTable, szFiledComment)

  Call Randomize(Timer)
  
  'add items to table
  For i = 0 To 99
    'get new row offset
    lRow = db.Add(dwTable)
    'set values
    db.Value(dwTable, dwFiledNum, lRow) = i
    db.Value(dwTable, dwFiledRandom, lRow) = Rnd
    db.Value(dwTable, dwFiledComment, lRow) = "Created by SCINER: lenar2003@mail.ru #" & CStr(i)
    'slow speed
    'Call db.Update
  Next
  
  'for flush db records to file
  Call db.Update
  
'-------------------------------------------------------
'read from table from database
'-------------------------------------------------------

  Call lvDB.ColumnHeaders.Clear
  Call lvDB.ListItems.Clear

  'get table offset
  dwTable = db.getTableOffset(szTableName)

  'get all table fields
  lField = db.getFirstField(dwTable)
  Do While lField > 0
    Call lvDB.ColumnHeaders.Add(, , db.getFieldName(lField))
    lField = db.getNextField(lField)
  Loop

  With db
    'get first table row offset
    lRow = db.getFirstTableRow(dwTable)
    'enumerate all rows in table
    Do While lRow > 0
      'get first table field offset
      lField = .getFirstField(dwTable)
      Set lItem = lvDB.ListItems.Add(, , .Value(dwTable, lField, lRow))
      i = 1
      'get next table field
      lField = .getNextField(lField)
      Do While lField > 0
        v = Empty
        'get field type
        Select Case .getFieldType(lField)
        Case fldByteArray: v = "<BINARY>"
        Case fldJpegFileBytes: v = "<BINARY>"
        Case Else: v = .Value(dwTable, lField, lRow)
        End Select
        If IsEmpty(v) Then v = "NULL"
        lItem.SubItems(i) = v
        i = i + 1
        'get next table field
        lField = .getNextField(lField)
      Loop
      'go to next row
      lRow = db.getNextRow(lRow)
    Loop
  End With

End Sub
