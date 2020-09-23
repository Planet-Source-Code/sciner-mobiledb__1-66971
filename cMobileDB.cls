VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMobileDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'© 2000-2007 SCINER
'Created by SCINER: lenar2003@mail.ru
'03/11/2006 20:41
'Custom DataBase Engine v1.0

'ìåõàíèçì âåäåíèÿ áàçû äàííûõ ñâîåãî ôîðìàòà v0.2.1
'19/10/2006 16:47

'âîçìîæíîñòè:
'1. íåîãðàíè÷åííîå ÷èñëî òàáëèö
'2. íåîãðàíè÷åííîå ÷èñëî ïîëåé â òàáëèöàõ
'3. íåîãðàíè÷åííîå ñè÷ëî çàïèñåé â òàáëèöå
'4. ðàöèîíàëüíîå èñïîëüçîâàíèå ïàìÿòè â ôàéëå
'   (ïîëå ôèçè÷åñêè íå ñóùåñòâóåò äî òåõ ïîð,
'   ïîêà â íåãî ÷òî íèáóäü íå çàïèñàëè)
'5. ïîääåðæêà 8 òèïîâ äàííûõ, âêëþ÷àÿ JpegFileBytes
'6. ññûëî÷íàÿ ñòðóêòóðà, îáåñïå÷èâàþùàÿ âûñîêóþ ñêîðîñòü îáðàáîòêè áàçû
'7. òèï string ÿâëÿåòñÿ áóôåðîì ïåðåìåííîãî ðàçìåðà aka memo

'ðåàëèçîâàíî
'1. ñîçäàíèå òàáëèö ñ ïðîèçâîëüíûì íàçâàíèåì äëèíîé äî 256 áàéò
'2. ñîçäàíèå ïîëåé ñ ïðîèçâîëüíûì íàçâàíèåì äëèíîé äî 256 áàéò è ñ âûáðàííûì òèïîì
'3. äîáàâëåíèå çàïèñåé
'4. çàïèñü/÷òåíèå ïîëåé

'áëàãîäàðÿ ññûëî÷íîé ñòðóêòóðå ëåãêî ðåàëèçîâàòü:
'1. óäàëåíèå òàáëèöû, ïîëÿ, çàïèñè
'2. ôðàãìåíòàöèþ ôàéëà áàçû äàííûõ

Const dbsign = 23658740
Const dbheadlength = 264
Const datasign = 255
Const rowsign = 12548

Event ExportProgress(ByVal lMax As Long, ByVal lValue As Long)
Event ExportComplete()

Public Enum MobDbAttributes
  mSystem = vbSystem
  mReadOnly = vbReadOnly
  mNormal = vbNormal
  mHidden = vbHidden
  mAlias = vbAlias
End Enum

Public Enum DBFieldType
  fldString = 1
  fldLong = 2
  fldDouble = 4
  fldByte = 8
  fldDate = 16
  fldInteger = 32
  fldByteArray = 64
  fldJpegFileBytes = 128
  'fld = 256
  'fld = 512
  'fld = 1024
  'fld = 2048
  'fld = 4196
  'fld = 8192
  'fld = 16384
  'fld = 32768
End Enum

Private Type data_type_2
  size As Long
  sizereal As Long
  next As Long
  lOffset As Long
  fieldoffset As Long
  sign As Byte
  Type As Integer
End Type

Private Type row_type_2
  sign As Integer
  reserved(0 To 1) As Byte
  Data As Long
  previous As Long
  next As Long
End Type

Private Type field_type_2
  sign As Long
  Type As Integer
  szName(255) As Byte
  next As Long
End Type

Private Type table_type_2
  sign As Long
  record As Long
  lastrecord As Long
  field As Long
  lastfield As Long
  next As Long
  recordcount As Long
  szName(127) As Byte
  attributes As Integer
  reserved(124) As Byte
End Type

Private Type last_update_type_2
  tableoffset As Long
  rowoffset As Long
  fieldoffset As Long
  datetime As Date
End Type

Private Type db_head_type_2
  sign As Long
  table As Long
  lasttable As Long
  lastupdate As last_update_type_2
  szName(63) As Byte
  uniqueid As Long 'not used
  reserved(187) As Byte
End Type

Dim db As db_head_type_2

Private Type fastfieldindex
  Key As String
  lOffset As Long
  Type As DBFieldType
End Type

Private Type fastindex
  Key As String
  lOffset As Long
  Fields() As fastfieldindex
  fieldscount As Long
End Type

Dim tables() As fastindex
Dim ltablecountarraysize As Long
Dim ltablecount As Long
Dim FF As Integer
Dim lpFileName As String

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Function getLastUpdate(dtDate As Date, lTableOffset As Long, lRowOffset As Long)
  Dim tempdb As db_head_type_2
  If FF <> 0 Then
    Get #FF, 1, tempdb
    dtDate = tempdb.lastupdate.datetime
    lTableOffset = tempdb.lastupdate.tableoffset
    lRowOffset = tempdb.lastupdate.rowoffset
  End If
End Function

'new--------
Function getFirstTable() As Long
  If db.sign = dbsign Then getFirstTable = db.table
End Function
Function getTableName(ByVal lTableOffset As Long) As String
  Dim i As Long
  For i = 0 To ltablecount - 1
    If tables(i).lOffset = lTableOffset Then
      getTableName = tables(i).Key
      Exit Function
    End If
  Next
End Function
Function getNextTable(ByVal lTableOffset As Long) As Long
  Dim table As table_type_2
  Get #FF, lTableOffset, table
  If table.sign <> dbsign Then Exit Function
  If table.next < 1 Then Exit Function
  lTableOffset = table.next
  Get #FF, lTableOffset, table
  If table.sign = dbsign Then getNextTable = lTableOffset
End Function
Function getFieldType(ByVal lFieldOffset As Long) As DBFieldType
  Dim field As field_type_2
  If lFieldOffset < 1 Then Exit Function
  Get #FF, lFieldOffset, field
  If field.sign = dbsign Then getFieldType = field.Type
End Function

Function RowExists(ByVal lOffset As Long) As Boolean

  Dim row As row_type_2
  Dim Data As data_type_2

  If lOffset < 1 Then Exit Function
  
  Get #FF, lOffset, row
  
  If row.sign = rowsign Then
    RowExists = True
  Else
    If row.Data > 0 Then
      Get #FF, row.Data, Data
      RowExists = Data.sign = datasign
    End If
  End If

End Function

Function getFirstTableRow(ByVal lOffset As Long) As Long
  Dim table As table_type_2
  Get #FF, lOffset, table
  getFirstTableRow = table.record
End Function

Function getLastTableRow(ByVal lOffset As Long) As Long
  Dim table As table_type_2
  Get #FF, lOffset, table
  getLastTableRow = table.lastrecord
End Function

Function getNextRow(ByVal lOffset As Long) As Long
  Dim row As row_type_2
  Get #FF, lOffset, row
  getNextRow = row.next
End Function

Function getPreviousRow(ByVal lOffset As Long)
  Dim row As row_type_2
  Get #FF, lOffset, row
  getPreviousRow = row.previous
End Function

Sub Update()
  If FF <> 0 Then
    Put #FF, 1, db
  End If
  Call flushDB
End Sub

Property Get recordcount(ByVal lTableOffset As Long)
  Dim table As table_type_2
  Get #FF, lTableOffset, table
  If table.sign = dbsign Then recordcount = table.recordcount
End Property

Function Add(ByVal lTableOffset As Long) As Long
  
  If FF = 0 Then Exit Function
  
  Dim loffsetlast As Long
  Dim lOffset As Long
  Dim table As table_type_2
  Dim row As row_type_2
  Dim row_last As row_type_2
  
  Get #FF, lTableOffset, table
  
  If table.field < 1 Then Exit Function

  loffsetlast = table.lastrecord
  lOffset = LOF(FF) + 1
  
  table.recordcount = table.recordcount + 1

  If loffsetlast < 1 Then
    table.record = lOffset
    'row.index = 0
    table.record = lOffset
    table.lastrecord = lOffset
  Else
    Get #FF, loffsetlast, row_last
    row_last.next = lOffset
    Put #FF, loffsetlast, row_last
    'row.index = row_last.index + 1
    row.previous = loffsetlast
    table.lastrecord = lOffset
  End If

  row.sign = rowsign

  Put #FF, lOffset, row
  Put #FF, lTableOffset, table
  Add = lOffset

End Function

Function getFieldTypeName(ByVal lType As DBFieldType) As String
  Select Case lType
    Case fldString: getFieldTypeName = "string"
    Case fldLong: getFieldTypeName = "long"
    Case fldDouble: getFieldTypeName = "double"
    Case fldByte: getFieldTypeName = "byte"
    Case fldDate: getFieldTypeName = "date"
    Case fldInteger: getFieldTypeName = "integer"
    Case fldByteArray: getFieldTypeName = "bytearray"
    Case fldJpegFileBytes: getFieldTypeName = "jpegfilebytearray"
    Case Else: getFieldTypeName = "unknown"
  End Select
End Function

Property Get TableCount() As Long
  TableCount = ltablecount
End Property

Private Sub ReadFields()

  Dim i As Long
  Dim lOffset As Long
  Dim lTableOffset As Long
  Dim table As table_type_2
  Dim field As field_type_2

  For i = 0 To TableCount - 1
    With tables(i)
      lTableOffset = .lOffset
      If lTableOffset > 0 Then
        Get #FF, lTableOffset, table
        lOffset = table.field
        Do While lOffset > 0
          ReDim Preserve .Fields(.fieldscount)
          Get #FF, lOffset, field
          .Fields(.fieldscount).Key = ByteToString(field.szName)
          .Fields(.fieldscount).Type = field.Type
          .Fields(.fieldscount).lOffset = lOffset
          .fieldscount = .fieldscount + 1
          lOffset = field.next
        Loop
      End If
    End With
  Next

End Sub

Function getNextField(ByVal lOffset As Long) As Long
  Dim field As field_type_2
  Get #FF, lOffset, field
  If field.sign = dbsign Then getNextField = field.next
End Function

Function getFieldOffset(ByVal lTableOffset As Long, ByVal szfieldname As String) As Long
  Dim i As Long
  Dim j As Long
  For i = 0 To ltablecount - 1 '()
    If tables(i).lOffset = lTableOffset Then
      For j = 0 To tables(i).fieldscount - 1
        If tables(i).Fields(j).Key = szfieldname Then
          getFieldOffset = tables(i).Fields(j).lOffset
        End If
      Next
      Exit Function
    End If
  Next
End Function

Function addField(ByVal lTableOffset As Long, ByVal szName As String, ByVal lType As DBFieldType) As Long

  Dim lOffset As Long
  Dim loffsetprevious As Long
  Dim table As table_type_2
  Dim field As field_type_2
  Dim fieldrevious As field_type_2

  If lTableOffset < 1 Then Exit Function
  If getFieldOffset(lTableOffset, szName) > 0 Then Exit Function

  Get #FF, lTableOffset, table
  lOffset = table.field
  
  If lOffset = 0 Then
    lOffset = LOF(FF) + 1
    table.field = lOffset
  Else
    lOffset = LOF(FF) + 1
    loffsetprevious = table.lastfield
    Get #FF, loffsetprevious, fieldrevious
    fieldrevious.next = lOffset
    Put #FF, loffsetprevious, fieldrevious
  End If

  table.lastfield = lOffset
  Put #FF, lTableOffset, table
  addField = lOffset

  With field
    .sign = dbsign
    StringToByte .szName, szName
    .Type = lType
  End With

  Put #FF, lOffset, field
  Call flushDB

End Function

Function getTableOffset(ByVal szTableName As String) As Long
  Dim i As Long
  For i = 0 To ltablecount - 1
    If tables(i).Key = szTableName Then
      getTableOffset = tables(i).lOffset
      Exit Function
    End If
  Next
End Function

Function EachTables() As String
  Dim i As Long
  Dim j As Long
  Dim Tmp As String
  For i = 0 To ltablecount - 1
    Tmp = Tmp & vbCrLf & _
          CStr(tables(i).lOffset) & vbTab & """" & tables(i).Key & """"
          For j = 0 To tables(i).fieldscount - 1
            Tmp = Tmp & vbCrLf & vbTab & CStr(tables(i).Fields(j).lOffset) & vbTab & """" & tables(i).Fields(j).Key & """," & vbTab & getFieldTypeName(tables(i).Fields(j).Type)
          Next
  Next
  EachTables = Tmp
End Function

Private Sub ReadTables()
  Dim lOffset As Long
  Dim table As table_type_2
  ltablecount = 0
  Erase tables
  ltablecountarraysize = 10
  ReDim tables(ltablecountarraysize)
  If FF = 0 Then Exit Sub
  lOffset = db.table
  Do While lOffset > 0
    If lOffset = 0 Then Exit Do
    Get #FF, lOffset, table
    If ltablecount + 10 > ltablecountarraysize Then
      ltablecountarraysize = ltablecount * 2 + 10
      ReDim Preserve tables(ltablecountarraysize)
    End If
    With tables(ltablecount)
      .lOffset = lOffset
      .Key = ByteToString(table.szName)
    End With
    lOffset = table.next
    ltablecount = ltablecount + 1
  Loop
  Call ReadFields
End Sub

Function addTable(ByVal szName As String, Optional ByVal mAttributes As MobDbAttributes = 0) As Long
  
  Dim lOffset As Long
  Dim loffsetprevious As Long
  Dim table As table_type_2
  Dim tableprevious As table_type_2
  
  If getTableOffset(szName) > 0 Then Exit Function
  lOffset = db.table
  
  If lOffset = 0 Then
    lOffset = LOF(FF) + 1
    db.table = lOffset
    db.lasttable = lOffset
    Put #FF, 1, db
  Else
    lOffset = LOF(FF) + 1
    loffsetprevious = db.lasttable
    Get #FF, loffsetprevious, tableprevious
    tableprevious.next = lOffset
    Put #FF, loffsetprevious, tableprevious
    db.lasttable = lOffset
    Put #FF, 1, db
  End If
  addTable = lOffset
  With table
    .sign = dbsign
    .attributes = mAttributes
    StringToByte .szName, szName
  End With
  Put #FF, lOffset, table
  Call flushDB

End Function

Private Sub flushDB()
  Call CloseDB
  Call OpenDB
End Sub

Sub CloseDB()
  Erase tables
  If FF <> 0 Then
    Close FF
    FF = 0
  End If
End Sub

Private Sub StringToByte(b() As Byte, ByVal szstring As String)
  Dim lLen As Long
  Dim sz() As Byte
  lLen = UBound(b) - LBound(b) + 1
  If lLen < 1 Then Exit Sub
  sz = StrConv(szstring, vbFromUnicode)
  ReDim Preserve sz(lLen - 1)
  Call CopyMemory(b(0), sz(0), lLen)
End Sub

Private Function ByteToString(b() As Byte) As String
  Dim lRet As Long
  Dim sz() As Byte
  Dim Tmp As String
  Tmp = StrConv(b, vbUnicode)
  lRet = InStr(Tmp, vbNullChar)
  If lRet > 0 Then Tmp = VBA.Left$(Tmp, lRet - 1)
  ByteToString = Tmp
End Function

Sub OpenDB()
  If FF <> 0 Then Call CloseDB
  FF = FreeFile
  Open lpFileName For Binary As #FF
  Get #FF, 1, db
  With db
    If .sign <> dbsign Then
      .sign = dbsign
      Erase .szName
      .table = 0
      .lasttable = 0
      Put #FF, 1, db
    End If
  End With
  ReDim tables(0)
  ltablecountarraysize = 0
  ltablecount = 0
  Call ReadTables
End Sub

'get database file
Public Property Get Filename() As String
  Call flushDB
  Filename = lpFileName
  Call OpenDB
End Property

'set database file
Public Property Let Filename(ByVal vNewValue As String)
  On Error Resume Next
  Call CloseDB
  'Kill vNewValue
  lpFileName = vNewValue
  Call OpenDB
End Property

Public Property Get Value(ByVal lTableOffset As Long, _
                          ByVal lFieldOffset As Long, _
                          ByVal loffsetrow As Long) As Variant

  Dim field As field_type_2
  Dim lType As Long
  Dim Data As data_type_2
  Dim loffsetdata As Long
  Dim lSize As Long

  Dim b() As Byte
  Dim lng As Long
  Dim dbl As Double
  Dim bt As Byte
  Dim dt As Date
  Dim it As Integer
  Dim v
  
  'âîçâðàùàåì òèï vbEmpty
  Value = v 'loffsetdata

  If lFieldOffset < 1 Then Exit Property
  Get #FF, lFieldOffset, field
  If field.sign <> dbsign Then Exit Property
  lType = field.Type

  loffsetdata = getDataByRowOffset(loffsetrow, lFieldOffset)
  If loffsetdata = 0 Then Exit Property

  Get #FF, loffsetdata, Data
  
  If Data.sign <> datasign Then Exit Property
  lType = Data.Type

  Select Case lType
  Case fldJpegFileBytes
    If Data.sizereal > 0 Then
      ReDim b(Data.sizereal - 1)
      Get #FF, Data.lOffset, b
    End If
    Value = b
  Case fldByteArray
    If Data.sizereal > 0 Then
      ReDim b(Data.sizereal - 1)
      Get #FF, Data.lOffset, b
    End If
    Value = b
  Case fldString
    Erase b
    If Data.sizereal > 0 Then
      ReDim b(Data.sizereal - 1)
      Get #FF, Data.lOffset, b
    End If
    Value = b
  Case fldLong
    ReDim b(Data.sizereal - 1)
    Get #FF, Data.lOffset, lng
    Value = lng
  Case fldDouble
    ReDim b(Data.sizereal - 1)
    Get #FF, Data.lOffset, dbl
    Value = dbl
  Case fldByte
    ReDim b(Data.sizereal - 1)
    Get #FF, Data.lOffset, bt
    Value = bt
  Case fldDate
    ReDim b(Data.sizereal - 1)
    Get #FF, Data.lOffset, dt
    Value = dt
  Case fldInteger
    ReDim b(Data.sizereal - 1)
    Get #FF, Data.lOffset, it
    Value = it
  End Select

End Property

Function getDataByRowOffset(ByVal loffsetrow As Long, ByVal loffsetfield As Long, Optional ByVal createifnotexists As Boolean = True) As Long

  Dim row As row_type_2
  Dim loffsetdata As Long
  Dim Data As data_type_2
  Dim loffsetlastdata As Long
  Dim dataadd As data_type_2
  
  Get #FF, loffsetrow, row

  loffsetdata = row.Data
  
  Do While loffsetdata > 0
    loffsetlastdata = loffsetdata
    Get #FF, loffsetdata, Data
    If Data.fieldoffset = loffsetfield Then
      getDataByRowOffset = loffsetdata
      Exit Function
    End If
    loffsetdata = Data.next
  Loop
  
  If Not createifnotexists Then Exit Function
  
  loffsetdata = LOF(FF) + 1
  dataadd.fieldoffset = loffsetfield
  Put #FF, loffsetdata, dataadd
  If row.Data = 0 Then
    row.Data = loffsetdata
    Put #FF, loffsetrow, row
  Else
    Get #FF, loffsetlastdata, Data
    Data.next = loffsetdata
    Put #FF, loffsetlastdata, Data
  End If
  getDataByRowOffset = loffsetdata

End Function

Public Property Let Value(ByVal lTableOffset As Long, _
                          ByVal lFieldOffset As Long, _
                          ByVal loffsetrow As Long, _
                           vNewValue)

  Dim field As field_type_2
  Dim lType As Long
  Dim Data As data_type_2
  Dim loffsetdata As Long
  Dim lSize As Long

  Dim b() As Byte
  Dim lng As Long
  Dim dbl As Double
  Dim bt As Byte
  Dim dt As Date
  Dim it As Integer

  If lFieldOffset < 1 Then Exit Property
  Get #FF, lFieldOffset, field
  If field.sign <> dbsign Then Exit Property
  lType = field.Type
  
  Select Case lType
  Case fldJpegFileBytes
    b = vNewValue
    lSize = UBound(b) + 1
  Case fldByteArray
    b = vNewValue
    lSize = UBound(b) + 1
  Case fldString
    b = CStr(vNewValue)
    lSize = UBound(b) + 1
  Case fldLong
    lng = CLng(Val(vNewValue))
    lSize = 4
  Case fldDouble
    dbl = CDbl(vNewValue)
    lSize = 8
  Case fldByte
    bt = CByte(Val(vNewValue))
    lSize = 1
  Case fldDate
    dt = CDate(vNewValue)
    lSize = 8
  Case fldInteger
    it = CInt(Val(vNewValue))
    lSize = 2
  End Select

  loffsetdata = getDataByRowOffset(loffsetrow, lFieldOffset)
  Get #FF, loffsetdata, Data

  If Data.lOffset < 1 Then Data.lOffset = LOF(FF) + 1

  If lSize > Data.size Then
    Data.size = lSize
    Data.sizereal = lSize
    Data.lOffset = LOF(FF) + 1
  Else
    Data.sizereal = lSize
  End If

  Select Case lType
  Case fldJpegFileBytes: Put #FF, Data.lOffset, b
  Case fldByteArray:     Put #FF, Data.lOffset, b
  Case fldString:        Put #FF, Data.lOffset, b
  Case fldLong:          Put #FF, Data.lOffset, lng
  Case fldDouble:        Put #FF, Data.lOffset, dbl
  Case fldByte:          Put #FF, Data.lOffset, bt
  Case fldDate:          Put #FF, Data.lOffset, dt
  Case fldInteger:       Put #FF, Data.lOffset, it
  End Select

  Data.sign = datasign
  Data.Type = lType

  Put #FF, loffsetdata, Data

  With db.lastupdate
    .datetime = Now
    .tableoffset = lTableOffset
    .rowoffset = loffsetrow
  End With

End Property

Function getFirstField(ByVal lOffsetTable As Long) As Long
  Dim table As table_type_2
  Get #FF, lOffsetTable, table
  If table.sign = dbsign Then getFirstField = table.field
End Function

Function getFieldName(ByVal lOffset As Long) As String
  Dim field As field_type_2
  Get #FF, lOffset, field
  If field.sign = dbsign Then getFieldName = ByteToString(field.szName)
End Function

'Ýêñïîðò òàáëèöû âî âíåøíèé ôàéë
'Ýòîé æå ôóíêöèåé ìîæíî âîñïîëüçîâàòüñÿ äëÿ óïîðÿäî÷èâàíèÿ(ñæàòèÿ) áàçû äàííûõ
Function ExportTable(ByVal lTableOffset As Long, ByVal szPath As String) As Long

  Dim FF As Long
  Dim pvDB As cMobileDB
  Dim lRow As Long
  Dim lField As Long
  Dim lData As Long
  Dim lTable As Long
  Dim Fields As Collection
  Dim ExportFields As Collection
  Dim lExportRow As Long
  Dim i As Long
  Dim szNewTableName As String
  Dim lCount As Long
  Dim lCountAll As Long
  Dim lRowLast As Long

  Set pvDB = New cMobileDB
  Set Fields = New Collection
  Set ExportFields = New Collection

  'Call DeleteFile(szPath)
  
  pvDB.Filename = szPath
  szNewTableName = Me.getTableName(lTableOffset)

  Call pvDB.addTable(szNewTableName)
  lTable = pvDB.getTableOffset(szNewTableName)
  lField = Me.getFirstField(lTableOffset)
  
  Do While lField > 0
    Call pvDB.addField(lTable, Me.getFieldName(lField), Me.getFieldType(lField))
    Call Fields.Add(lField)
    lField = Me.getNextField(lField)
  Loop
  
  lField = pvDB.getFirstField(lTable)
  Do While lField > 0
    Call ExportFields.Add(lField)
    lField = pvDB.getNextField(lField)
  Loop

  lRow = Me.getFirstTableRow(lTableOffset)
  lRowLast = Me.getLastTableRow(lTableOffset)
  lCountAll = Me.recordcount(lTableOffset)

  Do While lRow > 0
    RaiseEvent ExportProgress(lCountAll, lCount)
    lExportRow = pvDB.Add(lTable)
    For i = 1 To Fields.Count
      pvDB.Value(lTable, ExportFields(i), lExportRow) = Me.Value(lTableOffset, Fields(i), lRow)
    Next
    Call pvDB.Update
    lCount = lCount + 1
    RaiseEvent ExportProgress(lCountAll, lCount)
    If lRow = lRowLast Then Exit Do
    lRow = Me.getNextRow(lRow)
  Loop

  Call pvDB.CloseDB
  Set pvDB = Nothing
  Set Fields = Nothing
  Set ExportFields = Nothing
  
  ExportTable = lCount
  
  RaiseEvent ExportComplete
  
End Function

Public Property Get TableAttributes(ByVal lOffsetTable As Long) As MobDbAttributes
  Dim table As table_type_2
  Get #FF, lOffsetTable, table
  If table.sign = dbsign Then TableAttributes = table.attributes
End Property

Public Property Let TableAttributes(ByVal lOffsetTable As Long, ByVal vNewValue As MobDbAttributes)
  Dim table As table_type_2
  Get #FF, lOffsetTable, table
  If table.sign = dbsign Then
    table.attributes = vNewValue
    Put #FF, lOffsetTable, table
  End If
End Property
