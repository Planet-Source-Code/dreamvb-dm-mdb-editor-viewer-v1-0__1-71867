Attribute VB_Name = "ModDB"
Option Explicit

Public DBOpen As Boolean
Public db As Database

Private Type TField
    fName As String
    fType As Integer
    fSize As Integer
    fAttr As Integer
    fRequired As Boolean
    fAllowZeroLength As Boolean
End Type

'Used for Fields
Public TempField() As TField
Public TmpTable As String
Public FieldCounter As Integer

'Database error msg
Public DBErrMsg As String
'Current Oppened Database
Public DataBaseFile As String
'Current Selected Table
Public CurrentTable As String
'Current selected record
Public CurrentRecord As Integer

Private Sub TableDefauls()
    Erase TempField
    TmpTable = vbNullString
    FieldCounter = 0
End Sub

Public Function CreateBlankDataBase(ByVal Filename As String, Optional mOptions As DatabaseTypeEnum) As Integer
Dim db As Database
On Error GoTo ErrFlag:
    'Create a blank database
    Set db = CreateDatabase(Filename, dbLangGeneral, mOptions)
    
    CreateBlankDataBase = 1
    Set db = Nothing
    
    Exit Function
ErrFlag:
    CreateBlankDataBase = 0
    Set db = Nothing
End Function

Public Sub CloseDataBase()
    If (DBOpen) Then
        Call db.Close
        DBOpen = False
    End If
End Sub

Public Function OpenDB(ByVal Filename As String) As Boolean
On Error GoTo OpenErr:
    'Open the Database
    Set db = OpenDatabase(Filename, False)
    'Tells us the database is open
    DBOpen = True
    OpenDB = True
    Exit Function
OpenErr:
    OpenDB = False
End Function

Public Function GetFieldData(ByVal sQuery As String, TListView As ListView) As Integer
Dim fCount As Integer
Dim rc As Recordset
On Error GoTo DErr:

    'This sub adds the data to the listview
    'Clear the listview
    TListView.ListItems.Clear
    TListView.Sorted = False
    'First lets open the record set
    Set rc = db.OpenRecordset(sQuery)
    
    With rc
        While (Not rc.EOF)
            'Add the first item
            TListView.ListItems.Add , , rc.Fields(0).Value
            'Add the subitems
            For fCount = 1 To (rc.Fields.Count - 1)
                TListView.ListItems(TListView.ListItems.Count).SubItems(fCount) = rc.Fields(fCount).Value & vbNullChar
            Next fCount
            'Get the next record
            .MoveNext
        Wend
    End With
    
    GetFieldData = 1
    
    Set rc = Nothing
    Exit Function
    'Error flag
DErr:
    GetFieldData = 0
    DBErrMsg = Err.Description
End Function

Public Sub GetFieldNames(ByVal TableName As String, TListView As ListView)
Dim rc As Recordset
Dim TField As Field

    'This sub adds all the field names to the listview control
    'Clear the headers
    TListView.ColumnHeaders.Clear
    'First we need to open the record set
    Set rc = db.OpenRecordset(TableName)
    'Now we can add the field names
    For Each TField In rc.Fields
        TListView.ColumnHeaders.Add , , TField.Name
    Next TField
    
    Set rc = Nothing
    Set TField = Nothing
End Sub

Public Sub GetTables(vbTabC As TabStrip)
Dim Td As TableDef
    vbTabC.Tabs.Clear
    
    'Get all Table names
    For Each Td In db.TableDefs
        If (Td.Attributes And dbSystemObject) = 0 Then
            'If not system table add the table
            vbTabC.Tabs.Add , , Td.Name
        End If
    Next Td
    
    Set Td = Nothing
End Sub

Public Function CreateTable() As Integer
Dim Count As Integer
Dim Td As TableDef
On Error GoTo CreateErr:

    'This function creates the table
    If Len(TmpTable) = 0 Then
        Call TableDefauls
        CreateTable = 0
        Exit Function
    End If
    
    'Create TableDef
    Set Td = db.CreateTableDef(TmpTable)
    'Fill in TableDef info
    With Td
        For Count = 0 To UBound(TempField)
            'Create Fields
            .Fields.Append .CreateField(TempField(Count).fName, TempField(Count).fType, TempField(Count).fSize)
            'Set Feild required
            .Fields(Count).Required = TempField(Count).fRequired
            'Set allow zero length
            .Fields(Count).AllowZeroLength = TempField(Count).fAllowZeroLength
            'Set field Attributes
            .Fields(Count).Attributes = TempField(Count).fAttr
        Next Count
    End With
    'Append the TableDef to the database
    db.TableDefs.Append Td
    CreateTable = 2
    
    'Clear up
    Set Td = Nothing
    Call TableDefauls
    Count = 0
    
    Exit Function
    'Error flag
CreateErr:
    DBErrMsg = Err.Description
    CreateTable = 1
End Function

Public Function DeleteRecord(ByVal RecIndex As Integer) As Integer
On Error GoTo DeleteErr:
Dim rc As Recordset
    'Open the record set
    Set rc = db.OpenRecordset(CurrentTable)
    'Move to the record to delete
    rc.Move RecIndex
    'Delete the record
    rc.Delete
    
    Set rc = Nothing
    DeleteRecord = 1
    Exit Function
'Error flag
DeleteErr:
    DBErrMsg = Err.Description
    DeleteRecord = 0
End Function

Public Function RenameTable(ByVal NewTableName As String) As Integer
On Error GoTo RenameErr:
Dim rc As Recordset
Dim Td As TableDef
    'Open the tabledef
    Set Td = db.TableDefs(CurrentTable)
    'Rename the table
    Td.Name = NewTableName
    
    RenameTable = 1
    Exit Function
    'Error flag
RenameErr:
    RenameTable = 0
    DBErrMsg = Err.Description
End Function

Public Function DeleteTable() As Integer
On Error GoTo DelErr:
    'Delete a table form a database.
    Call db.TableDefs.Delete(CurrentTable)
    Call db.TableDefs.Refresh
    
    Exit Function
    'Error flag
DelErr:
    DeleteTable = 0
    DBErrMsg = Err.Description
End Function

Public Function GetRecordCount() As Integer
On Error GoTo ErrFlag:
Dim rc As Recordset
    'Open the recordset
    Set rc = db.OpenRecordset(CurrentTable)
    'Return number of records
    GetRecordCount = rc.RecordCount
    Set rc = Nothing
    
    Exit Function
    'Error flag
ErrFlag:
    GetRecordCount = -1
    DBErrMsg = Err.Description
End Function
