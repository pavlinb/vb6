Attribute VB_Name = "mIteSql"
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'ite
Private Declare Sub ite__free_table Lib "iteSql" (ByVal ResultPtr As Long)
Private Declare Function ite__get_table Lib "iteSql" (ByVal db As Long, ByVal SQLStatement As String, ByRef ResultPtr As Long, ByRef rRowCount As Long, ByRef rColCount As Long) As Long
Private Declare Function ite__get_data Lib "iteSql" (ByVal ResultPtr As Long, ByVal RowIdx As Long, ByVal ColIdx As Long, ByVal ColCount As Long) As Long

Private Declare Function ite__open Lib "iteSql" (ByVal dbname As String, ByRef db As Long) As Long
Private Declare Function ite__close Lib "iteSql" (ByVal db As Long) As Long
Private Declare Function ite__exec Lib "iteSql" (ByVal db As Long, ByVal SQLStatement As String) As Long
Private Declare Function ite__libversion Lib "iteSql" () As Long 'T
Private Declare Function ite__errmsg Lib "iteSql" (ByVal db As Long) As Long 'T
Private Declare Function ite__errcode Lib "iteSql" (ByVal db As Long) As Long
Private Declare Function ite__last_insert_rowid Lib "iteSql" (ByVal db As Long) As Long
Private Declare Function ite__changes Lib "iteSql" (ByVal db As Long) As Long
Private Declare Function ite__total_changes Lib "iteSql" (ByVal db As Long) As Long
Private Declare Function ite__prepare Lib "iteSql" (ByVal db As Long, ByVal SQLStatement As String, ByRef hStmt As Long) As Long
Private Declare Function ite__finalize Lib "iteSql" (ByVal hStmt As Long) As Long
Private Declare Function ite__reset Lib "iteSql" (ByVal hStmt As Long) As Long
Private Declare Function ite__step Lib "iteSql" (ByVal hStmt As Long) As Long
Private Declare Function ite__data_count Lib "iteSql" (ByVal hStmt As Long) As Long
Private Declare Function ite__column_count Lib "iteSql" (ByVal hStmt As Long) As Long
Private Declare Function ite__column_name Lib "iteSql" (ByVal hStmt As Long, ByVal ColNum As Long) As Long 'T
Private Declare Function ite__column_decltype Lib "iteSql" (ByVal hStmt As Long, ByVal ColNum As Long) As Long 'T
Private Declare Function ite__column_type Lib "iteSql" (ByVal hStmt As Long, ByVal ColNum As Long) As Long
Private Declare Function ite__column_blob Lib "iteSql" (ByVal hStmt As Long, ByVal ColNum As Long) As Long
Private Declare Function ite__column_bytes Lib "iteSql" (ByVal hStmt As Long, ByVal ColNum As Long) As Long
Private Declare Function ite__column_double Lib "iteSql" (ByVal hStmt As Long, ByVal ColNum As Long) As Double
Private Declare Function ite__column_int Lib "iteSql" (ByVal hStmt As Long, ByVal ColNum As Long) As Long
Private Declare Function ite__column_text Lib "iteSql" (ByVal hStmt As Long, ByVal ColNum As Long) As Long 'T
Private Declare Function ite__bind_blob Lib "iteSql" (ByVal hStmt As Long, ByVal ParamNum As Long, ByVal ptrData As Long, ByVal numBytes As Long) As Long
Private Declare Function ite__bind_int Lib "iteSql" (ByVal hStmt As Long, ByVal ParamNum As Long, ByVal nValue As Long) As Long
Private Declare Function ite__bind_double Lib "iteSql" (ByVal hStmt As Long, ByVal ParamNum As Long, ByVal nValue As Double) As Long
Private Declare Function ite__bind_text Lib "iteSql" (ByVal hStmt As Long, ByVal ParamNum As Long, ByVal szValue As String) As Long
Public Const SQLITE_OK = 0
Public Const SQLITE_ERROR = 1
Public Const SQLITE_INTERNAL = 2
Public Const SQLITE_PERM = 3
Public Const SQLITE_ABORT = 4
Public Const SQLITE_BUSY = 5
Public Const SQLITE_LOCKED = 6
Public Const SQLITE_NOMEM = 7
Public Const SQLITE_READONLY = 8 ' Attempt to write a readonly database
Public Const SQLITE_INTERRUPT = 9 ' Operation terminated by sqlite3_interrupt()
Public Const SQLITE_IOERR = 10 ' Some kind of disk I/O error occurred
Public Const SQLITE_CORRUPT = 11 ' The database disk image is malformed
Public Const SQLITE_NOTFOUND = 12 ' (Internal Only) Table or record not found
Public Const SQLITE_FULL = 13 ' Insertion failed because database is full
Public Const SQLITE_CANTOPEN = 14 ' Unable to open the database file
Public Const SQLITE_PROTOCOL = 15 ' Database lock protocol error
Public Const SQLITE_EMPTY = 16 ' Database is empty
Public Const SQLITE_SCHEMA = 17 ' The database schema changed
Public Const SQLITE_TOOBIG = 18 ' Too much data for one row of a table
Public Const SQLITE_CONSTRAINT = 19 ' Abort due to contraint violation
Public Const SQLITE_MISMATCH = 20 ' Data type mismatch
Public Const SQLITE_MISUSE = 21 ' Library used incorrectly
Public Const SQLITE_NOLFS = 22 ' Uses OS features not supported on host
Public Const SQLITE_AUTH = 23 ' Authorization denied
Public Const SQLITE_FORMAT = 24 ' Auxiliary database format error
Public Const SQLITE_RANGE = 25 ' 2nd parameter to sqlite3_bind out of range
Public Const SQLITE_NOTADB = 26 ' File opened that is not a database file
Public Const SQLITE_ROW = 100 ' sqlite3_step() has another row ready
Public Const SQLITE_DONE = 101 ' sqlite3_step() has finished executing
Public Enum SQLITETYPE
    SQLITE_INTEGER = 1
    SQLITE_FLOAT = 2
    SQLITE_TEXT = 3
    SQLITE_BLOB = 4
    SQLITE_NULL = 5
End Enum
Private Const JULIANDAY_OFFSET As Double = 2415018.5

Private Function UTF8StringFromPtr(ByVal pUtf8String As Long) As String
    Dim cSize As Long
    UTF8StringFromPtr = ""
    cSize = MultiByteToWideChar(65001, 0, pUtf8String, -1, 0, 0)
    If cSize > 1 Then
        UTF8StringFromPtr = String(cSize - 1, " ")
        MultiByteToWideChar 65001, 0, pUtf8String, -1, StrPtr(UTF8StringFromPtr), cSize
    End If
End Function

Private Function BytesFromPtr(ByVal lAddr As Long, ByVal lSize As Long) As Byte()
    ReDim bvData(lSize - 1) As Byte
    CopyMemory bvData(0), ByVal lAddr, lSize
    BytesFromPtr = bvData
End Function

Private Function ToJulianDay(oleDate As Date) As Double
    ToJulianDay = CDbl(oleDate) + JULIANDAY_OFFSET
End Function

Private Function FromJulianDay(julianDay As Double) As Date
    FromJulianDay = CDate(julianDay - JULIANDAY_OFFSET)
End Function
 
Public Function ite_errmsg(ByVal db As Long) As Long
    ite_errmsg = UTF8StringFromPtr(ite__errmsg(db))
End Function

Public Function ite_errcode(ByVal db As Long) As Long
    ite_errcode = ite__errcode(db)
End Function

Public Function ite_libversion() As String
    ite_libversion = UTF8StringFromPtr(ite__libversion)
End Function

Public Function ite_open(ByVal DBFile As String, ByRef db As Long) As Long
    ite_open = ite__open((DBFile), db)
End Function

Public Function ite_close(ByVal db As Long) As Long
    ite_close = ite__close(db)
End Function

Public Function ite_exec(ByVal db As Long, ByVal zSQL As String) As Long
    ite_exec = ite__exec(db, zSQL)
End Function

Public Function ite_get_table(ByVal db As Long, ByVal zSQL As String, ByRef ResultPtr As Long, ByRef rRowCount As Long, ByRef rColCount As Long) As Long
    ite_get_table = ite__get_table(db, zSQL, ResultPtr, rRowCount, rColCount)
End Function

Public Sub ite_free_table(ByVal ResultPtr As Long)
    Call ite__free_table(ResultPtr)
End Sub

Public Function ite_get_data(ByVal ResultPtr As Long, ByVal RowIdx As Long, ByVal ColIdx As Long, ByVal ColCount As Long) As String
    ite_get_data = UTF8StringFromPtr(ite__get_data(ResultPtr, RowIdx, ColIdx, ColCount))
End Function

Public Function ite_changes(ByVal db As Long) As Long
    ite_changes = ite__changes(db)
End Function

Public Function ite_total_changes(ByVal db As Long) As Long
    ite_total_changes = ite__total_changes(db)
End Function

Public Function ite_last_insert_rowid(ByVal db As Long) As Long
    ite_last_insert_rowid = ite__last_insert_rowid(db)
End Function

Public Function ite_prepare(ByVal db As Long, ByVal zSQL As String, ByRef hStmt As Long) As Long
    ite_prepare = ite__prepare(db, zSQL, hStmt)
End Function

Public Function ite_finalize(ByVal hStmt As Long) As Long
    ite_finalize = ite__finalize(hStmt)
End Function

Public Function ite_reset(ByVal hStmt As Long) As Long
    ite_reset = ite__reset(hStmt)
End Function

Public Function ite_step(ByVal hStmt As Long) As Long
    ite_step = ite__step(hStmt)
End Function

Public Function ite_next(ByVal hStmt As Long) As Boolean
    ite_next = (ite__step(hStmt) = 100)
End Function

Public Function ite_data_count(ByVal hStmt As Long) As Long
    ite_data_count = ite__data_count(hStmt)
End Function

Public Function ite_column_name(ByVal hStmt As Long, ByVal ColNum As Long) As String
    ite_column_name = UTF8StringFromPtr(ite__column_name(hStmt, ColNum))
End Function

Public Function ite_column_decltype(ByVal hStmt As Long, ByVal ColNum As Long) As String
    ite_column_decltype = UTF8StringFromPtr(ite__column_decltype(hStmt, ColNum))
End Function

Public Function ite_column_type(ByVal hStmt As Long, ByVal ColNum As Long) As Long
    ite_column_type = ite__column_type(hStmt, ColNum)
End Function

Public Function ite_column_count(ByVal hStmt As Long) As Long
    ite_column_count = ite__column_count(hStmt)
End Function

Public Function ite_column_double(ByVal hStmt As Long, ByVal ColNum As Long) As Double
    ite_column_double = ite__column_double(hStmt, ColNum)
End Function

Public Function ite_column_int(ByVal hStmt As Long, ByVal ColNum As Long) As Double
    ite_column_int = ite__column_int(hStmt, ColNum)
End Function

Public Function ite_column_text(ByVal hStmt As Long, ByVal ColNum As Long) As String
    ite_column_text = UTF8StringFromPtr(ite__column_text(hStmt, ColNum))
End Function

Public Function ite_column_blob(ByVal hStmt As Long, ByVal ColNum As Long) As Byte()
    Dim lAddr As Long, lSize As Long
    lAddr = ite__column_blob(hStmt, ColNum)
    lSize = ite__column_bytes(hStmt, ColNum)
    ite_column_blob = BytesFromPtr(lAddr, lSize)
End Function

Public Function ite_column_date(ByVal hStmt As Long, ByVal ColNum As Long) As Date
    ite_column_date = FromJulianDay(ite__column_double(hStmt, ColNum))
End Function

Public Function ite_bind_date(ByVal hStmt As Long, ByVal ParamNum As Long, ByVal nValue As Date) As Long
    ite_bind_date = ite__bind_double(hStmt, ParamNum, ToJulianDay(nValue))
End Function

Public Function ite_bind_blob(ByVal hStmt As Long, ByVal ParamNum As Long, pData() As Byte) As Long
    ite_bind_blob = ite__bind_blob(hStmt, ParamNum, VarPtr(pData(0)), UBound(pData) + 1)
End Function

Public Function ite_bind_int(ByVal hStmt As Long, ByVal ParamNum As Long, ByVal nValue As Long) As Long
    ite_bind_int = ite__bind_int(hStmt, ParamNum, nValue)
End Function

Public Function ite_bind_double(ByVal hStmt As Long, ByVal ParamNum As Long, ByVal nValue As Double) As Long
    ite_bind_double = ite__bind_double(hStmt, ParamNum, nValue)
End Function

Public Function ite_bind_text(ByVal hStmt As Long, ByVal ParamNum As Long, ByVal szValue As String) As Long
    ite_bind_text = ite__bind_text(hStmt, ParamNum, szValue)
End Function

Public Function ite_update_blob(ByVal db As Long, ByVal zSQL As String, pData() As Byte) As Long
    Dim sqlite3_stmt As Long
    If ite_prepare(db, zSQL, sqlite3_stmt) = 0 Then
        ite_update_blob = ite_bind_blob(sqlite3_stmt, 1, pData)
        ite_step sqlite3_stmt
        ite_finalize sqlite3_stmt
    End If
End Function

Public Function ite_compact(ByVal db As Long) As Long
    ite_compact = ite_exec(db, "VACUUM")
End Function

Public Function ite_row_number(ByVal db As Long, ByVal zTable As String, Optional ByVal zFilter As String) As Long
    Dim sqlite3_stmt As Long, zSQL As String
    zSQL = "select count(*) from " & zTable
    If Len(zFilter) > 0 Then zSQL = zSQL & " where " & zFilter
    If ite_prepare(db, zSQL, sqlite3_stmt) = 0 Then
        ite_step sqlite3_stmt
        ite_row_number = ite_column_int(sqlite3_stmt, 0)
        ite_finalize sqlite3_stmt
    End If
End Function


