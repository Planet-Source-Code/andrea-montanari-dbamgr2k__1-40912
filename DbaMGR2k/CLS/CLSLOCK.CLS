VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private iKey As Integer
Private iReq_Spid As Integer ' Process ID of the process requesting the lock.
Private sLockType As String '  nvarchar(36)    A text description of a locking mode. For more information about interpreting values, see the description of the system table syslockinfo column req_mode.
Private sTableName As String 'nvarchar(129)   If applicable, the name of the table against which the lock is applied.
Private sIndexName As String 'nvarchar(129)   If applicable, the name of the index against which the lock is applied.
Private sDbName As String 'nvarchar(129)   Name of the database in which the locked resource is defined.
Private iStatus As Integer 'tinyint An integer indicating lock application status. For more information about interpreting values, see the description of the system table syslockinfo column req_status.
Public Property Let Spid(ByVal iVal As Integer)
    iReq_Spid = iVal
End Property
Public Property Get Spid() As Integer
    Spid = iReq_Spid
End Property
Public Property Let LockType(ByVal sVal As String)
    sLockType = sVal
End Property
Public Property Get LockType() As String
    LockType = sLockType
End Property
Public Property Let LockTbName(ByVal sVal As String)
    sTableName = sVal
End Property
Public Property Get LockTbName() As String
    LockTbName = sTableName
End Property
Public Property Let LockIdxName(ByVal sVal As String)
    sIndexName = sVal
End Property
Public Property Get LockIdxName() As String
    LockIdxName = sIndexName
End Property
Public Property Let LockDbName(ByVal sVal As String)
    sDbName = sVal
End Property
Public Property Get LockDbName() As String
    LockDbName = sDbName
End Property
Public Property Let LockStatus(ByVal iVal As Integer)
    iStatus = iVal
End Property
Public Property Get LockStatus() As Integer
    LockStatus = iStatus
End Property
Public Property Get strLockStatus() As String
    If iStatus = 1 Then
        strLockStatus = "Granted"
    ElseIf iStatus = 2 Then
        strLockStatus = "Converting"
    ElseIf iStatus = 3 Then
        strLockStatus = "Waiting"
    End If

'1=Granted
'2=Converting
'3=Waiting

End Property
Public Property Let Key(ByVal iVal As Integer)
    iKey = iVal
End Property
Public Property Get Key() As Integer
    Key = iKey
End Property

