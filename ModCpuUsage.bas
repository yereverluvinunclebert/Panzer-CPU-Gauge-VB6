Attribute VB_Name = "ModCpuUsage"
Option Explicit

' v1.3 demo - changed to public
Public NumCores As Integer ' store number of cpu cores

' used to get num of cpu cores
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

' PDH
Private Const MAX_PATH As Integer = 260
Private Const COUNTERPERF_PROCESSOR = 238
Private Const COUNTERPERF_PERCENTPROCESSORTIME = 6

Private Type CounterInfo
    hCounter As Long
End Type

Private Counters() As CounterInfo
Private hQuery As Long

Private Enum PDH_STATUS
    PDH_CSTATUS_VALID_DATA = &H0
    PDH_CSTATUS_NEW_DATA = &H1
End Enum

Private Declare Function PdhOpenQuery Lib "PDH.DLL" (ByVal Reserved As Long, ByVal dwUserData As Long, ByRef hQuery As Long) As PDH_STATUS
Private Declare Function PdhCloseQuery Lib "PDH.DLL" (ByVal hQuery As Long) As PDH_STATUS
Private Declare Function PdhVbAddCounter Lib "PDH.DLL" (ByVal QueryHandle As Long, ByVal CounterPath As String, ByRef CounterHandle As Long) As PDH_STATUS
Private Declare Function PdhCollectQueryData Lib "PDH.DLL" (ByVal QueryHandle As Long) As PDH_STATUS
Private Declare Function PdhVbGetDoubleCounterValue Lib "PDH.DLL" (ByVal CounterHandle As Long, ByRef CounterStatus As Long) As Double
Private Declare Sub PdhLookupPerfNameByIndex Lib "PDH.DLL" Alias "PdhLookupPerfNameByIndexA" (ByVal szMachineName As String, ByVal dwNameIndex As Long, ByVal szNameBuffer As String, ByRef pcchNameBufferSize As Long)

Public Sub Close_CPU_Usage()
   If hQuery Then PdhCloseQuery (hQuery)  ' close
End Sub

Public Function InitializeCPU() As Boolean
    ' Add CPU counters
    Dim pdhStatus As PDH_STATUS
    Dim SysInfo As SYSTEM_INFO
    Dim CPU_Obj As String
    Dim hCounter As Long
    Dim i As Integer
    
    ' get # of cpus (cores)
    GetSystemInfo SysInfo
    NumCores = SysInfo.dwNumberOrfProcessors - 1
        
    ' we need at least 1 CPU core (Core 0) to proceed.
    If NumCores < 0 Then Exit Function
    
    ' set number of PDH counters needed
    ReDim Counters(NumCores)
        
    ' Query PDH
    pdhStatus = PdhOpenQuery(0, 1, hQuery)
    If pdhStatus <> PDH_CSTATUS_VALID_DATA Then Exit Function ' Query failed
    
    For i = 0 To NumCores ' add counter for each cpu core
        CPU_Obj = GetCPUCounter(CStr(i)) ' get CPU Process Object and Instance names for next cpu core
        pdhStatus = PdhVbAddCounter(hQuery, CPU_Obj, hCounter) ' add counter
        If pdhStatus = PDH_CSTATUS_VALID_DATA Then
            Counters(i).hCounter = hCounter
        Else ' add counter failed
            Exit Function
        End If
    Next i
    
    ' return success
    InitializeCPU = True

End Function

Public Sub Update_Cpu_Usage(ByRef dblArray() As Double)
    
    ' // quary counters  //
    Dim pdhStatus As PDH_STATUS
    Dim i As Integer
        
    ' Query Data
    PdhCollectQueryData (hQuery)
    
    ' get cpu usage per core, store in passed array
    For i = 0 To NumCores
        dblArray(i) = PdhVbGetDoubleCounterValue(Counters(i).hCounter, pdhStatus)
    Next

End Sub

Private Function GetCPUCounter(strInstance As String) As String
    ' // get Object & Counter names for CPU Usage //
    ' / Different languages of windows use different string names so we need a look-up! /
    Dim NameLen As Long
    Dim ObjectName As String, CounterName As String
    
    NameLen = MAX_PATH
    ObjectName = Space$(NameLen)
    PdhLookupPerfNameByIndex ByVal vbNullString, COUNTERPERF_PROCESSOR, ObjectName, NameLen
    ObjectName = Left$(ObjectName, NameLen - 1)
        
    NameLen = MAX_PATH
    CounterName = Space$(NameLen)
    
    PdhLookupPerfNameByIndex ByVal vbNullString, COUNTERPERF_PERCENTPROCESSORTIME, CounterName, NameLen
    CounterName = Left$(CounterName, NameLen - 1)
    GetCPUCounter = "\" & ObjectName & "(" & strInstance & ")\" & CounterName
    
End Function
