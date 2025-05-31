VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMultiCore 
   Caption         =   "CPU Usage Per Core"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   225
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   397
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2715
      Top             =   885
   End
   Begin VB.Label lblProg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   2520
      TabIndex        =   1
      Top             =   210
      Width           =   1935
   End
End
Attribute VB_Name = "frmMultiCore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmMultiCore
' Author    : EdgeMeal http://edgemeal.110mb.com
' Date      : 31/05/2025
' Purpose   : Get CPU Usage Per Core v1.3
'---------------------------------------------------------------------------------------

'Notes: Run as Administrator!
Option Explicit

'  array to hold CPU usage value for each CPU core
Private dblCpuUsage() As Double



'---------------------------------------------------------------------------------------
' Procedure : Form_Load
' Author    : EdgeMeal
' Date      : 31/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Load()
    
    Dim i As Integer
         
    ' setup CPU usage quary
   On Error GoTo Form_Load_Error

    If InitializeCPU = True Then
       
        ' redimension cpu usage value array to number of CPU cores
        ReDim dblCpuUsage(NumCores)
        
        ' add additional prog bar and lable for each cpu core (if more then 1 cpu core).
        For i = 1 To NumCores
            Load ProgBar(i)
            Load lblProg(i)
            ProgBar(i).Max = 100
            ProgBar(i).Top = ProgBar(i - 1).Top + ProgBar(i - 1).Height + 15
            lblProg(i).Top = ProgBar(i).Top
            ProgBar(i).Visible = True
            lblProg(i).Visible = True
        Next i
        
        ' initalize cpu usage
        Update_Cpu_Usage dblCpuUsage()
        
        ' update/display usage
        Timer1.Enabled = True
    Else
        ' failed...
        MsgBox "Sorry, Unable to initialize CPU usage.", vbApplicationModal + vbInformation
        Unload Me
        Exit Sub
    End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Form frmMultiCore"
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Form_Unload
' Author    : EdgeMeal
' Date      : 31/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

    Timer1.Enabled = False ' stop updating
    Close_CPU_Usage ' close PDH if used
    Set frmMultiCore = Nothing

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form frmMultiCore"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Timer1_Timer
' Author    : EdgeMeal
' Date      : 31/05/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Timer1_Timer()
    ' // update display//
    
    Dim i As Integer
    Dim TotalUsage As Double
    
    'quary current cpu usage / store in array
   On Error GoTo Timer1_Timer_Error

    Update_Cpu_Usage dblCpuUsage()
    
    ' display usage per core
    For i = 0 To NumCores
        TotalUsage = TotalUsage + dblCpuUsage(i)
        ProgBar(i).Value = CInt(dblCpuUsage(i))
        lblProg(i).Caption = "CPU " & i & ": " & Format(dblCpuUsage(i), "0.0") & "%"
    Next i
    
    ' display total (usage per core divided by number of cores)
    Me.Caption = "CPU Usage: " & Format(TotalUsage / (NumCores + 1), "0.0") & "%"

   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Form frmMultiCore"
End Sub
