VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMultiCore 
   Caption         =   "CPU Usage Per Core"
   ClientHeight    =   2085
   ClientLeft      =   1395
   ClientTop       =   2175
   ClientWidth     =   3960
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   3960
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
      Left            =   3930
      Top             =   30
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
'//////////////////////////////////////////
' Get CPU Usage Per Core v1.3            //
' Edgemeal - http://edgemeal.110mb.com   //
'//////////////////////////////////////////

'Notes: Run as Administrator!
Option Explicit

'  array to hold CPU usage value for each CPU core
Dim dblCpuUsage() As Double



Private Sub Form_Load()
    
    Dim i As Integer
    
    If App.PrevInstance = True Then
        MsgBox " A copy of CPU Usage Test is already loaded!", vbInformation
        Unload Me
        Exit Sub
    End If
        
    ' setup CPU usage quary
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
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False ' stop updating
    Close_CPU_Usage ' close PDH if used
    Set frmMultiCore = Nothing
End Sub

Private Sub Timer1_Timer()
    ' // update display//
    
    Dim i As Integer
    Dim TotalUsage As Double
    
    'quary current cpu usage / store in array
    Update_Cpu_Usage dblCpuUsage()
    
    ' display usage per core
    For i = 0 To NumCores
        TotalUsage = TotalUsage + dblCpuUsage(i)
        ProgBar(i).Value = CInt(dblCpuUsage(i))
        lblProg(i).Caption = "CPU " & i & ": " & Format(dblCpuUsage(i), "0.0") & "%"
    Next i
    
    ' display total (usage per core divided by number of cores)
    Me.Caption = "CPU Usage: " & Format(TotalUsage / (NumCores + 1), "0.0") & "%"
End Sub
