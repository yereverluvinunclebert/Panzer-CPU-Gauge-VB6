VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMultiCore 
   Caption         =   "CPU Usage Per Core"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   3900
   Icon            =   "frmMultiCore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   Begin VB.Timer multicorePositionTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2730
      Top             =   1425
   End
   Begin MSComctlLib.ProgressBar ProgBar 
      Height          =   225
      Index           =   0
      Left            =   150
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
   Begin VB.Timer tmrMultiCore 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2730
      Top             =   885
   End
   Begin VB.Label lblGenericLabel 
      Caption         =   "tmrMultiCore"
      Height          =   240
      Left            =   1665
      TabIndex        =   2
      Top             =   990
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblProg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "00.0"
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
    
    Dim I As Integer: I = 0
    Dim leftPoint As Long:  leftPoint = 0
    Dim topPoint  As Long:  topPoint = 0
    
    On Error GoTo Form_Load_Error
    
    ' set the initial small interval so that the sliders are displayed straight away, reset later
    tmrMultiCore.Interval = 1000
    
    Call readMulticorePosition

    ' setup CPU usage query
    If fInitializeCPU = True Then
    
        If Val(gblMulticoreXPosTwips) <> 0 Then
                
            Me.Top = Val(gblMulticoreYPosTwips)
            Me.Left = Val(gblMulticoreXPosTwips)
        Else
    
            leftPoint = fAlpha.gaugeForm.Widgets("housing/surround").Widget.Left
            topPoint = fAlpha.gaugeForm.Widgets("housing/stopButton").Widget.Top
        
            Me.Top = ((topPoint) * gblScreenTwipsPerPixelY) - 150
            Me.Left = (fAlpha.gaugeForm.Left + leftPoint) * gblScreenTwipsPerPixelX - 250
            
        End If
       
        ' redimension cpu usage value array to number of CPU cores
        ReDim dblCpuUsage(NumCores)
        
        ' set the font characteristics of the master label
        lblProg(0).Font.Name = gblPrefsFont
        lblProg(0).Font.Italic = CBool(gblPrefsFontItalics)
        lblProg(0).ForeColor = gblPrefsFontColour
        lblProg(0).Font.Size = Val(gblPrefsFontSizeLowDPI)
        
        ' add additional prog bar and lable for each cpu core (if more then 1 cpu core).
        For I = 1 To NumCores
            Load ProgBar(I)
            Load lblProg(I)
            
            lblProg(I).Font.Name = gblPrefsFont
            lblProg(I).Font.Italic = CBool(gblPrefsFontItalics)
            lblProg(I).ForeColor = gblPrefsFontColour
            lblProg(I).Font.Size = Val(gblPrefsFontSizeLowDPI)
            
            ProgBar(I).Max = 100
            ProgBar(I).Top = ProgBar(I - 1).Top + ProgBar(I - 1).Height + 15
            lblProg(I).Top = ProgBar(I).Top
            ProgBar(I).Visible = True
            lblProg(I).Visible = True
        Next I
        
        frmMultiCore.Height = ProgBar(I - 1).Top + 800 + 250
        
        ' initalize cpu usage
        Update_Cpu_Usage dblCpuUsage()
        
        ' start the main timer that displays cpu usage
        tmrMultiCore.Enabled = True
        
        ' start the position timer
        multicorePositionTimer.Enabled = True
        
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
   
    gblMultiCoreEnable = "0"
    
    Call writeMulticorePosition

    tmrMultiCore.Enabled = False ' stop updating
    Close_CPU_Usage ' close PDH if used
    Set frmMultiCore = Nothing

   On Error GoTo 0
   Exit Sub

Form_Unload_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Form frmMultiCore"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : tmrMultiCore_Timer
' Author    : EdgeMeal
' Date      : 31/05/2025
' Purpose   : update display
'---------------------------------------------------------------------------------------
'
Private Sub tmrMultiCore_Timer()

    On Error GoTo tmrMultiCore_Timer_Error
    
    Call updateCoreDisplay

   On Error GoTo 0
   Exit Sub

tmrMultiCore_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrMultiCore_Timer of Form frmMultiCore"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : updateCoreDisplay
' Author    : beededea
' Date      : 02/06/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub updateCoreDisplay()
    Dim I As Integer: I = 0
    Dim TotalUsage As Double: TotalUsage = 0
    
    'query current cpu usage / store in array
    On Error GoTo updateCoreDisplay_Error
    
    tmrMultiCore.Interval = Val(gblSamplingInterval) * 1000

    Update_Cpu_Usage dblCpuUsage()
    
    ' display usage per core
    For I = 0 To NumCores
        TotalUsage = TotalUsage + dblCpuUsage(I)
        ProgBar(I).Value = CInt(dblCpuUsage(I))
        lblProg(I).Caption = "CPU " & I & ": " & Format(dblCpuUsage(I), "0.0") & "%"
    Next I
    
    ' display total (usage per core divided by number of cores)
    Me.Caption = "CPU Usage: " & Format(TotalUsage / (NumCores + 1), "0.0") & "%"

   On Error GoTo 0
   Exit Sub

updateCoreDisplay_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure updateCoreDisplay of Form frmMultiCore"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : multicorePositionTimer_Timer
' Author    : beededea
' Date      : 27/05/2023
' Purpose   : periodically read the prefs form position and store
'---------------------------------------------------------------------------------------
'
Private Sub multicorePositionTimer_Timer()
    ' save the current X and y position of this form to allow repositioning when restarting
    On Error GoTo multicorePositionTimer_Timer_Error
   
    If frmMultiCore.IsVisible = True Then Call writeMulticorePosition

   On Error GoTo 0
   Exit Sub

multicorePositionTimer_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure multicorePositionTimer_Timer of Form frmMultiCore"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : writeMulticorePosition
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : save the current X and y position of this form to allow repositioning when restarting
'---------------------------------------------------------------------------------------
'
Private Sub writeMulticorePosition()
        
   On Error GoTo writeMulticorePosition_Error

    If frmMultiCore.WindowState = vbNormal Then ' when vbMinimised the value = -48000  !
        gblMulticoreXPosTwips = CStr(frmMultiCore.Left)
        gblMulticoreYPosTwips = CStr(frmMultiCore.Top)
        
        ' now write those params to the toolSettings.ini
        sPutINISetting "Software\PzCPUGauge", "multicoreXPosTwips", gblMulticoreXPosTwips, gblSettingsFile
        sPutINISetting "Software\PzCPUGauge", "multicoreYPosTwips", gblMulticoreYPosTwips, gblSettingsFile
    End If
    
    On Error GoTo 0
   Exit Sub

writeMulticorePosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure writeMulticorePosition of Form frmMultiCore"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : readMulticorePosition
' Author    : beededea
' Date      : 28/05/2023
' Purpose   : save the current X and y position of this form to allow repositioning when restarting
'---------------------------------------------------------------------------------------
'
Private Sub readMulticorePosition()
        
   On Error GoTo readMulticorePosition_Error
   
    gblMulticoreXPosTwips = fGetINISetting("Software\PzCPUGauge", "multicoreXPosTwips", gblSettingsFile)
    gblMulticoreYPosTwips = fGetINISetting("Software\PzCPUGauge", "multicoreYPosTwips", gblSettingsFile)
    
    On Error GoTo 0
   Exit Sub

readMulticorePosition_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure readMulticorePosition of Form frmMultiCore"
End Sub




'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : beededea
' Date      : 08/05/2023
' Purpose   : calling a manual property to a form allows external checks to the form to
'             determine whether it is loaded, without also activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    If Me.WindowState = vbNormal Then
        IsVisible = Me.Visible
    Else
        IsVisible = False
    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Form frmMultiCore"
            Resume Next
          End If
    End With
End Property
