VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cwOverlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule IntegerDataType, ModuleWithoutFolder
'---------------------------------------------------------------------------------------
' Module    : cwOverlay
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : create a fully click-through Widget, which allows for all kind of
'             things to render on top of the PSD-generated Widgets created in cfAlpha
'---------------------------------------------------------------------------------------

Option Explicit

Private Declare Function GetSystemTimes Lib "kernel32" (lpIdleTime As Any, lpKernelTime As Any, lpUserTime As Any) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

'these Public Properties are set in the Parent (here, cfAlpha)
Public cCollOverlayImages As cCollection
Public cxPerc As Single
Public cyPerc As Single

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents tmrSampler As cTimer
Attribute tmrSampler.VB_VarHelpID = -1
Private WithEvents tmrAnimator As cTimer
Attribute tmrAnimator.VB_VarHelpID = -1

Private SDeg As Double
Private cpuPercent As Currency
Private noOfProcessors As Currency
Private oldSDeg As Double
Private newSDeg As Double
Private animationQuanta As Double
Private animationAmount As Double
Private animationDivisions As Integer
Private animationInterval As Integer
Private animateCCW As Boolean

'------------------------------------------------------ ENDS

Private mTicking As Boolean ' property Ticking
Private mThisCPUDetail As String ' property
Private mPointerAnimate As Boolean ' property PointerAnimate
Private mLocked As Boolean ' property Locked
Private mHidden As Boolean ' property Hidden
Private mThisOpacity As String ' property ThisOpacity
Private mSamplingInterval As String ' property SamplingInterval



'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
    Dim SInfo As SYSTEM_INFO
    
    On Error GoTo Class_Initialize_Error

    Set W = Cairo.WidgetBase
        W.BackColor = -1  'we don't use any BackColor here, transparent
        W.ImplementsHitTest = True 'when at False, no HitTest-Event would be triggered
                    
    Set tmrSampler = Nothing
    Set tmrSampler = New_c.Timer(1000, mTicking)
    
    animationDivisions = 12
    animationInterval = 40
    
    Set tmrAnimator = Nothing
    Set tmrAnimator = New_c.Timer(animationInterval, False)
    
    'Get the system information
    GetSystemInfo SInfo
    noOfProcessors = SInfo.dwNumberOrfProcessors
    
    mThisCPUDetail = obtainCPUDetails
    
'    Dim processorType As String
'    processorType = _
'        "Processor:" + Str$(SInfo.dwProcessorType) & vbCrLf & _
'        "Low memory address:" + Str$(SInfo.lpMinimumApplicationAddress) & vbCrLf & _
'        "High memory address:" + Str$(SInfo.lpMaximumApplicationAddress)
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cwOverlay"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Widget
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widget() As cWidgetBase
   On Error GoTo Widget_Error

    Set Widget = W

   On Error GoTo 0
   Exit Property

Widget_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widget of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Procedure : Widgets
' Author    : olaf schmidt
' Date      : 21/09/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Widgets() As cWidgets
   On Error GoTo Widgets_Error

    Set Widgets = W.Widgets

   On Error GoTo 0
   Exit Property

Widgets_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Widgets of Class Module cwOverlay"
End Property


'---------------------------------------------------------------------------------------
' Property : Ticking
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Ticking() As Boolean
   On Error GoTo TickingGet_Error

  Ticking = mTicking

   On Error GoTo 0
   Exit Property

TickingGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Ticking
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Ticking(ByVal newValue As Boolean)
    'Dim intervalValue As Integer: intervalValue = 0
    On Error GoTo TickingLet_Error

    If mTicking <> newValue Then mTicking = newValue Else Exit Property
    
    If mTicking Then
        tmrSampler.Enabled = True
    Else
        tmrSampler.Enabled = False
    End If

   On Error GoTo 0
   Exit Property

TickingLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Ticking of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : thisCPUDetail
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisCPUDetail() As String
   On Error GoTo thisCPUDetailGet_Error

    thisCPUDetail = mThisCPUDetail

   On Error GoTo 0
   Exit Property

thisCPUDetailGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisCPUDetail of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisCPUDetail
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisCPUDetail(ByVal newValue As String)
    On Error GoTo thisCPUDetailLet_Error

    If mThisCPUDetail <> newValue Then mThisCPUDetail = newValue Else Exit Property
        
    On Error GoTo 0
    Exit Property

thisCPUDetailLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisCPUDetail of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property : thisOpacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get thisOpacity() As String
   On Error GoTo thisOpacityGet_Error

    thisOpacity = mThisOpacity

   On Error GoTo 0
   Exit Property

thisOpacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : thisOpacity
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let thisOpacity(ByVal newValue As String)
    On Error GoTo thisOpacityLet_Error

    If mThisOpacity <> newValue Then mThisOpacity = newValue Else Exit Property
        
    On Error GoTo 0
    Exit Property

thisOpacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property thisOpacity of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property : pointerAnimate
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get pointerAnimate() As Boolean
   On Error GoTo pointerAnimateGet_Error

    pointerAnimate = mPointerAnimate

   On Error GoTo 0
   Exit Property

pointerAnimateGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : PointerAnimate
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let pointerAnimate(ByVal newValue As Boolean)
    On Error GoTo pointerAnimateLet_Error

    If mPointerAnimate <> newValue Then mPointerAnimate = newValue Else Exit Property
        
   On Error GoTo 0
   Exit Property

pointerAnimateLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property pointerAnimate of Class Module cwOverlay"
End Property ' the earth pointerAnimate flag

'---------------------------------------------------------------------------------------
' Property : samplingInterval
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get samplingInterval() As String
   On Error GoTo samplingIntervalGet_Error

    samplingInterval = mSamplingInterval

   On Error GoTo 0
   Exit Property

samplingIntervalGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : samplingInterval
' Author    :
' Date      : 09/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let samplingInterval(ByVal newValue As String)
    On Error GoTo samplingIntervalLet_Error

    If mSamplingInterval <> newValue Then mSamplingInterval = newValue Else Exit Property
    
    tmrSampler.Interval = Val(mSamplingInterval) * 1000
    If mPointerAnimate = True Then
        If tmrSampler.Interval = 1000 Then tmrSampler.Interval = 1200 ' take into account the animator
    End If
    
    On Error GoTo 0
    Exit Property

samplingIntervalLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property samplingInterval of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Locked() As Boolean
   On Error GoTo LockedGet_Error

   Locked = mLocked

   On Error GoTo 0
   Exit Property

LockedGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Locked
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Locked(ByVal newValue As Boolean)
   On Error GoTo LockedLet_Error

   If mLocked <> newValue Then mLocked = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

LockedLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Locked of Class Module cwOverlay"
End Property

'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Hidden() As Boolean
   On Error GoTo HiddenGet_Error

   Hidden = mHidden

   On Error GoTo 0
   Exit Property

HiddenGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlay"
End Property
'---------------------------------------------------------------------------------------
' Property  : Hidden
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Hidden(ByVal newValue As Boolean)
   On Error GoTo HiddenLet_Error

    If mHidden <> newValue Then mHidden = newValue Else Exit Property
    If mHidden = True Then
        W.Visible = False
        
        W.Refresh
    Else
        W.Visible = True
        
        W.Refresh
    End If

   On Error GoTo 0
   Exit Property

HiddenLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Hidden of Class Module cwOverlay"
End Property
'
'---------------------------------------------------------------------------------------------
' Procedure : W_HitTest
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   : This is run continuously for the whole widget.
'             For this Overlay-Widget alone, we make everything "fully click-through"
'---------------------------------------------------------------------------------------------
'
Private Sub W_HitTest(ByVal x As Single, ByVal Y As Single, HitResultHit As Boolean) 'ensure ClickThrough-behaviour in ImagePixels which are "fully Alpha"
   On Error GoTo W_HitTest_Error

   HitResultHit = False

   On Error GoTo 0
   Exit Sub

W_HitTest_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_HitTest of Class Module cwOverlay"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : W_Paint
' Author    : olaf schmidt
' Date      : 01/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub W_Paint(CC As RC6.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    Dim I As Long: I = 0
    Dim Arr As Variant
    Dim Key As String: Key = vbNullString
    Dim Pat As cCairoPattern
    Dim txtNoOfCores As String: txtNoOfCores = vbNullString
    Dim txtCPUPercent As String: txtCPUPercent = vbNullString
    Dim textDrawn As Boolean:   textDrawn = False
    Dim someOpacity As Double: someOpacity = 0
    Dim cream As String: cream = vbNullString
    
    cream = RGB(235, 217, 179) ' cannot create a constant from an RGB value
   
    On Error GoTo W_Paint_Error
    
    'set some Font- and ForeColor-Properties and draw some Text on the Widget
    W.FontName = "arial"
    W.ForeColor = cream
    W.FontSize = 18
    W.FontBold = True
    W.SelectFontSettingsInto CC
    W.Refresh
    someOpacity = (Val(mThisOpacity) / 100) * 0.75
    
    txtCPUPercent = CStr(Int(cpuPercent))
    txtNoOfCores = CStr(noOfProcessors)

    For I = 0 To cCollOverlayImages.Count - 1 ' this is the ExcludedPaths collection populated in cfAlpha by duplicating collPSDNonUIElements.
      Arr = cCollOverlayImages.ItemByIndex(I)
      Key = cCollOverlayImages.KeyByIndex(I)
      
      If IsArray(Arr) Then
             CC.Save
             'derive a Pattern-Obj from each retrieved ImageList-Surface-reference
             Set Pat = Cairo.ImageList(Key).CreateSurfacePattern
                         
            ' for specific keys perform rotation actions, clock hands &c
             Select Case Key

                 Case "secondshadow", "secondhand"
                    
                     ' draw the date and day in the windows just once to retain opacity
                    CC.DrawText 265, 103, 200, 200, txtCPUPercent, , , 2, True, , someOpacity
                    CC.DrawText 461, 298, 200, 200, txtNoOfCores, , , 2, True, , someOpacity
                    textDrawn = True
                    
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(11, 246) 'shift to the "rotation-point" within the Pattern
                    CC.TranslateDrawings W.Width * cxPerc, W.Height * cyPerc 'shift to the center-point of the Overlay-Widget-area
                    
                    CC.RotateDrawingsDeg SDeg
                    CC.Paint Arr(2), Pat 'render the Pattern with the current PSD-LayerAlpha, present in Arr(2)
                    
                 Case "redlamptrue"
                    
                    If cpuPercent >= 90 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1)) 'shift the Pattern via the original X,Y PSD-Offsets (stored in Arr(0) and Arr(1))
                        CC.Paint Arr(2), Pat 'render the Pattern
                    End If
                    
                 Case "redlampfalse"
                    
                    If cpuPercent < 90 Then
                        Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                        CC.Paint Arr(2), Pat ' render the Pattern
                    End If
                    
                 Case Else 'plain output for everything else ... this includes the reflection-layers,
                     
                    Set Pat.Matrix = Pat.Matrix.TranslateCoords(-Arr(0), -Arr(1))
                    CC.Paint Arr(2), Pat 'render the Pattern
                   
             End Select
             
             CC.Restore
      End If
     Next

   On Error GoTo 0
   Exit Sub

W_Paint_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure W_Paint of Class Module cwOverlay"
End Sub
 
    
'---------------------------------------------------------------------------------------
' Procedure : tmrAnimator_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : animate timer, once every few (40) milliseconds
'---------------------------------------------------------------------------------------
'
Private Sub tmrAnimator_Timer()
    
    On Error GoTo tmrAnimator_Timer_Error
    '
    ' move the pointer using SDeg
    '
    animationAmount = animationAmount + animationQuanta

    If animateCCW = True Then
        If SDeg > newSDeg Then
            SDeg = oldSDeg - animationAmount
        Else
            tmrAnimator.Enabled = False ' when the animation is complete then stop the animator
            animationAmount = 0
            Exit Sub
        End If
    Else
        If SDeg < newSDeg Then
            SDeg = oldSDeg + animationAmount
        Else
            tmrAnimator.Enabled = False
            animationAmount = 0
            Exit Sub
        End If
    End If

    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

tmrAnimator_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrAnimator_Timer of Class Module cwOverlay"
End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : tmrSampler_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub tmrSampler_Timer()
    Dim degDiff As Double: degDiff = 0
    
    On Error GoTo tmrSampler_Timer_Error

    If mPointerAnimate = True Then
        If tmrAnimator.Enabled = False Then
        
            ' oldSDeg stored
            oldSDeg = SDeg
                
            ' obtain the new cpu amount
            cpuPercent = CPU_Usage_Percent
            
            ' calculate the new cpu amount in degrees
            newSDeg = (cpuPercent * 3) + 30
            
            If oldSDeg = newSDeg Then Exit Sub
            
            ' calculate the difference
            degDiff = newSDeg - oldSDeg
            
            ' if the result is negative animate CCW
            If degDiff < 0 Then
                animateCCW = True
            Else
                animateCCW = False
            End If
            
            Select Case degDiff
                Case Is <= 5
                    animationDivisions = 4
                Case Is <= 10
                    animationDivisions = 8
                Case Is > 10
                    animationDivisions = 16
                Case Is > 20
                    animationDivisions = 24
            End Select
            
            ' divide that by 24 (one second duration)
            animationQuanta = Abs(degDiff / animationDivisions)
            
            tmrAnimator.Enabled = True
            
        End If
    Else
        cpuPercent = CPU_Usage_Percent
        SDeg = (cpuPercent * 3) + 30
    End If
    
    W.Parent.Refresh

   On Error GoTo 0
   Exit Sub

tmrSampler_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampler_Timer of Class Module cwOverlay"
End Sub

' ----------------------------------------------------------------
' Procedure Name: CPU_Usage_Percent
' Purpose:
' Procedure Kind: Function
' Procedure Access: Private
' Return Type: Long
' Author: https://www.tek-tips.com/userinfo.cfm?member=Hypetia
' Date: 09/01/2024
' ----------------------------------------------------------------
Private Function CPU_Usage_Percent() As Currency
    Dim idleTimeLive As Currency: idleTimeLive = 0
    Dim kernelTimeLive As Currency: kernelTimeLive = 0
    Dim usermodeTimeLive As Currency: usermodeTimeLive = 0
    Static idleTimeStored As Currency
    Static kernelTimeStored As Currency
    Static usermodeTimeStored As Currency
    
    On Error GoTo CPU_Usage_Percent_Error

    GetSystemTimes idleTimeLive, kernelTimeLive, usermodeTimeLive
    
    idleTimeStored = idleTimeLive - idleTimeStored
    kernelTimeStored = kernelTimeLive - kernelTimeStored
    usermodeTimeStored = usermodeTimeLive - usermodeTimeStored
    
    CPU_Usage_Percent = ((kernelTimeStored + usermodeTimeStored - idleTimeStored) / (kernelTimeStored + usermodeTimeStored)) * 100
    
    idleTimeStored = idleTimeLive
    kernelTimeStored = kernelTimeLive
    usermodeTimeStored = usermodeTimeLive
    
    On Error GoTo 0
    Exit Function

CPU_Usage_Percent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CPU_Usage_Percent, line " & Erl & "."

End Function




' ----------------------------------------------------------------
' Procedure Name: obtainCPUDetails
' Purpose:
' Procedure Kind: Function
' Procedure Access: Public
' Return Type: String
' Author: beededea
' Date: 21/01/2024
' ----------------------------------------------------------------
Function obtainCPUDetails() As String
    Dim result As String
    Dim objCPUItem As Object
    Dim objCPU As Object
    
    On Error GoTo obtainCPUDetails_Error

    Set objCPUItem = GetObject("winmgmts:").InstancesOf("Win32_Processor")

    If Err.Number <> 0 Then
        result = "Error getting Win32_Processor " & _
        "information." & vbCrLf
    Else
        For Each objCPU In objCPUItem
            result = result & "Description: " & Trim$(objCPU.Name) & vbCrLf
            result = result & "Frequency (MHz): " & objCPU.MaxClockSpeed & vbCrLf
            result = result & "CPU-ID: " & objCPU.ProcessorId & vbCrLf
            result = result & vbCrLf
            Debug.Print result
        Next

        Set objCPUItem = Nothing
    End If

    obtainCPUDetails = result
    
    On Error GoTo 0
    Exit Function

obtainCPUDetails_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainCPUDetails, line " & Erl & "."

End Function



' ----------------------------------------------------------------
' Procedure Name: obtainCoreCPUPercent
' Purpose: not currently used
' Procedure Kind: Sub
' Procedure Access: Public
' Author: beededea
' Date: 21/01/2024
' ----------------------------------------------------------------
Sub obtainCoreCPUPercent()
    Dim strComputer$
    Dim I&
    Dim objWMIService As Object, colItems As Object, objItem As Object
    
        On Error GoTo obtainCoreCPUPercent_Error

    strComputer = "." 'Local machine, can be adjusted to access remote    workstations
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_PerfFormattedData_PerfOS_Processor", , 48)
    For Each objItem In colItems
        I = I + 1 ' 1 = cpu all
        Debug.Print "-----------------------------------"
        Debug.Print "Processor " & I
        'Debug.Print "-----------------------------------"
        Debug.Print "usage: " & objItem.PercentProcessorTime
    Next
    Set objWMIService = Nothing
    Set colItems = Nothing
    Set objItem = Nothing
    
    On Error GoTo 0
    Exit Sub

obtainCoreCPUPercent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainCoreCPUPercent, line " & Erl & "."

End Sub

