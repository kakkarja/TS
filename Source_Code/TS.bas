Attribute VB_Name = "Module1"
Option Explicit

'''SETUP TO MAKE TIME SPEECH AND THE ALARM SPEECH FUNCTIONS'''

Dim SWT As Date
Dim Ts As Date
Dim Ws As Worksheet
Dim H As Integer
Dim M As Integer
Dim s As Integer
Dim RT As Date
Dim TSU As UserForm
Dim Tx As Workbook, TXA As Workbook
Dim TXW As Worksheet, TXWA As Worksheet


Private Sub TSS()
    TimeSpeech.Show
End Sub

Sub TimeSp()
Set TSU = TimeSpeech
    If TSU.ActiveControl Is Nothing Then GoTo bye
    TV
    Application.Speech.Speak Text:= _
    "Now is" & Str(Hour(Time)) & "O'Clock" & Str(Minute(Time)) & _
    "Minutes" & Str(Second(Time)) & "Seconds" & TimeSpeech.SayAt.Text
    With TimeSpeech
        If Time > .SoundTime Then
            If .StopRoll.Locked = True Then
                .StopRoll.Locked = False
            End If
        End If
    End With
bye:
Set TSU = Nothing
End Sub

Private Sub TV()
        H = TimeSpeech.Jam
        M = TimeSpeech.Menit
        s = TimeSpeech.Detik
    TikTo H, M, s
bye:
H = 0
M = 0
s = 0
End Sub

Private Sub TikTo(Hs As Integer, Ms As Integer, Ss As Integer)
    RT = _
    Time() + TimeValue(Hs & ":" & Ms & ":" & Ss)
    If RT <> Time Then
        TimeSpeech.SoundTime.Caption = RT
        Application.OnTime RT, "TimeSp"
    Else
        TimeSpeech.SoundTime.Caption = RT
        Application.OnTime RT, "TimeSp"
        Application.OnTime RT, "TimeSp", , False
    End If
RT = 0
End Sub

Private Sub TimeScroll(Sec As Date)
    Application.OnTime Now + Sec, "WoTi"

End Sub
Private Sub TimeScrollStop(Sec As Date)
    Application.OnTime Now + Sec, "WoTi", , False
End Sub

Sub WoTi()
Fx1
Set TSU = TimeSpeech
Set Ws = ActiveSheet
    SWT = TimeValue("00:00:01")
    If TSU.ActiveControl Is Nothing Then GoTo bye
    If Not Ws.Range("A1") = "" Then
        Ws.Range("A1").Value = _
        Format(Time(), "hh:mm:ss")
        TimeScroll (SWT)
        GoTo bye
    Else
        TimeScroll (SWT)
        TimeScrollStop (SWT)
    End If
    Ws.Range("A1").Value = _
    Format(Time(), "hh:mm:ss")
bye:
SWT = 0
Set Ws = Nothing
Set TSU = Nothing
Fx2
End Sub


Private Sub Fast(SU As Boolean, DS As Boolean, C As String, EE As Boolean)
    Application.ScreenUpdating = SU
    Application.DisplayStatusBar = DS
    Application.Calculation = C
    Application.EnableEvents = EE
End Sub
Private Sub Fx1()
Call Fast(False, True, xlCalculationManual, False)
End Sub
Private Sub Fx2()
Call Fast(True, True, xlCalculationAutomatic, True)
End Sub

Sub MakeTimeS()
Dim Q1 As Variant
    If Range("A1:K10").MergeCells = True Then GoTo _
    bye

    Q1 = MsgBox( _
    "You are going to create Time Speech." & _
    " Make sure this worksheet is clear." _
    , vbOKCancel, "Time Speech")
    Select Case Q1
        Case Is = vbOK
            With Range("A1:K10")
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .WrapText = True
                .Merge
            End With
            With Range("A1")
                .Value = Format(Time(), "HH:MM:SS")
                With .Font
                    .Name = "Eras Bold ITC"
                    .Size = 105
                    .ColorIndex = 2
                End With
                With .Interior
                    .ColorIndex = 5
                    .Pattern = 16
                    .PatternColorIndex = 33
                End With
            End With
            Star
            Range("A1").Select
            ActiveSheet.Shapes("StarTime").OnAction = "TSS"
            TSS
        Case Is = vbCancel
            MsgBox "Please create new worksheet" & _
            " for Time Speech.", vbInformation, _
            "Time Speech"
    End Select
bye:
Q1 = vbNullString
End Sub
Private Sub Star()

    ActiveSheet.Shapes.AddShape(msoShape5pointStar, (Cells(1, 12).Left) + 2, 1, 14, 12.5).Select
    With Selection.ShapeRange.ThreeD
        .SetPresetCamera (msoCameraOrthographicFront)
        .RotationX = 0
        .RotationY = 0
        .RotationZ = 0
        .FieldOfView = 0
        .LightAngle = 145
        .PresetLighting = msoLightRigBalanced
        .PresetMaterial = msoMaterialWarmMatte
        .Depth = 0
        .ContourWidth = 0
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 15
        .BevelTopDepth = 3
        .BevelBottomType = msoBevelNone
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 3.5
        .OffsetX = 1.3471114791E-16
        .OffsetY = 2.2
        .RotateWithShape = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0.6800000072
        .Size = 100
    End With
    With Selection.ShapeRange.ThreeD
        .SetPresetCamera (msoCameraOrthographicFront)
        .RotationX = 0
        .RotationY = 0
        .RotationZ = 0
        .FieldOfView = 0
        .LightAngle = 145
        .PresetLighting = msoLightRigBalanced
        .PresetMaterial = msoMaterialWarmMatte
        .Depth = 0
        .ContourWidth = 0
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 15
        .BevelTopDepth = 3
        .BevelBottomType = msoBevelNone
    End With
    With Selection.ShapeRange.Shadow
        .Type = msoShadow25
        .Visible = msoTrue
        .Style = msoShadowStyleOuterShadow
        .Blur = 3.5
        .OffsetX = 1.3471114791E-16
        .OffsetY = 2.2
        .RotateWithShape = msoTrue
        .ForeColor.RGB = RGB(0, 0, 0)
        .Transparency = 0.6800000072
        .Size = 100
    End With
    Selection.ShapeRange.Line.Visible = msoFalse
    Selection.ShapeRange.Name = "StarTime"
End Sub

'''TIME SPEECH USERFORM'''

Private Sub Detik_Change()
    SpinButton3.Value = Detik.Text
End Sub

Private Sub Initiate_Click()
    If Detik = 0 And Menit = 0 And Jam = 0 Then Exit Sub
    StopRoll.Locked = True
    If OnSpeech = True And SoundTime.Caption = "" Then
        Call TV
    End If
End Sub

Private Sub Jam_Change()
    SpinButton1.Value = Jam.Text
End Sub

Private Sub Menit_Change()
    SpinButton2.Value = Menit.Text
End Sub

Private Sub SpinButton1_Change()
    Jam = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
    Menit = SpinButton2.Value
End Sub

Private Sub SpinButton3_Change()
    Detik = SpinButton3.Value
End Sub

Private Sub TikTo(Hs As Integer, Ms As Integer, Ss As Integer)
Dim RT As Date
    RT = _
    Time() + TimeValue(Hs & ":" & Ms & ":" & Ss)
    SoundTime.Caption = RT
    Application.OnTime RT, "TimeSp"
RT = 0
End Sub

Private Sub TV()
Dim H As Integer
Dim M As Integer
Dim s As Integer
    H = Jam
    M = Menit
    s = Detik
    If TimeValue(H & ":" & M & ":" & s) = TimeValue("00:00:00") Then
        Exit Sub
    Else
        TikTo H, M, s
    End If
bye:
H = 0
M = 0
s = 0
End Sub

Private Sub StopRoll_Click()
    On Error Resume Next
        If Initiate.Locked = False Then
            Initiate.Locked = True
        End If
        If Not Range("A1") = "" Then
            Range("A1").Value = ""
            SoundTime.Caption = ""
            OnSpeech.Locked = False
            OnSpeech = False
            OnSpeech.Locked = True
        End If
End Sub

Private Sub TimeRoll_Click()
    If Range("A1") = "" Then Range("A1").Value = Time
    If Initiate.Locked = True Then
        Initiate.Locked = False
    End If
    If OnSpeech = False Then
        OnSpeech.Locked = False
        OnSpeech = True
        OnSpeech.Locked = True
        WoTi
    Else
        Exit Sub
    End If
End Sub

Private Sub UserForm_Initialize()
Dim Ct As LongPtr, k As LongPtr
    ActiveSheet.Unprotect Environ("userprofile")
    Jam.Locked = True
    Menit.Locked = True
    Detik.Locked = True
    OnSpeech.Locked = True
    Initiate.Locked = True
    For Ct = 1 To Worksheets.Count
        If Left(Worksheets(Ct).Name, 11) _
        = "Time Speech" Then
            k = k + 1
        End If
    Next Ct
    If Left(ActiveSheet.Name, 11) <> "Time Speech" Then
        If k > 0 Then
            With ActiveWorkbook
                If .ProtectStructure = False Then
                    ActiveSheet.Name = "Time Speech " & k + 1
                Else
                    MsgBox "Please unlock the workbook first", vbInformation, _
                    "Time Speech"
                End If
            End With
        Else
            With ActiveWorkbook
                If .ProtectStructure = False Then
                    ActiveSheet.Name = "Time Speech"
                Else
                    MsgBox "Please unlock the workbook first", vbInformation, _
                    "Time Speech"
                End If
            End With
        End If
    End If
Ct = 0
k = 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If OnSpeech = True Then
        Cancel = True
    End If
End Sub

Private Sub UserForm_Terminate()
    With ActiveSheet
        .Protect Environ("userprofile")
        .EnableSelection = xlNoSelection
    End With
End Sub
