Attribute VB_Name = "modMain"

Option Explicit

Global objDX As New DirectX7
Global objDXEvent As DirectXEvent
Global objDI As DirectInput
Global objDIDev As DirectInputDevice
Global objDIDevC As DirectInputDevice
Global objDIDevK As DirectInputDevice
Global objDIEnum As DirectInputEnumDevices
Global joyCaps As DIDEVCAPS
Global js As DIJOYSTATE
Global diState As DIKEYBOARDSTATE
Global mState As DIMOUSESTATE
Global NumItems As Integer


'FFB
Global diEnumObjects As DirectInputEnumDeviceObjects       'DirectInput enumeration for objects on a device object
Global diDevObjInstance As DirectInputDeviceObjectInstance 'DirectInput object on a device object
Global diEffEnum As DirectInputEnumEffects                 'DirectInput enumeration for force feedback effects object
Global diFFEffect() As DirectInputEffect                   'Force feedback effects object
Global diEffectType As Long                                'Will be used to store the type of effect an effect object is
Global diFFStaticParams As Long                            'Will be used to store the static parameters of an effect object
Global EffectParams() As Long
Global Prop As DIPROPLONG
'
Global Effect As DIEFFECT

'Settings
Global DiProp_Dead As DIPROPLONG
Global DiProp_Range As DIPROPRANGE
Global DiProp_Saturation As DIPROPLONG

Global AxisPresent(1 To 11) As Boolean
Global joyIndex As Long
Global preIndex As Long
Global jLoop As Long


Global didInit As Boolean
Global didLoad As Boolean

Global oldTime As Long
Global oldTime2 As Long
Global ffinCar As Boolean
Global ffWait As Long
Global oinCar As Boolean
Global aMode As Long


Global manBut As Long

Global Joysticks() As String
Global JoyCount As Long

Public Const BufferSize = 10
Global diDeviceData(1 To BufferSize) As DIDEVICEOBJECTDATA

Public EventHandle As Long
Public EventHandle2 As Long
Public EventHandle3 As Long


Global ctlKeys(2, 53) As Long
Global ctlInputs(24) As Integer
Global orawInputs(308) As Integer
Global rawInputs(308) As Integer
Global zrawInputs(308) As Integer
Global ozrawInputs(308) As Integer
Global olrawInputs(308) As Integer
Global ctlIndex As Long
Global ctlSecond As Boolean
Global bDown(53) As Boolean


Global cPoint As Long
Global cPoint2 As Long
Global bListen As Boolean

Global Old1
Global Old2
Global Old3
Global Old4
Global r1
Global r2
Global r3
Global r4
Global rr
Global fr
Global rl
Global fl
Global rr2
Global fr2
Global rl2
Global fl2
Global sAll
'3
'3 =6
'2 =8
'3 =11


Global testing As Boolean

  
Public Function InitDirectInput()
On Error Resume Next
  ' Create DirectInput and set up the mouse
    Set objDI = objDX.DirectInputCreate
    
EventHandle = objDX.CreateEvent(frmMain)
EventHandle2 = objDX.CreateEvent(frmMain)
EventHandle3 = objDX.CreateEvent(frmMain)

  Set objDIDev = objDI.CreateDevice("guid_SysMouse")
  Call objDIDev.SetCommonDataFormat(DIFORMAT_MOUSE)
  Call objDIDev.SetCooperativeLevel(frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
  ' Set the buffer size
  Prop.lHow = DIPH_DEVICE
  Prop.lObj = 0
  Prop.lData = BufferSize
  Prop.lSize = Len(Prop)
  Call objDIDev.SetProperty("DIPROP_BUFFERSIZE", Prop)

  ' Ask for notifications
  Call objDIDevK.SetEventNotification(EventHandle3)
  
  ' Acquire the mouse
  objDIDev.Acquire


  'Acquire Keyboard
  Set objDIDevK = objDI.CreateDevice("GUID_SysKeyboard")
  Call objDIDevK.SetCommonDataFormat(DIFORMAT_KEYBOARD)
  Call objDIDevK.SetCooperativeLevel(frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
  ' Ask for notification of events
  Call objDIDevK.SetEventNotification(EventHandle2)

  objDIDevK.Acquire
    
    'Acquire Joystick 0
    'Set objDIEnum = objDI.GetDIDevices(DI8DEVCLASS_GAMECTRL, DIEDFL_ATTACHEDONLY)
    Set objDIEnum = objDI.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
    JoyCount = objDIEnum.GetCount
    If Not objDIEnum.GetCount = 0 Then
    ReDim Joysticks(1 To objDIEnum.GetCount)
    Dim i As Integer
    For i = 1 To objDIEnum.GetCount
        Joysticks(i) = objDIEnum.GetItem(i).GetInstanceName
    Next
    'ReDim Joysticks(0)
    End If
    

  
    'AcquireJoystick joyIndex

'SetDefaultText



'Screen.MousePointer = vbNormal
End Function
Sub IdentifyAxes(diDev As DirectInputDevice)
On Error Resume Next
   ' It's not enough to count axes; we need to know which in particular
   ' are present.
   
   Dim didoEnum As DirectInputEnumDeviceObjects
   Dim dido As DirectInputDeviceObjectInstance
   Dim i As Integer
   
   For i = 1 To 8
     AxisPresent(i) = False
   Next
   
   ' Enumerate the axes
   Set didoEnum = diDev.GetDeviceObjectsEnum(DIDFT_AXIS)
   
   ' Check data offset of each axis to learn what it is
   Dim sGuid As String, SliderCount As Long, POVCount As Long
   For i = 1 To didoEnum.GetCount
     Set dido = didoEnum.GetItem(i)
         Select Case dido.GetOfs
            Case DIJOFS_X
              AxisPresent(1) = True
            Case DIJOFS_Y
              AxisPresent(2) = True
            Case DIJOFS_Z
              AxisPresent(3) = True
            Case DIJOFS_RX
              AxisPresent(4) = True
            Case DIJOFS_RY
              AxisPresent(5) = True
            Case DIJOFS_RZ
              AxisPresent(6) = True
            Case DIJOFS_SLIDER0
              AxisPresent(7) = True
            Case DIJOFS_SLIDER1
              AxisPresent(8) = True
         End Select
 
   Next
End Sub
Public Function AcquireJoystick(jIndex As Long)
On Error Resume Next


    If Not objDIDevC Is Nothing Then
      objDIDevC.Unacquire
    End If
    
    Dim fCount As Long
    Dim intCount As Integer
    Dim ctlCount As Long
    Dim ctIndex() As Long
    
    
    'Create the joystick device
    Set objDIDevC = Nothing
    Set objDIDevC = objDI.CreateDevice(objDIEnum.GetItem(jIndex).GetGuidInstance)
    objDIDevC.SetCommonDataFormat DIFORMAT_JOYSTICK
    objDIDevC.SetCooperativeLevel frmMain.hwnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
    'Call objDIDevC.SetEventNotification(EventHandle)
    ' Find out what device objects it has
    objDIDevC.GetCapabilities joyCaps
    Call IdentifyAxes(objDIDevC)
    
    ' Ask for notification of events
    Call objDIDevC.SetEventNotification(EventHandle)
    
    SetPropRange 'set range :)
    
    ' Set deadzone for X and Y axis to 10 percent of the range of travel
    With DiProp_Dead
        .lData = 0
        .lHow = DIPH_BYOFFSET
        .lSize = Len(DiProp_Dead)
        .lObj = DIJOFS_X
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        .lObj = DIJOFS_Y
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        .lObj = DIJOFS_Z
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        .lObj = DIJOFS_RX
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        .lObj = DIJOFS_RY
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
        .lObj = DIJOFS_RZ
        objDIDevC.SetProperty "DIPROP_DEADZONE", DiProp_Dead
    End With
    
    ' Set saturation zones for X and Y axis to 5 percent of the range
    With DiProp_Saturation
        .lData = 10000
        .lHow = DIPH_BYOFFSET
        .lSize = Len(DiProp_Saturation)
        .lObj = DIJOFS_X
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        .lObj = DIJOFS_Y
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        .lObj = DIJOFS_Z
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        .lObj = DIJOFS_RX
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        .lObj = DIJOFS_RY
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
        .lObj = DIJOFS_RZ
         objDIDevC.SetProperty "DIPROP_SATURATION", DiProp_Saturation
         
    End With
    

    
    
    'Force Feedback
    ReDim ctIndex(frmMain.fType.UBound)
    For ctlCount = 0 To frmMain.fType.UBound
        frmMain.fType(ctlCount).Clear
        ctIndex(ctlCount) = frmMain.fType(ctlCount).ListIndex
    Next ctlCount
    
    'Enumerate available Effects
    Set diEffEnum = objDIDevC.GetEffectsEnum(DIEFT_ALL)
        

          For intCount = 1 To diEffEnum.GetCount      'Loop through all the effects
            diEffectType = diEffEnum.GetType(intCount) And &HFF
                                                    'Filter out the major type of effect it is
            diFFStaticParams = diEffEnum.GetStaticParams(intCount)
                                                    'Get the static parameters of this effect
            If (diEffectType = DIEFT_HARDWARE) And _
            (diFFStaticParams And DIEP_TYPESPECIFICPARAMS) <> 0 Then
                                                    'If this is a hardware effect that has type-specific parameters,
                GoTo ignore                         'ignore it and skip to the next effect
            ElseIf diEffectType = DIEFT_CUSTOMFORCE Then
                                                    'If this effect is a custom effect,
                GoTo ignore                         'ignore it and skip to the next effect
            End If
            fCount = fCount + 1
            'Debug.Print diEffEnum.GetName(intCount)
            

            
            For ctlCount = 0 To frmMain.fType.Count
                frmMain.fType(ctlCount).AddItem diEffEnum.GetName(intCount)
            Next ctlCount
                                                    'Add this effect to the listbox, displaying the name of the
                                                    'effect
            ReDim Preserve EffectParams(fCount - 1)
                                                    'Redimension the array that stores the type of effect this
                                                    'effect is
            EffectParams(fCount - 1) = diEffectType
                                                    'store the type of effect in the EffectParams array
            ReDim Preserve diFFEffect(fCount - 1)
    
                                                    'Redimension the effect object array
            
                    'Catch any errors when creating the effect object
            Set diFFEffect(UBound(diFFEffect)) = objDIDevC.CreateEffect(diEffEnum.GetEffectGuid(intCount), _
            CreateFFEffect(intCount))               'Create the effect, using the return value from the
                                                    'CreateFFEffect function, which returns a generic effect
                                                    'structure
            diFFEffect(UBound(diFFEffect)).Unload   'Since creating an effect automtically downloads it, unload
                                                    'it so we don't run out of room on the device.
                                                    
                                                    

            
ignore:
        Next

    'set old stuff
    For ctlCount = 0 To frmMain.fType.UBound
        If ctIndex(ctlCount) = -1 Then
        frmMain.fType(ctlCount).ListIndex = 0
        Else
        frmMain.fType(ctlCount).ListIndex = ctIndex(ctlCount)
        End If
    Next ctlCount
    frmMain.SetOld

'SetRTCGAIN
End Function
Public Function CreateFFEffect(Index As Integer) As DIEFFECT
On Error Resume Next
    With CreateFFEffect
        .lDuration = -1                               'Infinite duration
        .lGain = 10000                                  'Full gain
        .lSamplePeriod = 0                              'Default sample period
        .lTriggerButton = DIEB_NOTRIGGER                     'Use button 1 on the joystick as the trigger
        'Debug.Print Button(1)
        .lTriggerRepeatInterval = -1                    'Turn off trigger repeat interval
                
        .constantForce.lMagnitude = 10000               'Make the magnitude of a constant force effect at full
        .rampForce.lRangeStart = 0                      'Make the magnitude at the start of a ramp force 0
        .rampForce.lRangeEnd = 0                        'Make the magnitude at the end of a ramp force 0
        .conditionFlags = DICONDITION_USE_BOTH_AXES     'Use both axis when using a conditional force
        With .conditionX                                'For the X axis
            .lDeadBand = 0                              'Make an effect with no deadband
            .lNegativeSaturation = 10000                'Turn the negative saturation all the way up
            .lOffset = 0                                'Zero the offset
            .lPositiveSaturation = 10000                'Turn the positive saturation all the way up
        End With
        With .conditionY                                'For the Y axis
            .lDeadBand = 0                              'Make an effect with no deadband
            .lNegativeSaturation = 10000                'Turn the negative saturation all the way up
            .lOffset = 0                                'Zero the offset
            .lPositiveSaturation = 10000                'Turn the positive saturation all the way up
        End With
        With .periodicForce                             'For a periodic force
            .lMagnitude = 10000                         'Turn the magnitude of the force all the way up
            .lOffset = 0                                'Zero the offset
            .lPeriod = 1                                'Set the length of a cycle to 1
            .lPhase = 0                                 'Zero the starting phase. Phase is something that has very
                                                        'limited support, so changing this parameter will almost always
                                                        'fail. Be prepared to catch the error this will return.
        End With
    End With
End Function

Sub SetPropRange()
    ' NOTE Some devices do not let you set the range
    On Local Error Resume Next

    ' Set range for all axes
    With DiProp_Range
        .lHow = DIPH_DEVICE
        .lMin = 0
        .lMax = 510
        .lSize = LenB(DiProp_Range)
    End With
    objDIDevC.SetProperty "DIPROP_RANGE", DiProp_Range
End Sub

Public Sub JoyPoll()
Dim i As Long
On Error Resume Next
    'If objDIDevC Is Nothing Then Exit Sub
objDIDevC.Poll
    '' Get the device info

    objDIDevC.GetDeviceStateJoystick js
    If Err.Number = DIERR_NOTACQUIRED Or Err.Number = DIERR_INPUTLOST Then
        objDIDevC.Acquire
        Exit Sub
    End If



End Sub


Public Function ReInit()
On Error Resume Next
  Set objDI = objDX.DirectInputCreate

  
  'Mouse Stuff
  'Set objDIDev = objDI.CreateDevice("guid_SysMouse")
  'Call objDIDev.SetCommonDataFormat(DIFORMAT_MOUSE)
  'Call objDIDev.SetCooperativeLevel(frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
  ' Set the buffer size
  'Dim diProp As DIPROPLONG
  'diProp.lHow = DIPH_DEVICE
  'diProp.lObj = 0
  'diProp.lData = BufferSize
  'diProp.lSize = Len(diProp)
  'Call objDIDev.SetProperty("DIPROP_BUFFERSIZE", diProp)

  ' Ask for notifications
  'EventHandle = objDX.CreateEvent(frmMain)
  'Call objDIDev.SetEventNotification(EventHandle)
  'Call objDIDevK.SetEventNotification(EventHandle)
  
  ' Acquire the mouse
  'objDIDev.Acquire

  
  'Keyboard Stuff
  Set objDIDevK = objDI.CreateDevice("GUID_SysKeyboard")
  Call objDIDevK.SetCommonDataFormat(DIFORMAT_KEYBOARD)
  Call objDIDevK.SetCooperativeLevel(frmMain.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
  ' Ask for notification of events
  Call objDIDevK.SetEventNotification(EventHandle2)
  objDIDevK.Acquire
  
    'Joystick Stuff
    Set objDIEnum = objDI.GetDIEnumDevices(DIDEVTYPE_JOYSTICK, DIEDFL_ATTACHEDONLY)
    JoyCount = objDIEnum.GetCount
    If Not objDIEnum.GetCount = 0 Then
    ReDim Joysticks(1 To objDIEnum.GetCount)
    Dim i As Integer
    For i = 1 To objDIEnum.GetCount
        Joysticks(i) = objDIEnum.GetItem(i).GetInstanceName
    Next
    'ReDim Joysticks(0)
    End If


AcquireJoystick joyIndex


End Function

Public Function DeInit()
On Error Resume Next
If EventHandle <> 0 Then objDX.DestroyEvent EventHandle
If EventHandle2 <> 0 Then objDX.DestroyEvent EventHandle2
objDIDevK.Unacquire
objDIDevC.Unacquire
End Function
Public Function SetToDefaultSysColors(frm As Form, sErrorMsg As _
String) As Boolean
    On Error GoTo ErrSetToDefaultSysColors
    Dim cntrl As Control
    
    'init to failure
    SetToDefaultSysColors = False
    
    'set form properties
    frm.Font = vbWindowText
    
    'set control properties
    For Each cntrl In frm.Controls
        If (TypeOf cntrl Is ComboBox) Or (TypeOf cntrl Is _
         TextBox) Then
            cntrl.BackColor = vbWindowBackground
            cntrl.Font = vbWindowText
        'NOTE: MUST HAVE REFERENCE TO FLEX GRID CONTROL
        'FOR THIS TO WORK
        'ElseIf TypeOf cntrl Is MSHFlexGrid Then
         '   cntrl.CellBackColor = vbWindowBackground
          '  cntrl.CellFontName = vbWindowText
        ElseIf (TypeOf cntrl Is Label Or TypeOf cntrl Is CheckBox Or TypeOf cntrl Is OptionButton Or TypeOf cntrl Is Frame) Then
            cntrl.Font = vbWindowText
        ElseIf TypeOf cntrl Is ListBox Then
            cntrl.Font = vbWindowText
            cntrl.BackColor = vbWindowBackground
        ElseIf TypeOf cntrl Is CommandButton Then
            cntrl.Font = vbWindowText
        End If
    Next cntrl
    
    'indicate success
    SetToDefaultSysColors = True
    
SetToDefaultSysColorsExit:
    Exit Function
ErrSetToDefaultSysColors:
    SetToDefaultSysColors = False
    Resume SetToDefaultSysColorsExit
End Function

Public Function InterpretPOV(pos2, outputUp, outputDown, outputLeft, outputRight)
On Error Resume Next
Dim pos As Long
If pos2 = -1 Then outputUp = 0: outputDown = 0: outputLeft = 0: outputRight = 0: Exit Function
If pos2 > 360 Then
pos = pos2 / 100
End If
Select Case pos
Case 0: outputUp = 255: outputDown = 0: outputLeft = 0: outputRight = 0
Case 45: outputUp = 255: outputDown = 0: outputLeft = 0: outputRight = 255
Case 90:  outputUp = 0: outputDown = 0: outputLeft = 0: outputRight = 255
Case 135: outputUp = 0: outputDown = 255: outputLeft = 0: outputRight = 255
Case 180:  outputUp = 0: outputDown = 255: outputLeft = 0: outputRight = 0
Case 225: outputUp = 0: outputDown = 255: outputLeft = 255: outputRight = 0
Case 270: outputUp = 0: outputDown = 0: outputLeft = 255: outputRight = 0
Case 315: outputUp = 255: outputDown = 0: outputLeft = 255: outputRight = 0
End Select
End Function

Public Function WriteControls(Block As Boolean)
On Error Resume Next
If didInit = False Then Exit Function
For jLoop = 0 To 307
If rawInputs(jLoop) < 0 Then rawInputs(jLoop) = 0
If zrawInputs(jLoop) < 0 Then zrawInputs(jLoop) = 0
Next jLoop



'SetHex &H541C40, "C20400"
'SetHex &H541A70, "C20800"

'SetHex &H541C40, "83EC30"
'SetHex &H541A70, "568BF1"


If (rawInputs(ctlKeys(0, 53)) Or rawInputs(ctlKeys(1, 53))) > 0 And GetByte(&HB7CB49) = 0 Then SetByte &HBA677B, 1


If inCar = False Then

    'jump/zoom in
    If inZoom Then
        If Not (zrawInputs(ctlKeys(0, 11)) Or zrawInputs(ctlKeys(1, 11))) = (ozrawInputs(ctlKeys(0, 11)) Or ozrawInputs(ctlKeys(1, 11))) Then
        SetInteger (cPoint2 + &H1C), (zrawInputs(ctlKeys(0, 11)) Or zrawInputs(ctlKeys(1, 11)))
        End If
        rawInputs(ctlKeys(0, 11)) = 0
        rawInputs(ctlKeys(1, 11)) = 0
        
    Else
        If Not (rawInputs(ctlKeys(0, 11)) Or rawInputs(ctlKeys(1, 15))) = (olrawInputs(ctlKeys(0, 15)) Or olrawInputs(ctlKeys(1, 15))) Then
        SetInteger (cPoint2 + &H1C), (rawInputs(ctlKeys(0, 15)) Or rawInputs(ctlKeys(1, 15)))
        End If
    
    End If


    'sprint/zoom out
    If inZoom Then
        If Not (zrawInputs(ctlKeys(0, 12)) Or zrawInputs(ctlKeys(1, 12))) = (ozrawInputs(ctlKeys(0, 12)) Or ozrawInputs(ctlKeys(1, 12))) Then
        SetInteger (cPoint2 + &H20), (zrawInputs(ctlKeys(0, 12)) Or zrawInputs(ctlKeys(1, 12)))
        End If
        rawInputs(ctlKeys(0, 12)) = 0
        rawInputs(ctlKeys(1, 12)) = 0
        
    Else
        If Not (rawInputs(ctlKeys(0, 16)) Or rawInputs(ctlKeys(1, 16))) = (olrawInputs(ctlKeys(0, 16)) Or olrawInputs(ctlKeys(1, 16))) Then
        SetInteger (cPoint2 + &H20), (rawInputs(ctlKeys(0, 16)) Or rawInputs(ctlKeys(1, 16)))
        End If
    End If
    
    


    'Foot

    'Walk Left/Right
    If Not (rawInputs(ctlKeys(0, 10)) Or rawInputs(ctlKeys(1, 10))) = (olrawInputs(ctlKeys(0, 10)) Or olrawInputs(ctlKeys(1, 10))) Or Not (rawInputs(ctlKeys(0, 9)) Or rawInputs(ctlKeys(1, 9))) = (olrawInputs(ctlKeys(0, 9)) Or olrawInputs(ctlKeys(1, 9))) Then
    If (rawInputs(ctlKeys(0, 9)) Or rawInputs(ctlKeys(1, 9))) > (rawInputs(ctlKeys(0, 10)) Or rawInputs(ctlKeys(1, 10))) Then
        SetInteger (cPoint2 + &H0), -Fix((rawInputs(ctlKeys(0, 9)) Or rawInputs(ctlKeys(1, 9))) / 2)
    ElseIf (rawInputs(ctlKeys(0, 9)) Or rawInputs(ctlKeys(1, 9))) < (rawInputs(ctlKeys(0, 10)) Or rawInputs(ctlKeys(1, 10))) Then
        SetInteger (cPoint2 + &H0), Fix((rawInputs(ctlKeys(0, 10)) Or rawInputs(ctlKeys(1, 10))) / 2)
    Else
        SetInteger (cPoint2 + &H0), 0
    End If
    End If
    
    'Walk Forward/Backward
    If Not (rawInputs(ctlKeys(0, 8)) Or rawInputs(ctlKeys(1, 8))) = (olrawInputs(ctlKeys(0, 8)) Or olrawInputs(ctlKeys(1, 8))) Or Not (rawInputs(ctlKeys(0, 7)) Or rawInputs(ctlKeys(1, 7))) = (olrawInputs(ctlKeys(0, 7)) Or olrawInputs(ctlKeys(1, 7))) Then
    If (rawInputs(ctlKeys(0, 7)) Or rawInputs(ctlKeys(1, 7))) > (rawInputs(ctlKeys(0, 8)) Or rawInputs(ctlKeys(1, 8))) Then
        SetInteger (cPoint2 + &H2), -Fix((rawInputs(ctlKeys(0, 7)) Or rawInputs(ctlKeys(1, 7))) / 2)
    ElseIf (rawInputs(ctlKeys(0, 7)) Or rawInputs(ctlKeys(1, 7))) < (rawInputs(ctlKeys(0, 8)) Or rawInputs(ctlKeys(1, 8))) Then
        SetInteger (cPoint2 + &H2), Fix((rawInputs(ctlKeys(0, 8)) Or rawInputs(ctlKeys(1, 8))) / 2)
    Else
        SetInteger (cPoint2 + &H2), 0
    End If
    End If
    
    'Camera Left/Right

    If Not (rawInputs(ctlKeys(0, 23)) Or rawInputs(ctlKeys(1, 23))) = (olrawInputs(ctlKeys(0, 23)) Or olrawInputs(ctlKeys(1, 23))) Or Not (rawInputs(ctlKeys(0, 22)) Or rawInputs(ctlKeys(1, 22))) = (olrawInputs(ctlKeys(0, 22)) Or olrawInputs(ctlKeys(1, 22))) Then
    If (rawInputs(ctlKeys(0, 22)) Or rawInputs(ctlKeys(1, 22))) > (rawInputs(ctlKeys(0, 23)) Or rawInputs(ctlKeys(1, 23))) Then
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H4), -Fix((rawInputs(ctlKeys(0, 22)) Or rawInputs(ctlKeys(1, 22))) / 2)
    ElseIf (rawInputs(ctlKeys(0, 22)) Or rawInputs(ctlKeys(1, 22))) < (rawInputs(ctlKeys(0, 23)) Or rawInputs(ctlKeys(1, 23))) Then
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H4), Fix((rawInputs(ctlKeys(0, 23)) Or rawInputs(ctlKeys(1, 23))) / 2)
    Else
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H4), 0
    End If
    End If
    
    'Camera Up/Down
    If Not (rawInputs(ctlKeys(0, 25)) Or rawInputs(ctlKeys(1, 25))) = (olrawInputs(ctlKeys(0, 25)) Or olrawInputs(ctlKeys(1, 25))) Or Not (rawInputs(ctlKeys(0, 24)) Or rawInputs(ctlKeys(1, 24))) = (olrawInputs(ctlKeys(0, 24)) Or olrawInputs(ctlKeys(1, 24))) Then
    If (rawInputs(ctlKeys(0, 24)) Or rawInputs(ctlKeys(1, 24))) > (rawInputs(ctlKeys(0, 25)) Or rawInputs(ctlKeys(1, 25))) Then
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H6), -Fix((rawInputs(ctlKeys(0, 24)) Or rawInputs(ctlKeys(1, 24))) / 2)
    ElseIf (rawInputs(ctlKeys(0, 24)) Or rawInputs(ctlKeys(1, 24))) < (rawInputs(ctlKeys(0, 25)) Or rawInputs(ctlKeys(1, 25))) Then
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H6), Fix((rawInputs(ctlKeys(0, 25)) Or rawInputs(ctlKeys(1, 25))) / 2)
    Else
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H6), 0
    End If

        'SetInteger (cPoint2 + &H6), 0
    End If
    'center camera
    If Not (rawInputs(ctlKeys(0, 26)) Or rawInputs(ctlKeys(1, 26))) = (olrawInputs(ctlKeys(0, 26)) Or olrawInputs(ctlKeys(1, 26))) Then
    SetInteger (cPoint2 + &H8), (rawInputs(ctlKeys(0, 26)) Or rawInputs(ctlKeys(1, 26)))
    End If

    'action
    If Not (rawInputs(ctlKeys(0, 19)) Or rawInputs(ctlKeys(1, 19))) = (olrawInputs(ctlKeys(0, 19)) Or olrawInputs(ctlKeys(1, 19))) Then
    SetInteger (cPoint2 + &H8), (rawInputs(ctlKeys(0, 19)) Or rawInputs(ctlKeys(1, 19)))
    End If

    'prev weapon
    If Not (rawInputs(ctlKeys(0, 2)) Or rawInputs(ctlKeys(1, 2))) = (olrawInputs(ctlKeys(0, 2)) Or olrawInputs(ctlKeys(1, 2))) Then
    SetInteger (cPoint2 + &HA), (rawInputs(ctlKeys(0, 2)) Or rawInputs(ctlKeys(1, 2)))
    End If
    
    'target
    If Not (rawInputs(ctlKeys(0, 17)) Or rawInputs(ctlKeys(1, 17))) = (olrawInputs(ctlKeys(0, 17)) Or olrawInputs(ctlKeys(1, 17))) Then
    SetInteger (cPoint2 + &HC), (rawInputs(ctlKeys(0, 17)) Or rawInputs(ctlKeys(1, 17)))
    End If
    
    'next weapon
    If Not (rawInputs(ctlKeys(0, 1)) Or rawInputs(ctlKeys(1, 1))) = (olrawInputs(ctlKeys(0, 1)) Or olrawInputs(ctlKeys(1, 1))) Then
    SetInteger (cPoint2 + &HE), (rawInputs(ctlKeys(0, 1)) Or rawInputs(ctlKeys(1, 1)))
    End If
    
    'group forward
    If Not (rawInputs(ctlKeys(0, 3)) Or rawInputs(ctlKeys(1, 3))) = (olrawInputs(ctlKeys(0, 3)) Or olrawInputs(ctlKeys(1, 3))) Then
    SetInteger (cPoint2 + &H10), (rawInputs(ctlKeys(0, 3)) Or rawInputs(ctlKeys(1, 3)))
    End If
    
    'group back
    If Not (rawInputs(ctlKeys(0, 4)) Or rawInputs(ctlKeys(1, 4))) = (olrawInputs(ctlKeys(0, 4)) Or olrawInputs(ctlKeys(1, 4))) Then
    SetInteger (cPoint2 + &H12), (rawInputs(ctlKeys(0, 4)) Or rawInputs(ctlKeys(1, 4)))
    End If
    
    'convo no
    If Not (rawInputs(ctlKeys(0, 5)) Or rawInputs(ctlKeys(1, 5))) = (olrawInputs(ctlKeys(0, 5)) Or olrawInputs(ctlKeys(1, 5))) Then
    SetInteger (cPoint2 + &H14), (rawInputs(ctlKeys(0, 5)) Or rawInputs(ctlKeys(1, 5)))
    End If
    
    'SetInteger (cPoint + &H14), (rawInputs(ctlKeys(0, 5)) Or rawInputs(ctlKeys(1, 5)))

    'convo yes
    If Not (rawInputs(ctlKeys(0, 6)) Or rawInputs(ctlKeys(1, 6))) = (olrawInputs(ctlKeys(0, 6)) Or olrawInputs(ctlKeys(1, 6))) Then
    SetInteger (cPoint2 + &H16), (rawInputs(ctlKeys(0, 6)) Or rawInputs(ctlKeys(1, 6)))
    End If
        
    'reserved
    'SetInteger (cPoint + &H18), (rawInputs(ctlKeys(0, 4)) Or rawInputs(ctlKeys(1, 4)))
    
    'change view
    If Not (rawInputs(ctlKeys(0, 14)) Or rawInputs(ctlKeys(1, 14))) = (olrawInputs(ctlKeys(0, 14)) Or olrawInputs(ctlKeys(1, 14))) Then
    SetInteger (cPoint2 + &H1A), (rawInputs(ctlKeys(0, 14)) Or rawInputs(ctlKeys(1, 14)))
    End If
    
    'enter+exit
    If Not (rawInputs(ctlKeys(0, 13)) Or rawInputs(ctlKeys(1, 13))) = (olrawInputs(ctlKeys(0, 13)) Or olrawInputs(ctlKeys(1, 13))) Then
        SetInteger (cPoint2 + &H1E), (rawInputs(ctlKeys(0, 13)) Or rawInputs(ctlKeys(1, 13)))
    End If
    'fire
    If Not (rawInputs(ctlKeys(0, 0)) Or rawInputs(ctlKeys(1, 0))) = (olrawInputs(ctlKeys(0, 0)) Or olrawInputs(ctlKeys(1, 0))) Then
        SetInteger (cPoint2 + &H22), (rawInputs(ctlKeys(0, 0)) Or rawInputs(ctlKeys(1, 0)))
    End If
    'SetInteger (cPoint + &H22), (rawInputs(ctlKeys(0, 0)) Or rawInputs(ctlKeys(1, 0)))

    'crouch
    If Not (rawInputs(ctlKeys(0, 18)) Or rawInputs(ctlKeys(1, 18))) = (olrawInputs(ctlKeys(0, 18)) Or olrawInputs(ctlKeys(1, 18))) Then
    SetInteger (cPoint2 + &H24), (rawInputs(ctlKeys(0, 18)) Or rawInputs(ctlKeys(1, 18)))
    End If
    'look back
    If Not (rawInputs(ctlKeys(0, 21)) Or rawInputs(ctlKeys(1, 21))) = (olrawInputs(ctlKeys(0, 21)) Or olrawInputs(ctlKeys(1, 21))) Then
    SetInteger (cPoint2 + &H26), (rawInputs(ctlKeys(0, 21)) Or rawInputs(ctlKeys(1, 21)))
    End If
    'reserved
    'SetInteger (cPoint + &H26), (rawInputs(ctlKeys(0, 21)) Or rawInputs(ctlKeys(1, 21)))

    'walk
    If Not (rawInputs(ctlKeys(0, 20)) Or rawInputs(ctlKeys(1, 20))) = (olrawInputs(ctlKeys(0, 20)) Or olrawInputs(ctlKeys(1, 20))) Then
    SetInteger (cPoint2 + &H2A), (rawInputs(ctlKeys(0, 20)) Or rawInputs(ctlKeys(1, 20)))
    End If








Else
    'Car
    'cPoint = &HB73488

    'Steer Left/Right
    If Not (rawInputs(ctlKeys(0, 33)) Or rawInputs(ctlKeys(1, 33))) = (olrawInputs(ctlKeys(0, 33)) Or olrawInputs(ctlKeys(1, 33))) Or Not (rawInputs(ctlKeys(0, 32)) Or rawInputs(ctlKeys(1, 32))) = (olrawInputs(ctlKeys(0, 32)) Or olrawInputs(ctlKeys(1, 32))) Then
    If (rawInputs(ctlKeys(0, 32)) Or rawInputs(ctlKeys(1, 32))) > (rawInputs(ctlKeys(0, 33)) Or rawInputs(ctlKeys(1, 33))) Then
        SetInteger (cPoint2 + &H0), -Fix((rawInputs(ctlKeys(0, 32)) Or rawInputs(ctlKeys(1, 32))) / 2)
    ElseIf (rawInputs(ctlKeys(0, 32)) Or rawInputs(ctlKeys(1, 32))) < (rawInputs(ctlKeys(0, 33)) Or rawInputs(ctlKeys(1, 33))) Then
        SetInteger (cPoint2 + &H0), Fix((rawInputs(ctlKeys(0, 33)) Or rawInputs(ctlKeys(1, 33))) / 2)
    Else
        SetInteger (cPoint2 + &H0), 0
    End If
    End If
    
    'Steer Up/Down
    If Not (rawInputs(ctlKeys(0, 35)) Or rawInputs(ctlKeys(1, 35))) = (olrawInputs(ctlKeys(0, 35)) Or olrawInputs(ctlKeys(1, 35))) Or Not (rawInputs(ctlKeys(0, 34)) Or rawInputs(ctlKeys(1, 34))) = (olrawInputs(ctlKeys(0, 34)) Or olrawInputs(ctlKeys(1, 34))) Then
    If (rawInputs(ctlKeys(0, 34)) Or rawInputs(ctlKeys(1, 34))) > (rawInputs(ctlKeys(0, 35)) Or rawInputs(ctlKeys(1, 35))) Then
        SetInteger (cPoint2 + &H2), -Fix((rawInputs(ctlKeys(0, 34)) Or rawInputs(ctlKeys(1, 34))) / 2)
    ElseIf (rawInputs(ctlKeys(0, 34)) Or rawInputs(ctlKeys(1, 34))) < (rawInputs(ctlKeys(0, 35)) Or rawInputs(ctlKeys(1, 35))) Then
        SetInteger (cPoint2 + &H2), Fix((rawInputs(ctlKeys(0, 35)) Or rawInputs(ctlKeys(1, 35))) / 2)
    Else
        SetInteger (cPoint2 + &H2), 0
    End If
    End If
    
    
    'Special Left/Right
    If Not (rawInputs(ctlKeys(0, 50)) Or rawInputs(ctlKeys(1, 50))) = (olrawInputs(ctlKeys(0, 50)) Or olrawInputs(ctlKeys(1, 50))) Or Not (rawInputs(ctlKeys(0, 49)) Or rawInputs(ctlKeys(1, 49))) = (olrawInputs(ctlKeys(0, 49)) Or olrawInputs(ctlKeys(1, 49))) Then
    If (rawInputs(ctlKeys(0, 49)) Or rawInputs(ctlKeys(1, 49))) > (rawInputs(ctlKeys(0, 50)) Or rawInputs(ctlKeys(1, 50))) Then
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H4), -Fix((rawInputs(ctlKeys(0, 49)) Or rawInputs(ctlKeys(1, 49))) / 2)
    ElseIf (rawInputs(ctlKeys(0, 49)) Or rawInputs(ctlKeys(1, 49))) < (rawInputs(ctlKeys(0, 50)) Or rawInputs(ctlKeys(1, 50))) Then
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H4), Fix((rawInputs(ctlKeys(0, 50)) Or rawInputs(ctlKeys(1, 50))) / 2)
    Else
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H4), 0
    End If
    End If
    'Special Up/Down
    If Not (rawInputs(ctlKeys(0, 52)) Or rawInputs(ctlKeys(1, 52))) = (olrawInputs(ctlKeys(0, 52)) Or olrawInputs(ctlKeys(1, 52))) Or Not (rawInputs(ctlKeys(0, 51)) Or rawInputs(ctlKeys(1, 51))) = (olrawInputs(ctlKeys(0, 51)) Or olrawInputs(ctlKeys(1, 51))) Then
    If (rawInputs(ctlKeys(0, 51)) Or rawInputs(ctlKeys(1, 51))) > (rawInputs(ctlKeys(0, 52)) Or rawInputs(ctlKeys(1, 52))) Then
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H6), -Fix((rawInputs(ctlKeys(0, 51)) Or rawInputs(ctlKeys(1, 51))) / 2)
    ElseIf (rawInputs(ctlKeys(0, 51)) Or rawInputs(ctlKeys(1, 51))) < (rawInputs(ctlKeys(0, 52)) Or rawInputs(ctlKeys(1, 52))) Then
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H6), Fix((rawInputs(ctlKeys(0, 52)) Or rawInputs(ctlKeys(1, 52))) / 2)
    Else
        If frmMain.chkSec.value = Unchecked Then SetByte &HB6EC2E, 0
        SetInteger (cPoint2 + &H6), 0
    End If
    End If
    
    'second fire
    If Not (rawInputs(ctlKeys(0, 29)) Or rawInputs(ctlKeys(1, 29))) = (olrawInputs(ctlKeys(0, 29)) Or olrawInputs(ctlKeys(1, 29))) Then
    SetInteger (cPoint2 + &H8), (rawInputs(ctlKeys(0, 29)) Or rawInputs(ctlKeys(1, 29)))
    End If
    
    'look back
    
    If Not (rawInputs(ctlKeys(0, 45)) Or rawInputs(ctlKeys(1, 45))) > 0 Then

    'Else
    'look left
    'If GetInteger(cPoint2 + &HA) > 0 And GetInteger(cPoint2 + &HE) > 0 Then GoTo skipllr
    If Not (rawInputs(ctlKeys(0, 47)) Or rawInputs(ctlKeys(1, 47))) = (olrawInputs(ctlKeys(0, 47)) Or olrawInputs(ctlKeys(1, 47))) Then
    SetInteger (cPoint2 + &HA), (rawInputs(ctlKeys(0, 47)) Or rawInputs(ctlKeys(1, 47)))
    End If
    
    'look right
    If Not (rawInputs(ctlKeys(0, 48)) Or rawInputs(ctlKeys(1, 48))) = (olrawInputs(ctlKeys(0, 48)) Or olrawInputs(ctlKeys(1, 48))) Then
    SetInteger (cPoint2 + &HE), (rawInputs(ctlKeys(0, 48)) Or rawInputs(ctlKeys(1, 48)))
    End If
    End If
'skipllr:
    
'    End If
     If Not (rawInputs(ctlKeys(0, 45)) Or rawInputs(ctlKeys(1, 45))) = (olrawInputs(ctlKeys(0, 45)) Or olrawInputs(ctlKeys(1, 45))) Then
        SetInteger (cPoint2 + &HA), (rawInputs(ctlKeys(0, 45)) Or rawInputs(ctlKeys(1, 45)))
        SetInteger (cPoint2 + &HE), (rawInputs(ctlKeys(0, 45)) Or rawInputs(ctlKeys(1, 45)))
    End If
    

    
    'handbrake
    If Not (rawInputs(ctlKeys(0, 44)) Or rawInputs(ctlKeys(1, 44))) = (olrawInputs(ctlKeys(0, 44)) Or olrawInputs(ctlKeys(1, 44))) Then
    SetInteger (cPoint2 + &HC), (rawInputs(ctlKeys(0, 44)) Or rawInputs(ctlKeys(1, 44)))
    End If
    
    
    'next radio
    If Not (rawInputs(ctlKeys(0, 38)) Or rawInputs(ctlKeys(1, 38))) = (olrawInputs(ctlKeys(0, 38)) Or olrawInputs(ctlKeys(1, 38))) Then
    SetInteger (cPoint2 + &H10), (rawInputs(ctlKeys(0, 38)) Or rawInputs(ctlKeys(1, 38)))
    End If
    
    'prev radio
    If Not (rawInputs(ctlKeys(0, 39)) Or rawInputs(ctlKeys(1, 39))) = (olrawInputs(ctlKeys(0, 39)) Or olrawInputs(ctlKeys(1, 39))) Then
    SetInteger (cPoint2 + &H12), (rawInputs(ctlKeys(0, 39)) Or rawInputs(ctlKeys(1, 39)))
    End If

    'prev radio???????????????
    'SetInteger (cPoint + &H12), (rawInputs(ctlKeys(0, 39)) Or rawInputs(ctlKeys(1, 39)))


    'change view
    If Not (rawInputs(ctlKeys(0, 43)) Or rawInputs(ctlKeys(1, 43))) = (olrawInputs(ctlKeys(0, 43)) Or olrawInputs(ctlKeys(1, 43))) Then
    SetInteger (cPoint2 + &H1A), (rawInputs(ctlKeys(0, 43)) Or rawInputs(ctlKeys(1, 43)))
    End If
    
    'If Block = False Then
    'SetInteger (cPoint + &H1A), (rawInputs(ctlKeys(0, 43)) Or rawInputs(ctlKeys(1, 43)))
    'End If
    'brake
    If Not (rawInputs(ctlKeys(0, 31)) Or rawInputs(ctlKeys(1, 31))) = (olrawInputs(ctlKeys(0, 31)) Or olrawInputs(ctlKeys(1, 31))) Then
    SetInteger (cPoint2 + &H1C), (rawInputs(ctlKeys(0, 31)) Or rawInputs(ctlKeys(1, 31)))
    End If

    'Enter+exit
    If Not (rawInputs(ctlKeys(0, 36)) Or rawInputs(ctlKeys(1, 36))) = (olrawInputs(ctlKeys(0, 36)) Or olrawInputs(ctlKeys(1, 36))) Then
    SetInteger (cPoint2 + &H1E), (rawInputs(ctlKeys(0, 36)) Or rawInputs(ctlKeys(1, 36)))
    End If
    'If Block = False Then
    'SetInteger (cPoint + &H1E), (rawInputs(ctlKeys(0, 36)) Or rawInputs(ctlKeys(1, 36)))
    'End If
    
    'throttle
    If Not (rawInputs(ctlKeys(0, 30)) Or rawInputs(ctlKeys(1, 30))) = (olrawInputs(ctlKeys(0, 30)) Or olrawInputs(ctlKeys(1, 30))) Then
    SetInteger (cPoint2 + &H20), (rawInputs(ctlKeys(0, 30)) Or rawInputs(ctlKeys(1, 30)))
    End If
    'fire
    If Not (rawInputs(ctlKeys(0, 28)) Or rawInputs(ctlKeys(1, 28))) = (olrawInputs(ctlKeys(0, 28)) Or olrawInputs(ctlKeys(1, 28))) Then
    SetInteger (cPoint2 + &H22), (rawInputs(ctlKeys(0, 28)) Or rawInputs(ctlKeys(1, 28)))
    End If
    'brake??
    'SetInteger (cPoint + &H1C), (rawInputs(ctlKeys(0, 31)) Or rawInputs(ctlKeys(1, 31)))

    'horn
    If Not (rawInputs(ctlKeys(0, 41)) Or rawInputs(ctlKeys(1, 41))) = (olrawInputs(ctlKeys(0, 41)) Or olrawInputs(ctlKeys(1, 41))) Then
    SetInteger (cPoint2 + &H24), (rawInputs(ctlKeys(0, 41)) Or rawInputs(ctlKeys(1, 41)))
    End If

    'submission
    If Not (rawInputs(ctlKeys(0, 42)) Or rawInputs(ctlKeys(1, 42))) = (olrawInputs(ctlKeys(0, 42)) Or olrawInputs(ctlKeys(1, 42))) Then
    SetInteger (cPoint2 + &H26), (rawInputs(ctlKeys(0, 42)) Or rawInputs(ctlKeys(1, 42)))
    End If
    

    
    'trip skip
    
    If Not (rawInputs(ctlKeys(0, 37)) Or rawInputs(ctlKeys(1, 37))) = (olrawInputs(ctlKeys(0, 37)) Or olrawInputs(ctlKeys(1, 37))) Then
    SetInteger (cPoint2 + &H14), (rawInputs(ctlKeys(0, 37)) Or rawInputs(ctlKeys(1, 37)))
    End If

    'mouse look
    If Not (rawInputs(ctlKeys(0, 46)) Or rawInputs(ctlKeys(1, 46))) = (olrawInputs(ctlKeys(0, 46)) Or olrawInputs(ctlKeys(1, 46))) Then
    SetInteger (cPoint2 + &H2C), (rawInputs(ctlKeys(0, 46)) Or rawInputs(ctlKeys(1, 46)))
    End If
    
    'user track skip
    If Not (rawInputs(ctlKeys(0, 40)) Or rawInputs(ctlKeys(1, 40))) = (olrawInputs(ctlKeys(0, 40)) Or olrawInputs(ctlKeys(1, 40))) Then
    SetInteger (cPoint2 + &H2E), (rawInputs(ctlKeys(0, 40)) Or rawInputs(ctlKeys(1, 40)))
    End If
    
    'convo no
    If Not (rawInputs(ctlKeys(0, 5)) Or rawInputs(ctlKeys(1, 5))) = (olrawInputs(ctlKeys(0, 5)) Or olrawInputs(ctlKeys(1, 5))) Then
    SetInteger (cPoint2 + &H14), (rawInputs(ctlKeys(0, 5)) Or rawInputs(ctlKeys(1, 5)))
    End If
    
    'SetInteger (cPoint + &H14), (rawInputs(ctlKeys(0, 5)) Or rawInputs(ctlKeys(1, 5)))

    'convo yes
    If Not (rawInputs(ctlKeys(0, 6)) Or rawInputs(ctlKeys(1, 6))) = (olrawInputs(ctlKeys(0, 6)) Or olrawInputs(ctlKeys(1, 6))) Then
    SetInteger (cPoint2 + &H16), (rawInputs(ctlKeys(0, 6)) Or rawInputs(ctlKeys(1, 6)))
    End If

End If

End Function

Public Function inCar() As Boolean
On Error Resume Next
'If frmMain.chkSec = Checked Then
'If GetByte(&H8CB6F9) = 0 Then inCar = False Else inCar = True
'Else
If GetLong(&HBA18FC) = 0 Then inCar = False Else inCar = True
'End If
End Function

Public Function StartEffect(eID As Long, eMag As Single, eDur As Single, ePer As Single, oBox As Object)
On Error Resume Next
If eMag > 10000 Then eMag = 10000
If eMag < 0 Then eMag = 0
Effect.lDuration = eDur * 100

If Effect.lDuration = 0 Then Effect.lDuration = -1

If oBox(0) Then
   Effect.X = 9000
ElseIf oBox(1) Then
   Effect.X = 27000
ElseIf oBox(2) Then
   Effect.X = 18000
ElseIf oBox(3) Then
   Effect.X = 0
ElseIf oBox(4) Then
   Effect.X = 13500
ElseIf oBox(5) Then
   Effect.X = 22500
ElseIf oBox(6) Then
   Effect.X = 4500
ElseIf oBox(7) Then
   Effect.X = 31500
End If

With Effect.conditionX
    .lDeadBand = 0
    .lNegativeCoefficient = 10000
    .lNegativeSaturation = Val(eMag)
    .lOffset = 0
    .lPositiveCoefficient = 10000
    .lPositiveSaturation = Val(eMag)
End With
Effect.constantForce.lMagnitude = eMag 'CLng(mAmt) * 300 'Val(sWep(0).Value)
Effect.rampForce.lRangeEnd = 10000
Effect.rampForce.lRangeStart = 0
With Effect.conditionY
    .lDeadBand = 0
    .lNegativeCoefficient = 10000
    .lNegativeSaturation = Val(eMag)
    .lOffset = 0
    .lPositiveCoefficient = 10000
    .lPositiveSaturation = Val(eMag)
End With
Effect.periodicForce.lMagnitude = eMag '10000
Effect.periodicForce.lPeriod = ePer * 100
diFFEffect(eID).SetParameters Effect, DIEP_DURATION
diFFEffect(eID).SetParameters Effect, DIEP_TYPESPECIFICPARAMS
diFFEffect(eID).SetParameters Effect, DIEP_DIRECTION
diFFEffect(eID).Start 1, 0
End Function
Public Function StopSteer()
On Error Resume Next
diFFEffect(7).Stop
diFFEffect(10).Stop
diFFEffect(0).Stop
End Function
Public Function StartSteer(eMag As Single, eMulti As Single, eTurn As Single, eSpring As Single, eFriction As Single, oBox As Object)
On Error Resume Next
If eMag > 10000 Then eMag = 10000
If eMag < 0 Then eMag = 0

If eSpring > 10000 Then eSpring = 10000
If eSpring < 0 Then eSpring = 0

If eFriction > 10000 Then eFriction = 10000
If eFriction < 0 Then eFriction = 0


'eSpring
'Spring, Friction
Effect.lDuration = -1
If oBox(0) Then
   Effect.X = 9000
ElseIf oBox(1) Then
   Effect.X = 27000
ElseIf oBox(2) Then
   Effect.X = 18000
ElseIf oBox(3) Then
   Effect.X = 0
ElseIf oBox(4) Then
   Effect.X = 13500
ElseIf oBox(5) Then
   Effect.X = 22500
ElseIf oBox(6) Then
   Effect.X = 4500
ElseIf oBox(7) Then
   Effect.X = 31500
End If


'spring
With Effect.conditionX
    .lDeadBand = 0
    .lNegativeCoefficient = 10000
    .lNegativeSaturation = Val(eSpring)
    .lOffset = 0
    .lPositiveCoefficient = 10000
    .lPositiveSaturation = Val(eSpring)
End With

With Effect.conditionY
    .lDeadBand = 0
    .lNegativeCoefficient = 10000
    .lNegativeSaturation = Val(eSpring)
    .lOffset = 0
    .lPositiveCoefficient = 10000
    .lPositiveSaturation = Val(eSpring)
End With
'7 = spring
diFFEffect(7).SetParameters Effect, DIEP_DURATION
diFFEffect(7).SetParameters Effect, DIEP_TYPESPECIFICPARAMS
diFFEffect(7).SetParameters Effect, DIEP_DIRECTION
diFFEffect(7).Start 1, 0


'friction
With Effect.conditionX
    .lDeadBand = 0
    .lNegativeCoefficient = 10000
    .lNegativeSaturation = Val(eFriction)
    .lOffset = 0
    .lPositiveCoefficient = 10000
    .lPositiveSaturation = Val(eFriction)
End With

With Effect.conditionY
    .lDeadBand = 0
    .lNegativeCoefficient = 10000
    .lNegativeSaturation = Val(eFriction)
    .lOffset = 0
    .lPositiveCoefficient = 10000
    .lPositiveSaturation = Val(eFriction)
End With
'10 = friction
diFFEffect(10).SetParameters Effect, DIEP_DURATION
diFFEffect(10).SetParameters Effect, DIEP_TYPESPECIFICPARAMS
diFFEffect(10).SetParameters Effect, DIEP_DIRECTION
'diFFEffect(10).Start 1, 0

'Constant
eTurn = ((eTurn * eMulti))
If Abs(eTurn) > 9000 Then
If eTurn < 0 Then eTurn = -9000
If eTurn > 0 Then eTurn = 9000
End If
If eTurn < 0 Then eTurn = eTurn - 36000
Effect.X = eTurn
Effect.constantForce.lMagnitude = eMag
With Effect.conditionX
    .lDeadBand = 0
    .lNegativeCoefficient = 10000
    .lNegativeSaturation = 10000
    .lOffset = 0
    .lPositiveCoefficient = 10000
    .lPositiveSaturation = 10000
End With

With Effect.conditionY
    .lDeadBand = 0
    .lNegativeCoefficient = 10000
    .lNegativeSaturation = 10000
    .lOffset = 0
    .lPositiveCoefficient = 10000
    .lPositiveSaturation = 10000
End With
'0 = friction
diFFEffect(0).SetParameters Effect, DIEP_DURATION
diFFEffect(0).SetParameters Effect, DIEP_TYPESPECIFICPARAMS
diFFEffect(0).SetParameters Effect, DIEP_DIRECTION
diFFEffect(0).Start 1, 0
End Function


Public Function InitGame()
On Error Resume Next
oldTime = 0
SetHex &H53F547, "89 1D 9C F6 53 00 BB A0 F6 53 00"
SetHex &H53F552, " 66 8B 03 66 89 01 66 8B 43 02 66"
SetHex &H53F55D, " 89 41 02 66 8B 43 04 66 89 41 04"
SetHex &H53F568, " 66 8B 43 06 66 89 41 06 66 8B 43"
SetHex &H53F573, " 08 66 89 41 08 66 8B 43 0A 66 89"
SetHex &H53F57E, " 41 0A 66 8B 43 0C 66 89 41 0C 66"
SetHex &H53F589, " 8B 43 0E 66 89 41 0E 66 8B 43 10"
SetHex &H53F594, " 66 89 41 10 66 8B 43 12 66 89 41"
SetHex &H53F59F, " 12 66 8B 43 14 66 89 41 14 66 8B"
SetHex &H53F5AA, " 43 16 66 89 41 16 66 8B 43 18 66"
SetHex &H53F5B5, " 89 41 18 66 8B 43 1A 66 89 41 1A"
SetHex &H53F5C0, "66 8B 43 1C 66 89 41 1C 66 8B 43"
SetHex &H53F5CB, "1E 66 89 41 1E 66 8B 43 20 66 89"
SetHex &H53F5D6, "41 20 66 8B 43 22 66 89 41 22 66"
SetHex &H53F5E1, "8B 43 24 66 89 41 24 66 8B 43 26"
SetHex &H53F5EC, "66 89 41 26 66 8B 43 28 66 89 41"
SetHex &H53F5F7, "28 66 8B 43 2A 66 89 41 2A 66 8B"
SetHex &H53F602, "43 2E 66 89 41 2E 66 8B 43 2C 66"
SetHex &H53F60D, "89 41 2C 66 8B 43 2E 66 89 41 2E"
SetHex &H53F618, "8B 1D 9C F6 53 00 E9 CF 02 00 00"
SetHex &H53F6A0, "00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6AB, "00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6B6, "00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6C1, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6CC, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6D7, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6E2, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6ED, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6F8, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F6D7, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F703, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F70E, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F719, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F724, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F72F, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F73A, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F745, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F750, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H53F75B, " 00 00 00 00 00 00 00 00 00 00 00"
SetHex &H73DC64, "909090"
SetAim aMode


'2p hack

SetHex &H53F720, "8B 1D 58 F7 53 00 83 C3 01 83 FB"
SetHex &H53F72B, "04 72 05 BB 00 00 00 00 83 FB 01"
SetHex &H53F736, "76 10 89 1D 58 F7 53 00 BB E0 F6"
SetHex &H53F741, "53 00 E9 0A FE FF FF 89 1D 58 F7"
SetHex &H53F74C, " 53 00 BB A0 F6 53 00 E9 FA FD FF"
SetHex &H53F757, "FF"
SetHex &H53F758, "00 00 00 00"
SetHex &H53F54D, "E9 CE 01 00 00"


'0060B65F
End Function

Public Function DeInitGame()
 SetHex &H53F540, "14 8B 54 24 18 33 FF 66-39 7E 08 75 06 66 39 7A "   '¶ïT$3 f9~uf9z"
 SetHex &H53F550, "08 74 0D B8 FF 00 00 00-66 A3 C8 36 B7 00 EB 05 "   't
 SetHex &H53F560, " B8 FF 00 00 00 66 39 7E-0A 75 06 66 39 7A 0A 74 "  '+ ...f9~
 SetHex &H53F570, "06 66 A3 CA 36 B7 00 66-39 7E 0C 75 06 66 39 7A "   'fú-6+.f9~uf9z"
 SetHex &H53F580, "0C 74 06 66 A3 CC 36 B7-00 66 39 7E 0E 75 06 66 "   'tfú¦6+.f9~uf"
 SetHex &H53F590, "39 7A 0E 74 06 66 A3 CE-36 B7 00 66 39 7E 18 75 "   '9ztfú+6+.f9~u"
 SetHex &H53F5A0, "06 66 39 7A 18 74 06 66-A3 D8 36 B7 00 66 39 7E "   'f9ztfú+6+.f9~"
 SetHex &H53F5B0, "1A 75 06 66 39 7A 1A 74-06 66 A3 DA 36 B7 00 66 "   'uf9ztfú+6+.f"
 SetHex &H53F5C0, "39 7E 1C 75 06 66 39 7A-1C 74 06 66 A3 DC 36 B7 "   '9~uf9ztfú_6+"
 SetHex &H53F5D0, "00 66 39 7E 1E 75 06 66-39 7A 1E 74 06 66 A3 DE "   '.f9~uf9ztfú¦"
 SetHex &H53F5E0, "36 B7 00 66 39 7E 20 75-06 66 39 7A 20 74 06 66 "   '6+.f9~ uf9z tf"
 SetHex &H53F5F0, "A3 E0 36 B7 00 66 39 7E-22 75 06 66 39 7A 22 74 "   'úa6+.f9~"uf9z"t"
 SetHex &H53F600, "06 66 A3 E2 36 B7 00 66-39 7E 24 75 06 66 39 7A "   'fúG6+.f9~$uf9z"
 SetHex &H53F610, "24 74 06 66 A3 E4 36 B7-00 66 39 7E 26 75 06 66 "   '$tfúS6+.f9~&uf"
 SetHex &H53F620, "39 7A 26 74 06 66 A3 E6-36 B7 00 66 39 7E 28 75 "   '9'z&tfúµ6+.f9~(u"
 SetHex &H53F630, "06 66 39 7A 28 74 06 66-A3 E8 36 B7 00 66 39 7E "   ''f9z(tfúF6+.f9~"
 SetHex &H53F640, "2A 75 06 66 39 7A 2A 74-06 66 A3 EA 36 B7 00 66 "   '*''uf9z*tfúO6+.f"
 SetHex &H53F650, "39 7E 2E 75 06 66 39 7A-2E 74 06 66 A3 EE 36 B7 "   '9'~.uf9z.tfúe6+"
 SetHex &H53F660, "00 66 39 7E 2C 75 06 66-39 7A 2C 74 06 66 A3 EC "   '.f9~,uf9z,tfú8"
 SetHex &H53F670, "36 B7 00 66 8B 0E 66 3B-CF 7C 25 66 8B 02 66 3B "   '6+.fïf;-|%fïf;"
 SetHex &H53F680, "C7 7C 1D 66 3B C8 7E 0C-66 8B E9 66 89 2D C0 36 "   '¦'|f;+~fïTfë-+6"
 SetHex &H53F690, " B7 00 EB 13 66 8B E8 66-89 2D C0 36 B7 00 EB 07 "  '+.dfïFfë-+6+.d"
 SetHex &H53F6A0, " 66 8B 2D C0 36 B7 00 66-8B 0E 66 3B CF 7F 1A 66 "  'fï-+6+.fïf;-f"
 SetHex &H53F6B0, "8B 02 66 3B C7 7F 12 66-3B C8 66 8B E9 7C 03 66 "   'ïf;¦f;+fïT|f"
 SetHex &H53F6C0, "8B E8 66 89 2D C0 36 B7-00 66 8B 4E 02 66 3B CF "   'ïFfë-+6+.fïNf;-"
 SetHex &H53F6D0, "53 7C 26 66 8B 42 02 66-3B C7 7C 1D 66 3B C8 7E "   'S|&fïBf;¦|f;+~"
 SetHex &H53F6E0, "0C 66 8B D9 66 89 1D C2-36 B7 00 EB 13 66 8B D8 "    'fï+fë-6+.dfï+"
 SetHex &H53F6F0, "66 89 1D C2 36 B7 00 EB-07 66 8B 1D C2 36 B7 00 "    'fë-6+.dfï-6+."
 SetHex &H53F700, "66 8B 4E 02 66 3B CF 7F-1B 66 8B 42 02 66 3B C7 "    'fïNf;-fïBf;¦"
 SetHex &H53F710, "7F 12 66 3B C8 66 8B D9-7C 03 66 8B D8 66 89 1D "    'f;+fï+|fï+fë"
 SetHex &H53F720, "C2 36 B7 00 66 8B 06 66-3B C7 7E 08 66 39 3A 7C "    '-6+.fïf;¦~f9:|"
 SetHex &H53F730, "0A 66 3B C7 7D 0F 66 39-3A 7E 0A 66 33 ED 66 89 "
 SetHex &H53F740, "2D C0 36 B7 00 66 8B 46-02 66 3B C7 7E 09 66 39 "    '-+6+.fïFf;¦~  f9"
 SetHex &H53F750, "7A 02 7C 0B 66 3B C7 7D-10 66 39 7A 02 7E 0A 66 "    'z|f;¦}f9z~
 SetHex &H53F760, "33 DB 66 89 1D C2 36 B7-00 66 8B 4E 04 66 3B CF "    '3¦fë-6+.fïNf;-"
 SetHex &H53F770, "7C 1B 66 8B 42 04 66 3B-C7 7C 12 66 3B C8 66 89 "    '|fïBf;¦|f;+fë"
 SetHex &H53F780, "0D C4 36 B7 00 7F 06 66-A3 C4 36 B7 00 66 8B 4E "
 SetHex &H53F790, "04 66 3B CF 7F 1B 66 8B-42 04 66 3B C7 7F 12 66 "    'f;-fïBf;¦f"


SetHex &H73DC64, "89 5E 04"

End Function
Public Function isBike(carID) As Boolean
On Error Resume Next
Select Case carID
Case 448: isBike = True
Case 468: isBike = True
Case 461: isBike = True
Case 462: isBike = True
Case 463: isBike = True
Case 481: isBike = True
Case 509: isBike = True
Case 510: isBike = True
Case 521: isBike = True
Case 522: isBike = True
Case 523: isBike = True
Case 581: isBike = True
Case 586: isBike = True
Case Else: isBike = False
End Select
End Function

Public Function inZoom() As Boolean
On Error Resume Next
Dim curMode As Long
curMode = GetLong(&HB6F1A8)
If curMode = 7 Or curMode = 46 Then
inZoom = True
Else
inZoom = False
End If
End Function
Public Function SetRTCGAIN(ffRTC As Boolean, ffGain As Long)
'On Error Resume Next
objDIDevC.Unacquire
If ffRTC = True Then
Prop.lData = 1
Else
Prop.lData = 0
End If
Prop.lHow = DIPH_DEVICE
Prop.lObj = 0
Prop.lSize = Len(Prop)
Call objDIDevC.SetProperty("DIPROP_AUTOCENTER", Prop)  'Turn off autocenter

Prop.lData = CLng(ffGain)
Prop.lHow = DIPH_DEVICE
Prop.lObj = 0
Prop.lSize = Len(Prop)
Call objDIDevC.SetProperty("DIPROP_FFGAIN", Prop)
DoEvents
objDIDevC.Acquire
End Function

Public Function findDirection(oBox As Object)
On Error Resume Next
Dim cLoop As Long
For cLoop = 0 To oBox.UBound
If oBox(cLoop).value = True Then findDirection = cLoop: Exit For
Next cLoop
End Function

Public Function LSus(inVal)
Dim nVal
nVal = (inVal + 10) * 0.1
'LSus = (16.2 * (inVal * 0.01) ^ 2 - 14.8 * (inVal * 0.01) - 1.4)
'LSus = 1.4476 * (inVal * 0.1) ^ 2 - 7.9779 * (inVal * 0.1) + 13.689
'LSus = 0.0223 * (inVal * 0.1) ^ 6 - 0.7236 * (inVal * 0.1) ^ 5 + 9.1196 * (inVal * 0.1) ^ 4 - 56.188 * (inVal * 0.1) ^ 3 + 174.05 * (inVal * 0.1) ^ 2 - 247.83 * (inVal * 0.1) + 122.88
'LSus = (0.0188 * (inVal * 0.1) ^ 6 - 0.5495 * (inVal * 0.1) ^ 5 + 6.2668 * (inVal * 0.1) ^ 4 - 35.085 * (inVal * 0.1) ^ 3 + 99.779 * (inVal * 0.1) ^ 2 - 133.21 * (inVal * 0.1) + 73)
'Debug.Print LSus
LSus = 0.008 * nVal ^ 6 - 0.2604 * nVal ^ 5 + 3.2812 * nVal ^ 4 - 20.088 * nVal ^ 3 + 61.135 * nVal ^ 2 - 81.257 * nVal + 47.411
End Function

Public Function PressEscape()
On Error Resume Next
If GetByte(&HB7CB49) = 1 Or isFocused = False Then Exit Function
'SetByte (&HBA67A4), 1
'SetByte (&HB7CB49), 1

keybd_event 27, 0, 0, 0
keybd_event 27, 0, KEYEVENTF_KEYUP, 0
End Function
Public Function isFocused()
On Error Resume Next
If GetForegroundWindow = hWin Then
isFocused = True
Else
isFocused = False
End If
End Function

Public Function SetAim(aimVal)
If aimVal = 0 Then
SetLong &H521A17, &HB6EC2E
SetLong &H686906, &HB6EC2E
SetLong &H6869B1, &HB6EC2E
SetLong &H686CE7, &HB6EC2E
SetLong &H60B65F, &HB6EC2E
SetLong &H52A93D, &HB6EC2E
ElseIf aimVal = 1 Then
SetLong &H521A17, &HB6EC2F
SetLong &H686906, &HB6EC2F
SetLong &H6869B1, &HB6EC2F
SetLong &H686CE7, &HB6EC2F
SetLong &H60B65F, &HB6EC2F
SetLong &H52A93D, &HB6EC2F
SetByte &HB6EC2F, 1
ElseIf aimVal = 2 Then
SetLong &H521A17, &HB6EC2F
SetLong &H686906, &HB6EC2F
SetLong &H6869B1, &HB6EC2F
SetLong &H686CE7, &HB6EC2F
SetLong &H60B65F, &HB6EC2F
SetLong &H52A93D, &HB6EC2F
SetByte &HB6EC2F, 0
End If
End Function
