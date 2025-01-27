Attribute VB_Name = "KeyMod"
Option Explicit




Public Function KeyDown(vKey) As Boolean
   On Error Resume Next
   
   
       objDIDevK.GetDeviceStateKeyboard diState
If vKey <= 237 Then
    If diState.Key(vKey) = 128 Then KeyDown = True
ElseIf vKey > 237 And vKey <= 269 Then

    If js.buttons(vKey - 238) = 128 Then If Not objDIDevC Is Nothing Then KeyDown = True
ElseIf vKey >= 270 And vKey <= 297 Then
        If (rawInputs(vKey) - orawInputs(vKey)) > 100 Then

        If Not objDIDevC Is Nothing Then KeyDown = True
        
        End If
ElseIf vKey > 297 Then
If (rawInputs(vKey)) > 0 Then

        If Not objDIDev Is Nothing Then KeyDown = True
        
        End If
End If
End Function



'0 enable
'1 disable
'2 extend arm
'3 retract arm
'4 rotate ccw
'5 rotate cw
'6 move up
'7 move down
'8 mouse-look
'9 zoom in
'10 zoom out
'11 fov up
'12 fov down






Function GetKeyboardString(iNum) As String
On Error Resume Next
Select Case iNum
    Case &H0: GetKeyboardString = "Undefined / Disabled"
    Case &H1: GetKeyboardString = "DIK_ESCAPE"
    Case &H2: GetKeyboardString = "DIK_1"
    Case &H3: GetKeyboardString = "DIK_2"
    Case &H4: GetKeyboardString = "DIK_3"
    Case &H5: GetKeyboardString = "DIK_4"
    Case &H6: GetKeyboardString = "DIK_5"
    Case &H7: GetKeyboardString = "DIK_6"
    Case &H8: GetKeyboardString = "DIK_7"
    Case &H9: GetKeyboardString = "DIK_8"
    Case &HA: GetKeyboardString = "DIK_9"
    Case &HB: GetKeyboardString = "DIK_0"
    Case &HC: GetKeyboardString = "DIK_MINUS"              ' - on main keyboard
    Case &HD: GetKeyboardString = "DIK_EQUALS"
    Case &HE: GetKeyboardString = "DIK_BACK"                   ' backspace
    Case &HF: GetKeyboardString = "DIK_TAB"
    Case &H10: GetKeyboardString = "DIK_Q"
    Case &H11: GetKeyboardString = "DIK_W"
    Case &H12: GetKeyboardString = "DIK_E"
    Case &H13: GetKeyboardString = "DIK_R"
    Case &H14: GetKeyboardString = "DIK_T"
    Case &H15: GetKeyboardString = "DIK_Y"
    Case &H16: GetKeyboardString = "DIK_U"
    Case &H17: GetKeyboardString = "DIK_I"
    Case &H18: GetKeyboardString = "DIK_O"
    Case &H19: GetKeyboardString = "DIK_P"
    Case &H1A: GetKeyboardString = "DIK_LBRACKET"
    Case &H1B: GetKeyboardString = "DIK_RBRACKET"
    Case &H1C: GetKeyboardString = "DIK_RETURN"  ' Enter on main keyboard
    Case &H1D: GetKeyboardString = "DIK_LCONTROL"
    Case &H1E: GetKeyboardString = "DIK_A"
    Case &H1F: GetKeyboardString = "DIK_S"
    Case &H20: GetKeyboardString = "DIK_D"
    Case &H21: GetKeyboardString = "DIK_F"
    Case &H22: GetKeyboardString = "DIK_G"
    Case &H23: GetKeyboardString = "DIK_H"
    Case &H24: GetKeyboardString = "DIK_J"
    Case &H25: GetKeyboardString = "DIK_K"
    Case &H26: GetKeyboardString = "DIK_L"
    Case &H27: GetKeyboardString = "DIK_SEMICOLON"
    Case &H28: GetKeyboardString = "DIK_APOSTROPHE"
    Case &H29: GetKeyboardString = "DIK_GRAVE"  ' accent grave
    Case &H2A: GetKeyboardString = "DIK_LSHIFT"
    Case &H2B: GetKeyboardString = "DIK_BACKSLASH"
    Case &H2C: GetKeyboardString = "DIK_Z"
    Case &H2D: GetKeyboardString = "DIK_X"
    Case &H2E: GetKeyboardString = "DIK_C"
    Case &H2F: GetKeyboardString = "DIK_V"
    Case &H30: GetKeyboardString = "DIK_B"
    Case &H31: GetKeyboardString = "DIK_N"
    Case &H32: GetKeyboardString = "DIK_M"
    Case &H33: GetKeyboardString = "DIK_COMMA"
    Case &H34: GetKeyboardString = "DIK_PERIOD"  ' . on main keyboard
    Case &H35: GetKeyboardString = "DIK_SLASH"  ' / on main keyboard
    Case &H36: GetKeyboardString = "DIK_RSHIFT"
    Case &H37: GetKeyboardString = "DIK_MULTIPLY"  ' * on numeric keypad
    Case &H38: GetKeyboardString = "DIK_LMENU"  ' left Alt
    Case &H39: GetKeyboardString = "DIK_SPACE"
    Case &H3A: GetKeyboardString = "DIK_CAPITAL"
    Case &H3B: GetKeyboardString = "DIK_F1"
    Case &H3C: GetKeyboardString = "DIK_F2"
    Case &H3D: GetKeyboardString = "DIK_F3"
    Case &H3E: GetKeyboardString = "DIK_F4"
    Case &H3F: GetKeyboardString = "DIK_F5"
    Case &H40: GetKeyboardString = "DIK_F6"
    Case &H41: GetKeyboardString = "DIK_F7"
    Case &H42: GetKeyboardString = "DIK_F8"
    Case &H43: GetKeyboardString = "DIK_F9"
    Case &H44: GetKeyboardString = "DIK_F10"
    Case &H45: GetKeyboardString = "DIK_NUMLOCK"
    Case &H46: GetKeyboardString = "DIK_SCROLL"  ' Scroll Lock
    Case &H47: GetKeyboardString = "DIK_NUMPAD7"
    Case &H48: GetKeyboardString = "DIK_NUMPAD8"
    Case &H49: GetKeyboardString = "DIK_NUMPAD9"
    Case &H4A: GetKeyboardString = "DIK_SUBTRACT"  ' - on numeric keypad
    Case &H4B: GetKeyboardString = "DIK_NUMPAD4"
    Case &H4C: GetKeyboardString = "DIK_NUMPAD5"
    Case &H4D: GetKeyboardString = "DIK_NUMPAD6"
    Case &H4E: GetKeyboardString = "DIK_ADD"  ' + on numeric keypad
    Case &H4F: GetKeyboardString = "DIK_NUMPAD1"
    Case &H50: GetKeyboardString = "DIK_NUMPAD2"
    Case &H51: GetKeyboardString = "DIK_NUMPAD3"
    Case &H52: GetKeyboardString = "DIK_NUMPAD0"
    Case &H53: GetKeyboardString = "DIK_DECIMAL"  ' . on numeric keypad
    Case &H56: GetKeyboardString = "DIK_OEM_102 < > | on UK/Germany keyboards"
    Case &H57: GetKeyboardString = "DIK_F11"
    Case &H58: GetKeyboardString = "DIK_F12"
    Case &H64: GetKeyboardString = "DIK_F13 on (NEC PC98: getkeyboardstring  "
    Case &H65: GetKeyboardString = "DIK_F14 on (NEC PC98: getkeyboardstring  "
    Case &H66: GetKeyboardString = "DIK_F15 on (NEC PC98: getkeyboardstring  "
    Case &H70: GetKeyboardString = "DIK_KANA on (Japanese keyboard: getkeyboardstring "
    Case &H73: GetKeyboardString = "DIK_ABNT_C1 / ? on Portugese (Brazilian: getkeyboardstring  keyboards "
    Case &H79: GetKeyboardString = "DIK_CONVERT on (Japanese keyboard: getkeyboardstring "
    Case &H7B: GetKeyboardString = "DIK_NOCONVERT on (Japanese keyboard: getkeyboardstring "
    Case &H7D: GetKeyboardString = "DIK_YEN on (Japanese keyboard: getkeyboardstring "
    Case &H7E: GetKeyboardString = "DIK_ABNT_C2 on Numpad . on Portugese (Brazilian: getkeyboardstring  keyboards "
    Case &H8D: GetKeyboardString = "DIK_NUMPADEQUALS = on numeric keypad (NEC PC98: getkeyboardstring  "
    Case &H90: GetKeyboardString = "DIK_PREVTRACK on Previous Track (DIK_CIRCUMFLEX on Japanese keyboard: getkeyboardstring  "
    Case &H91: GetKeyboardString = "DIK_AT (NEC PC98: getkeyboardstring  "
    Case &H92: GetKeyboardString = "DIK_COLON (NEC PC98: getkeyboardstring  "
    Case &H93: GetKeyboardString = "DIK_UNDERLINE (NEC PC98: getkeyboardstring  "
    Case &H94: GetKeyboardString = "DIK_KANJI on (Japanese keyboard: getkeyboardstring "
    Case &H95: GetKeyboardString = "DIK_STOP (NEC PC98: getkeyboardstring  "
    Case &H96: GetKeyboardString = "DIK_AX (Japan AX: getkeyboardstring  "
    Case &H97: GetKeyboardString = "DIK_UNLABELED (J3100: getkeyboardstring  "
    Case &H99: GetKeyboardString = "DIK_NEXTTRACK"  ' Next Track
    Case &H9C: GetKeyboardString = "DIK_NUMPADENTER"  ' Enter on numeric keypad
    Case &H9D: GetKeyboardString = "DIK_RCONTROL"
    Case &HA0: GetKeyboardString = "DIK_MUTE"  ' Mute
    Case &HA1: GetKeyboardString = "DIK_CALCULATOR"  ' Calculator
    Case &HA2: GetKeyboardString = "DIK_PLAYPAUSE"  ' Play / Pause
    Case &HA4: GetKeyboardString = "DIK_MEDIASTOP"  ' Media Stop
    Case &HAE: GetKeyboardString = "DIK_VOLUMEDOWN"  ' Volume -
    Case &HB0: GetKeyboardString = "DIK_VOLUMEUP"  ' Volume +
    Case &HB2: GetKeyboardString = "DIK_WEBHOME"  ' Web home
    Case &HB3: GetKeyboardString = "DIK_NUMPADCOMMA"  ' , on numeric keypad (NEC PC98 getkeyboardstring
    Case &HB5: GetKeyboardString = "DIK_DIVIDE"  ' / on numeric keypad
    Case &HB7: GetKeyboardString = "DIK_SYSRQ"
    Case &HB8: GetKeyboardString = "DIK_RMENU"  ' right Alt
    Case &HC5: GetKeyboardString = "DIK_PAUSE"  ' Pause
    Case &HC7: GetKeyboardString = "DIK_HOME"  ' Home on arrow keypad
    Case &HC8: GetKeyboardString = "DIK_UP"  ' UpArrow on arrow keypad
    Case &HC9: GetKeyboardString = "DIK_PRIOR"  ' PgUp on arrow keypad
    Case &HCB: GetKeyboardString = "DIK_LEFT"  ' LeftArrow on arrow keypad
    Case &HCD: GetKeyboardString = "DIK_RIGHT"  ' RightArrow on arrow keypad
    Case &HCF: GetKeyboardString = "DIK_END"  ' End on arrow keypad
    Case &HD0: GetKeyboardString = "DIK_DOWN"  ' DownArrow on arrow keypad
    Case &HD1: GetKeyboardString = "DIK_NEXT"  ' PgDn on arrow keypad
    Case &HD2: GetKeyboardString = "DIK_INSERT"  ' Insert on arrow keypad
    Case &HD3: GetKeyboardString = "DIK_DELETE"  ' Delete on arrow keypad
    Case &HDB: GetKeyboardString = "DIK_LWIN"  ' Left Windows key
    Case &HDC: GetKeyboardString = "DIK_RWIN"  ' Right Windows key
    Case &HDD: GetKeyboardString = "DIK_APPS"  ' AppMenu key
    Case &HDE: GetKeyboardString = "DIK_POWER"  ' System Power
    Case &HDF: GetKeyboardString = "DIK_SLEEP"  ' System Sleep
    Case &HE3: GetKeyboardString = "DIK_WAKE"  ' System Wake
    Case &HE5: GetKeyboardString = "DIK_WEBSEARCH"  ' Web Search
    Case &HE6: GetKeyboardString = "DIK_WEBFAVORITES"  ' Web Favorites
    Case &HE7: GetKeyboardString = "DIK_WEBREFRESH"  ' Web Refresh
    Case &HE8: GetKeyboardString = "DIK_WEBSTOP"  ' Web Stop
    Case &HE9: GetKeyboardString = "DIK_WEBFORWARD"  ' Web Forward
    Case &HEA: GetKeyboardString = "DIK_WEBBACK"  ' Web Back
    Case &HEB: GetKeyboardString = "DIK_MYCOMPUTER"  ' My Computer
    Case &HEC: GetKeyboardString = "DIK_MAIL"  ' Mail
    Case &HED: GetKeyboardString = "DIK_MEDIASELECT"  ' Media Select
    Case 238: GetKeyboardString = "Joystick Button 1"
    Case 239: GetKeyboardString = "Joystick Button 2"
    Case 240: GetKeyboardString = "Joystick Button 3"
    Case 241: GetKeyboardString = "Joystick Button 4"
    Case 242: GetKeyboardString = "Joystick Button 5"
    Case 243: GetKeyboardString = "Joystick Button 6"
    Case 244: GetKeyboardString = "Joystick Button 7"
    Case 245: GetKeyboardString = "Joystick Button 8"
    Case 246: GetKeyboardString = "Joystick Button 9"
    Case 247: GetKeyboardString = "Joystick Button 10"
    Case 248: GetKeyboardString = "Joystick Button 11"
    Case 249: GetKeyboardString = "Joystick Button 12"
    Case 250: GetKeyboardString = "Joystick Button 13"
    Case 251: GetKeyboardString = "Joystick Button 14"
    Case 252: GetKeyboardString = "Joystick Button 15"
    Case 253: GetKeyboardString = "Joystick Button 16"
    Case 254: GetKeyboardString = "Joystick Button 17"
    Case 255: GetKeyboardString = "Joystick Button 18"
    Case 256: GetKeyboardString = "Joystick Button 19"
    Case 257: GetKeyboardString = "Joystick Button 20"
    Case 258: GetKeyboardString = "Joystick Button 21"
    Case 259: GetKeyboardString = "Joystick Button 22"
    Case 260: GetKeyboardString = "Joystick Button 23"
    Case 261: GetKeyboardString = "Joystick Button 24"
    Case 262: GetKeyboardString = "Joystick Button 25"
    Case 263: GetKeyboardString = "Joystick Button 26"
    Case 264: GetKeyboardString = "Joystick Button 27"
    Case 265: GetKeyboardString = "Joystick Button 28"
    Case 266: GetKeyboardString = "Joystick Button 29"
    Case 267: GetKeyboardString = "Joystick Button 30"
    Case 268: GetKeyboardString = "Joystick Button 31"
    Case 269: GetKeyboardString = "Joystick Button 32"
    'transfer
    Case 270: GetKeyboardString = "X-Axis Positive"
    Case 271: GetKeyboardString = "X-Axis Negative"
    Case 272: GetKeyboardString = "Y-Axis Positive"
    Case 273: GetKeyboardString = "Y-Axis Negative"
    Case 274: GetKeyboardString = "Z-Axis Positive"
    Case 275: GetKeyboardString = "Z-Axis Negative"
    Case 276: GetKeyboardString = "RX-Axis Positive"
    Case 277: GetKeyboardString = "RX-Axis Negative"
    Case 278: GetKeyboardString = "RY-Axis Positive"
    Case 279: GetKeyboardString = "RY-Axis Negative"
    Case 280: GetKeyboardString = "RZ-Axis Positive"
    Case 281: GetKeyboardString = "RZ-Axis Negative"
    Case 282: GetKeyboardString = "Slider1 Positive"
    Case 283: GetKeyboardString = "Slider1 Negative"
    Case 284: GetKeyboardString = "Slider2 Positive"
    Case 285: GetKeyboardString = "Slider2 Negative"
    Case 286: GetKeyboardString = "POV1 Up"
    Case 287: GetKeyboardString = "POV1 Down"
    Case 288: GetKeyboardString = "POV1 Left"
    Case 289: GetKeyboardString = "POV1 Right"
    Case 290: GetKeyboardString = "POV2 Up"
    Case 291: GetKeyboardString = "POV2 Down"
    Case 292: GetKeyboardString = "POV2 Left"
    Case 293: GetKeyboardString = "POV2 Right"
    Case 294: GetKeyboardString = "POV3 Up"
    Case 295: GetKeyboardString = "POV3 Down"
    Case 296: GetKeyboardString = "POV3 Left"
    Case 297: GetKeyboardString = "POV3 Right"
    Case 298: GetKeyboardString = "MOUSE_BUTTON0"
    Case 299: GetKeyboardString = "MOUSE_BUTTON1"
    Case 300: GetKeyboardString = "MOUSE_BUTTON2"
    Case 301: GetKeyboardString = "MOUSE_BUTTON3"
    Case 302: GetKeyboardString = ""
    Case 303: GetKeyboardString = ""
    Case 304: GetKeyboardString = ""
    Case 305: GetKeyboardString = ""
    Case 306: GetKeyboardString = "MOUSE_WHEELUP"
    Case 307: GetKeyboardString = "MOUSE_WHEELDOWN"
    Case Else: GetKeyboardString = "Unknown " & iNum
End Select
End Function
