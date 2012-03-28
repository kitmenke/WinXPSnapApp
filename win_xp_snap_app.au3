#Include <C:\Apps\AutoIt\Include\Constants.au3>

;******************************
;       Global variables
;******************************

Global $VERSION = "2.0"
Global $WINDOW_HEIGHT = GetHeight()
Global $WINDOW_HALF_WIDTH = @DesktopWidth/2

;******************************
;           Setup
;******************************

#NoTrayIcon
; 1 = Default tray menu items (Script Paused/Exit) will not be shown.
; 2 = Items in the menu aren't "checked" when you click them
Opt("TrayMenuMode",3)   
$snapleftitem = TrayCreateItem("Snap Left")
$snaprightitem = TrayCreateItem("Snap Right")
$customsnapitem = TrayCreateItem("Custom Snap")
TrayCreateItem("")
$aboutitem  = TrayCreateItem("About")
TrayCreateItem("")
$exititem   = TrayCreateItem("Exit")

TraySetState()

If HotKeySet("^!{LEFT}", "SnapLeft") = 0 OR HotKeySet("^!{RIGHT}", "SnapRight") = 0 Then
    MsgBox(0, "Hotkey", "Unable to set hotkeys. Check to see they are not already being used.")
EndIf

;******************************
;           Main loop
;******************************

While 1
    $msg = TrayGetMsg()
    Select
        Case $msg = 0
            ContinueLoop
        Case $msg = $aboutitem
            Msgbox(64, "About", "CTRL+ALT+LEFT  - snap the current window left" & @CR & @LF & "CTRL+ALT+RIGHT - snap the current window right" & @CR & @LF & "Windows XP Snap App by Kit Menke " & $VERSION)
        Case $msg = $customsnapitem
            $dims = GetCustomDimensions()
            If $dims <> "" Then
               WinWaitNotActive(GetActiveWindow(), "", 3)
               Snap($dims[0], $dims[1], $dims[2], $dims[3])
            EndIf
        Case $msg = $snapleftitem
            WinWaitNotActive(GetActiveWindow(), "", 3)
            SnapLeft()
        Case $msg = $snaprightitem
            WinWaitNotActive(GetActiveWindow(), "", 3)
            SnapRight()
        Case $msg = $exititem
            ExitLoop
    EndSelect
WEnd

Exit


;******************************
;           Functions
;******************************

Func GetHeight()
    $height = @DesktopHeight
    ; adjust height to account for start menu bar
    $handle = WinGetHandle("[CLASS:Shell_TrayWnd]", "")
    If $handle <> "" Then
        $size = WinGetClientSize($handle)
        $height = $height - $size[1]
    EndIf
    Return $height
EndFunc

Func SnapLeft()
    Snap(0, 0, $WINDOW_HALF_WIDTH, $WINDOW_HEIGHT)
EndFunc

Func SnapRight()
    Snap($WINDOW_HALF_WIDTH, 0, $WINDOW_HALF_WIDTH, $WINDOW_HEIGHT)
EndFunc


Func Snap($x, $y, $w, $h)
    $activeWindowHandle = GetActiveWindow()
    
    ; special case for OC
    If WinGetTitle($activeWindowHandle) = "Office Communicator" Then
        $w = 250
        If $x <> 0 Then
            ; snap all the way right
            $x = @DesktopWidth - $w
        EndIf
    EndIf
    
    If $activeWindowHandle <> "" Then
        ;x,y,w,h
        WinMove($activeWindowHandle, "", $x, $y, $w, $h)
    EndIf
EndFunc

; Gets the handle of the currently active window
; Or "" if no window is active
Func GetActiveWindow()
    $listWindows = WinList()
    For $i = 1 to $listWindows[0][0]
        ; Search for the active window handle
        If $listWindows[$i][0] <> "" AND IsActive($listWindows[$i][1]) Then
            Return $listWindows[$i][1]
        EndIf
    Next
    Return ""
EndFunc

Func IsActive($handle)
    If BitAnd( WinGetState($handle), 8 ) Then 
        Return 1
    Else
        Return 0
    EndIf
EndFunc

Func GetCustomDimensions()
    $dimensions = InputBox("Custom Snap", "Enter the dimensions to set the next active window: (x,y,width,height)", "0,0,800,600")
    $dimensions = StringSplit($dimensions, ",", 2)
    If IsArray($dimensions) AND UBound($dimensions) = 4 Then
        Dim $size[4]
        For $i = 0 To 3
            $size[$i] = Int($dimensions[$i])
        Next
        return $size
    EndIf
    return ""
EndFunc

