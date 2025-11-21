Attribute VB_Name = "modHotKey"
Option Explicit

Public Const HOTKEY_IDENTIFIER_START As Long = 42000

Private Const GWL_WNDPROC As Long = (-4)
Private Const WM_HOTKEY As Long = &H312

Dim HotkeyhWnds As Long
Dim HotkeyhWnd() As Long
Dim OldHotkeyWindowProc() As Long

Dim HotKeys As Long
Dim HotkeyTrigger() As VBRUN.KeyCodeConstants
Dim HotkeyModifier() As RegisterHotKeyModifiers
Dim HotKeyID() As Long
Dim HotkeyReggedhWnd() As Long
Dim HotkeyClass() As HotKey

Public lpPrevWndProc As Long

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal ID As Long, ByVal fsModifiers As RegisterHotKeyModifiers, ByVal vk As KeyCodeConstants) As Long
Private Declare Function UnRegisterHotKey Lib "user32" Alias "UnregisterHotKey" (ByVal hWnd As Long, ByVal ID As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Long
  Private Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Long) As Long

Public Function RegisterNewHotKey(hWnd As Long, ClassObject As HotKey, TriggerKey As VBRUN.KeyCodeConstants, ModifierKeys As RegisterHotKeyModifiers) As Boolean
  Dim NewID As Long
  Dim hResult As Boolean
  NewID = 0
  If Not hWnd = 0 Then
    If Not GethWndIndex(hWnd) = 0 Then
      If GetNewUniqueIdentifier(NewID) Then
        If RegisterHotKey(hWnd, NewID, ModifierKeys, TriggerKey) = 0 Then
          hResult = False
        Else
          hResult = True
        End If
      Else
        hResult = False
      End If
    Else
      hResult = False
    End If
  Else
    hResult = False
  End If
  If hResult Then
    AddHotkey hWnd, ClassObject, TriggerKey, ModifierKeys, NewID
  Else
    If Not NewID = 0 Then
      ReleaseIdentifier NewID
    End If
  End If
End Function

Public Function ReleaseHotkey(TriggerKey As VBRUN.KeyCodeConstants, ModifierKeys As RegisterHotKeyModifiers) As Boolean
  Dim hIndex As Long
  hIndex = GetHotkeyIndex(TriggerKey, ModifierKeys)
  If Not hIndex = 0 Then
    If UnRegisterHotKey(HotkeyReggedhWnd(hIndex), HotKeyID(hIndex)) = 0 Then
      ReleaseHotkey = False
    Else
      ReleaseHotkey = True
      RemoveHotkey hIndex
    End If
  Else
    ReleaseHotkey = False
  End If
End Function

Private Sub AddHotkey(hWnd As Long, ClassObject As HotKey, TriggerKey As VBRUN.KeyCodeConstants, ModifierKeys As RegisterHotKeyModifiers, ID As Long)
  If (HotKeys Mod 10) = 0 Then
    ReDim Preserve HotkeyTrigger(1 To (HotKeys + 10)) As VBRUN.KeyCodeConstants
    ReDim Preserve HotkeyModifier(1 To (HotKeys + 10)) As RegisterHotKeyModifiers
    ReDim Preserve HotKeyID(1 To (HotKeys + 10)) As Long
    ReDim Preserve HotkeyReggedhWnd(1 To (HotKeys + 10)) As Long
    ReDim Preserve HotkeyClass(1 To (HotKeys + 10)) As HotKey
  End If
  HotKeys = (HotKeys + 1)
  HotkeyTrigger(HotKeys) = TriggerKey
  HotkeyModifier(HotKeys) = ModifierKeys
  HotKeyID(HotKeys) = ID
  HotkeyReggedhWnd(HotKeys) = hWnd
  Set HotkeyClass(HotKeys) = ClassObject
End Sub

Private Sub RemoveHotkey(Index As Long)
  Dim i As Long
  For i = Index To (HotKeys - 1)
    HotkeyTrigger(i) = HotkeyTrigger((i + 1))
    HotkeyModifier(i) = HotkeyModifier((i + 1))
    HotKeyID(i) = HotKeyID((i + 1))
    HotkeyReggedhWnd(i) = HotkeyReggedhWnd((i + 1))
    Set HotkeyClass(i) = HotkeyClass((i + 1))
  Next
  HotkeyTrigger(HotKeys) = 0
  HotkeyModifier(HotKeys) = 0
  HotKeyID(HotKeys) = 0
  HotkeyReggedhWnd(HotKeys) = 0
  Set HotkeyClass(HotKeys) = Nothing
  HotKeys = (HotKeys - 1)
  If (HotKeys Mod 10) = 0 Then
    If HotKeys = 0 Then
      Erase HotkeyTrigger
      Erase HotkeyModifier
      Erase HotKeyID
      Erase HotkeyReggedhWnd
      Erase HotkeyClass
    Else
      ReDim Preserve HotkeyTrigger(1 To HotKeys) As VBRUN.KeyCodeConstants
      ReDim Preserve HotkeyModifier(1 To HotKeys) As RegisterHotKeyModifiers
      ReDim Preserve HotKeyID(1 To HotKeys) As Long
      ReDim Preserve HotkeyReggedhWnd(1 To HotKeys) As Long
      ReDim Preserve HotkeyClass(1 To HotKeys) As HotKey
    End If
  End If
End Sub

Private Function GetHotkeyIndex(TriggerKey As VBRUN.KeyCodeConstants, ModifierKeys As RegisterHotKeyModifiers) As Long
  Dim i As Long
  For i = 1 To HotKeys
    If HotkeyTrigger(i) = TriggerKey Then
      If HotkeyModifier(i) = ModifierKeys Then
        GetHotkeyIndex = i
        Exit For
      End If
    End If
  Next
End Function

Private Function GetHotkeyIndexFromID(ID As Long) As Long
  Dim i As Long
  For i = 1 To HotKeys
    If HotKeyID(i) = ID Then
      GetHotkeyIndexFromID = i
      Exit For
    End If
  Next
End Function

Private Function GetHotkeyIndexFromhWnd(hWnd As Long) As Long
  Dim i As Long
  For i = 1 To HotKeys
    If HotkeyReggedhWnd(i) = hWnd Then
      GetHotkeyIndexFromhWnd = i
      Exit For
    End If
  Next
End Function

Private Function GethWndIndex(hWnd As Long, Optional DontCreate As Boolean = False) As Long
  Dim i As Long
  Dim hIndex As Long
  hIndex = 0
  For i = 1 To HotkeyhWnds
    If HotkeyhWnd(i) = hWnd Then
      hIndex = i
      Exit For
    End If
  Next
  If Not DontCreate Then
    If hIndex = 0 Then
      hIndex = RegisterHotkeyhWnd(hWnd)
    End If
  End If
  GethWndIndex = hIndex
End Function

Public Function IsHotkeyhWndRegistered(hWnd As Long) As Boolean
  IsHotkeyhWndRegistered = Not (GethWndIndex(hWnd, True) = 0)
End Function

Public Function RegisterHotkeyhWnd(hWnd As Long) As Long
  HotkeyhWnds = (HotkeyhWnds + 1)
  ReDim Preserve HotkeyhWnd(1 To HotkeyhWnds) As Long
  ReDim Preserve OldHotkeyWindowProc(1 To HotkeyhWnds) As Long
  HotkeyhWnd(HotkeyhWnds) = hWnd
  OldHotkeyWindowProc(HotkeyhWnds) = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf HotKeyWinProc)
  RegisterHotkeyhWnd = HotkeyhWnds
End Function

Public Function UnregisterHotkeyhWnd(hWnd As Long) As Boolean
  Dim i As Long
  Dim hIndex As Long
  hIndex = GethWndIndex(hWnd, True)
  If Not hIndex = 0 Then
    Call SetWindowLong(hWnd, GWL_WNDPROC, OldHotkeyWindowProc(hIndex))
    For i = hIndex To (HotkeyhWnds - 1)
      HotkeyhWnd(i) = HotkeyhWnd((i + 1))
      OldHotkeyWindowProc(i) = OldHotkeyWindowProc((i + 1))
    Next
    HotkeyhWnd(HotkeyhWnds) = 0
    OldHotkeyWindowProc(HotkeyhWnds) = 0
    HotkeyhWnds = (HotkeyhWnds - 1)
    If HotkeyhWnds = 0 Then
      Erase HotkeyhWnd
      Erase OldHotkeyWindowProc
    Else
      ReDim Preserve HotkeyhWnd(1 To HotkeyhWnds) As Long
      ReDim Preserve OldHotkeyWindowProc(1 To HotkeyhWnds) As Long
    End If
    UnregisterHotkeyhWnd = True
  Else
    UnregisterHotkeyhWnd = False
  End If
End Function

Private Sub HotkeyPressed(ID As Long)
  Dim hIndex As Long
  hIndex = GetHotkeyIndexFromID(ID)
  If Not hIndex = 0 Then
  End If
End Sub

Private Function GetNewUniqueIdentifier(NewID As Long) As Boolean
  Dim nID As Long
  nID = GlobalAddAtom("ube_hotkey" & HotKeys & Now & Timer)
  If nID = 0 Then
    NewID = 0
    GetNewUniqueIdentifier = False
  Else
    NewID = nID
    GetNewUniqueIdentifier = True
  End If
End Function

Private Sub ReleaseIdentifier(ID As Long)
  Call GlobalDeleteAtom(ID)
End Sub

Public Function HotKeyWinProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim hIndex As Long
  hIndex = 0
  If uMsg = WM_HOTKEY Then
    HotkeyPressed wParam
    HotKeyWinProc = 0
  Else
    hIndex = GethWndIndex(hWnd, True)
    If Not hIndex = 0 Then
      HotKeyWinProc = CallWindowProc(OldHotkeyWindowProc(hIndex), hWnd, uMsg, wParam, lParam)
    Else
      HotKeyWinProc = 0
    End If
  End If
End Function
