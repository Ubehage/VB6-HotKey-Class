VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum RegisterHotKeyModifiers
  hkAlt = &H1
  hkControl = &H2
  hkShift = &H4
End Enum

Public Type HOTKEY_Item
  Key As VBRUN.KeyCodeConstants
  Modifier As RegisterHotKeyModifiers
End Type

Dim m_hWnd As Long
Dim m_HotKeys As Long
Dim m_HotKey() As HOTKEY_Item

Public Event HotkeyPressed(HotKey As HOTKEY_Item)

Public Property Get hWnd() As Long
  hWnd = m_hWnd
End Property
Public Property Get HotKeys() As Long
  HotKeys = m_HotKeys
End Property
Public Property Get HotKey(Optional Index As Long, Optional KeyCode As VBRUN.KeyCodeConstants, Optional SpecialKeys As RegisterHotKeyModifiers = hkControl Or hkShift) As HOTKEY_Item
  Dim hkIndex As Long
  If Index = 0 Then
    hkIndex = GetHotkeyIndex(KeyCode, SpecialKeys)
  Else
    hkIndex = Index
  End If
  If (hkIndex > 0 And hkIndex <= m_HotKeys) Then
    HotKey = m_HotKey(Index)
  End If
End Property
Public Property Let hWnd(New_hWnd As Long)
  If Not m_hWnd = New_hWnd Then
    If IsHotkeyhWndRegistered(m_hWnd) Then
      UnregisterAllHotKeys
      Call UnregisterHotkeyhWnd(m_hWnd)
    End If
    m_hWnd = New_hWnd
    If Not m_hWnd = 0 Then
      Call RegisterHotkeyhWnd(m_hWnd)
    End If
  End If
End Property

Public Function RegisterHotKey(KeyCode As VBRUN.KeyCodeConstants, Optional SpecialKeys As RegisterHotKeyModifiers = hkControl Or hkShift) As Boolean
  If Not m_hWnd = 0 Then
    If Not IsHotKeyRegistered(KeyCode, SpecialKeys) Then
      If RegisterNewHotKey(m_hWnd, Me, KeyCode, SpecialKeys) Then
        RegisterHotKey = True
        If (m_HotKeys Mod 10) = 0 Then
          ReDim Preserve m_HotKey(1 To (m_HotKeys + 10)) As HOTKEY_Item
        End If
        m_HotKeys = (m_HotKeys + 1)
        With m_HotKey(m_HotKeys)
          .Key = KeyCode
          .Modifier = SpecialKeys
        End With
      End If
    End If
  End If
End Function

Public Function UnRegisterHotKey(KeyCode As VBRUN.KeyCodeConstants, Optional SpecialKeys As RegisterHotKeyModifiers = hkControl Or hkShift) As Boolean
  Dim i As Long
  Dim hIndex As Long
  hIndex = GetHotkeyIndex(KeyCode, SpecialKeys)
  If Not hIndex = 0 Then
    If ReleaseHotkey(KeyCode, SpecialKeys) Then
      For i = hIndex To (m_HotKeys - 1)
        With m_HotKey(i)
          .Key = m_HotKey((i + 1)).Key
          .Modifier = m_HotKey((i + 1)).Modifier
        End With
      Next
      With m_HotKey(m_HotKeys)
        .Key = 0
        .Modifier = 0
      End With
      m_HotKeys = (m_HotKeys - 1)
      If (m_HotKeys Mod 10) = 0 Then
        If m_HotKeys = 0 Then
          Erase m_HotKey
        Else
          ReDim Preserve m_HotKey(1 To m_HotKeys) As HOTKEY_Item
        End If
      End If
      UnRegisterHotKey = True
    Else
      UnRegisterHotKey = False
    End If
  Else
    UnRegisterHotKey = False
  End If
End Function

Public Function UnregisterAllHotKeys() As Boolean
  Dim uSuccess As Boolean
  uSuccess = True
  Do Until m_HotKeys = 0
    If Not UnregisterHotKeyByIndex(1) Then
      uSuccess = False
      Exit Do
    End If
  Loop
  UnregisterAllHotKeys = uSuccess
End Function

Friend Sub HotKeyPressedNow(HotkeyTrigger As VBRUN.KeyCodeConstants, HotkeyModifier As RegisterHotKeyModifiers)
  Dim hIndex As Long
  hIndex = GetHotkeyIndex(HotkeyTrigger, HotkeyModifier)
  If (hIndex > 0 And hIndex <= m_HotKeys) Then
    RaiseEvent HotkeyPressed(m_HotKey(hIndex))
  End If
End Sub

Private Function UnregisterHotKeyByIndex(Index As Long) As Boolean
  Dim hKey As VBRUN.KeyCodeConstants
  Dim hMod As RegisterHotKeyModifiers
  If (Index > 0 And Index <= m_HotKeys) Then
    With m_HotKey(Index)
      hKey = .Key
      hMod = .Modifier
    End With
    UnregisterHotKeyByIndex = UnRegisterHotKey(hKey, hMod)
  Else
    UnregisterHotKeyByIndex = False
  End If
End Function

Private Function IsHotKeyRegistered(KeyCode As VBRUN.KeyCodeConstants, SpecialKeys As RegisterHotKeyModifiers) As Boolean
  IsHotKeyRegistered = Not (GetHotkeyIndex(KeyCode, SpecialKeys) = 0)
End Function

Private Function GetHotkeyIndex(KeyCode As VBRUN.KeyCodeConstants, SpecialKeys As RegisterHotKeyModifiers) As Long
  Dim i As Long
  For i = 1 To m_HotKeys
    With m_HotKey(i)
      If (.Key = KeyCode And .Modifier = SpecialKeys) Then
        GetHotkeyIndex = i
        Exit For
      End If
    End With
  Next
End Function

Private Function GetKeyModifierFromCode(Modifier As Integer) As RegisterHotKeyModifiers
  Select Case Modifier
    Case hkAlt
      GetKeyModifierFromCode = hkAlt
    Case hkControl
      GetKeyModifierFromCode = hkControl
    Case hkShift
      GetKeyModifierFromCode = hkShift
  End Select
End Function

Private Sub Class_Terminate()
  UnregisterAllHotKeys
  If Not m_hWnd = 0 Then
    If IsHotkeyhWndRegistered(m_hWnd) Then
      Call UnregisterHotkeyhWnd(m_hWnd)
    End If
  End If
End Sub

