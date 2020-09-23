VERSION 5.00
Begin VB.UserControl Anchor 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Anchor.ctx":0000
   PropertyPages   =   "Anchor.ctx":0BE2
   ScaleHeight     =   495
   ScaleWidth      =   495
   ToolboxBitmap   =   "Anchor.ctx":0BF2
End
Attribute VB_Name = "Anchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'******************************************
'* Developer : Hamed Oveisi               *
'******************************************

'------------------------------------------
'API Function for freezing the form
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const WM_PAINT = &HF
'-------------------------------------------

Dim LastHeight As Long, LastWidth As Long    'Form Last Height and Width values
Dim WithEvents Frm As Form
Attribute Frm.VB_VarHelpID = -1
Dim Ctrls() As String                        'Saves Controls Tag property @ RunTime

Public RegString As String
Attribute RegString.VB_VarMemberFlags = "400"
Public DontRefresh As Boolean                'Use Freezeing to increase the speed of resizing
Attribute DontRefresh.VB_VarMemberFlags = "400"
Public MinHeight As Long, MinWidth As Long   'Min Height and Width of Form so that form never get
                                             'smaller than these values

Public Event AfterResize(ByVal HeightChanged As Long, ByVal WidthChanged As Long)
Public Event BeforeResize()

Public Sub Freeze(ByVal ohWnd As Long)
    'Freezing the Application
    SendMessage ohWnd, WM_SETREDRAW, False, 0
End Sub

Public Sub UnFreeze(ByVal ohWnd As Long, Optional ByVal ForceUpdate As Boolean = True)
    'Unfreezing the Application
    SendMessage ohWnd, WM_SETREDRAW, 1&, 0
    'SendMessage ohWnd, WM_PAINT, 1&, 0
    If ForceUpdate Then
        RedrawWindow ohWnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN
    End If
End Sub

'********************************************
' DoResize checks the control tag and resize
'it.
'
' HeightChange and WidthChange are both save
'changing in height or width
'
'
' Use Tag this way: Left,Top,Right and Bottom
' If U want to resize your control from all 4 sides
'use Tag this way: TTTT (T means True)
'********************************************
Public Sub DoResize()
    On Error Resume Next
    Dim HeightChange As Long, WidthChange As Long
    Dim Tg As String
    Dim i As Integer
    
    Set Frm = Extender.Parent
    Set CtrlParent = Extender.Parent
    ' Exit sub on Minimize
    If Frm.WindowState = vbMinimized Then Exit Sub
    
    'Freezing the Controls Redarw, makes the application more faster
    'in resizing controls. It's Now faster than any other Anchoring
    'methods Like the one Used in Delphi!
    If Not DontRefresh Then Freeze Frm.hwnd
    
    ' Check the form for Min Values
    If Frm.Height <= MinHeight Then Frm.Height = MinHeight
    If Frm.Width <= MinWidth Then Frm.Width = MinWidth
    
    'Calculate the Changes
    HeightChange = Frm.Height - LastHeight
    WidthChange = Frm.Width - LastWidth
     
    'If this is not the first time of resize
    If LastHeight <> 0 And LastWidth <> 0 Then
        RaiseEvent BeforeResize
        
        For i = 0 To Frm.Controls.Count - 1
            Tg = Ctrls(i)
            If Tg = "" Then GoTo Nxt
            If Right(Tg, 2) <> "FF" Then
            'Checking Tag
             If Right(Tg, 1) = "T" Then
                If Mid(Tg, 2, 1) = "T" Then
                    Frm.Controls(i).Height = Frm.Controls(i).Height + HeightChange
                Else
                    Frm.Controls(i).Top = Frm.Controls(i).Top + HeightChange
                End If
             End If
            
             If Mid(Tg, 3, 1) = "T" Then
                If Left(Tg, 1) = "T" Then
                    Frm.Controls(i).Width = Frm.Controls(i).Width + WidthChange
                Else
                    Frm.Controls(i).Left = Frm.Controls(i).Left + WidthChange
                End If
             End If
            End If
Nxt:
         Next i

         RaiseEvent AfterResize(HeightChange, WidthChange)
    Else
        If MinHeight = 0 Then
          'This is the first Resize
           MinHeight = Frm.Height
           MinWidth = Frm.Width
        Else
           Frm.Height = MinHeight
           Frm.Width = MinWidth
        End If
    End If
    'Save Last values
    LastHeight = Frm.Height
    LastWidth = Frm.Width
    If Not DontRefresh Then UnFreeze Frm.hwnd, True
End Sub

Private Sub Frm_Resize()
    DoResize
End Sub

Private Sub UserControl_Initialize()
   On Error Resume Next
   Set Frm = Extender.Parent
   Set CtrlParent = Extender.Parent
   
End Sub

Private Sub UserControl_InitProperties()
   On Error Resume Next
   Set Frm = Extender.Parent
   Set CtrlParent = Extender.Parent
   
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    Set Frm = Extender.Parent
    Set CtrlParent = Extender.Parent
    Width = 480
    Height = 465
   
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
   Dim msg As String
   msg = "Anchor Property for Visual Basic V3.0" & vbCr
   msg = msg & "Written by : Hamed Oveisi"
   MsgBox msg, vbInformation, "About Anchor"
   
End Sub

Private Sub UserControl_Terminate()
   On Error Resume Next
   Dim Pos As String, RegItem
   
   If Frm.WindowState = vbNormal Then
      RegItem = Split(RegString, ",")
      Pos = Frm.Left & "|" & Frm.Top
      SaveSetting RegItem(0), RegItem(1), RegItem(2) & "Pos", Pos
   End If
   
   Set Frm = Nothing
   Set CtrlParent = Nothing
End Sub

Public Sub DoInit(Optional ByVal FormHeight As Long, Optional ByVal FormWidth As Long)
    On Error Resume Next
    Dim i As Integer
    Dim Tg As String
    Set Frm = Extender.Parent
    
    'Form Positioner
    Dim Pos, RegItem
    RegItem = Split(RegString, ",")
    Pos = Split(GetSetting(RegItem(0), RegItem(1), RegItem(2) & "Pos", RegItem(3)), "|")
    Frm.Left = Pos(0)
    Frm.Top = Pos(1)
    '---------------
    
    ReDim Ctrls(Frm.Controls.Count)
    'Save Tag Properties of Controls so that Tag
    'Can be use in runtime for other reasons
    For i = 0 To Frm.Controls.Count - 1
         Tg = Frm.Controls(i).Tag
         ' Every anchor information ends with */
         If InStr(1, Tg, "*/") > 0 Then
            Ctrls(i) = Left(Tg, 4)
            ' Now eliminate anchors from Tag property!
            ' So there is no dependency to
            ' the TAG property of the object @ RunTime.
            Frm.Controls(i).Tag = Right(Tg, Len(Tg) - 6)
         End If
    Next i
    
    If Not IsMissing(FormHeight) Then MinHeight = FormHeight
    If Not IsMissing(FormWidth) Then MinWidth = FormWidth
    
    LastHeight = Frm.Height
    LastWidth = Frm.Width
End Sub
