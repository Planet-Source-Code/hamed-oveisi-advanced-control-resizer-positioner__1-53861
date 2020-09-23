VERSION 5.00
Object = "*\AProject1.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin prjAnchor.Anchor Anchor1 
      Left            =   2460
      Tag             =   "TTFF*/"
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   820
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2205
      Left            =   120
      TabIndex        =   0
      Tag             =   "TTTT*/"
      Top             =   180
      Width           =   4395
      Begin VB.CommandButton Command1 
         Caption         =   "Click!"
         Height          =   315
         Left            =   3450
         TabIndex        =   2
         Tag             =   "FFTT*/This is a button"
         Top             =   1740
         Width           =   885
      End
      Begin VB.TextBox Text1 
         Height          =   1215
         Left            =   180
         MultiLine       =   -1  'True
         TabIndex        =   1
         Tag             =   "TTTT*/"
         Text            =   "AnchDemo.frx":0000
         Top             =   330
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'* New Advanced Anchor Control
'* Version 3 (May 2004)
'*
'* With these new functionality:
'* - Resizes controls same as .NET and Delphi
'* - Using Advanced Freezing Technic to increase the speed of
'*   Resizing Controls (2 to 3 times faster than previous versions!) -New
'* - Eliminates Fixed Size Controls nad Invisible At RunTime Controls
'*   From PropertyPage to increase the speed -New
'* - Saves the posion of form and open it at the exact saved position
'*   automaticaly on next show of the form -NEW
'* - Checks form Size and don't allow it to gets smaller
'* - And Lots of more....
'*
'* Developed By : Hamed Oveisi
'* Please leave your feedbacks and votes on PSC
'**********************************************

Private Sub Command1_Click()
   'Show that Anchor informations removed at runtime from Tag property
   MsgBox "Command Tag is :" & Command1.Tag & vbCr & "Anchors information eliminated at runtime!", vbInformation
End Sub

Private Sub Form_Load()
   'Do Your runtime resize and checks before calling the
   'DoInit method
   
   With Anchor1
      'Simply provide the RegString to save the position of form
      'and use it for next show of the form
      'Seprate the AppName,Section,Key (Same as VB GetSetting) with Comma
      'And Default value (Top and Left) with | Like below:
      .RegString = "AnchorCtrl,Positions,Form1,1110|540"
      .DoInit ' Has 2 optional arg, that can sets
              'Height and Width of the form when showing the form
   End With
End Sub
