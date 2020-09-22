VERSION 5.00
Begin VB.UserControl UpDown 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   ScaleHeight     =   2400
   ScaleWidth      =   3645
   ToolboxBitmap   =   "UpDown.ctx":0000
   Begin VB.Image UpDOwn_R 
      Height          =   255
      Index           =   1
      Left            =   740
      Picture         =   "UpDown.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Label LblValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image UpDOwn_L 
      Height          =   255
      Index           =   1
      Left            =   0
      Picture         =   "UpDown.ctx":081E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
   Begin VB.Image UpDOwn_L 
      Height          =   255
      Index           =   0
      Left            =   0
      Picture         =   "UpDown.ctx":0D2A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Image UpDOwn_R 
      Height          =   255
      Index           =   0
      Left            =   740
      Picture         =   "UpDown.ctx":1236
      Stretch         =   -1  'True
      Top             =   0
      Width           =   360
   End
End
Attribute VB_Name = "UpDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: UpDown
' Nome dei Files..:
' Data............: 27/10/2004
' Versione........: 1.0
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'=====================================================
'
'                Not For Commercial Use
'=====================================================
'
Option Explicit

Private M_Value As Integer
Private M_MinValue As Integer
Private M_MaxValue As Integer
'                                Dichiarazione Eventi
Public Event Change(Value As Integer)

'
'      Inizializza le Variabili ( Solo Progetazione )
'
Private Sub UserControl_InitProperties()
     M_Value = 0
     M_MinValue = 0
     M_MaxValue = 10
     UserControl.Height = 255
     UserControl.Width = 1080
End Sub
'
'                        Resizing
'
Private Sub UserControl_Resize()
 Dim I As Integer
  For I = 0 To 1
    UpDOwn_L(I).Left = 0
    UpDOwn_L(I).Top = 0

    UpDOwn_L(I).Height = ScaleHeight
    UpDOwn_R(I).Left = ScaleWidth - 360
    UpDOwn_R(I).Top = 0
    UpDOwn_R(I).Height = ScaleHeight
 Next I
    LblValue.Left = 360
    LblValue.Top = 0
    LblValue.Width = ScaleWidth - (360 * 2)
    LblValue.Height = ScaleHeight
 
End Sub
'
'                       inizializa
'
Private Sub UserControl_Initialize()
  LblValue.Caption = M_Value
End Sub
'
'                         Eventi
'
Private Sub ChangeEvent(Valore As Integer)
    RaiseEvent Change(Valore)
End Sub
'
'                                Property
'
'
Public Property Get Value() As Long
   Value = M_Value
End Property
Public Property Let Value(ByVal NewValue As Long)
   M_Value = NewValue
   PropertyChanged "Value"
   LblValue.Caption = Value
End Property
'
Public Property Get MinValue() As Long
   MinValue = M_MinValue
End Property
Public Property Let MinValue(ByVal NewValue As Long)
   M_MinValue = NewValue
   PropertyChanged "MinValue"
End Property
'
Public Property Get MaxValue() As Long
   MaxValue = M_MaxValue
End Property
Public Property Let MaxValue(ByVal NewValue As Long)
   M_MaxValue = NewValue
   PropertyChanged "MaxValue"
End Property
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  M_Value = PropBag.ReadProperty("Value", 0)
  M_MinValue = PropBag.ReadProperty("MinValue", 0)
  M_MaxValue = PropBag.ReadProperty("MaxValue", 5)
  LblValue.Caption = M_Value
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Value", M_Value, 0)
  Call PropBag.WriteProperty("MinValue", M_MinValue, 0)
  Call PropBag.WriteProperty("MaxValue", M_MaxValue, 5)
End Sub
'                  Cursori Updown
'
'                     Sinistro
Private Sub UpDOwn_L_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   UpDOwn_L(1).Visible = False
End Sub
Private Sub UpDOwn_L_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   UpDOwn_L(1).Visible = True
    If M_Value = M_MinValue Then Exit Sub
    M_Value = M_Value - 1
    LblValue.Caption = M_Value
    ChangeEvent Value

End Sub
'                     Destro
Private Sub UpDOwn_R_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   UpDOwn_R(1).Visible = False
End Sub
Private Sub UpDOwn_R_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   UpDOwn_R(1).Visible = True
   If M_Value = M_MaxValue Then Exit Sub
   M_Value = M_Value + 1
   LblValue.Caption = M_Value
   ChangeEvent Value
End Sub



