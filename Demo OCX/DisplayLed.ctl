VERSION 5.00
Begin VB.UserControl DisplayLed 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   Picture         =   "DisplayLed.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   ToolboxBitmap   =   "DisplayLed.ctx":0464
   Begin VB.PictureBox Image1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   0
      Picture         =   "DisplayLed.ctx":0776
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   0
      Width           =   240
   End
   Begin VB.Image ImgBack 
      Appearance      =   0  'Flat
      Height          =   1320
      Index           =   3
      Left            =   1920
      Picture         =   "DisplayLed.ctx":0BDA
      Top             =   1200
      Width           =   960
   End
   Begin VB.Image ImgBack 
      Appearance      =   0  'Flat
      Height          =   990
      Index           =   2
      Left            =   1080
      Picture         =   "DisplayLed.ctx":4E1E
      Top             =   1200
      Width           =   720
   End
   Begin VB.Image ImgBack 
      Appearance      =   0  'Flat
      Height          =   660
      Index           =   1
      Left            =   480
      Picture         =   "DisplayLed.ctx":7382
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image ImgBack 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   0
      Left            =   120
      Picture         =   "DisplayLed.ctx":8446
      Top             =   1200
      Width           =   240
   End
End
Attribute VB_Name = "DisplayLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================
' Descrizione.....: Display Led
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

Dim I As Integer
Dim Matrice(255, 4) As Byte

Private M_Value As Integer
Private M_Zoom As Integer
Private M_Colore As Long
Private M_Style As Integer
'                                Dichiarazione Eventi
Public Event Change(Value As Integer)

'
'      Inizializza le Variabili ( Solo Progetazione )
'
Private Sub UserControl_InitProperties()
     M_Value = 0
     M_Zoom = 1
     M_Style = 1
     M_Colore = RGB(168, 255, 0)
     UserControl.Height = 330
     UserControl.Width = 240
End Sub
'
'                        Resizing
'
Private Sub UserControl_Resize()
    Image1.Left = 0
    Image1.Top = 0
    
    UserControl.Height = 330 * M_Zoom
    UserControl.Width = 240 * M_Zoom
    
    Image1.Width = ScaleWidth
    Image1.Height = ScaleHeight
End Sub
'
'                       inizializa
'
Private Sub UserControl_Initialize()
  UserControl.Height = 330 * M_Zoom
  UserControl.Width = 240 * M_Zoom
  Call LeggeMatrici
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
   Call Scrive(M_Value)
End Property
'
Public Property Get Zoom() As Long
   Zoom = M_Zoom
End Property
Public Property Let Zoom(ByVal NewValue As Long)
   
   If NewValue > 4 Then NewValue = 4
   If NewValue < 1 Then NewValue = 1
   
   M_Zoom = NewValue
   PropertyChanged "Zoom"
   '
   UserControl.Height = 330 * M_Zoom
   UserControl.Width = 240 * M_Zoom
   Call CaricaFondo(M_Zoom)
   Call Scrive(M_Value)
End Property
'
Public Property Get Style() As Long
   Style = M_Style
End Property
Public Property Let Style(ByVal NewValue As Long)
   
   If NewValue > 4 Then NewValue = 4
   If NewValue < 1 Then NewValue = 1
   
   M_Style = NewValue
   PropertyChanged "Style"
   '
 '  UserControl.Height = 330 * M_Zoom
 '  UserControl.Width = 240 * M_Zoom
   Call CaricaFondo(M_Zoom)
   Call Scrive(M_Value)
End Property
'
Public Property Get Colore() As Long
   Colore = M_Colore
End Property
Public Property Let Colore(ByVal NewValue As Long)
   M_Colore = NewValue
   PropertyChanged "Colore"
   Call Scrive(M_Value)
End Property
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  M_Value = PropBag.ReadProperty("Value", 0)
  M_Zoom = PropBag.ReadProperty("Zoom", 1)
  M_Style = PropBag.ReadProperty("Style", 1)
  M_Colore = PropBag.ReadProperty("Colore", RGB(168, 255, 0))
 '
  UserControl.Height = 330 * M_Zoom
  UserControl.Width = 240 * M_Zoom
  Call CaricaFondo(M_Zoom)
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("Value", M_Value, 0)
  Call PropBag.WriteProperty("Zoom", M_Zoom, 1)
  Call PropBag.WriteProperty("Style", M_Style, 1)
  Call PropBag.WriteProperty("Colore", M_Colore, RGB(168, 255, 0))
End Sub
'
'
'         Inizio Routine DisplayLed
'
'
Private Sub CaricaFondo(Zm As Integer)
 Image1.Picture = ImgBack(Zm - 1).Picture
End Sub
'
'                 Scrive
'
Private Sub Scrive(Valore As Integer)
 Dim Vrt As Integer
 Dim Rig As Integer
 Dim Col As Integer
 Dim Nib As Integer
 '
 Dim Tmp(4) As Byte
 Dim Scr As Integer
 Dim St As Integer
 '
 Dim PauseTime, start, Finish, TotalTime
 '
 Select Case Style
 
 Case 1 ' ===============  Style 1 Standard
  
  Image1.Cls
  For Col = 0 To 4
    Nib = 1
    For Rig = 0 To 6
      If (Matrice(Valore, Col) Or Nib) = Matrice(Valore, Col) Then
       Call Plot(Rig, Col)
      End If
     Nib = Nib * 2
    Next Rig
  Next Col
 
 Case 2 ' ===============  Style 2 Scroll Da Destra a Sinistra

  For Scr = 0 To 4 Step 1         ' Loop Pricipale Da Destra a Sinistra
     Call Pause(0.05)             ' Ritardo
    For St = 0 To 3               ' Loop Shift Carattere
      Tmp(St) = Tmp(St + 1)
    Next St                       ' Fine Loop Shift Carattere
    Tmp(4) = Matrice(Valore, Scr) '
    '
    Image1.Cls
    For Col = 0 To 4             ' Loop Scrive Carattere
      Nib = 1
     For Rig = 0 To 6
      
      If (Tmp(Col) Or Nib) = Tmp(Col) Then
       Call Plot(Rig, Col)
      End If
     
      Nib = Nib * 2
     Next Rig
    Next Col                     ' Fine Loop Scrive Carattere
  Next Scr                       ' Fine Loop Pricipale
  
Case 3 ' ===============  Style 2 Scroll Da Sinistra a Destra

 For Scr = 4 To 0 Step -1        ' Loop Pricipale Da Sinistra a Destra
   Call Pause(0.05)              ' Ritardo
  For St = 4 To 1 Step -1        ' Loop Shift Carattere
     Tmp(St) = Tmp(St - 1)
  Next St                        ' Fine Loop Shift Carattere
  Tmp(0) = Matrice(Valore, Scr)  '
  '
  Image1.Cls
  For Col = 0 To 4               ' Loop Scrive Carattere
    Nib = 1
    For Rig = 0 To 6
      
      If (Tmp(Col) Or Nib) = Tmp(Col) Then
       Call Plot(Rig, Col)
      End If
     
     Nib = Nib * 2
    Next Rig
  Next Col                       ' Fine Loop Scrive Carattere
 Next Scr                        ' Fine Loop Pricipale

 End Select
End Sub
'
'                       Pausa
'
Private Sub Pause(Tempo As Double)
Dim start As Double

  ' Tempo = 0.05                        ' Imposta la durata.
   start = Timer                        ' Imposta l'ora di inizio.
   Do While Timer < start + Tempo
      DoEvents                          ' Passa il controllo ad altri processi.
   Loop
End Sub

'
'                       Ploting
'
Private Sub Plot(Rig As Integer, Col As Integer)
 Dim Vrt As Integer
 Dim Hr As Integer
 Dim Lr As Integer
 
 Lr = 3 * M_Zoom
 Hr = Lr * Col
 
 For I = 1 To M_Zoom * 2
   
   Vrt = (M_Zoom + I) - 1        ' Base 0
   Vrt = (Vrt + (Lr * Rig))
   
   Image1.Line (Hr + M_Zoom, Vrt)-(Hr + Lr, Vrt), M_Colore
 Next I
'
End Sub

'------------------------------------------------------
'
'                       Matrice
'
'------------------------------------------------------
Private Sub LeggeMatrici()
'
'                           Vari
'
'                                   "
Matrice(34, 0) = 0
Matrice(34, 1) = 7
Matrice(34, 2) = 0
Matrice(34, 3) = 7
Matrice(34, 4) = 0
'                                   &
Matrice(38, 0) = 50
Matrice(38, 1) = 77
Matrice(38, 2) = 89
Matrice(38, 3) = 38
Matrice(38, 4) = 80
'                                   (
Matrice(40, 0) = 0
Matrice(40, 1) = 62
Matrice(40, 2) = 65
Matrice(40, 3) = 0
Matrice(40, 4) = 0
'                                   )
Matrice(41, 0) = 0
Matrice(41, 1) = 0
Matrice(41, 2) = 65
Matrice(41, 3) = 62
Matrice(41, 4) = 0
'                                   +
Matrice(43, 0) = 8
Matrice(43, 1) = 8
Matrice(43, 2) = 62
Matrice(43, 3) = 8
Matrice(43, 4) = 8
'                                   -
Matrice(45, 0) = 8
Matrice(45, 1) = 8
Matrice(45, 2) = 8
Matrice(45, 3) = 8
Matrice(45, 4) = 8
'                                   .
Matrice(46, 0) = 0
Matrice(46, 1) = 0
Matrice(46, 2) = 32
Matrice(46, 3) = 0
Matrice(46, 4) = 0
'                                   /
Matrice(47, 0) = 32
Matrice(47, 1) = 16
Matrice(47, 2) = 8
Matrice(47, 3) = 4
Matrice(47, 4) = 2
'
'                                   :
Matrice(58, 0) = 0
Matrice(58, 1) = 0
Matrice(58, 2) = 20
Matrice(58, 3) = 0
Matrice(58, 4) = 0
'                                   =
Matrice(61, 0) = 20
Matrice(61, 1) = 20
Matrice(61, 2) = 20
Matrice(61, 3) = 20
Matrice(61, 4) = 20
'
'                          Numerici
'
'                                   0
Matrice(48, 0) = 62
Matrice(48, 1) = 65
Matrice(48, 2) = 65
Matrice(48, 3) = 65
Matrice(48, 4) = 62
'                                   1
Matrice(49, 0) = 4
Matrice(49, 1) = 2
Matrice(49, 2) = 127
Matrice(49, 3) = 0
Matrice(49, 4) = 0
'                                   2
Matrice(50, 0) = 121
Matrice(50, 1) = 73
Matrice(50, 2) = 73
Matrice(50, 3) = 73
Matrice(50, 4) = 79
'                                   3
Matrice(51, 0) = 73
Matrice(51, 1) = 73
Matrice(51, 2) = 73
Matrice(51, 3) = 73
Matrice(51, 4) = 127
'                                   4
Matrice(52, 0) = 15
Matrice(52, 1) = 8
Matrice(52, 2) = 8
Matrice(52, 3) = 8
Matrice(52, 4) = 127
'                                   5
Matrice(53, 0) = 79
Matrice(53, 1) = 73
Matrice(53, 2) = 73
Matrice(53, 3) = 73
Matrice(53, 4) = 121
'                                   6
Matrice(54, 0) = 127
Matrice(54, 1) = 73
Matrice(54, 2) = 73
Matrice(54, 3) = 73
Matrice(54, 4) = 121
'                                   7
Matrice(55, 0) = 65
Matrice(55, 1) = 33
Matrice(55, 2) = 17
Matrice(55, 3) = 9
Matrice(55, 4) = 7
'                                   8
Matrice(56, 0) = 127
Matrice(56, 1) = 73
Matrice(56, 2) = 73
Matrice(56, 3) = 73
Matrice(56, 4) = 127
'                                   9
Matrice(57, 0) = 79
Matrice(57, 1) = 73
Matrice(57, 2) = 73
Matrice(57, 3) = 73
Matrice(57, 4) = 127
'
'                             AlfaNumerici
'
'                                   A
Matrice(65, 0) = 126
Matrice(65, 1) = 9
Matrice(65, 2) = 9
Matrice(65, 3) = 9
Matrice(65, 4) = 126
'                                   B
Matrice(66, 0) = 127
Matrice(66, 1) = 73
Matrice(66, 2) = 73
Matrice(66, 3) = 73
Matrice(66, 4) = 54
'                                   C
Matrice(67, 0) = 62
Matrice(67, 1) = 65
Matrice(67, 2) = 65
Matrice(67, 3) = 65
Matrice(67, 4) = 65
'                                   D
Matrice(68, 0) = 127
Matrice(68, 1) = 65
Matrice(68, 2) = 65
Matrice(68, 3) = 65
Matrice(68, 4) = 62
'                                   E
Matrice(69, 0) = 127
Matrice(69, 1) = 73
Matrice(69, 2) = 73
Matrice(69, 3) = 73
Matrice(69, 4) = 65
'                                   F
Matrice(70, 0) = 127
Matrice(70, 1) = 9
Matrice(70, 2) = 9
Matrice(70, 3) = 9
Matrice(70, 4) = 1
'                                   G
Matrice(71, 0) = 62
Matrice(71, 1) = 65
Matrice(71, 2) = 65
Matrice(71, 3) = 73
Matrice(71, 4) = 121
'                                   H
Matrice(72, 0) = 127
Matrice(72, 1) = 8
Matrice(72, 2) = 8
Matrice(72, 3) = 8
Matrice(72, 4) = 127
'                                   I
Matrice(73, 0) = 0
Matrice(73, 1) = 65
Matrice(73, 2) = 127
Matrice(73, 3) = 65
Matrice(73, 4) = 0
'                                   J
Matrice(74, 0) = 48
Matrice(74, 1) = 64
Matrice(74, 2) = 65
Matrice(74, 3) = 65
Matrice(74, 4) = 63
'                                   K
Matrice(75, 0) = 127
Matrice(75, 1) = 8
Matrice(75, 2) = 20
Matrice(75, 3) = 34
Matrice(75, 4) = 65
'                                   L
Matrice(76, 0) = 127
Matrice(76, 1) = 64
Matrice(76, 2) = 64
Matrice(76, 3) = 64
Matrice(76, 4) = 64
'                                   M
Matrice(77, 0) = 127
Matrice(77, 1) = 4
Matrice(77, 2) = 8
Matrice(77, 3) = 4
Matrice(77, 4) = 127
'                                   N
Matrice(78, 0) = 127
Matrice(78, 1) = 4
Matrice(78, 2) = 8
Matrice(78, 3) = 16
Matrice(78, 4) = 127
'                                   O
Matrice(79, 0) = 127
Matrice(79, 1) = 65
Matrice(79, 2) = 65
Matrice(79, 3) = 65
Matrice(79, 4) = 127
'                                   P
Matrice(80, 0) = 127
Matrice(80, 1) = 9
Matrice(80, 2) = 9
Matrice(80, 3) = 9
Matrice(80, 4) = 6
'                                   Q
Matrice(81, 0) = 127
Matrice(81, 1) = 65
Matrice(81, 2) = 81
Matrice(81, 3) = 33
Matrice(81, 4) = 95
'                                   R
Matrice(82, 0) = 127
Matrice(82, 1) = 9
Matrice(82, 2) = 25
Matrice(82, 3) = 41
Matrice(82, 4) = 70
'                                   S
Matrice(83, 0) = 79
Matrice(83, 1) = 73
Matrice(83, 2) = 73
Matrice(83, 3) = 73
Matrice(83, 4) = 121
'                                   T
Matrice(84, 0) = 1
Matrice(84, 1) = 1
Matrice(84, 2) = 127
Matrice(84, 3) = 1
Matrice(84, 4) = 1
'                                   U
Matrice(85, 0) = 127
Matrice(85, 1) = 64
Matrice(85, 2) = 64
Matrice(85, 3) = 64
Matrice(85, 4) = 127
'                                   V
Matrice(86, 0) = 31
Matrice(86, 1) = 32
Matrice(86, 2) = 64
Matrice(86, 3) = 32
Matrice(86, 4) = 31
'                                   W
Matrice(87, 0) = 127
Matrice(87, 1) = 32
Matrice(87, 2) = 16
Matrice(87, 3) = 32
Matrice(87, 4) = 127
'                                   X
Matrice(88, 0) = 99
Matrice(88, 1) = 20
Matrice(88, 2) = 8
Matrice(88, 3) = 20
Matrice(88, 4) = 99
'                                   Y
Matrice(89, 0) = 7
Matrice(89, 1) = 8
Matrice(89, 2) = 120
Matrice(89, 3) = 8
Matrice(89, 4) = 7
'                                   Z
Matrice(90, 0) = 97
Matrice(90, 1) = 81
Matrice(90, 2) = 73
Matrice(90, 3) = 69
Matrice(90, 4) = 67
End Sub
