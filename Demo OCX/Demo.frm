VERSION 5.00
Begin VB.Form Demo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5520
   Icon            =   "Demo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin Progetto1.UpDown UpDown6 
      Height          =   255
      Left            =   1320
      TabIndex        =   51
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Value           =   1
      MinValue        =   1
      MaxValue        =   4
   End
   Begin VB.Timer TimerDisplay 
      Interval        =   90
      Left            =   3480
      Top             =   3840
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   0
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed DisplayLed2 
      Height          =   1320
      Left            =   2640
      TabIndex        =   28
      Top             =   2520
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   2328
      Zoom            =   4
   End
   Begin Progetto1.UpDown UpDown5 
      Height          =   255
      Left            =   1320
      TabIndex        =   24
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      Value           =   4
      MinValue        =   1
      MaxValue        =   4
   End
   Begin Progetto1.UpDown UpDown4 
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   2880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
   End
   Begin Progetto1.UpDown UpDown3 
      Height          =   255
      Left            =   1320
      TabIndex        =   22
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      MaxValue        =   9
   End
   Begin Progetto1.DisplayLed DisplayLed1 
      Height          =   660
      Index           =   0
      Left            =   2280
      TabIndex        =   18
      Top             =   960
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1164
      Zoom            =   2
   End
   Begin Progetto1.UpDown UpDown2 
      Height          =   255
      Left            =   2280
      TabIndex        =   17
      Top             =   1680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      Value           =   100
      MinValue        =   100
      MaxValue        =   200
   End
   Begin Progetto1.Slider Slider2 
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   4800
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      Value           =   200
      MinValue        =   200
      MaxValue        =   1000
      Picture         =   "Demo.frx":030A
   End
   Begin Progetto1.Slider Slider1 
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
      Value           =   5
      Picture         =   "Demo.frx":1876
   End
   Begin VB.TextBox Txt_VSlider2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Text            =   "0"
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox Txt_VSlider1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4080
      TabIndex        =   8
      Text            =   "0"
      Top             =   3000
      Width           =   615
   End
   Begin Progetto1.VSlider VSlider1 
      Height          =   1935
      Left            =   4200
      TabIndex        =   6
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3413
      Value           =   500
      MaxValue        =   1000
      Picture         =   "Demo.frx":2DE2
   End
   Begin VB.TextBox TxtSlider2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Text            =   "0"
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox TxtSlider1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Text            =   "0"
      Top             =   4440
      Width           =   615
   End
   Begin Progetto1.UpDown UpDown1 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      MaxValue        =   10
   End
   Begin Progetto1.VSlider VSlider2 
      Height          =   1935
      Left            =   4920
      TabIndex        =   10
      Top             =   960
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   3413
      Value           =   1000
      MinValue        =   100
      MaxValue        =   1000
      Picture         =   "Demo.frx":46CA
   End
   Begin Progetto1.DisplayLed DisplayLed1 
      Height          =   660
      Index           =   1
      Left            =   2760
      TabIndex        =   19
      Top             =   960
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1164
      Zoom            =   2
   End
   Begin Progetto1.DisplayLed DisplayLed1 
      Height          =   660
      Index           =   2
      Left            =   3240
      TabIndex        =   20
      Top             =   960
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   1164
      Zoom            =   2
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   1
      Left            =   360
      TabIndex        =   30
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   2
      Left            =   600
      TabIndex        =   31
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   3
      Left            =   840
      TabIndex        =   32
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   4
      Left            =   1080
      TabIndex        =   33
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   5
      Left            =   1320
      TabIndex        =   34
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   6
      Left            =   1560
      TabIndex        =   35
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   7
      Left            =   1800
      TabIndex        =   36
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   8
      Left            =   2040
      TabIndex        =   37
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   9
      Left            =   2280
      TabIndex        =   38
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   10
      Left            =   2520
      TabIndex        =   39
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   11
      Left            =   2760
      TabIndex        =   40
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   12
      Left            =   3000
      TabIndex        =   41
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   13
      Left            =   3240
      TabIndex        =   42
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   14
      Left            =   3480
      TabIndex        =   43
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   15
      Left            =   3720
      TabIndex        =   44
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   16
      Left            =   3960
      TabIndex        =   45
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   17
      Left            =   4200
      TabIndex        =   46
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   18
      Left            =   4440
      TabIndex        =   47
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   19
      Left            =   4680
      TabIndex        =   48
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   20
      Left            =   4920
      TabIndex        =   49
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin Progetto1.DisplayLed LedRiga 
      Height          =   330
      Index           =   21
      Left            =   5160
      TabIndex        =   50
      Top             =   120
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   52
      Top             =   3600
      Width           =   975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1785
      Left            =   4035
      Picture         =   "Demo.frx":5FB2
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1290
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Zoom"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Colore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Valore"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   2160
      Width           =   975
   End
   Begin VB.Shape Shape9 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   1695
      Left            =   120
      Top             =   2280
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   600
      Width           =   975
   End
   Begin VB.Shape Shape7 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   2160
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label LblPerc 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   13
      Top             =   2685
      Width           =   285
   End
   Begin VB.Label LblPerc 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   12
      Top             =   1845
      Width           =   255
   End
   Begin VB.Label LblPerc 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4560
      TabIndex        =   11
      Top             =   975
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   4080
      X2              =   5280
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   4080
      X2              =   5280
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Line Line1 
      X1              =   4080
      X2              =   5280
      Y1              =   1065
      Y2              =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "V Sliders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      Width           =   975
   End
   Begin VB.Label LblUpDown 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UpDown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label LblSlider 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H Sliders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   4035
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4035
      Width           =   1215
   End
   Begin VB.Shape ShapeSlider 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   120
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   120
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape5 
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   4080
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      Height          =   2655
      Left            =   3960
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================
' Descrizione.....: Demo OCX
' Nome dei Files..: VSlider - HSlider - UpDown - DisplayLed
' Data............: 27/10/2004
' Versione........: 1.0
' Sistema.........: Windows
' Scritto da......: Bruno Crepaldi ®
' E-Mail..........: bruno.crepax@libero.it
'===========================================================
'
'                    Not For Commercial Use
'===========================================================
Option Explicit

Private I As Integer
Private I1 As Integer
Private CntDs As Integer
Private StrDisplay As String
'
'
'
Private Sub Form_Load()
   
 TimerDisplay.Enabled = False
   
   TxtSlider1 = Slider1.Value           ' Slider1
   TxtSlider2 = Slider2.Value           ' Slider2
   
   Txt_VSlider1 = VSlider1.Value        ' Vslider1
   Txt_VSlider2 = VSlider2.Value        ' Vslider2
   
   LblUpDown = UpDown1.Value            ' UpDown
   
   Call UpDown2_Change(100)             ' Display 1
    
   DisplayLed2.Value = Asc("0")         ' Display 2
   '
   StrDisplay = "BY BRUNO CREPALDI - 2004 - "
   StrDisplay = StrDisplay + "HELLO FROM VENICE - ITALY - "
   
 TimerDisplay.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

'-------------------------------------------------------
'                 Horizontal Sliders
'-------------------------------------------------------
Private Sub Slider1_Change(Value As Long)
  TxtSlider1 = Value
End Sub

Private Sub Slider2_Change(Value As Long)
  TxtSlider2 = Value
End Sub

'
Private Sub TxtSlider1_Change()
 Slider1.Value = Val(TxtSlider1.Text)
End Sub
Private Sub TxtSlider2_Change()
 Slider2.Value = Val(TxtSlider2.Text)
End Sub
'--------------------------------------------------------
'                    DisplayLed Esempio 1
'--------------------------------------------------------
Private Sub UpDown2_Change(Value As Integer)
Dim Ch As String

Ch = Trim(Str(Value))

For I = 1 To Len(Ch)
 DisplayLed1(I - 1).Value = Asc(Mid$(Ch, I)) ' Ascii Value
Next I
End Sub
Private Sub CambiaColore(Colore As Long)
  For I = 0 To 2
    DisplayLed1(I).Colore = Colore       ' Colore
  Next I
End Sub
'--------------------------------------------------------
'                    DisplayLed Esempio 2
'--------------------------------------------------------
Private Sub UpDown3_Change(Value As Integer)
 DisplayLed2.Value = Asc(Trim(Str(Value))) '    Valore Ascii
End Sub
Private Sub UpDown4_Change(Value As Integer)
 Select Case Value
   Case 0
     DisplayLed2.Colore = RGB(168, 255, 0)
   Case 1
     DisplayLed2.Colore = &HFFFFFF
   Case 2
     DisplayLed2.Colore = RGB(255, 90, 0)
   Case 3
     DisplayLed2.Colore = RGB(252, 255, 0)
   Case 4
     DisplayLed2.Colore = RGB(168, 250, 255)
   Case 5
     DisplayLed2.Colore = RGB(255, 150, 200)
 End Select
End Sub
Private Sub UpDown5_Change(Value As Integer)
 DisplayLed2.Zoom = Value
End Sub
Private Sub UpDown6_Change(Value As Integer)
 DisplayLed2.Style = Value
End Sub
'--------------------------------------------------------
'                  Vertical Sliders
'--------------------------------------------------------
Private Sub VSlider1_Change(Value As Long)
  Txt_VSlider1 = Value
End Sub
Private Sub VSlider2_Change(Value As Long)
  Txt_VSlider2 = Value
End Sub
Private Sub Txt_VSlider1_Change()
 VSlider1.Value = Val(Txt_VSlider1.Text)
End Sub
Private Sub Txt_VSlider2_Change()
 VSlider2.Value = Val(Txt_VSlider2.Text)
End Sub
'--------------------------------------------------------
'                       UpDown
'--------------------------------------------------------
Private Sub UpDown1_Change(Value As Integer)
 LblUpDown = Value
End Sub
'--------------------------------------------------------
'                     Riga Display
'--------------------------------------------------------
Private Sub TimerDisplay_Timer()
  Dim Ch As String * 1
  Dim Lstr As String
     
   Lstr = Len(StrDisplay)
   
   CntDs = CntDs + 1: If CntDs > Lstr Then CntDs = 1
     
     For I1 = 0 To 20
       LedRiga(I1).Value = LedRiga(I1 + 1).Value
     Next I1
    
    Ch = Mid$(StrDisplay, CntDs, 1)
    LedRiga(21).Value = Asc(Ch)
   
End Sub

