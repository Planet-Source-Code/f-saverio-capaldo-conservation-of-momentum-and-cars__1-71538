VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Conservation of Momentum"
   ClientHeight    =   11580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   772
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   992
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   " Data and Results "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   720
      TabIndex        =   1
      Top             =   8520
      Width           =   8895
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   7
         Text            =   "2"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Text            =   "4"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   5
         Text            =   "4"
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   4
         Text            =   "2"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   6120
         TabIndex        =   3
         Text            =   "2,5"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   6120
         TabIndex        =   2
         Text            =   "0,2"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "CAR Left"
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
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Mass"
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
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Speed"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "CAR Right"
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
         Left            =   2400
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Mass"
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
         Left            =   2400
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Speed"
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
         Left            =   2400
         TabIndex        =   16
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Deceleration"
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
         Left            =   4680
         TabIndex        =   15
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Elasticity"
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
         Left            =   4680
         TabIndex        =   14
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label Label2 
         Caption         =   "Speed 2"
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
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Speed 2"
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
         Index           =   3
         Left            =   2400
         TabIndex        =   12
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Speed"
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
         Index           =   4
         Left            =   1200
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Speed"
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
         Index           =   4
         Left            =   3480
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Time 2"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label4 
         Caption         =   "Elasticity"
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
         Index           =   3
         Left            =   6120
         TabIndex        =   8
         Top             =   1440
         Width           =   1350
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "START"
      Height          =   615
      Left            =   11640
      TabIndex        =   0
      Top             =   10560
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1560
      Top             =   360
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
Xcol = 0
Ycol = 0
    
Pig = 4 * Atn(1)
    
'Mass car 1 be positive
m1 = Abs(CDbl(Text1(0).Text))
'Speed car 1 be positive
v1 = Abs(CDbl(Text1(1).Text))
'Mass car 2 be positive
m2 = Abs(CDbl(Text2(0).Text))
'Speed car 2 be negative
v2 = -Abs(CDbl(Text2(1).Text))

'Debug.Print m1, m2, v1, v2
   
'Position
    x1 = 50
    x2 = 900
    
'Radii, now dimension, as function of mass
If m1 < m2 Then
   r1 = 20
   r2 = Int(r1 * m2 ^ (1 / 3) / m1 ^ (1 / 2))
Else
   r2 = 20
   r1 = Int(r2 * m1 ^ (1 / 3) / m2 ^ (1 / 2))
End If

    
    Coll = 0
    Tcol = 0
    B1 = 0
    B2 = 0
    vt1 = 0
    vt2 = 0
    
Dec = CDbl(Text3(0).Text)
Ec = CDbl(Text3(1).Text)

Call Punti_Auto_1(Veic_01(), r1)
Call Punti_Auto_2(Veic_21(), r1)
Call Punti_Auto_1(Veic_02(), r2)
Call Punti_Auto_2(Veic_22(), r2)
   
    
   Timer1.Enabled = True
   
End Sub

Private Sub Command2_Click()



Call Dis_Auto(Form1, 750, Form1.ScaleHeight / 2 - 70, 3.14 / 2, 9)
'Form1.PSet (750, Form1.ScaleHeight / 2 - 70), QBColor(9)
'Form1.DrawWidth = 2

'Form1.Line -(250, Form1.ScaleHeight / 2 - 70), QBColor(9)
'Form1.DrawWidth = 1

End Sub


Private Sub Form_Load()
   
   Call Centra_Form(Form1)
   Timer1.Enabled = False
   
End Sub

Private Sub Timer1_Timer()
    
    
    'update position
    Select Case Coll
      Case 0
         x1 = x1 + v1
         x2 = x2 + v2
      Case 1
         ' the two balls have collided, spped sign change
         ' acceleration Dec be consider on Timer interval
         '(10 millisecondi)
         Tcol = Tcol + 0.01
         
         If B1 = 0 Then Vc1 = (v1 + (Dec * Tcol))
         If B2 = 0 Then Vc2 = (v2 - (Dec * Tcol))
         
         If Abs(Vc1) <= 0.25 Then
            Vc1 = 0
            B1 = 1
         End If
         If Abs(Vc2) <= 0.25 Then
            Vc2 = 0
            B2 = 1
         End If
         
         If B1 = 0 Then x1 = x1 + Vc1
         If B2 = 0 Then x2 = x2 + Vc2
      Case Else
   End Select
      
         
    
    'store temporary velocities
    vt1 = v1
    vt2 = v2
    
    'the two balls have collided
    If (x1 + (2.25 * r1)) >= (x2 - (2.25 * r2)) And Coll = 0 Then
        'calculate speeds after collision (conservation of momentum)
        v1 = ((m1 - m2) / (m1 + m2)) * vt1 + ((2 * m2) / (m1 + m2)) * vt2
        v2 = ((2 * m1) / (m1 + m2)) * vt1 + ((m2 - m1) / (m1 + m2)) * vt2
        'consider elasticity coefficient Ec
        v1 = v1 * Ec
        v2 = v2 * Ec
        Coll = 1
        Tcol = 0
        Xcol = ((x1 + (2.25 * r1)) + (x2 - (2.25 * r2))) / 2
        Ycol = Me.ScaleHeight / 2 - 70
        'Call Marca_Punto(Form1, Xcol, Ycol, 12)
    End If
    
    Label2(4).Caption = v1
    Label3(4).Caption = v2
    Label4(3).Caption = Tcol
    
    Me.Refresh
    'Me.DrawStyle = 1
    Call Dis_Auto(Form1, x1, Me.ScaleHeight / 2 - 70, Veic_01(), Veic_21(), Pig * 3 / 2, 4)
    Call Dis_Auto(Form1, x2, Me.ScaleHeight / 2 - 70, Veic_02(), Veic_22(), Pig / 2, 2)
    'Me.DrawStyle = 3
    Me.Line (30, Me.ScaleHeight / 2 - 120)-(950, Me.ScaleHeight / 2 - 120), QBColor(0)
    Me.Line (30, Me.ScaleHeight / 2 - 20)-(950, Me.ScaleHeight / 2 - 20), QBColor(0)
    'Me.DrawStyle = 1
    Call Marca_Punto(Form1, Xcol, Ycol, 12)

    If B1 = 1 And B2 = 1 Then Timer1.Enabled = False
    
    
    
End Sub
