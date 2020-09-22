VERSION 5.00
Begin VB.Form LineCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Line Coding v1.1"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13035
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   13035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaxLength       =   64
      TabIndex        =   14
      Top             =   3360
      Width           =   5955
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2000
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   12735
      TabIndex        =   13
      Top             =   600
      Width           =   12800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   495
      Left            =   10920
      TabIndex        =   2
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   10920
      TabIndex        =   1
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Refresh"
      Height          =   495
      Left            =   10920
      TabIndex        =   0
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Line Coding Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   12800
      Begin VB.OptionButton Option1 
         Caption         =   "Unipolar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Non Return to Zero Level (NRZ-L)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   11
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Manchester"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   2280
         Width           =   3015
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Bipolar Alternate Mark Invertion (AMI)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   840
         Width           =   3015
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Bipolar Multiline Transmission, three level (MLT-3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   1800
         Width           =   4095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Return to Zero (RZ)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1800
         Width           =   3015
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Non Return to Zero Invert (NRZ-I)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1320
         Width           =   3015
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Bipolar Two Binary, One Quaternary (2B1Q)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   1320
         Width           =   3495
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Differential Manchester"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   4
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Label Label8 
      Caption         =   "0 or 1 input only"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   85
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "arthurobot@yahoo.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   84
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Philippines"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   83
      Top             =   3480
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   82
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Number of Bits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   81
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   80
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   315
      TabIndex        =   79
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   525
      TabIndex        =   78
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   77
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   915
      TabIndex        =   76
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1125
      TabIndex        =   75
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   1320
      TabIndex        =   74
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   1515
      TabIndex        =   73
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   1725
      TabIndex        =   72
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   1920
      TabIndex        =   71
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   2115
      TabIndex        =   70
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   2325
      TabIndex        =   69
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2520
      TabIndex        =   68
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   2715
      TabIndex        =   67
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   2925
      TabIndex        =   66
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   3120
      TabIndex        =   65
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   3315
      TabIndex        =   64
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   3525
      TabIndex        =   63
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   3720
      TabIndex        =   62
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   3915
      TabIndex        =   61
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   4125
      TabIndex        =   60
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   4320
      TabIndex        =   59
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   4515
      TabIndex        =   58
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   4725
      TabIndex        =   57
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   4920
      TabIndex        =   56
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   5115
      TabIndex        =   55
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   5325
      TabIndex        =   54
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   5520
      TabIndex        =   53
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   5715
      TabIndex        =   52
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   29
      Left            =   5925
      TabIndex        =   51
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   30
      Left            =   6120
      TabIndex        =   50
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   31
      Left            =   6315
      TabIndex        =   49
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   32
      Left            =   6525
      TabIndex        =   48
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   33
      Left            =   6720
      TabIndex        =   47
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   34
      Left            =   6915
      TabIndex        =   46
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   35
      Left            =   7125
      TabIndex        =   45
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   36
      Left            =   7320
      TabIndex        =   44
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   37
      Left            =   7515
      TabIndex        =   43
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   38
      Left            =   7725
      TabIndex        =   42
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   39
      Left            =   7920
      TabIndex        =   41
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   40
      Left            =   8115
      TabIndex        =   40
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   41
      Left            =   8325
      TabIndex        =   39
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   42
      Left            =   8520
      TabIndex        =   38
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   43
      Left            =   8715
      TabIndex        =   37
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   44
      Left            =   8925
      TabIndex        =   36
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   45
      Left            =   9120
      TabIndex        =   35
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   46
      Left            =   9315
      TabIndex        =   34
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   47
      Left            =   9525
      TabIndex        =   33
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   48
      Left            =   9720
      TabIndex        =   32
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   49
      Left            =   9915
      TabIndex        =   31
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   50
      Left            =   10125
      TabIndex        =   30
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   51
      Left            =   10320
      TabIndex        =   29
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   52
      Left            =   10515
      TabIndex        =   28
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   53
      Left            =   10725
      TabIndex        =   27
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   54
      Left            =   10920
      TabIndex        =   26
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   55
      Left            =   11115
      TabIndex        =   25
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   56
      Left            =   11325
      TabIndex        =   24
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   57
      Left            =   11520
      TabIndex        =   23
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   58
      Left            =   11715
      TabIndex        =   22
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   59
      Left            =   11925
      TabIndex        =   21
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   60
      Left            =   12120
      TabIndex        =   20
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   61
      Left            =   12315
      TabIndex        =   19
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   62
      Left            =   12525
      TabIndex        =   18
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   63
      Left            =   12720
      TabIndex        =   17
      Top             =   360
      Width           =   195
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Program by: Arthur S. Agustin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   16
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Angeles University Foundation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      TabIndex        =   15
      Top             =   3240
      Width           =   3015
   End
End
Attribute VB_Name = "LineCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bit As String
Dim NextLine As Integer

Private Sub Command1_Click()
Text1 = ""
GenerateLineCode
Guidelines
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Guidelines
GenerateLineCode
End Sub

Private Sub Form_Activate()
Guidelines
End Sub

Private Sub Form_Load()
For cnt = 0 To 63
Label3(cnt).Caption = ""
Next cnt
End Sub

Private Sub Option1_Click()
GenerateLineCode
End Sub

Private Sub Option2_Click()
GenerateLineCode
End Sub

Private Sub Option3_Click()
GenerateLineCode
End Sub

Private Sub Option4_Click()
GenerateLineCode
End Sub

Private Sub Option5_Click()
GenerateLineCode
End Sub

Private Sub Option6_Click()
GenerateLineCode
End Sub

Private Sub Option7_Click()
GenerateLineCode
End Sub

Private Sub Option8_Click()
GenerateLineCode
End Sub

Private Sub Option9_Click()
GenerateLineCode
End Sub

Private Sub Text1_Change()
Label1.Caption = Len(Text1)
GenerateLineCode
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 48 Or KeyAscii = 49 Or KeyAscii = 8 Then
Else
    KeyAscii = 0
End If
End Sub


'*****************************************************Display Functions*****************************************

Function GenerateLineCode()
If Option1.Value = True Then
Unipolar
ElseIf Option2.Value = True Then
NRZL
ElseIf Option3.Value = True Then
NRZI
ElseIf Option4.Value = True Then
RZ
ElseIf Option5.Value = True Then
Manchester
ElseIf Option6.Value = True Then
DifferentialManchester
ElseIf Option7.Value = True Then
BipolarAMI
ElseIf Option8.Value = True Then
Bipolar2B1Q
ElseIf Option9.Value = True Then
BipolarMLT3
End If
End Function


Function Guidelines()
Picture1.Line (0, 1000)-(13155, 1000), &HC0&
For cnt1 = 0 To 13155
    If cnt1 Mod 200 = 0 Then
        Picture1.Line (cnt1, 0)-(cnt1, 2000), &HC0&
    End If
Next cnt1
For cnt = 0 To 63
    Label3(cnt).Caption = ""
Next cnt
For cnt = 0 To Len(Text1) - 1
    Label3(cnt).Caption = Mid(Text1, cnt + 1, 1)
Next cnt
Text1.SetFocus
Text1.SelStart = 64
End Function


'*****************************************************Line Coding Functions*************************************

Function Unipolar()
Picture1.Cls
Guidelines
NextLine = 0
For cnt = 1 To Len(Text1)
    Bit = Mid(Text1, cnt, 1)
    If Bit = "1" Then
        Picture1.Line (0 + NextLine, 500)-(200 + NextLine, 500), vbYellow
    ElseIf Bit = "0" Then
        Picture1.Line (0 + NextLine, 1000)-(200 + NextLine, 1000), vbYellow
    End If
    NextLine = NextLine + 200
    
    If cnt = Len(Text1) Then
    Else
        If Mid(Text1, cnt, 1) <> Mid(Text1, cnt + 1, 1) Then
            Picture1.Line (NextLine, 500)-(NextLine, 1000), vbYellow
        End If
    End If
Next cnt
End Function

Function NRZL()
Picture1.Cls
Guidelines
NextLine = 0
For cnt = 1 To Len(Text1)
    Bit = Mid(Text1, cnt, 1)
    If Bit = "0" Then
        Picture1.Line (0 + NextLine, 1500)-(200 + NextLine, 1500), vbYellow
    ElseIf Bit = "1" Then
        Picture1.Line (0 + NextLine, 500)-(200 + NextLine, 500), vbYellow
    End If
    NextLine = NextLine + 200
    
    If cnt = Len(Text1) Then
    Else
        If Mid(Text1, cnt, 1) <> Mid(Text1, cnt + 1, 1) Then
            Picture1.Line (NextLine, 500)-(NextLine, 1500), vbYellow
        End If
    End If
Next cnt
End Function

Function NRZI()
Dim PositionBit As Integer
Picture1.Cls
Guidelines
NextLine = 0
If Len(Text1) > 0 Then
    If Mid(Text1, 1, 1) = "0" Then
        PositionBit = 500
    Else
        PositionBit = 1500
    End If
End If
For cnt = 1 To Len(Text1)
    Bit = Mid(Text1, cnt, 1)
    If Bit = "0" Then
        Picture1.Line (0 + NextLine, PositionBit)-(200 + NextLine, PositionBit), vbYellow
    ElseIf Bit = "1" Then
        If PositionBit = 500 Then
            PositionBit = 1500
        ElseIf PositionBit = 1500 Then
            PositionBit = 500
        End If
        Picture1.Line (0 + NextLine, PositionBit)-(200 + NextLine, PositionBit), vbYellow
        Picture1.Line (NextLine, 500)-(NextLine, 1500), vbYellow
    End If
    NextLine = NextLine + 200
    
    If cnt = Len(Text1) Then
    Else
    
    End If
Next cnt
End Function

Function RZ()
Picture1.Cls
Guidelines
NextLine = 0
For cnt = 1 To Len(Text1)
    Bit = Mid(Text1, cnt, 1)
    If Bit = "0" Then
        Picture1.Line (0 + NextLine, 1500)-(100 + NextLine, 1500), vbYellow
        Picture1.Line (100 + NextLine, 1000)-(200 + NextLine, 1000), vbYellow
        Picture1.Line (NextLine + 100, 1000)-(NextLine + 100, 1500), vbYellow
        Picture1.Line (NextLine, 1000)-(NextLine, 1500), vbYellow
    ElseIf Bit = "1" Then
        Picture1.Line (0 + NextLine, 500)-(100 + NextLine, 500), vbYellow
        Picture1.Line (100 + NextLine, 1000)-(200 + NextLine, 1000), vbYellow
        Picture1.Line (NextLine + 100, 500)-(NextLine + 100, 1000), vbYellow
        Picture1.Line (NextLine, 500)-(NextLine, 1000), vbYellow
    End If
    NextLine = NextLine + 200
Next cnt
End Function

Function Manchester()
Picture1.Cls
Guidelines
NextLine = 0
For cnt = 1 To Len(Text1)
    Bit = Mid(Text1, cnt, 1)
    If Bit = "1" Then
        Picture1.Line (0 + NextLine, 1500)-(100 + NextLine, 1500), vbYellow
        Picture1.Line (100 + NextLine, 500)-(200 + NextLine, 500), vbYellow
        Picture1.Line (NextLine + 100, 500)-(NextLine + 100, 1500), vbYellow
    ElseIf Bit = "0" Then
        Picture1.Line (0 + NextLine, 500)-(100 + NextLine, 500), vbYellow
        Picture1.Line (100 + NextLine, 1500)-(200 + NextLine, 1500), vbYellow
        Picture1.Line (NextLine + 100, 500)-(NextLine + 100, 1500), vbYellow
    End If
    NextLine = NextLine + 200
    If cnt = Len(Text1) Then
    Else
        If Mid(Text1, cnt, 1) = Mid(Text1, cnt + 1, 1) Then
            Picture1.Line (NextLine, 500)-(NextLine, 1500), vbYellow
        End If
    End If
Next cnt
End Function

Function DifferentialManchester()
Dim x1, x2 As Integer
Picture1.Cls
Guidelines
NextLine = 0
If Len(Text1) > 0 Then
    If Mid(Text1, 1, 1) = "0" Then
        Picture1.Line (0 + NextLine, 1500)-(100 + NextLine, 1500), vbYellow
        Picture1.Line (100 + NextLine, 500)-(200 + NextLine, 500), vbYellow
        Picture1.Line (NextLine + 100, 500)-(NextLine + 100, 1500), vbYellow
        x1 = 1500
        x2 = 500
    Else
        Picture1.Line (0 + NextLine, 500)-(100 + NextLine, 500), vbYellow
        Picture1.Line (100 + NextLine, 1500)-(200 + NextLine, 1500), vbYellow
        Picture1.Line (NextLine + 100, 500)-(NextLine + 100, 1500), vbYellow
        x1 = 500
        x2 = 1500
    End If
Picture1.Line (0, 500)-(0, 1500), vbYellow
End If
NextLine = NextLine + 200
For cnt = 2 To Len(Text1)
    Bit = Mid(Text1, cnt, 1)
        If Mid(Text1, cnt, 1) = "0" Then
            Picture1.Line (0 + NextLine, x1)-(100 + NextLine, x1), vbYellow
            Picture1.Line (100 + NextLine, x2)-(200 + NextLine, x2), vbYellow
            Picture1.Line (NextLine + 100, 500)-(NextLine + 100, 1500), vbYellow
            Picture1.Line (NextLine, 500)-(NextLine, 1500), vbYellow
        Else
            If x1 = 500 Then
            x1 = 1500
            x2 = 500
            Else
            x1 = 500
            x2 = 1500
            End If
            Picture1.Line (0 + NextLine, x1)-(100 + NextLine, x1), vbYellow
            Picture1.Line (100 + NextLine, x2)-(200 + NextLine, x2), vbYellow
            Picture1.Line (NextLine + 100, 500)-(NextLine + 100, 1500), vbYellow
        End If
NextLine = NextLine + 200
Next cnt
End Function

Function BipolarAMI()
Dim BitPosition As String
Picture1.Cls
Guidelines
NextLine = 0
If Len(Text1) > 0 Then
    If Mid(Text1, 1, 1) = "0" Then
        Picture1.Line (0 + NextLine, 1000)-(200 + NextLine, 1000), vbYellow
        BitPosition = "up"
    Else
        Picture1.Line (0 + NextLine, 500)-(200 + NextLine, 500), vbYellow
        Picture1.Line (0 + NextLine, 500)-(0 + NextLine, 1000), vbYellow
        Picture1.Line (200 + NextLine, 500)-(200 + NextLine, 1000), vbYellow
        BitPosition = "down"
    End If
End If
NextLine = NextLine + 200
For cnt = 2 To Len(Text1)
    Bit = Mid(Text1, cnt, 1)
        If Mid(Text1, cnt, 1) = "0" Then
            Picture1.Line (0 + NextLine, 1000)-(200 + NextLine, 1000), vbYellow
        Else
            If BitPosition = "up" Then
            BitPosition = "down"
            Picture1.Line (0 + NextLine, 500)-(200 + NextLine, 500), vbYellow
            Picture1.Line (0 + NextLine, 500)-(0 + NextLine, 1000), vbYellow
            Picture1.Line (200 + NextLine, 500)-(200 + NextLine, 1000), vbYellow
            Else
            BitPosition = "up"
            Picture1.Line (0 + NextLine, 1500)-(200 + NextLine, 1500), vbYellow
            Picture1.Line (0 + NextLine, 1000)-(0 + NextLine, 1500), vbYellow
            Picture1.Line (200 + NextLine, 1000)-(200 + NextLine, 1500), vbYellow
            End If
        End If
NextLine = NextLine + 200
Next cnt
End Function

Function Bipolar2B1Q()
Dim y1, y2 As Integer
Picture1.Cls
Guidelines
If Len(Text1) Mod 2 <> 0 Then
MsgBox "The bit number must be divisible by 2"
Exit Function
End If
NextLine = 0
y1 = 1000
y2 = 1000
For cnt = 1 To Len(Text1)
    If cnt Mod 2 = 0 Then
        If Mid(Text1, cnt - 1, 1) = "0" And Mid(Text1, cnt, 1) = "0" Then
            Picture1.Line (0 + NextLine, 1500)-(400 + NextLine, 1500), vbYellow
            y1 = y2
            y2 = 1500
        ElseIf Mid(Text1, cnt - 1, 1) = "1" And Mid(Text1, cnt, 1) = "1" Then
            Picture1.Line (0 + NextLine, 750)-(400 + NextLine, 750), vbYellow
            y1 = y2
            y2 = 750
        ElseIf Mid(Text1, cnt - 1, 1) = "0" And Mid(Text1, cnt, 1) = "1" Then
            Picture1.Line (0 + NextLine, 1250)-(400 + NextLine, 1250), vbYellow
            y1 = y2
            y2 = 1250
        ElseIf Mid(Text1, cnt - 1, 1) = "1" And Mid(Text1, cnt, 1) = "0" Then
            Picture1.Line (0 + NextLine, 500)-(400 + NextLine, 500), vbYellow
            y1 = y2
            y2 = 500
        End If
        Picture1.Line (NextLine, y1)-(NextLine, y2), vbYellow
    NextLine = NextLine + 400
    End If
Next cnt

End Function

Function BipolarMLT3()
Dim BitPosition As String
Dim x As Integer
Picture1.Cls
Guidelines
NextLine = 0
BitPosition = "up"
If Len(Text1) > 0 Then
    If Mid(Text1, 1, 1) = "0" Then
        x = 1000
        Picture1.Line (0 + NextLine, 1000)-(200 + NextLine, 1000), vbYellow
        BitPosition = "up"
    Else
        x = 500
        Picture1.Line (0 + NextLine, 500)-(200 + NextLine, 500), vbYellow
        BitPosition = "down"
    End If
End If
NextLine = NextLine + 200
For cnt = 2 To Len(Text1)
    Bit = Mid(Text1, cnt, 1)
        If Mid(Text1, cnt, 1) = "0" Then
            Picture1.Line (0 + NextLine, x)-(200 + NextLine, x), vbYellow
        Else
            If x = 500 Then
            x = 1000
            BitPosition = "down"
            Picture1.Line (0 + NextLine, x)-(200 + NextLine, x), vbYellow
            Picture1.Line (0 + NextLine, 500)-(0 + NextLine, 1000), vbYellow
            ElseIf x = 1500 Then
            x = 1000
            BitPosition = "up"
            Picture1.Line (0 + NextLine, x)-(200 + NextLine, x), vbYellow
            Picture1.Line (0 + NextLine, 1000)-(0 + NextLine, 1500), vbYellow
            ElseIf x = 1000 And BitPosition = "down" Then
            x = 1500
            Picture1.Line (0 + NextLine, x)-(200 + NextLine, x), vbYellow
            Picture1.Line (0 + NextLine, 1000)-(0 + NextLine, 1500), vbYellow
            ElseIf x = 1000 And BitPosition = "up" Then
            x = 500
            Picture1.Line (0 + NextLine, x)-(200 + NextLine, x), vbYellow
            Picture1.Line (0 + NextLine, 500)-(0 + NextLine, 1000), vbYellow
            End If
        End If
NextLine = NextLine + 200
Next cnt
End Function


