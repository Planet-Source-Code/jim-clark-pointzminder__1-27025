VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PointzMinder"
   ClientHeight    =   5235
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4335
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab PtzPage 
      Height          =   5235
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   9234
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "&Minder"
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblRange(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblRange(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "UpDownCalories"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtCalories"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraPoints(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraFatGroup"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraFiberGroup"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "&Buster"
      TabPicture(1)   =   "frmMain.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label5"
      Tab(1).Control(2)=   "lblRange(4)"
      Tab(1).Control(3)=   "lblRange(5)"
      Tab(1).Control(4)=   "lblRange(6)"
      Tab(1).Control(5)=   "lblRange(7)"
      Tab(1).Control(6)=   "UpDownMinutes"
      Tab(1).Control(7)=   "fraPoints(1)"
      Tab(1).Control(8)=   "fraIntensityGroup"
      Tab(1).Control(9)=   "txtMinutes"
      Tab(1).Control(10)=   "txtWeight"
      Tab(1).Control(11)=   "UpDownWeight"
      Tab(1).ControlCount=   12
      Begin VB.Frame fraFiberGroup 
         Caption         =   "&Fiber Gramz"
         Height          =   1455
         Left            =   660
         TabIndex        =   0
         Top             =   1140
         Width           =   1335
         Begin VB.Frame fraFiber 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1095
            Left            =   300
            TabIndex        =   1
            ToolTipText     =   "Set Fiber Gramz"
            Top             =   240
            Width           =   675
            Begin VB.Frame fraFiberMask 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   990
               Left            =   225
               TabIndex        =   2
               ToolTipText     =   "Set Fiber Gramz"
               Top             =   0
               Width           =   435
               Begin MSComctlLib.Slider sldFiber 
                  Height          =   1170
                  Left            =   -60
                  TabIndex        =   3
                  ToolTipText     =   "Set Fiber Gramz"
                  Top             =   -60
                  Width           =   630
                  _ExtentX        =   1111
                  _ExtentY        =   2064
                  _Version        =   393216
                  Orientation     =   1
                  LargeChange     =   4
                  Min             =   1
                  Max             =   9
                  SelStart        =   1
                  TickStyle       =   1
                  TickFrequency   =   2
                  Value           =   1
               End
            End
            Begin VB.Label lblFiber 
               Alignment       =   1  'Right Justify
               Caption         =   "4"
               Height          =   165
               Index           =   4
               Left            =   -45
               TabIndex        =   72
               Top             =   810
               Width           =   180
            End
            Begin VB.Label lblFiber 
               Alignment       =   1  'Right Justify
               Caption         =   "3"
               Height          =   165
               Index           =   3
               Left            =   -45
               TabIndex        =   71
               Top             =   615
               Width           =   180
            End
            Begin VB.Label lblFiber 
               Alignment       =   1  'Right Justify
               Caption         =   "2"
               Height          =   165
               Index           =   2
               Left            =   -45
               TabIndex        =   70
               Top             =   420
               Width           =   180
            End
            Begin VB.Label lblFiber 
               Alignment       =   1  'Right Justify
               Caption         =   "1"
               Height          =   165
               Index           =   1
               Left            =   -45
               TabIndex        =   69
               Top             =   225
               Width           =   180
            End
            Begin VB.Label lblFiber 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               Height          =   165
               Index           =   0
               Left            =   -45
               TabIndex        =   68
               Top             =   30
               Width           =   180
            End
         End
      End
      Begin VB.Frame fraFatGroup 
         Caption         =   "F&at Gramz"
         Height          =   4575
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   1335
         Begin VB.Frame fraFat 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4275
            Left            =   240
            TabIndex        =   8
            ToolTipText     =   "Set Fat Gramz"
            Top             =   240
            Width           =   765
            Begin VB.Frame fraFatMask 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   4170
               Left            =   315
               TabIndex        =   7
               ToolTipText     =   "Set Fat Gramz"
               Top             =   0
               Width           =   435
               Begin MSComctlLib.Slider sldFat 
                  Height          =   4305
                  Left            =   -60
                  TabIndex        =   9
                  ToolTipText     =   "Set Fat Gramz"
                  Top             =   -60
                  Width           =   630
                  _ExtentX        =   1111
                  _ExtentY        =   7594
                  _Version        =   393216
                  Orientation     =   1
                  LargeChange     =   4
                  Min             =   1
                  Max             =   41
                  SelStart        =   1
                  TickStyle       =   1
                  TickFrequency   =   2
                  Value           =   1
               End
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "20"
               Height          =   165
               Index           =   20
               Left            =   0
               TabIndex        =   67
               Top             =   3930
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "19"
               Height          =   165
               Index           =   19
               Left            =   0
               TabIndex        =   66
               Top             =   3735
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "18"
               Height          =   165
               Index           =   18
               Left            =   0
               TabIndex        =   65
               Top             =   3540
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "17"
               Height          =   165
               Index           =   17
               Left            =   0
               TabIndex        =   64
               Top             =   3345
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "16"
               Height          =   165
               Index           =   16
               Left            =   0
               TabIndex        =   63
               Top             =   3150
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "15"
               Height          =   165
               Index           =   15
               Left            =   0
               TabIndex        =   62
               Top             =   2955
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "14"
               Height          =   165
               Index           =   14
               Left            =   0
               TabIndex        =   61
               Top             =   2760
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "13"
               Height          =   165
               Index           =   13
               Left            =   0
               TabIndex        =   60
               Top             =   2565
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "12"
               Height          =   165
               Index           =   12
               Left            =   0
               TabIndex        =   59
               Top             =   2370
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "11"
               Height          =   165
               Index           =   11
               Left            =   0
               TabIndex        =   58
               Top             =   2175
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "10"
               Height          =   165
               Index           =   10
               Left            =   0
               TabIndex        =   57
               Top             =   1980
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "9"
               Height          =   165
               Index           =   9
               Left            =   0
               TabIndex        =   56
               Top             =   1785
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "8"
               Height          =   165
               Index           =   8
               Left            =   0
               TabIndex        =   55
               Top             =   1590
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "7"
               Height          =   165
               Index           =   7
               Left            =   0
               TabIndex        =   54
               Top             =   1395
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "6"
               Height          =   165
               Index           =   6
               Left            =   0
               TabIndex        =   53
               Top             =   1200
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "5"
               Height          =   165
               Index           =   5
               Left            =   0
               TabIndex        =   52
               Top             =   1005
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "4"
               Height          =   165
               Index           =   4
               Left            =   0
               TabIndex        =   51
               Top             =   810
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "3"
               Height          =   165
               Index           =   3
               Left            =   0
               TabIndex        =   50
               Top             =   615
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "2"
               Height          =   165
               Index           =   2
               Left            =   0
               TabIndex        =   49
               Top             =   420
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "1"
               Height          =   165
               Index           =   1
               Left            =   0
               TabIndex        =   48
               Top             =   225
               Width           =   225
            End
            Begin VB.Label lblFat 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               Height          =   165
               Index           =   0
               Left            =   0
               TabIndex        =   47
               Top             =   30
               Width           =   225
            End
         End
      End
      Begin MSComCtl2.UpDown UpDownWeight 
         Height          =   315
         Left            =   -73320
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Use these up/down buttons to increase/decrease Entered Weight"
         Top             =   2580
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtWeight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74220
         TabIndex        =   13
         ToolTipText     =   "Type in Your Weight"
         Top             =   2580
         Width           =   915
      End
      Begin VB.TextBox txtMinutes 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   -74220
         TabIndex        =   11
         ToolTipText     =   "Type in Minutez of Activity"
         Top             =   1500
         Width           =   915
      End
      Begin VB.Frame fraIntensityGroup 
         Caption         =   "&Intensity Level"
         Height          =   3195
         Left            =   -72480
         TabIndex        =   14
         Top             =   960
         Width           =   1695
         Begin VB.Frame fraIntMain 
            BorderStyle     =   0  'None
            Height          =   2715
            Left            =   180
            TabIndex        =   15
            Top             =   300
            Width           =   1395
            Begin VB.Frame fraIntensity 
               BorderStyle     =   0  'None
               Height          =   2760
               Left            =   240
               TabIndex        =   16
               Top             =   -15
               Width           =   450
               Begin MSComctlLib.Slider sldIntensity 
                  Height          =   2790
                  Left            =   -15
                  TabIndex        =   27
                  ToolTipText     =   "Adjust to reflect your activity level"
                  Top             =   -15
                  Width           =   480
                  _ExtentX        =   847
                  _ExtentY        =   4921
                  _Version        =   393216
                  Orientation     =   1
                  Min             =   1
                  Max             =   21
                  SelStart        =   1
                  TickStyle       =   1
                  TickFrequency   =   2
                  Value           =   1
               End
            End
            Begin VB.Label lblActDesc 
               Caption         =   "High"
               Height          =   195
               Index           =   2
               Left            =   720
               TabIndex        =   41
               ToolTipText     =   "Definite sweat, jogging, running, competitive swimming/biking."
               Top             =   2460
               Width           =   675
            End
            Begin VB.Label lblActDesc 
               Caption         =   "Moderate"
               Height          =   195
               Index           =   1
               Left            =   720
               TabIndex        =   40
               ToolTipText     =   "Probable sweat, walking at a fast pace or biking."
               Top             =   1260
               Width           =   675
            End
            Begin VB.Label lblActDesc 
               Caption         =   "&Light"
               Height          =   195
               Index           =   0
               Left            =   720
               TabIndex        =   39
               ToolTipText     =   "No sweat, stretching or walking at a leisurely pace."
               Top             =   60
               Width           =   675
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "0"
               Height          =   195
               Index           =   0
               Left            =   0
               TabIndex        =   38
               Top             =   60
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "1"
               Height          =   195
               Index           =   1
               Left            =   0
               TabIndex        =   37
               Top             =   300
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "2"
               Height          =   195
               Index           =   2
               Left            =   0
               TabIndex        =   36
               Top             =   540
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "3"
               Height          =   195
               Index           =   3
               Left            =   0
               TabIndex        =   35
               Top             =   780
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "4"
               Height          =   195
               Index           =   4
               Left            =   0
               TabIndex        =   34
               Top             =   1020
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "5"
               Height          =   195
               Index           =   5
               Left            =   0
               TabIndex        =   33
               Top             =   1260
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "6"
               Height          =   195
               Index           =   6
               Left            =   0
               TabIndex        =   32
               Top             =   1500
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "7"
               Height          =   195
               Index           =   7
               Left            =   0
               TabIndex        =   31
               Top             =   1740
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "8"
               Height          =   195
               Index           =   8
               Left            =   0
               TabIndex        =   30
               Top             =   1980
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "9"
               Height          =   195
               Index           =   9
               Left            =   0
               TabIndex        =   29
               Top             =   2220
               Width           =   195
            End
            Begin VB.Label lblAct 
               Alignment       =   1  'Right Justify
               Caption         =   "10"
               Height          =   195
               Index           =   10
               Left            =   0
               TabIndex        =   28
               Top             =   2460
               Width           =   195
            End
         End
      End
      Begin VB.Frame fraPoints 
         Caption         =   "Pointz:"
         Height          =   915
         Index           =   1
         Left            =   -74580
         TabIndex        =   23
         ToolTipText     =   "Right Click for Rounding Options"
         Top             =   3960
         Width           =   1875
         Begin VB.Label lblPoints 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   795
            Index           =   1
            Left            =   60
            TabIndex        =   24
            ToolTipText     =   "Right Click for Rounding Options"
            Top             =   120
            Width           =   1755
         End
      End
      Begin VB.Frame fraPoints 
         Caption         =   "Pointz:"
         Height          =   915
         Index           =   0
         Left            =   420
         TabIndex        =   18
         ToolTipText     =   "Right Click for Rounding Options"
         Top             =   3960
         Width           =   1875
         Begin VB.Label lblPoints 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   795
            Index           =   0
            Left            =   60
            TabIndex        =   19
            ToolTipText     =   "Right Click for Rounding Options"
            Top             =   120
            Width           =   1755
         End
      End
      Begin VB.TextBox txtCalories 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   780
         TabIndex        =   5
         ToolTipText     =   "Type in caloriez"
         Top             =   3060
         Width           =   915
      End
      Begin MSComCtl2.UpDown UpDownMinutes 
         Height          =   315
         Left            =   -73320
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Use these up/down buttons to increase/decrease Activity Minutes"
         Top             =   1500
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDownCalories 
         Height          =   315
         Left            =   1680
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Use these up/down buttons to increase/decrease Caloriez"
         Top             =   3060
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "&Weight:"
         Height          =   195
         Index           =   7
         Left            =   -74280
         TabIndex        =   12
         ToolTipText     =   "Set Caloriez"
         Top             =   2340
         Width           =   675
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "(min-max)"
         Height          =   195
         Index           =   6
         Left            =   -74100
         TabIndex        =   43
         ToolTipText     =   "Caloriez per serving must not exceed 550"
         Top             =   2940
         Width           =   795
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "Minu&tez:"
         Height          =   195
         Index           =   5
         Left            =   -74280
         TabIndex        =   10
         ToolTipText     =   "Set Caloriez"
         Top             =   1260
         Width           =   675
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "(min-max)"
         Height          =   195
         Index           =   4
         Left            =   -74100
         TabIndex        =   42
         ToolTipText     =   "Caloriez per serving must not exceed 550"
         Top             =   1860
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "PointzBuster"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -74580
         TabIndex        =   26
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Calculates Activity Pointz"
         Height          =   255
         Left            =   -74580
         TabIndex        =   25
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "PointzMinder"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   420
         TabIndex        =   22
         Top             =   480
         Width           =   915
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "(min-max)"
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   20
         ToolTipText     =   "Caloriez per serving must not exceed 550"
         Top             =   3420
         Width           =   795
      End
      Begin VB.Label lblRange 
         Alignment       =   1  'Right Justify
         Caption         =   "&Caloriez:"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   4
         ToolTipText     =   "Set Caloriez"
         Top             =   2820
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Calculates Food Pointz"
         Height          =   255
         Left            =   420
         TabIndex        =   21
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Menu mnuRounding 
      Caption         =   "&Rounding"
      Visible         =   0   'False
      Begin VB.Menu mnuWhole 
         Caption         =   "&Whole Numbers Only"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOneDec 
         Caption         =   "&One Decimal Place"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTwoDec 
         Caption         =   "&Two Decimal Places"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'  This program written by WeightWizard
'  for educational purposes only.
'
'  The formulas and program are accurate & complete,
'  but are also patented by Weight Watchers(R)
'  You have my permission to use any part of this
'  code for whatever purposes you desire, but I do
'  not condone violating Weight Watchers(R)
'  copyright and trademarks. If you compile and
'  publish this program, Weight Watchers(R) will
'  likely hunt you down and threaten legal action
'  aginst you.
'

'Constant declaration and assignment
Private Const App_Title As String = "PointzMinder"
Private Const WEIGHT_MAX As Integer = 350
Private Const WEIGHT_MIN As Integer = 100
Private Const CALORIES_MAX As Integer = 550
Private Const CALORIES_MIN As Integer = 10
Private Const MINUTES_MAX As Integer = 120
Private Const MINUTES_MIN As Integer = 10
Private Const ACTIVITY_POINTZ_MAX As Integer = 13

'Variable declaration (integers)
Private mDecPlaces As Integer
Private lastSelStart As Integer
Private mWeight As Integer
Private mMinutes As Integer
Private mCalories As Integer

'Variable declaration (doubles)
Private mFiber As Double
Private mFat As Double
Private mIntensity As Double

'Variable declaration (strings)
Private lastValueStr As String
Private Const regKey As String = "SOFTWARE\Pointz"
Private errorString As String

'Class declaration
Private cReg As cRegistry
'

Private Sub Form_Load()
   'This sub executes once during initial load of program
   '
   'Create new instance of class
   Set cReg = New cRegistry
   
   'Attempt to retrieve values from registry
   'If no values stored in registry then use default values
   If Not getRegVars Then setRegFailedVals
   
   'Set Calorie, Minutes, and Weight captions (min - max)
   setCaptions
   
   'Set Sliders and User Text Box values
   setControlValues
   
   'Manually fire change events for all user
   'controls to force calcuation of points answer
   callChangeEventForAllDataControls
   
   'Set the checkmarks on the decimal place popup menu
   'And then calculate pointz
   setRounding
  
End Sub

Private Sub Form_Activate()
   'This sub is executed every time the form gets focus,
   'including initially when the program is launched, but
   'only after the form load event fires.
   '
   'Using this static variable, we can allow this code
   'to only run one time after "form load"
   Static bActivated As Boolean
   
   If Not bActivated Then
      bActivated = True
      
      'Change Tabs and then switch to First Tab (Tab 0)
      'This will force tab change event to fire
      PtzPage_Click Abs(PtzPage.Tab - 1)
      PtzPage.Tab = 0
      
   End If
End Sub

Private Sub setControlValues()
   'The values that were loaded into variables retrieved from
   'the registry need to be loaded into the controls
   
   'First load the food pointz tab
   '
   'Calories
   txtCalories.Text = mCalories
   '
   'Fat Grams
   sldFat_Move (mFat - 1) / 2
   '
   'Fiber Grams
   sldFiber_Move (mFiber - 1) / 2
   
   
   'Then load the activity pointz tab
   '
   'Minutes
   txtMinutes.Text = mMinutes
   '
   'Weight (in U.S. pounds)
   txtWeight.Text = mWeight
   '
   'Intensity
   sldIntensity_Move (mIntensity - 1) / 2
   
End Sub

Private Sub callChangeEventForAllDataControls()
   'Manually call the events that cause the
   'values of the user input controls to change
   
   lastValueStr = "0"
   lastSelStart = 0
   
   txtMinutes_Change
   txtWeight_Change
   txtCalories_Change
   
   sldIntensity_Scroll
   sldFat_Scroll
   sldFiber_Scroll
   
End Sub

Private Sub setRegFailedVals()
   'If unable to retrieve values from the registry
   'then use these default values
   
   mDecPlaces = 2
   mCalories = 10
   mFat = 0
   mFiber = 0
   mWeight = 100
   mIntensity = 0
   mMinutes = 10
End Sub

Private Sub setCaptions()
   'Set Calorie, Minutes, and Weight captions (min - max)
   
   lblRange(0) = "(" & CALORIES_MIN & "-" & CALORIES_MAX & ")"
   lblRange(4) = "(" & MINUTES_MIN & "-" & MINUTES_MAX & ")"
   lblRange(6) = "(" & WEIGHT_MIN & "-" & WEIGHT_MAX & ")"
   Me.Caption = App_Title & " Ver " & App.Major & "." & App.Minor
   
End Sub

Private Sub fraPoints_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   'There are two frames that display pointz,
   'one for food pointz, and one for activity pointz - fraPoints(0) & fraPoints(1)
   'This event fires if the user clicks the mouse on either of these frames
   '
   'If the right mouse button was clicked on the frame,
   'then display the rounding menu
   
   If Button = vbRightButton Then Me.PopupMenu mnuRounding
   
End Sub

Private Sub lblActDesc_Click(Index As Integer)
   'There are three labels that describe activity levels
   'Light, Moderate, & High named lblActDesc(0) through lblActDesc(2)
   'If the user clicks one of these then move the slider to that position
   '
   Select Case Index
      Case 0 'Light was clicked
         sldIntensity_Move 1
      Case 1 'Moderate was clicked
         sldIntensity_Move 11
      Case 2 'High was clicked
         sldIntensity_Move 21
   End Select
   
End Sub

Private Sub PtzPage_Click(PreviousTab As Integer)
   'This is the click event of the Microsoft Tabbed Dialog Control 6.0
   'This event fires when a new tab is clicked on,
   'or programatically changed
   '
   '
   Select Case PtzPage.Tab
      Case 0 'Food Pointz
         highLightTextBox txtCalories
      Case 1 'Activity Pointz
         highLightTextBox txtMinutes
   End Select
   
   'Calling enableTabStops will enable keyboard tabbing between the
   'controls of the Tab that was selected and disable tabbing for the
   'controls of the Tab that is not visible.
   'This is necessary whenever frames are used on the MS Tabbed Dialog Control
   'due to a bug in the control.  If this is not done, then tabbing can move
   'the focus to controls not currently visible
   enableTabStops PtzPage.Tab
   
End Sub

Private Sub highLightTextBox(mTextBox As VB.TextBox)
   'This function is passed a text box by reference
   'and the focus is set and the text in the textbox highlighted
   
   mTextBox.SetFocus
   mTextBox.SelStart = 0
   mTextBox.SelLength = 255
End Sub

Private Sub enableTabStops(mTab As Integer)
   'Calling enableTabStops will enable keyboard tabbing between the
   'controls of the Tab that was selected and disable tabbing for the
   'controls of the Tab that is not visible.
   'This is necessary whenever frames are used on the MS Tabbed Dialog Control
   'due to a bug in the control.  If this is not done, then tabbing can move
   'the focus to controls not currently visible
   '
   Dim bControlState As Boolean
   
   'If the current tab is 0 then bControlState = False
   'otherwise is will be true
   bControlState = CBool(mTab)
   
   'Enable or disable the controls on the Activity Pointz Tab
   txtMinutes.TabStop = bControlState
   txtWeight.TabStop = bControlState
   sldIntensity.TabStop = bControlState
   
   'Toggle the value of bControlState for the Food Pointz Tab
   bControlState = Not bControlState
   
   'Enable or disable the controls on the Food Pointz Tab
   sldFiber.TabStop = bControlState
   sldFat.TabStop = bControlState
   txtCalories.TabStop = bControlState
   
End Sub

Private Sub sldIntensity_Scroll()
   'This event fires right after these slider events:
   'MouseDown
   'MouseUp
   'MouseDrag
   'MouseWheelSpin
   'ArrowKey Press
   'ArrowKey HeldDown (fires repeatedly)
   '
   'It does not fire from the sldIntensity_Change event
   
   Dim mTmp As Double
   mTmp = (sldIntensity.Value - 1) / 2
   sldIntensity.Text = mTmp
   mIntensity = mTmp / 10
   calcPointsBuster
   
End Sub

Private Sub sldIntensity_Move(newPosition As Integer)
   'This sub updates the position of the slider,
   'and then calls the scroll event in order to update
   'the pointz answer
   
   sldIntensity.Value = newPosition
   sldIntensity_Scroll
   
End Sub

Private Sub sldFat_Scroll()
   'This event fires right after these slider events:
   'MouseDown
   'MouseUp
   'MouseDrag
   'MouseWheelSpin
   'ArrowKey Press
   'ArrowKey HeldDown (fires repeatedly)
   '
   'It does not fire from the sldFat_Change event
   
   mFat = (sldFat - 1) / 2
   sldFat.Text = mFat
   calcPointsMinder
   
End Sub

Private Sub sldFat_Move(newPosition As Integer)
   'This sub updates the position of the slider,
   'and then calls the scroll event in order to update
   'the pointz answer
   
   sldFat.Value = newPosition
   sldFat_Scroll
   
End Sub


Private Sub sldFiber_Scroll()
   'This event fires right after these slider events:
   'MouseDown
   'MouseUp
   'MouseDrag
   'MouseWheelSpin
   'ArrowKey Press
   'ArrowKey HeldDown (fires repeatedly)
   '
   'It does not fire from the sldFiber_Change event
   
   mFiber = (sldFiber - 1) / 2
   sldFiber.Text = mFiber
   calcPointsMinder
   
End Sub

Private Sub sldFiber_Move(newPosition As Integer)
   'This sub updates the position of the slider,
   'and then calls the scroll event in order to update
   'the pointz answer
   
   sldFiber.Value = newPosition
   sldFiber_Scroll
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Called when program is terminated
   '
   'Save the last values the user entered in the registry,
   'so they can be re-loaded when the program is started next time
   saveRegVars
   
   'Destroy instance of the cReg class to free up memory
   Set cReg = Nothing
   
End Sub

Private Sub lblFat_Click(Index As Integer)
   'User has clicked on one of the labels (0 - 20)
   'on the Fat Slider so the slider needs to be
   'moved to that position
   
   sldFat_Move Index * 2 + 1
   
End Sub

Private Sub lblFiber_Click(Index As Integer)
   'User has clicked on one of the labels (0 - 4)
   'on the Fiber Slider so the slider needs to be
   'moved to that position
   
   sldFiber_Move Index * 2 + 1
   
End Sub

Private Sub lblAct_Click(Index As Integer)
   'User has clicked on one of the labels (Low, Moderate, or High)
   'on the Intensity Slider so the slider needs to be
   'moved to that position

   sldIntensity_Move Index * 2 + 1
   
End Sub

Private Sub lblPoints_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   'There are two labels that display pointz,
   'one for food pointz, and one for activity pointz - lblPoints(0) & lblPoints(1)
   'This event fires if the user clicks the mouse on either of these labels
   '
   'If the right mouse button was clicked on the label,
   'then display the rounding menu
   If Button = vbRightButton Then Me.PopupMenu mnuRounding
   
End Sub

Private Sub saveRegDecimalPlaces(newDecimalPlaces As Integer)
   On Error Resume Next
   
   'Save the current Rounding (0, 1, or 2) option to the registry
   'so that it will be remembered the next time the program is executed
   cReg.SetStringValue "DecimalPlaces", CStr(newDecimalPlaces)
   
End Sub

Private Sub mnuWhole_Click()
   'User clicked "Whole Numbers Only" on the popup menu
   '
   'Assign 0 to the mDecPlaces variable so it can be accessed before an answer
   'is displayed on the form
   mDecPlaces = 0
   
   'Save the 0 to the registry
   saveRegDecimalPlaces 0
   
   'Set the checkmarks on the decimal place popup menu
   'And then recalculate pointz
   setRounding
   
End Sub

Private Sub mnuOneDec_Click()
   'User clicked "One Decimal Place" on the popup menu
   '
   'Assign 1 to the mDecPlaces variable so it can be accessed before an answer
   'is displayed on the form
   mDecPlaces = 1
   
   'Save the one to the registry
   saveRegDecimalPlaces 1
   
   'Set the checkmarks on the decimal place popup menu
   'And then recalculate pointz
   setRounding
   
End Sub

Private Sub mnuTwoDec_Click()
   'User clicked "Two Decimal Places" on the popup menu
   '
   'Assign 2 to the mDecPlaces variable so it can be accessed before an answer
   'is displayed on the form
   mDecPlaces = 2
   
   'Save the 2 to the registry
   saveRegDecimalPlaces 2
   
   'Set the checkmarks on the decimal place popup menu
   'And then recalculate pointz
   setRounding
   
End Sub

Private Sub setRounding()
   'Set the checkmarks on the decimal place popup menu
   'mnuX.Checked are boolans, if true then the check mark appears on the menu
   '(mDecPlaces = n) will evaluate to a boolean true or false value

   mnuWhole.Checked = (mDecPlaces = 0)
   mnuOneDec.Checked = (mDecPlaces = 1)
   mnuTwoDec.Checked = (mDecPlaces = 2)
   
   'Recalculate and display pointz with the current rounding
   calcPointsMinder
   calcPointsBuster
   
End Sub

Private Sub calcPointsMinder()
   '
   'This sub calculates food pointz based on the data
   'that the user has input.  It is called with the change event
   'of one of the text boxes, or the scroll event of a slider
   '
   Dim mVal As Double
   Dim PointzNumberString As String
   Dim DecimalPointPosition As Integer
   
   'Calculate Food Pointz
   mVal = (mCalories - mFiber * 10) / 50 + mFat / 12
   
   'Error check -> if Pointz are less than zero then make zero
   If mVal < 0 Then mVal = 0
   
   'Round the answer to the number of decimal places the user has specified
   PointzNumberString = Round(mVal, mDecPlaces)
   
   'Find the position of the decimal point in the Pointz String (0 if no decimal)
   DecimalPointPosition = InStr(1, PointzNumberString, ".")
   
   'Fill out blank decimal places with zeros
   'In other words, if user specifies 2 decimal places, and the answer is "2",
   'then change answer to "2.00"
   '
   '
   Select Case mDecPlaces
      Case 1
         'One decimal place
         If DecimalPointPosition = 0 Then PointzNumberString = PointzNumberString & ".0"
      Case 2
         'Two decimal places
         If DecimalPointPosition = 0 Then
            'If there is no "." in the answer then add ".00"
            PointzNumberString = PointzNumberString & ".00"
         ElseIf DecimalPointPosition = Len(PointzNumberString) - 1 Then
            'If the "." is in the next to the last position like "2.5"
            'then add a "0" to make it "2.50"
            PointzNumberString = PointzNumberString & "0"
         End If
      'If zero decimal places then answer will already be correct
   End Select
   
   'Display the Pointz answer on the form
   lblPoints(0).Caption = PointzNumberString
   
End Sub

Private Sub calcPointsBuster()
   '
   'This sub calculates activity pointz based on the data
   'that the user has input.  It is called with the change event
   'of one of the text boxes, or the scroll event of a slider
   '
   Dim mVal As Double
   Dim PointzNumberString As String
   Dim DecimalPointPosition As Integer
       
   'Calculate Activity Pointz
   mVal = mMinutes * mWeight * (2.5 * mIntensity ^ 2.66 + 1) / 4300
   
   'Error check -> Make sure Pointz answer is within range of 0 to ACTIVITY_POINTZ_MAX
   Select Case mVal
      Case Is < 0
         mVal = 0
      Case Is > ACTIVITY_POINTZ_MAX
         mVal = ACTIVITY_POINTZ_MAX
   End Select
   
   'Round the answer to the number of decimal places the user has specified
   PointzNumberString = Round(mVal, mDecPlaces)
   
   'Find the position of the decimal point in the Pointz String (0 if no decimal)
   DecimalPointPosition = InStr(1, PointzNumberString, ".")
   
   'Fill out blank decimal places with zeros
   'In other words, if user specifies 2 decimal places, and the answer is "2",
   'then change answer to "2.00"
   '
   'Get the position of the decimal point in the Pointz Answer
   Select Case mDecPlaces
      Case 1
         'One decimal place
         If DecimalPointPosition = 0 Then PointzNumberString = PointzNumberString & ".0"
      Case 2
         'Two decimal places
         If DecimalPointPosition = 0 Then
            'If there is no "." in the answer then add ".00"
            PointzNumberString = PointzNumberString & ".00"
         ElseIf DecimalPointPosition = Len(PointzNumberString) - 1 Then
            'If the "." is in the next to the last position like "2.5"
            'then add a "0" to make it "2.50"
            PointzNumberString = PointzNumberString & "0"
         'If zero decimal places then answer will already be correct
         End If
   End Select
   
   'Display the Pointz answer on the form
   lblPoints(1).Caption = PointzNumberString
   
End Sub

Private Sub txtCalories_LostFocus()
   'Make sure the user hasn't entered less calories than the CALORIES_MIN
   If Val(txtCalories.Text) < CALORIES_MIN Then txtCalories.Text = CALORIES_MIN
   
End Sub

Private Sub txtCalories_Change()
   'The user has changed the value in the Calories Text Box
   '
   Dim CalsTemp As Integer
   
   'Get the string from the text box and convert to an number
   CalsTemp = Val(txtCalories.Text)
   
   'Process the change
   Select Case CalsTemp
      Case Is > CALORIES_MAX
         'last digit typed in made number too big
         'so disregard the number by replacing it with
         'the last valid number entered, and put the cursor
         'back where it was
         txtCalories.Text = lastValueStr
         txtCalories.SelStart = lastSelStart
      Case Is >= CALORIES_MIN
         'If program execution makes it to this point, then
         'the entered calories was within the range allowed (CALORIES_MIN to CALORIES_MAX)
         'Load the value into the mCalories variable for future reference, and then
         'recalculate the pointz
         mCalories = CalsTemp
         calcPointsMinder
   End Select
   
End Sub

Private Sub txtCalories_KeyPress(KeyAscii As Integer)
   'This filters non-numeric values so that only numbers can be entered here
   'It also passes navigation keys like backspace, delete, etc.
   
   Select Case KeyAscii
      Case 3, 8, 22, 48 To 57
         'accept
         lastValueStr = txtCalories.Text
         lastSelStart = txtCalories.SelStart
         'remember value before changing in case it needs to be invalidated
      Case Else
         KeyAscii = 0
   End Select
   
End Sub

Private Sub txtMinutes_LostFocus()
   'Make sure the user hasn't entered fewer minutes than the MINUTES_MIN
   If Val(txtMinutes.Text) < MINUTES_MIN Then txtMinutes.Text = MINUTES_MIN

End Sub

Private Sub txtMinutes_Change()
   'The user has changed the value in the Minutes Text Box
   '
   Dim mTemp As Integer
   
   'Get the string from the text box and convert to an number
   mTemp = Val(txtMinutes.Text)
   
   'Process the change
   Select Case mTemp
      Case Is > MINUTES_MAX
         'last digit typed in made number too big
         'so disregard the number by replacing it with
         'the last valid number entered, and put the cursor
         'back where it was
         txtMinutes.Text = lastValueStr
         txtMinutes.SelStart = lastSelStart
      Case Is >= MINUTES_MIN
         'If program execution makes it to this point, then
         'the entered minutes were within the range allowed (MINUTES_MIN to MINUTES_MAX)
         'Load the value into the mMinutes variable for future reference, and then
         'recalculate the pointz
         mMinutes = mTemp
         calcPointsBuster
      End Select
   
End Sub

Private Sub txtMinutes_KeyPress(KeyAscii As Integer)
   'This filters non-numeric values so that only numbers can be entered here
   'It also passes navigation keys like backspace, delete, etc.
   
   Select Case KeyAscii
      Case 3, 8, 22, 48 To 57
         'accept
         lastValueStr = txtMinutes.Text
         lastSelStart = txtMinutes.SelStart
         'remember value before changing in case it needs to be invalidated
      Case Else
         KeyAscii = 0
   End Select
   
End Sub

Private Sub txtWeight_Change()
   'The user has changed the value in the Weight Text Box
   '
   Dim mTemp As Integer
   
   'Get the string from the text box and convert to an number
   mTemp = Val(txtWeight.Text)
   
   'Process the change
   Select Case mTemp
      Case Is > WEIGHT_MAX
         'last digit typed in made number too big
         'so disregard the number by replacing it with
         'the last valid number entered, and put the cursor
         'back where it was
         txtWeight.Text = lastValueStr
         txtWeight.SelStart = lastSelStart
      Case Is >= WEIGHT_MIN
         'If program execution makes it to this point, then
         'the entered weight was within the range allowed (WEIGHT_MIN to WEIGHT_MAX)
         'Load the value into the mWeight variable for future reference, and then
         'recalculate the pointz
         mWeight = mTemp
         calcPointsBuster
   End Select
   
End Sub

Private Sub txtWeight_KeyPress(KeyAscii As Integer)
   'This filters non-numeric values so that only numbers can be entered here
   'It also passes navigation keys like backspace, delete, etc.
   
   Select Case KeyAscii
      Case 3, 8, 22, 48 To 57
         'accept
         lastValueStr = txtWeight.Text
         lastSelStart = txtWeight.SelStart
         'remember value before changing in case it needs to be invalidated
      Case Else
         KeyAscii = 0
   End Select
End Sub

Private Sub txtWeight_LostFocus()
   'Make sure the user hasn't entered less weight than the WEIGHT_MIN
   If Val(txtWeight.Text) < WEIGHT_MIN Then txtWeight.Text = WEIGHT_MIN

End Sub

Private Sub UpDownCalories_DownClick()
   'User has clicked the down arrow next to the Calories text box so the
   'value in the text box must be decremented if it's not alread at the minimum
   '
   Dim CalsTemp As Integer
   
   'get the value from the textbox and decrment by 1
   CalsTemp = txtCalories.Text - 1
   
   'if the temporary number is not too small then put it back in text box
   If CalsTemp >= CALORIES_MIN Then txtCalories.Text = CalsTemp
   
End Sub

Private Sub UpDownCalories_UpClick()
   'User has clicked the up arrow next to the Calories text box so the
   'value in the text box must be incremented if it's not alread at the maximum
   '
   Dim CalsTemp As Integer
   
   'get the value from the textbox and increment by 1
   CalsTemp = txtCalories.Text + 1
   
   'if the temporary number is not too large then put it back in text box
   If CalsTemp <= CALORIES_MAX Then txtCalories.Text = CalsTemp
   
End Sub

Private Sub UpDownMinutes_DownClick()
   'User has clicked the down arrow next to the Minutes text box so the
   'value in the text box must be decremented if it's not alread at the minimum
   '
   Dim CalsTemp As Integer
   
   'get the value from the textbox and decrment by 1
   CalsTemp = txtMinutes.Text - 1
   
   'if the temporary number is not too small then put it back in text box
   If CalsTemp >= MINUTES_MIN Then txtMinutes.Text = CalsTemp
   
End Sub

Private Sub UpDownMinutes_UpClick()
   'User has clicked the up arrow next to the Minutes text box so the
   'value in the text box must be incremented if it's not alread at the maximum
   '
   Dim CalsTemp As Integer
   
   'get the value from the textbox and increment by 1
   CalsTemp = txtMinutes.Text + 1
   
   'if the temporary number is not too large then put it back in text box
   If CalsTemp <= MINUTES_MAX Then txtMinutes.Text = CalsTemp
   
End Sub

Private Sub UpDownWeight_DownClick()
   'User has clicked the down arrow next to the Weight text box so the
   'value in the text box must be decremented if it's not alread at the minimum
   '
   Dim CalsTemp As Integer
   
   'get the value from the textbox and decrment by 1
   CalsTemp = txtWeight.Text - 1
   
   'if the temporary number is not too small then put it back in text box
   If CalsTemp >= WEIGHT_MIN Then txtWeight.Text = CalsTemp
   
End Sub

Private Sub UpDownWeight_UpClick()
   'User has clicked the up arrow next to the Weight text box so the
   'value in the text box must be incremented if it's not alread at the maximum
   '
   Dim x As Integer
   
   'get the value from the textbox and increment by 1
   x = txtWeight.Text + 1
   
   'if the temporary number is not too large then put it back in text box
   If x <= WEIGHT_MAX Then txtWeight.Text = x
   
End Sub

Private Sub saveRegVars()
   On Error Resume Next
   'Save all the values the user has input so they will show up
   'next time the program is executed
   '
   cReg.SetStringValue "Weight", CStr(mWeight)
   cReg.SetStringValue "Intensity", CStr(mIntensity * 20 + 1)
   cReg.SetStringValue "ActivityMinutes", CStr(mMinutes)
   cReg.SetStringValue "Calories", CStr(mCalories)
   cReg.SetStringValue "FatGrams", CStr(mFat * 2 + 1)
   cReg.SetStringValue "FiberGrams", CStr(mFiber * 2 + 1)
   'DecimalPlaces is set when user right clicks and sets manually
   'so you don't have to save it here
   
End Sub

Private Function getRegVars() As Boolean
   'Retrieve the values stored in the Registry last time the program was ran, or else
   'load default values if no values found in Registry
   '
   On Error GoTo getRegVarsErrHdlr
   
   'By default this function will return true
   getRegVars = True
   
   If Not cReg.OpenKey(regKey) Then
      'Key doesn't exist
      If cReg.CreateKey(regKey) Then
         'Successfully created key
         'so store default values in registry
         cReg.SetStringValue "Weight", "100"
         cReg.SetStringValue "Intensity", "0"
         cReg.SetStringValue "ActivityMinutes", "10"
         cReg.SetStringValue "Calories", "10"
         cReg.SetStringValue "FatGrams", "0"
         cReg.SetStringValue "FiberGrams", "0"
         cReg.SetStringValue "DecimalPlaces", "2"
      Else
         'Unable to create key
         GoTo getRegVarsError
      End If
   End If
   
   'Key was opened or created with default values
   'Now read in values from registry and load into variables
   mDecPlaces = CInt(cReg.GetStringValue("DecimalPlaces", "2"))
   mWeight = CInt(cReg.GetStringValue("Weight", "100"))
   mIntensity = CInt(cReg.GetStringValue("Intensity", "0")) * 2 + 1
   mMinutes = CInt(cReg.GetStringValue("ActivityMinutes", "10"))
   mFat = CInt(cReg.GetStringValue("FatGrams", "0")) * 2 + 1
   mFiber = CInt(cReg.GetStringValue("FiberGrams", "0")) * 2 + 1
   mCalories = CInt(cReg.GetStringValue("Calories", "10"))
   
getRegVarsExit:
   On Error GoTo 0
   Exit Function

getRegVarsError:
   getRegVars = False
   GoTo getRegVarsExit

getRegVarsErrHdlr:
   Resume getRegVarsError

End Function
