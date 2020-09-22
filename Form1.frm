VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "FruitLand"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8340
   DrawWidth       =   2
   FillColor       =   &H008080FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Form1.frx":2982
   ScaleHeight     =   443
   ScaleMode       =   3  'ÇÈ¼¿
   ScaleWidth      =   556
   StartUpPosition =   2  'È­¸é °¡¿îµ¥
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1155
      ScaleWidth      =   2955
      TabIndex        =   24
      Top             =   600
      Width           =   3015
      Begin VB.Image ÇÇ¸Ánumber 
         Height          =   300
         Left            =   2160
         Picture         =   "Form1.frx":782E
         Top             =   840
         Width           =   300
      End
      Begin VB.Image Image10 
         Height          =   300
         Left            =   1320
         Picture         =   "Form1.frx":7A51
         Top             =   840
         Width           =   300
      End
      Begin VB.Image Image9 
         Height          =   360
         Left            =   720
         Picture         =   "Form1.frx":7CA7
         Stretch         =   -1  'True
         Top             =   790
         Width           =   360
      End
      Begin VB.Image µþ±ânumber 
         Height          =   300
         Index           =   1
         Left            =   1920
         Picture         =   "Form1.frx":8598
         Top             =   480
         Width           =   300
      End
      Begin VB.Image Image8 
         Height          =   300
         Left            =   1320
         Picture         =   "Form1.frx":87AA
         Top             =   480
         Width           =   300
      End
      Begin VB.Image Image5 
         Height          =   360
         Left            =   720
         Picture         =   "Form1.frx":8A00
         Stretch         =   -1  'True
         Top             =   420
         Width           =   360
      End
      Begin VB.Image µþ±ânumber 
         Height          =   300
         Index           =   0
         Left            =   2160
         Picture         =   "Form1.frx":94CD
         Top             =   480
         Width           =   300
      End
      Begin VB.Image Image6 
         Height          =   450
         Left            =   0
         Picture         =   "Form1.frx":9709
         Top             =   0
         Width           =   3000
      End
   End
   Begin VB.CommandButton Command8 
      DisabledPicture =   "Form1.frx":9FB5
      Height          =   450
      Left            =   960
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Form1.frx":A2A5
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   15
      ToolTipText     =   "³ôÀ½À» °áÁ¤ÇÕ´Ï´Ù."
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Appearance      =   0  'Æò¸é
      BackColor       =   &H80000016&
      DisabledPicture =   "Form1.frx":A645
      Height          =   450
      Left            =   6360
      MaskColor       =   &H80000013&
      Picture         =   "Form1.frx":A908
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   14
      ToolTipText     =   "¹èÆÃÀ» ÇÕ´Ï´Ù."
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      DisabledPicture =   "Form1.frx":AC8D
      Height          =   450
      Left            =   5280
      Picture         =   "Form1.frx":AF8A
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   13
      ToolTipText     =   "½½·Ô¸Ó½Å ÀÛµ¿½ÃÅµ´Ï´Ù."
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      DisabledPicture =   "Form1.frx":B363
      Height          =   450
      Left            =   4200
      Picture         =   "Form1.frx":B63C
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   12
      ToolTipText     =   "½ÀµæÇÑ µ·À» °¡Áý´Ï´Ù."
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      DisabledPicture =   "Form1.frx":B9C7
      Height          =   450
      Left            =   3120
      Picture         =   "Form1.frx":BCA8
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   11
      ToolTipText     =   "´õºí ¹èÆÃÀ» ÇÕ´Ï´Ù."
      Top             =   6000
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      DisabledPicture =   "Form1.frx":C041
      Height          =   450
      Left            =   2040
      Picture         =   "Form1.frx":C33B
      Style           =   1  '±×·¡ÇÈ
      TabIndex        =   10
      ToolTipText     =   "³·À½À» °áÁ¤ÇÕ´Ï´Ù."
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   1410
      Index           =   8
      Left            =   6840
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   9
      Top             =   3480
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   9
         Left            =   -30
         Picture         =   "Form1.frx":C6F9
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   22
         Top             =   0
         Width           =   1410
         Begin VB.Line line8 
            BorderColor     =   &H000000C0&
            Index           =   3
            X1              =   30
            X2              =   1380
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line3 
            BorderColor     =   &H000000C0&
            Index           =   3
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Line line6 
            BorderColor     =   &H000000C0&
            Index           =   3
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   1410
      Index           =   7
      Left            =   5400
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   8
      Top             =   3480
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   8
         Left            =   -30
         Picture         =   "Form1.frx":D1C6
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   21
         Top             =   0
         Width           =   1410
         Begin VB.Line line5 
            BorderColor     =   &H000000C0&
            Index           =   3
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line3 
            BorderColor     =   &H000000C0&
            Index           =   2
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   1410
      Index           =   6
      Left            =   3960
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   7
      Top             =   3480
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   7
         Left            =   -30
         Picture         =   "Form1.frx":DC93
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   20
         Top             =   0
         Width           =   1410
         Begin VB.Line line7 
            BorderColor     =   &H000000C0&
            Index           =   1
            X1              =   30
            X2              =   1380
            Y1              =   1350
            Y2              =   0
         End
         Begin VB.Line line4 
            BorderColor     =   &H000000C0&
            Index           =   3
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line3 
            BorderColor     =   &H000000C0&
            Index           =   1
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   1410
      Index           =   5
      Left            =   6840
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   6
      Top             =   2040
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   6
         Left            =   -30
         Picture         =   "Form1.frx":E760
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   19
         Top             =   0
         Width           =   1410
         Begin VB.Line line6 
            BorderColor     =   &H000000C0&
            Index           =   2
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line1 
            BorderColor     =   &H000000C0&
            Index           =   3
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   1410
      Index           =   4
      Left            =   5400
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   5
      Top             =   2040
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   5
         Left            =   -30
         Picture         =   "Form1.frx":F22D
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   23
         Top             =   0
         Width           =   1410
         Begin VB.Line line1 
            BorderColor     =   &H000000C0&
            Index           =   2
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Line line5 
            BorderColor     =   &H000000C0&
            Index           =   2
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line7 
            BorderColor     =   &H000000C0&
            Index           =   2
            X1              =   30
            X2              =   1380
            Y1              =   1350
            Y2              =   0
         End
         Begin VB.Line line8 
            BorderColor     =   &H000000C0&
            Index           =   2
            X1              =   30
            X2              =   1380
            Y1              =   0
            Y2              =   1350
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   1410
      Index           =   3
      Left            =   3960
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   4
      Top             =   2040
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   4
         Left            =   -30
         Picture         =   "Form1.frx":FCFA
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   18
         Top             =   0
         Width           =   1410
         Begin VB.Image bonuslastnum 
            Height          =   375
            Left            =   0
            Picture         =   "Form1.frx":107C7
            Top             =   480
            Width           =   375
         End
         Begin VB.Line line4 
            BorderColor     =   &H000000C0&
            Index           =   2
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line1 
            BorderColor     =   &H000000C0&
            Index           =   1
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   1410
      Index           =   2
      Left            =   6840
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   3
      Top             =   600
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   3
         Left            =   -30
         Picture         =   "Form1.frx":10BE9
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   17
         Top             =   0
         Width           =   1410
         Begin VB.Line line7 
            BorderColor     =   &H000000C0&
            Index           =   3
            X1              =   30
            X2              =   1380
            Y1              =   1350
            Y2              =   0
         End
         Begin VB.Line line6 
            BorderColor     =   &H000000C0&
            Index           =   1
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line2 
            BorderColor     =   &H000000C0&
            Index           =   3
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   5
      Height          =   1410
      Index           =   1
      Left            =   5400
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   2
      Top             =   600
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   2
         Left            =   -30
         Picture         =   "Form1.frx":116B6
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   16
         Top             =   0
         Width           =   1410
         Begin VB.Line line5 
            BorderColor     =   &H000000C0&
            Index           =   1
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line2 
            BorderColor     =   &H000000C0&
            Index           =   2
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawStyle       =   2  'Á¡
      DrawWidth       =   5
      Height          =   1410
      Index           =   0
      Left            =   3960
      ScaleHeight     =   1350
      ScaleWidth      =   1350
      TabIndex        =   0
      Top             =   600
      Width           =   1410
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  '¾øÀ½
         DrawStyle       =   2  'Á¡
         DrawWidth       =   5
         Height          =   1410
         Index           =   1
         Left            =   -30
         Picture         =   "Form1.frx":12183
         ScaleHeight     =   1410
         ScaleWidth      =   1410
         TabIndex        =   1
         Top             =   0
         Width           =   1410
         Begin VB.Line line2 
            BorderColor     =   &H000000C0&
            Index           =   1
            X1              =   30
            X2              =   1380
            Y1              =   675
            Y2              =   675
         End
         Begin VB.Line line4 
            BorderColor     =   &H000000C0&
            Index           =   1
            X1              =   705
            X2              =   705
            Y1              =   0
            Y2              =   1350
         End
         Begin VB.Line line8 
            BorderColor     =   &H000000C0&
            Index           =   1
            X1              =   30
            X2              =   1380
            Y1              =   0
            Y2              =   1350
         End
      End
   End
   Begin VB.Image Odd 
      Height          =   300
      Index           =   6
      Left            =   960
      Picture         =   "Form1.frx":12C50
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image pickup 
      Height          =   300
      Index           =   6
      Left            =   3360
      Picture         =   "Form1.frx":12E9E
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image credit 
      Height          =   300
      Index           =   6
      Left            =   1320
      Picture         =   "Form1.frx":130EC
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image pickup 
      Height          =   300
      Index           =   5
      Left            =   3600
      Picture         =   "Form1.frx":1333A
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image Card 
      Height          =   1020
      Index           =   4
      Left            =   1080
      Picture         =   "Form1.frx":13588
      Top             =   4440
      Width           =   750
   End
   Begin VB.Image Card 
      Height          =   1020
      Index           =   3
      Left            =   120
      Picture         =   "Form1.frx":13CD8
      Top             =   4440
      Width           =   750
   End
   Begin VB.Image Card 
      Height          =   1020
      Index           =   2
      Left            =   2040
      Picture         =   "Form1.frx":14428
      Top             =   3240
      Width           =   750
   End
   Begin VB.Image Card 
      Height          =   1020
      Index           =   1
      Left            =   1080
      Picture         =   "Form1.frx":14B78
      Top             =   3240
      Width           =   750
   End
   Begin VB.Image Card 
      Height          =   1020
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":152C8
      Top             =   3240
      Width           =   750
   End
   Begin VB.Image Odd 
      Height          =   300
      Index           =   5
      Left            =   1200
      Picture         =   "Form1.frx":15A18
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image Odd 
      Height          =   300
      Index           =   4
      Left            =   1440
      Picture         =   "Form1.frx":15C66
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image Odd 
      Height          =   300
      Index           =   3
      Left            =   1680
      Picture         =   "Form1.frx":15EB4
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image Odd 
      Height          =   300
      Index           =   2
      Left            =   1920
      Picture         =   "Form1.frx":16102
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image Odd 
      Height          =   300
      Index           =   1
      Left            =   2160
      Picture         =   "Form1.frx":16350
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image Odd 
      Height          =   300
      Index           =   0
      Left            =   2400
      Picture         =   "Form1.frx":1659E
      Top             =   2880
      Width           =   300
   End
   Begin VB.Image Image7 
      Height          =   300
      Left            =   2040
      Picture         =   "Form1.frx":167EC
      Top             =   2520
      Width           =   750
   End
   Begin VB.Image pool 
      Height          =   300
      Index           =   3
      Left            =   6960
      Picture         =   "Form1.frx":16B9D
      Top             =   5280
      Width           =   300
   End
   Begin VB.Image pool 
      Height          =   300
      Index           =   2
      Left            =   7200
      Picture         =   "Form1.frx":16DEB
      Top             =   5280
      Width           =   300
   End
   Begin VB.Image pool 
      Height          =   300
      Index           =   1
      Left            =   7440
      Picture         =   "Form1.frx":17039
      Top             =   5280
      Width           =   300
   End
   Begin VB.Image pool 
      Height          =   300
      Index           =   0
      Left            =   7680
      Picture         =   "Form1.frx":17287
      Top             =   5280
      Width           =   300
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   5760
      Picture         =   "Form1.frx":174D5
      Top             =   5280
      Width           =   1200
   End
   Begin VB.Image pickup 
      Height          =   300
      Index           =   4
      Left            =   3840
      Picture         =   "Form1.frx":17883
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image pickup 
      Height          =   300
      Index           =   3
      Left            =   4080
      Picture         =   "Form1.frx":17AD1
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image pickup 
      Height          =   300
      Index           =   2
      Left            =   4320
      Picture         =   "Form1.frx":17D1F
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image pickup 
      Height          =   300
      Index           =   1
      Left            =   4560
      Picture         =   "Form1.frx":17F6D
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image pickup 
      Height          =   300
      Index           =   0
      Left            =   4800
      Picture         =   "Form1.frx":181BB
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3720
      Picture         =   "Form1.frx":18409
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Image linebet 
      Height          =   450
      Index           =   6
      Left            =   3600
      Picture         =   "Form1.frx":189FA
      Top             =   4800
      Width           =   450
   End
   Begin VB.Image linebet 
      Height          =   450
      Index           =   7
      Left            =   3600
      Picture         =   "Form1.frx":18C71
      Top             =   240
      Width           =   450
   End
   Begin VB.Image linebet 
      Height          =   450
      Index           =   5
      Left            =   7320
      Picture         =   "Form1.frx":18EE6
      Top             =   120
      Width           =   450
   End
   Begin VB.Image linebet 
      Height          =   450
      Index           =   4
      Left            =   5880
      Picture         =   "Form1.frx":1919D
      Top             =   120
      Width           =   450
   End
   Begin VB.Image linebet 
      Height          =   450
      Index           =   3
      Left            =   4440
      Picture         =   "Form1.frx":19454
      Top             =   120
      Width           =   450
   End
   Begin VB.Image linebet 
      Height          =   450
      Index           =   2
      Left            =   3480
      Picture         =   "Form1.frx":1970B
      Top             =   3960
      Width           =   450
   End
   Begin VB.Image linebet 
      Height          =   450
      Index           =   1
      Left            =   3480
      Picture         =   "Form1.frx":199AE
      Top             =   1080
      Width           =   450
   End
   Begin VB.Image linebet 
      Height          =   450
      Index           =   0
      Left            =   3480
      Picture         =   "Form1.frx":19C51
      Top             =   2520
      Width           =   450
   End
   Begin VB.Image betting 
      Height          =   300
      Index           =   1
      Left            =   7440
      Picture         =   "Form1.frx":19EF4
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image betting 
      Height          =   300
      Index           =   0
      Left            =   7680
      Picture         =   "Form1.frx":1A142
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   5880
      Picture         =   "Form1.frx":1A390
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   120
      Picture         =   "Form1.frx":1A6AF
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Image credit 
      Height          =   300
      Index           =   5
      Left            =   1560
      Picture         =   "Form1.frx":1AB1B
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image credit 
      Height          =   300
      Index           =   4
      Left            =   1800
      Picture         =   "Form1.frx":1AD69
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image credit 
      Height          =   300
      Index           =   3
      Left            =   2040
      Picture         =   "Form1.frx":1AFB7
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image credit 
      Height          =   300
      Index           =   2
      Left            =   2280
      Picture         =   "Form1.frx":1B205
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image credit 
      Height          =   300
      Index           =   1
      Left            =   2520
      Picture         =   "Form1.frx":1B453
      Top             =   5640
      Width           =   300
   End
   Begin VB.Image credit 
      Height          =   300
      Index           =   0
      Left            =   2760
      Picture         =   "Form1.frx":1B6A1
      Top             =   5640
      Width           =   300
   End
   Begin VB.Menu menufile 
      Caption         =   "&File"
      Begin VB.Menu filenewgame 
         Caption         =   "&NewGame"
         Shortcut        =   ^N
      End
      Begin VB.Menu fileexit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menuoption 
      Caption         =   "&Option"
      Begin VB.Menu optionsound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const NUMFRAMES = 15
Dim bonusmode%, bonus2last%, bonuslast%, barnum%
Dim picArray(1 To NUMFRAMES) As Picture
Dim creditnum As Long, bettingnum%, linebetnum%
Dim linebetting(0 To 7) As Integer
Dim chrpick(1 To 3, 1 To 3) As Integer
Dim chrpickup(1 To 8) As String
Dim pickupnum As Long, poolnum As Integer, ÇÇ¸Ánum As Integer, µþ±ânum As Integer, oddnum As Long
Dim startbt As Integer, stopbt As Integer, allfruitnum As Integer, allbar As Integer
Dim soundonoff As Integer
Dim cardstyle%, cardnum%, cardcho%, cards$, cardn$
'if soundonoff = 1 Then sndPlaySound App.Path + "\sound\click.wav", &H1

Private Sub pooltocredit()
 If poolnum <= 0 Then Exit Sub
 
 Do
  poolnum = poolnum - 10:
  poolnumprint 0
  creditnum = creditnum + 10: creditnumprint 0
  If soundonoff = 1 Then sndPlaySound App.Path + "\sound\ptoc.wav", &H1
 If poolnum = 0 Then Exit Do
 Loop

End Sub
Private Sub pickuptocredit()
 If pickupnum <= 0 Then Exit Sub
 
 Do
  pickupnum = pickupnum - 10:
  If pickupnum < 0 Then
   pickupnum = pickupnum + 10:
   creditnum = creditnum + pickupnum: creditnumprint 0
   pickupnum = 0: pickupprint 0
   If soundonoff = 1 Then sndPlaySound App.Path + "\sound\ptoc.wav", &H1
   Exit Do
  End If
  pickupprint 0
  creditnum = creditnum + 10: creditnumprint 0
  If soundonoff = 1 Then sndPlaySound App.Path + "\sound\ptoc.wav", &H1
 If pickupnum = 0 Then Exit Do
 Loop
 
End Sub
Private Sub bonuslastnumprint()
Dim blnchr$
 If bonus2last = 0 Then Exit Sub
 blnchr$ = LTrim$(Str$(bonus2last))
 bonuslastnum.Picture = LoadPicture(App.Path & "\image\n-" & blnchr$ & ".gif")
 bonuslastnum.Refresh
End Sub
Private Sub poolnumprint(ifcls As Integer)
 Dim poolchr$, i%, j%
 
 poolchr$ = LTrim$(Str$(poolnum))
 If Len(poolchr$) <> 4 Then
  If Len(poolchr$) = 3 Then
   pool(3).Picture = LoadPicture("")
  Else
   For i = Len(poolchr$) To 3
    pool(i).Picture = LoadPicture("")
   Next
  End If
 End If
 j = 0
 For i = Len(poolchr$) To 1 Step -1
  j = j + 1
  pool(j - 1).Picture = LoadPicture(App.Path & "\image\n" & Mid$(poolchr$, i, 1) & ".jpg")
  pool(j - 1).Refresh
 Next
 
End Sub
Private Sub µþ±ânumprint()
 Dim µþ±âchr$, i%, j%
  µþ±âchr$ = LTrim$(Str$(µþ±ânum))
  If Len(µþ±âchr$) = 1 Then µþ±ânumber(1).Picture = LoadPicture("")
  j = 0
  For i = Len(µþ±âchr$) To 1 Step -1
   j = j + 1
   µþ±ânumber(j - 1).Picture = LoadPicture(App.Path & "\image\n2" & Mid$(µþ±âchr$, i, 1) & ".jpg")
   µþ±ânumber(j - 1).Refresh
  Next
  If µþ±ânum = 0 Then If soundonoff = 1 Then sndPlaySound App.Path + "\sound\1barbonus.wav", &H0
End Sub
Private Sub ÇÇ¸Ánumprint()
 Dim ÇÇ¸Áchr$, i%
 ÇÇ¸Áchr$ = LTrim$(Str$(ÇÇ¸Ánum))
 ÇÇ¸Ánumber.Picture = LoadPicture(App.Path & "\image\n2" & ÇÇ¸Áchr$ & ".jpg")
 ÇÇ¸Ánumber.Refresh
End Sub
Private Sub bettingnumprint(ifcls As Integer)
 Dim bettingchr$, i%, j%
 bettingchr$ = LTrim$(Str$(bettingnum))
 If Len(bettingchr$) = 1 Then betting(1).Picture = LoadPicture("")
 
 j = 0
 For i = Len(bettingchr$) To 1 Step -1
  j = j + 1
  betting(j - 1).Picture = LoadPicture(App.Path & "\image\n" & Mid$(bettingchr$, i, 1) & ".jpg")
 Next
End Sub
Private Sub pickupprint(ifcls As Integer)
Dim creditchr$, i%, j%
If stopbt = 0 Then
 If pickupnum = 0 Then
  Image3.Picture = LoadPicture(App.Path & "\image\youlose.jpg")
 Else
  Image3.Picture = LoadPicture(App.Path & "\image\youwin.jpg")
 End If
End If

creditchr$ = LTrim$(Str$(pickupnum))
If Len(creditchr$) <> 7 Then
 If Len(creditchr$) = 6 Then
  pickup(6).Picture = LoadPicture("")
 Else
  For i = Len(creditchr$) To 6
   pickup(i).Picture = LoadPicture("")
  Next
 End If
End If
 j = 0
 For i = Len(creditchr$) To 1 Step -1
  j = j + 1
  pickup(j - 1).Picture = LoadPicture(App.Path & "\image\n" & Mid$(creditchr$, i, 1) & ".jpg")
  pickup(j - 1).Refresh
 Next
 
End Sub
Private Sub oddnumprint()
Dim oddchr$, i%, j%

oddchr$ = LTrim$(Str$(oddnum))
If Len(oddchr$) <> 7 Then
 If Len(oddchr$) = 6 Then
  Odd(6).Picture = LoadPicture("")
 Else
  For i = Len(oddchr$) To 6
   Odd(i).Picture = LoadPicture("")
  Next
 End If
 
End If

 j = 0
 For i = Len(oddchr$) To 1 Step -1
  j = j + 1
  Odd(j - 1).Picture = LoadPicture(App.Path & "\image\n" & Mid$(oddchr$, i, 1) & ".jpg")
  Odd(j - 1).Refresh
 Next
End Sub
Private Sub creditnumprint(ifcls As Integer)
Dim creditchr$, i%, j%

creditchr$ = LTrim$(Str$(creditnum))
If Len(creditchr$) <> 7 Then
 If Len(creditchr$) = 6 Then
  credit(6).Picture = LoadPicture("")
 Else
  For i = Len(creditchr$) To 6
   credit(i).Picture = LoadPicture("")
  Next
 End If
 
End If

 j = 0
 For i = Len(creditchr$) To 1 Step -1
  j = j + 1
  credit(j - 1).Picture = LoadPicture(App.Path & "\image\n" & Mid$(creditchr$, i, 1) & ".jpg")
  credit(j - 1).Refresh
 Next
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command8_Click()
 
 If cardnum >= 6 Then
  Image3.Picture = LoadPicture(App.Path & "\image\youwin.jpg")
  If cardcho = 0 Then
   oddnum = pickupnum * 50: oddnumprint: pickupnum = oddnum
   Card(cardcho).Picture = LoadPicture(App.Path & "\image\" + cards + cardn + ".jpg")
   Card(cardcho).Refresh
   If soundonoff = 1 Then sndPlaySound App.Path + "\sound\double50.wav", &H0
   Command5_Click
   Exit Sub
  End If
  If cardnum <> 6 Then oddnum = oddnum * 2: oddnumprint
  Card(cardcho).Picture = LoadPicture(App.Path & "\image\" + cards + cardn + ".jpg")
  Card(cardcho).Refresh
  If soundonoff = 1 Then sndPlaySound App.Path + "\sound\doublewin.wav", &H0
  If cardnum = 6 Then Card(cardcho).Picture = LoadPicture(App.Path & "\image\cardback.jpg"): Card(cardcho).Refresh Else cardcho = cardcho - 1
  buttontype 3
 Else
  Image3.Picture = LoadPicture(App.Path & "\image\youlose.jpg")
  oddnum = 0: pickupnum = 0: oddnumprint: pickupprint 0
  Card(cardcho).Picture = LoadPicture(App.Path & "\image\" + cards + cardn + ".jpg")
  Card(cardcho).Refresh
  If soundonoff = 1 Then sndPlaySound App.Path + "\sound\doublelose.wav", &H0
  cardbackprint
  cardcho = 4
   If bonusmode = 4 Then bonus3: Exit Sub
   If bonusmode = 3 Then bonus2: Exit Sub
   If µþ±ânum = 0 Then
    bonusmode = 2
    µþ±ânum = 12: µþ±ânumprint
    bonus1
    Exit Sub
   End If
  buttontype 5
 End If
   
End Sub

Private Sub Command3_Click()
 
 If cardnum <= 6 Then
  Image3.Picture = LoadPicture(App.Path & "\image\youwin.jpg")
  If cardcho = 0 Then
   oddnum = pickupnum * 50: oddnumprint: pickupnum = oddnum
   Card(cardcho).Picture = LoadPicture(App.Path & "\image\" + cards + cardn + ".jpg")
   Card(cardcho).Refresh
   If soundonoff = 1 Then sndPlaySound App.Path + "\sound\double50.wav", &H0
   Command5_Click
   Exit Sub
  End If
  If cardnum <> 6 Then oddnum = oddnum * 2: oddnumprint
  Card(cardcho).Picture = LoadPicture(App.Path & "\image\" + cards + cardn + ".jpg")
  Card(cardcho).Refresh
  If soundonoff = 1 Then sndPlaySound App.Path + "\sound\doublewin.wav", &H0
  If cardnum = 6 Then Card(cardcho).Picture = LoadPicture(App.Path & "\image\cardback.jpg"): Card(cardcho).Refresh Else cardcho = cardcho - 1
  buttontype 3
 Else
  Image3.Picture = LoadPicture(App.Path & "\image\youlose.jpg")
  oddnum = 0: pickupnum = 0: oddnumprint: pickupprint 0
  Card(cardcho).Picture = LoadPicture(App.Path & "\image\" + cards + cardn + ".jpg")
  Card(cardcho).Refresh
  If soundonoff = 1 Then sndPlaySound App.Path + "\sound\doublelose.wav", &H0
  cardbackprint
  cardcho = 4
   If bonusmode = 4 Then bonus3: Exit Sub
   If bonusmode = 3 Then bonus2: Exit Sub
   If µþ±ânum = 0 Then
    bonusmode = 2
    µþ±ânum = 12: µþ±ânumprint
    bonus1
    Exit Sub
   End If
  buttontype 5
 End If
End Sub
Private Sub cardbackprint()
 Dim i%
 For i = 0 To 4
   Card(i).Picture = LoadPicture(App.Path & "\image\cardback.jpg")
   
  Next
End Sub

Private Sub Command4_Click()
Image3.Picture = LoadPicture(App.Path & "\image\pickup.jpg")
 If cardcho = 4 Then
  cardbackprint
  oddnum = pickupnum
  oddnumprint
 End If
 buttontype 4
 cardstyle = Int(Rnd(1) * 3)
 Select Case cardstyle
  Case 0: cards = "ÇÏÆ®"
  Case 1: cards = "´ÙÀÌ¾Æ"
  Case 2: cards = "½ºÆäÀÌ½º"
  Case 3: cards = "Å©·Î¹Ù"
 End Select
 cardnum = Int(Rnd(1) * 12)
 Select Case cardnum
  Case 0: cardn = "a"
  Case 1 To 9: cardn = LTrim$(Str$(cardnum + 1))
  Case 10: cardn = "j"
  Case 11: cardn = "q"
  Case 12: cardn = "k"
 End Select
 
End Sub

Private Sub Command5_Click()
 If oddnum > 0 Then
  pickupnum = oddnum: cardcho = 4: oddnum = 0: cardbackprint: oddnumprint
 End If
 Image3.Picture = LoadPicture(App.Path & "\image\pickup.jpg")
 stopbt = 1
 pickuptocredit
 stopbt = 0
 buttontype 5
 If bonusmode = 4 Then bonus3
 If bonusmode = 3 Then bonus2
 If µþ±ânum = 0 Then
  bonusmode = 2
  µþ±ânum = 12: µþ±ânumprint
  bonus1
 End If
End Sub
Private Sub bonus3()
bonuslastnumprint
allfruitnum = 1: 'Label1.Caption = Str$(allfruitnum)
Image3.Picture = LoadPicture(App.Path & "\image\pickup.jpg")
'If startbt = 1 Then creditnum = creditnum - bettingnum
' bettingnumprint 1
 creditnumprint 0
'startbt = 1
Dim i%, j%, chrnum%, l%, j1%, j2%, imsi$, k%
 
 allbar = 1
 buttontype 2
 lineclear
 boxclear
 k = 1
 For l = 1 To 45
  For j = k To 9
    chrnum = Int(Rnd(1) * 39)
   Select Case chrnum
    Case 0 To 5
     chrnum = 1  '1ºü
    Case 6 To 10
     chrnum = 2  '2ºü
    Case 11 To 15
     chrnum = 3  '3ºü
    Case 16 To 18
     chrnum = 4  'super bonanja
    Case 19
     chrnum = 5  '1¼¼ºì
    Case 20
     chrnum = 6  '3¼¼ºì
    Case 21 To 40
     chrnum = 7  'Ä®¶ó
   End Select
   'chrnum = 2
   'if soundonoff = 1 Then sndPlaySound App.Path + "\sound\ing.wav", &H1
   Picture2(j).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + "-1" + ".jpg")
   'Picture2(j).Refresh
   
  Next j
  Select Case l
   Case 5: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(1, 1) = chrnum
    Picture2(1).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(1).Refresh
   Case 10: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(2, 1) = chrnum
    Picture2(2).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(2).Refresh
   Case 15: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(3, 1) = chrnum
    Picture2(3).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(3).Refresh
   Case 20: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(1, 2) = chrnum
    Picture2(4).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(4).Refresh
   Case 25: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(2, 2) = chrnum
    Picture2(5).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(5).Refresh
   Case 30: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(3, 2) = chrnum
    Picture2(6).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(6).Refresh
   Case 35: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(1, 3) = chrnum
    Picture2(7).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(7).Refresh
   Case 40: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(2, 3) = chrnum
    Picture2(8).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(8).Refresh
   Case 45: k = k + 1
    If chrnum > 4 Then allbar = 0
    chrpick(3, 3) = chrnum
    Picture2(9).Picture = LoadPicture(App.Path & "\image\bonus2-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(9).Refresh
  End Select
  
 Next l
chrpickupprint4
': Label1.Caption = Str$(allfruitnum)


End Sub
Private Sub bonus2()
bonuslastnumprint
allfruitnum = 1: 'Label1.Caption = Str$(allfruitnum)
Image3.Picture = LoadPicture(App.Path & "\image\pickup.jpg")
'If startbt = 1 Then creditnum = creditnum - bettingnum
' bettingnumprint 1
 creditnumprint 0
'startbt = 1
Dim i%, j%, chrnum%, l%, j1%, j2%, imsi$, k%
 
 allbar = 1
 buttontype 2
 lineclear
 boxclear
 k = 1
 For l = 1 To 45
  For j = k To 9
    chrnum = Int(Rnd(1) * 14)
   Select Case chrnum
    Case 0, 1
     chrnum = 1  '1ºü
    Case 2, 3
     chrnum = 2  '2ºü
    Case 4, 5
     chrnum = 3  '3ºü
    Case 6
     chrnum = 4  '»ç°ú
    Case 7 To 15
     chrnum = 5  'Ä®¶ó
    
   End Select
   'chrnum = 5
   'if soundonoff = 1 Then sndPlaySound App.Path + "\sound\ing.wav", &H1
   Picture2(j).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + "-1" + ".jpg")
   'Picture2(j).Refresh
   
  Next j
  Select Case l
   Case 5: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(1, 1) = chrnum
    Picture2(1).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(1).Refresh
   Case 10: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(2, 1) = chrnum
    Picture2(2).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(2).Refresh
   Case 15: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(3, 1) = chrnum
    Picture2(3).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(3).Refresh
   Case 20: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(1, 2) = chrnum
    Picture2(4).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(4).Refresh
   Case 25: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(2, 2) = chrnum
    Picture2(5).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(5).Refresh
   Case 30: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(3, 2) = chrnum
    Picture2(6).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(6).Refresh
   Case 35: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(1, 3) = chrnum
    Picture2(7).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(7).Refresh
   Case 40: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(2, 3) = chrnum
    Picture2(8).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(8).Refresh
   Case 45: k = k + 1
    If chrnum > 3 Then allbar = 0
    chrpick(3, 3) = chrnum
    Picture2(9).Picture = LoadPicture(App.Path & "\image\bonus3-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(9).Refresh
  End Select
  
 Next l
chrpickupprint3
': Label1.Caption = Str$(allfruitnum)

End Sub
Private Sub bonus1()
allfruitnum = 1: ' Label1.Caption = Str$(allfruitnum)
Image3.Picture = LoadPicture(App.Path & "\image\pickup.jpg")
'If startbt = 1 Then creditnum = creditnum - bettingnum
' bettingnumprint 1
' creditnumprint 0
'startbt = 1
Dim i%, j%, chrnum%, l%, j1%, j2%, imsi$, k%
 
 allbar = 1
 buttontype 2
 lineclear
 boxclear
 k = 1
 For l = 1 To 45
  For j = k To 9
    chrnum = Int(Rnd(1) * 33)
   Select Case chrnum
    Case 0
     chrnum = 1  'µþ±â
    Case 1 To 5
     chrnum = 2  '1¼¼ºì
    Case 6 To 12
     chrnum = 3  '3¼¼ºì
    Case 13 To 19
     chrnum = 4  'Ä®¶ó
    Case 20 To 24
     chrnum = 5  '1ºü
    Case 25 To 29
     chrnum = 6  '2ºü
    Case 30 To 34
     chrnum = 7  '3ºü
   End Select
   'chrnum = 2
   'if soundonoff = 1 Then sndPlaySound App.Path + "\sound\ing.wav", &H1
   Picture2(j).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + "-1" + ".jpg")
   'Picture2(j).Refresh
  Next j
  Select Case l
   Case 5: k = k + 1
    chrpick(1, 1) = chrnum
    Picture2(1).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(1).Refresh
   Case 10: k = k + 1
    chrpick(2, 1) = chrnum
    Picture2(2).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(2).Refresh
   Case 15: k = k + 1
    chrpick(3, 1) = chrnum
    Picture2(3).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(3).Refresh
   Case 20: k = k + 1
    chrpick(1, 2) = chrnum
    Picture2(4).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(4).Refresh
   Case 25: k = k + 1
    chrpick(2, 2) = chrnum
    Picture2(5).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(5).Refresh
   Case 30: k = k + 1
    chrpick(3, 2) = chrnum
    Picture2(6).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(6).Refresh
   Case 35: k = k + 1
    chrpick(1, 3) = chrnum
    Picture2(7).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(7).Refresh
   Case 40: k = k + 1
    chrpick(2, 3) = chrnum
    Picture2(8).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(8).Refresh
   Case 45: k = k + 1
    chrpick(3, 3) = chrnum
    Picture2(9).Picture = LoadPicture(App.Path & "\image\bonus1-" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(9).Refresh
  End Select
  
 Next l
chrpickupprint2
': Label1.Caption = Str$(allfruitnum)

End Sub
Private Sub nobonus()
allfruitnum = 1 ': Label1.Caption = Str$(allfruitnum)
Image3.Picture = LoadPicture(App.Path & "\image\pickup.jpg")
'If startbt = 1 Then creditnum = creditnum - bettingnum
 bettingnumprint 1
 creditnumprint 0
startbt = 1
Dim i%, j%, chrnum%, l%, j1%, j2%, imsi$, k%
 
 allbar = 1
 buttontype 2
 lineclear
 boxclear
 k = 1
 For l = 1 To 45
  For j = k To 9
    chrnum = Int(Rnd(1) * 79)
   Select Case chrnum
    Case 0 To 11    '12
     chrnum = 6
    Case 12 To 22   '11
     chrnum = 3
    Case 23 To 32   '10
     chrnum = 4
    Case 33 To 41   '9
     chrnum = 5
    Case 42 To 49   '8
     chrnum = 2
    Case 50 To 56   '7
     chrnum = 12
    Case 57 To 62   '6
     chrnum = 1
    Case 63 To 67   '5
     chrnum = 7
    Case 68 To 71   '4
     chrnum = 8
    Case 72 To 74   '3
     chrnum = 9
    Case 75 To 77   '2
     chrnum = 10
    Case 78 To 80         '1
     chrnum = 11
   End Select
   'chrnum = 6
   'if soundonoff = 1 Then sndPlaySound App.Path + "\sound\ing.wav", &H1
   Picture2(j).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + "-1" + ".jpg")
   'Picture2(j).Refresh
  Next j
  Select Case l
   Case 5: k = k + 1
    'chrnum = 2
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(1, 1) = chrnum
    Picture2(1).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(1).Refresh
   Case 10: k = k + 1
    'chrnum = 7
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(2, 1) = chrnum
    Picture2(2).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(2).Refresh
   Case 15: k = k + 1
    'chrnum = 2
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(3, 1) = chrnum
    Picture2(3).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(3).Refresh
   Case 20: k = k + 1
    'chrnum = 7
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(1, 2) = chrnum
    Picture2(4).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(4).Refresh
   Case 25: k = k + 1
    'chrnum = 7
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(2, 2) = chrnum
    Picture2(5).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(5).Refresh
   Case 30: k = k + 1
    'chrnum = 7
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(3, 2) = chrnum
    Picture2(6).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(6).Refresh
   Case 35: k = k + 1
    'chrnum = 2
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(1, 3) = chrnum
    Picture2(7).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(7).Refresh
   Case 40: k = k + 1
    'chrnum = 7
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(2, 3) = chrnum
    Picture2(8).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(8).Refresh
   Case 45: k = k + 1
    'chrnum = 2
    If chrnum < 7 Or chrnum > 9 Then allbar = 0
    If chrnum > 5 And chrnum < 12 Then allfruitnum = 0
    chrpick(3, 3) = chrnum
    Picture2(9).Picture = LoadPicture(App.Path & "\image\f" + LTrim$(Str$(chrnum)) + ".jpg")
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\stop.wav", &H1
    'Picture2(9).Refresh
  End Select
  
 Next l
chrpickupprint
': Label1.Caption = Str$(allfruitnum)

End Sub
Private Sub command6_click()
 If bettingnum > creditnum Then
  If bonusmode = 1 And startbt = 1 Then
   bettingnum = creditnum: creditnum = 0
  End If
 Else
  If bonusmode = 1 And startbt = 1 Then creditnum = creditnum - bettingnum
 End If
 Select Case bonusmode
  Case 1: nobonus
  Case 2: bonus1
  Case 3: bonus2
  Case 4: bonus3
 End Select
  
 
End Sub
Private Sub chrpickupprint4()
Dim afind%, bfind%, abfind%, i%, all$, imsi$
 afind% = 0: bfind% = 0: abfind% = 0
 imsi$ = ""
 pickupnum = 0
 For i = 1 To 8
  chrpickup(i) = ""
 Next
 For i = 1 To 3
  chrpickup(1) = chrpickup(1) + Hex(chrpick(i, 2))
  chrpickup(2) = chrpickup(2) + Hex(chrpick(i, 1))
  chrpickup(3) = chrpickup(3) + Hex(chrpick(i, 3))
  chrpickup(4) = chrpickup(4) + Hex(chrpick(1, i))
  chrpickup(5) = chrpickup(5) + Hex(chrpick(2, i))
  chrpickup(6) = chrpickup(6) + Hex(chrpick(3, i))
  chrpickup(7) = chrpickup(7) + Hex(chrpick(i, (i Xor 3) + 1))
  chrpickup(8) = chrpickup(8) + Hex(chrpick(i, i))
 all = chrpickup(2) + chrpickup(1) + chrpickup(3)
 Next
 For i = 1 To 8
  
  'Label2(i - 1).Caption = chrpickup(i)
 Next
   

 
 For i = 1 To 8
  Select Case chrpickup(i)
   Case "111": pickupnum = pickupnum + (linebetting(i - 1) * 30): lineprint i  '1ºü
    imsi$ = imsi$ + "¿øºüÇÑÁÙ "
   Case "222": pickupnum = pickupnum + (linebetting(i - 1) * 50): lineprint i  '2ºü
    imsi$ = imsi$ + "ÅõºüÇÑÁÙ "
   Case "333": pickupnum = pickupnum + (linebetting(i - 1) * 100): lineprint i '3ºü
    imsi$ = imsi$ + "¾²¸®ºüÇÑÁÙ "
   Case "444": pickupnum = pickupnum + (linebetting(i - 1) * 400): lineprint i  '½´ÆÛº¸³­ÀÚ
    imsi$ = imsi$ + "½´ÆÛº¸³­ÀÚ "
   Case "555": pickupnum = pickupnum + (linebetting(i - 1) * 200): lineprint i '¼¼ºì
    imsi$ = imsi$ + "¿ø½êºìÇÑÁÙ "
   Case "666": pickupnum = pickupnum + (linebetting(i - 1) * 400): lineprint i '¾²¸®¼¼ºì
    imsi$ = imsi$ + "¾²¸®½êºÐÇÑÁÙ "
    
   Case Else
    Dim j%, nojapbar%
    nojapbar = 0
    For j = 1 To 3
     Select Case Mid$(chrpickup(i), j, 1)
      Case "5", "6", "7": nojapbar = 1
     End Select
    Next
    'Label3.Caption = Str$(nojapbar)
    If nojapbar = 0 Then pickupnum = pickupnum + (linebetting(i - 1) * 10): lineprint i
    'aa$ = Mid$(chrpickup(i), 1, 1)
    'bb$ = Mid$(chrpickup(i), 2, 1)
    'cc$ = Mid$(chrpickup(i), 3, 1)
    'If aa$ = "1" Or aa$ = "2" Or aa$ = "3" Or aa$ = "4" Then
    ' If bb$ = "1" Or bb$ = "2" Or bb$ = "3" Or bb$ = "4" Then
    '  If cc$ = "1" Or cc$ = "2" Or cc$ = "3" Or cc$ = "4" Then
    '   imsi$ = imsi$ + "ÀâºüÇÑÁÙ "
    '   pickupnum = pickupnum + (linebetting(i - 1) * 10): lineprint i
    '  End If
    ' End If
    'End If
  End Select
 Next
   For i = 1 To 9
    If Mid$(all, i, 1) = "4" Then
     pickupnum = pickupnum + 10
    End If
   Next
   boxprint all, &HFF
 For i = 1 To 9
  If Mid$(all, i, 1) = "5" Then
    abfind% = abfind% + 1
  ElseIf Mid$(all, i, 1) = "6" Then
    abfind% = abfind% + 1: bfind = bfind + 1
  End If
 Next
 If abfind% > 1 Then
  Dim plusnum%
  plusnum = bfind * bettingnum
  Select Case abfind
   Case 2: pickupnum = pickupnum + (bettingnum * 2) + plusnum
    imsi$ = imsi$ + "½êºìµÎ°³ "
   Case 3: pickupnum = pickupnum + (bettingnum * 5) + plusnum
    imsi$ = imsi$ + "½êºì¼¼°³ "
   Case 4: pickupnum = pickupnum + (bettingnum * 20) + plusnum
    imsi$ = imsi$ + "½êºì³×°³ "
   Case 5: pickupnum = pickupnum + (bettingnum * 50) + plusnum
    imsi$ = imsi$ + "½êºì´Ù¼¸°³ "
   Case 6: pickupnum = pickupnum + (bettingnum * 100) + plusnum
    imsi$ = imsi$ + "½êºì¿©¼¸°³ "
   Case 7: pickupnum = pickupnum + (bettingnum * 200) + plusnum
    imsi$ = imsi$ + "½êºìÀÏ°ö°³ "
   Case 8: pickupnum = pickupnum + (bettingnum * 400) + plusnum
    imsi$ = imsi$ + "½êºì¿©´ü°³ "
   Case 9: pickupnum = pickupnum + (bettingnum * 1000) + plusnum
    imsi$ = imsi$ + "½êºì¾ÆÈ©°³ "
  End Select
 boxprint all, &H80FF
 End If
 
 Select Case all
  Case "333333333"  'all 3ºü
   pickupnum = pickupnum + (bettingnum * 500): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã¾²¸®ºü "
  Case "222222222"  'all 2ºü
   pickupnum = pickupnum + (bettingnum * 400): boxprint all, &H0
   imsi$ = imsi$ + "¿ÃÅõºü "
  Case "111111111"  'all 1ºü
   pickupnum = pickupnum + (bettingnum * 300): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã¿øºü "
  Case "555555555"  'Ä®¶ó
   pickupnum = pickupnum + (bettingnum * 50): boxprint all, &H0
   imsi$ = imsi$ + "¿ÃÄ®¶ó "
 End Select
 

 If allbar = 1 Then pickupnum = pickupnum + (bettingnum * 50): boxprint all, &H0: imsi$ = imsi$ + "¿Ãºü "
   

 pickupprint 1
If soundonoff = 1 And pickupnum > 0 Then sndPlaySound App.Path + "\sound\pickup.wav", &H1
'Label1.Caption = imsi$
bonus2last = bonus2last - 1
If bonus2last < 1 Then
 bonus2last = 7
 barnum = barnum - 1
 If barnum <= 0 Then
  barnum = 0:  bonusmode = 1
  bonuslastnum.Picture = LoadPicture("")
  bonuslastnum.Refresh
 End If
End If
If bonusmode <> 1 Then
 If pickupnum = 0 Then buttontype 6 Else buttontype 3
Else
 If pickupnum = 0 Then buttontype 5 Else buttontype 3
End If


End Sub
Private Sub chrpickupprint3()
Dim afind%, bfind%, abfind%, i%, all$, imsi$
 afind% = 0: bfind% = 0: abfind% = 0
 imsi$ = ""
 pickupnum = 0
 For i = 1 To 8
  chrpickup(i) = ""
 Next
 For i = 1 To 3
  chrpickup(1) = chrpickup(1) + Hex(chrpick(i, 2))
  chrpickup(2) = chrpickup(2) + Hex(chrpick(i, 1))
  chrpickup(3) = chrpickup(3) + Hex(chrpick(i, 3))
  chrpickup(4) = chrpickup(4) + Hex(chrpick(1, i))
  chrpickup(5) = chrpickup(5) + Hex(chrpick(2, i))
  chrpickup(6) = chrpickup(6) + Hex(chrpick(3, i))
  chrpickup(7) = chrpickup(7) + Hex(chrpick(i, (i Xor 3) + 1))
  chrpickup(8) = chrpickup(8) + Hex(chrpick(i, i))
 all = chrpickup(2) + chrpickup(1) + chrpickup(3)
 Next
 For i = 1 To 8
  
  'Label2(i - 1).Caption = chrpickup(i)
 Next
   

 
 For i = 1 To 8
  Select Case chrpickup(i)
   Case "111": pickupnum = pickupnum + (linebetting(i - 1) * 30): lineprint i  '1ºü
    imsi$ = imsi$ + "¿øºüÇÑÁÙ "
   Case "222": pickupnum = pickupnum + (linebetting(i - 1) * 50): lineprint i  '2ºü
    imsi$ = imsi$ + "ÅõºüÇÑÁÙ "
   Case "333": pickupnum = pickupnum + (linebetting(i - 1) * 100): lineprint i '3ºü
    imsi$ = imsi$ + "¾²¸®ºüÇÑÁÙ "
   Case "444": pickupnum = pickupnum + (linebetting(i - 1) * 20): lineprint i  'µþ±â--µþ±â
    imsi$ = imsi$ + "ÇÇ¸ÁÇÑÁÙ "
   Case Else
    Dim j%, nojapbar%
    nojapbar = 0
    For j = 1 To 3
     Select Case Mid$(chrpickup(i), j, 1)
      Case "4": nojapbar = 1
     End Select
    Next
    'Label3.Caption = Str$(nojapbar)
    If nojapbar = 0 Then pickupnum = pickupnum + (linebetting(i - 1) * 10): lineprint i
  End Select
 Next
   For i = 1 To 9
    If Mid$(all, i, 1) = "4" Then
     pickupnum = pickupnum + 10
    End If
   Next
   boxprint all, &HFF
 Select Case all
  Case "333333333"  'all 3ºü
   pickupnum = pickupnum + (bettingnum * 500): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã¾²¸®ºü "
  Case "222222222"  'all 2ºü
   pickupnum = pickupnum + (bettingnum * 400): boxprint all, &H0
   imsi$ = imsi$ + "¿ÃÅõºü "
  Case "111111111"  'all 1ºü
   pickupnum = pickupnum + (bettingnum * 300): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã¿øºü "
  Case "444444444"  '»ç°ú
   pickupnum = pickupnum + (bettingnum * 200): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã»ç°ú "
  Case "555555555"  'Ä®¶ó
   pickupnum = pickupnum + (bettingnum * 50): boxprint all, &H0
   imsi$ = imsi$ + "¿ÃÄ®¶ó "
 End Select
 

 If allbar = 1 Then pickupnum = pickupnum + (bettingnum * 50): boxprint all, &H0: imsi$ = imsi$ + "¿Ãºü "
   

 pickupprint 1
If soundonoff = 1 And pickupnum > 0 Then sndPlaySound App.Path + "\sound\pickup.wav", &H1
'Label1.Caption = imsi$
bonus2last = bonus2last - 1
If bonus2last < 1 Then
 ÇÇ¸Ánum = 7: ÇÇ¸Ánumprint: bonus2last = 7: bonusmode = 1:
 bonuslastnum.Picture = LoadPicture("")
 bonuslastnum.Refresh
Else
 'bonuslastnumprint
End If
If bonusmode <> 1 Then
 If pickupnum = 0 Then buttontype 6 Else buttontype 3
Else
 If pickupnum = 0 Then buttontype 5 Else buttontype 3
End If

End Sub
Private Sub chrpickupprint2()
Dim afind%, bfind%, abfind%, i%, all$, imsi$
 afind% = 0: bfind% = 0: abfind% = 0
 imsi$ = ""
 pickupnum = 0
 For i = 1 To 8
  chrpickup(i) = ""
 Next
 For i = 1 To 3
  chrpickup(1) = chrpickup(1) + Hex(chrpick(i, 2))
  chrpickup(2) = chrpickup(2) + Hex(chrpick(i, 1))
  chrpickup(3) = chrpickup(3) + Hex(chrpick(i, 3))
  chrpickup(4) = chrpickup(4) + Hex(chrpick(1, i))
  chrpickup(5) = chrpickup(5) + Hex(chrpick(2, i))
  chrpickup(6) = chrpickup(6) + Hex(chrpick(3, i))
  chrpickup(7) = chrpickup(7) + Hex(chrpick(i, (i Xor 3) + 1))
  chrpickup(8) = chrpickup(8) + Hex(chrpick(i, i))
 all = chrpickup(2) + chrpickup(1) + chrpickup(3)
 Next
 For i = 1 To 8
  
  'Label2(i - 1).Caption = chrpickup(i)
 Next
 
 
 For i = 1 To 8
  Select Case chrpickup(i)
   Case "111": pickupnum = pickupnum + (linebetting(i - 1) * 10): lineprint i  'µþ±â--µþ±â
    imsi$ = imsi$ + "µþ±âÇÑÁÙ "
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\1barbonus.wav", &H0
    µþ±ânum = µþ±ânum - 2: If µþ±ânum < 0 Then µþ±ânum = 0
    µþ±ânumprint
   Case "222": pickupnum = pickupnum + (linebetting(i - 1) * 200): lineprint i '¼¼ºì
    imsi$ = imsi$ + "¿ø½êºìÇÑÁÙ "
   Case "333": pickupnum = pickupnum + (linebetting(i - 1) * 400): lineprint i '¾²¸®¼¼ºì
    imsi$ = imsi$ + "¾²¸®½êºìÇÑÁÙ "
   Case "555": pickupnum = pickupnum + (linebetting(i - 1) * 30): lineprint i  '1ºü
    imsi$ = imsi$ + "¿øºüÇÑÁÙ "
   Case "666": pickupnum = pickupnum + (linebetting(i - 1) * 50): lineprint i  '2ºü
    imsi$ = imsi$ + "ÅõºüÇÑÁÙ "
   Case "777": pickupnum = pickupnum + (linebetting(i - 1) * 100): lineprint i '3ºü
    imsi$ = imsi$ + "¾²¸®ºüÇÑÁÙ "
   Case Else
    Dim j%, nojapbar%
    nojapbar = 0
    For j = 1 To 3
     Select Case Mid$(chrpickup(i), j, 1)
      Case "1", "2", "3", "4": nojapbar = 1
     End Select
    Next
    'Label3.Caption = Str$(nojapbar)
    If nojapbar = 0 Then pickupnum = pickupnum + (linebetting(i - 1) * 10): lineprint i
  End Select
 Next
 For i = 1 To 9
  If Mid$(all, i, 1) = "2" Then
    abfind% = abfind% + 1
  ElseIf Mid$(all, i, 1) = "3" Then
    abfind% = abfind% + 1: bfind = bfind + 1
  End If
 Next
 
 
 If abfind% > 1 Then
  
  'For i = 1 To 9
  ' If Mid$(all, i, 1) = "A" Then
  '  boxprint i
  ' ElseIf Mid$(all, i, 1) = "B" Then
  '  boxprint i
  ' End If
  'Next
  Dim plusnum%
  plusnum = bfind * bettingnum
  Select Case abfind
   Case 2: pickupnum = pickupnum + (bettingnum * 2) + plusnum
    imsi$ = imsi$ + "½êºìµÎ°³ "
   Case 3: pickupnum = pickupnum + (bettingnum * 5) + plusnum
    imsi$ = imsi$ + "½êºì¼¼°³ "
   Case 4: pickupnum = pickupnum + (bettingnum * 20) + plusnum
    imsi$ = imsi$ + "½êºì³×°³ "
   Case 5: pickupnum = pickupnum + (bettingnum * 50) + plusnum
    imsi$ = imsi$ + "½êºì´Ù¼¸°³ "
   Case 6: pickupnum = pickupnum + (bettingnum * 100) + plusnum
    imsi$ = imsi$ + "½êºì¿©¼¸°³ "
   Case 7: pickupnum = pickupnum + (bettingnum * 200) + plusnum
    imsi$ = imsi$ + "½êºìÀÏ°ö°³ "
   Case 8: pickupnum = pickupnum + (bettingnum * 400) + plusnum
    imsi$ = imsi$ + "½êºì¿©´ü°³ "
   Case 9: pickupnum = pickupnum + (bettingnum * 1000) + plusnum
    imsi$ = imsi$ + "½êºì¾ÆÈ©°³ "
  End Select
 boxprint all, &HFF
 End If
 
 
   For i = 1 To 9
    If Mid$(all, i, 1) = "1" Then
     µþ±ânum = µþ±ânum - 1: µþ±ânumprint
    End If
   Next
   bonusmode = 1
If pickupnum = 0 Then buttontype 5 Else buttontype 3
 pickupprint 1
If soundonoff = 1 And pickupnum > 0 Then sndPlaySound App.Path + "\sound\pickup.wav", &H1
'Label1.Caption = imsi$
End Sub
Private Sub chrpickupprint()
 Dim afind%, bfind%, abfind%, i%, all$, imsi$
 afind% = 0: bfind% = 0: abfind% = 0
 imsi$ = ""
 pickupnum = 0
 For i = 1 To 8
  chrpickup(i) = ""
 Next
 For i = 1 To 3
  chrpickup(1) = chrpickup(1) + Hex(chrpick(i, 2))
  chrpickup(2) = chrpickup(2) + Hex(chrpick(i, 1))
  chrpickup(3) = chrpickup(3) + Hex(chrpick(i, 3))
  chrpickup(4) = chrpickup(4) + Hex(chrpick(1, i))
  chrpickup(5) = chrpickup(5) + Hex(chrpick(2, i))
  chrpickup(6) = chrpickup(6) + Hex(chrpick(3, i))
  chrpickup(7) = chrpickup(7) + Hex(chrpick(i, (i Xor 3) + 1))
  chrpickup(8) = chrpickup(8) + Hex(chrpick(i, i))
 all = chrpickup(2) + chrpickup(1) + chrpickup(3)
 Next
 For i = 1 To 8
  
  'Label2(i - 1).Caption = chrpickup(i)
 Next
 
 
 For i = 1 To 8
  Select Case chrpickup(i)
   Case "111": pickupnum = pickupnum + (linebetting(i - 1) * 10): lineprint i  'µþ±â--µþ±â
    imsi$ = imsi$ + "µþ±âÇÑÁÙ "
    If soundonoff = 1 Then sndPlaySound App.Path + "\sound\f1up.wav", &H0
    µþ±ânum = µþ±ânum - 2: If µþ±ânum <= 0 Then µþ±ânum = 0
   Case "222": pickupnum = pickupnum + (linebetting(i - 1) * 12): lineprint i  '±âÅ¸--Ã¼¸®
    imsi$ = imsi$ + "Ã¼¸®ÇÑÁÙ "
   Case "333": pickupnum = pickupnum + (linebetting(i - 1) * 12): lineprint i  '±âÅ¸--·¹¸ó
    imsi$ = imsi$ + "·¹¸óÇÑÁÙ "
   Case "444": pickupnum = pickupnum + (linebetting(i - 1) * 10): lineprint i  'Æ÷µµ--Æ÷µµ
    imsi$ = imsi$ + "Æ÷µµÇÑÁÙ "
   Case "555": pickupnum = pickupnum + (linebetting(i - 1) * 14): lineprint i  'º¹¼þ¾Æ--¹Ù³ª³ª
    imsi$ = imsi$ + "¹Ù³ª³ªÇÑÁÙ "
   Case "666": pickupnum = pickupnum + (linebetting(i - 1) * 20): lineprint i  'ÇÜ¹ö°Å--ÇÇ¸Á
    imsi$ = imsi$ + "ÇÜ¹ö°ÅÇÑÁÙ "
    ÇÇ¸Ánum = ÇÇ¸Ánum - 1: If ÇÇ¸Ánum < 0 Then ÇÇ¸Ánum = 0: bonusmode = 3
    ÇÇ¸Ánumprint
   Case "777": pickupnum = pickupnum + (linebetting(i - 1) * 30): lineprint i  '1ºü
    imsi$ = imsi$ + "¿øºüÇÑÁÙ "
    bonusmode = 4: barnum = barnum + 1
   Case "888": pickupnum = pickupnum + (linebetting(i - 1) * 50): lineprint i  '2ºü
    imsi$ = imsi$ + "ÅõºüÇÑÁÙ "
   Case "999": pickupnum = pickupnum + (linebetting(i - 1) * 100): lineprint i '3ºü
    imsi$ = imsi$ + "¾²¸®ºüÇÑÁÙ "
   Case "AAA": pickupnum = pickupnum + (linebetting(i - 1) * 200): lineprint i '¼¼ºì
    imsi$ = imsi$ + "¿ø½êºìÇÑÁÙ "
   Case "BBB": pickupnum = pickupnum + (linebetting(i - 1) * 400): lineprint i '¾²¸®¼¼ºì
    imsi$ = imsi$ + "¾²¸®½êºÐÇÑÁÙ "
   Case "CCC": pickupnum = pickupnum + (linebetting(i - 1) * 18): lineprint i  '»ç°ú--»ç°ú
    imsi$ = imsi$ + "»ç°úÇÑÁÙ"
   Case Else
    Dim j%, nojapbar%
    nojapbar = 0
    For j = 1 To 3
     Select Case Mid$(chrpickup(i), j, 1)
      Case "7", "8", "9"
      Case Else: nojapbar = 1
     End Select
    Next
'    Label3.Caption = Str$(nojapbar)
    If nojapbar = 0 Then pickupnum = pickupnum + (linebetting(i - 1) * 10): lineprint i
    If Left$(chrpickup(i), 2) = "11" Then
     pickupnum = pickupnum + (linebetting(i - 1) * 5): lineprint i
     µþ±ânum = µþ±ânum - 1: If µþ±ânum <= 0 Then µþ±ânum = 0
     
     imsi$ = imsi$ + "µþ±â2°³ "
     If soundonoff = 1 Then sndPlaySound App.Path + "\sound\f1up.wav", &H0
    ElseIf Left$(chrpickup(i), 1) = "1" Then
     imsi$ = imsi$ + "µþ±âÇÑ°³ "
     pickupnum = pickupnum + (linebetting(i - 1) * 2): lineprint i
    End If
    
  End Select
 Next
 For i = 1 To 9
  If Mid$(all, i, 1) = "A" Then
    abfind% = abfind% + 1
  ElseIf Mid$(all, i, 1) = "B" Then
    abfind% = abfind% + 1: bfind = bfind + 1
  End If
 Next
 
 If abfind% = 1 Then
  If Mid$(all, 5, 1) = "A" Or Mid$(all, 5, 1) = "B" Then
   If Mid$(all, 5, 1) = "A" Then poolnum = poolnum + 10
   If Mid$(all, 5, 1) = "B" Then poolnum = poolnum + 20
   poolnumprint 1: boxprint all, &HFF
   If soundonoff = 1 Then sndPlaySound App.Path + "\sound\pool.wav", &H0
  End If
 End If
 If abfind% > 1 Then
  
  'For i = 1 To 9
  ' If Mid$(all, i, 1) = "A" Then
  '  boxprint i
  ' ElseIf Mid$(all, i, 1) = "B" Then
  '  boxprint i
  ' End If
  'Next
  Dim plusnum%
  plusnum = bfind * bettingnum
  Select Case abfind
   Case 2: pickupnum = pickupnum + (bettingnum * 2) + plusnum
    imsi$ = imsi$ + "½êºìµÎ°³ "
   Case 3: pickupnum = pickupnum + (bettingnum * 5) + plusnum
    imsi$ = imsi$ + "½êºì¼¼°³ "
   Case 4: pickupnum = pickupnum + (bettingnum * 20) + plusnum
    imsi$ = imsi$ + "½êºì³×°³ "
   Case 5: pickupnum = pickupnum + (bettingnum * 50) + plusnum
    imsi$ = imsi$ + "½êºì´Ù¼¸°³ "
   Case 6: pickupnum = pickupnum + (bettingnum * 100) + plusnum
    imsi$ = imsi$ + "½êºì¿©¼¸°³ "
   Case 7: pickupnum = pickupnum + (bettingnum * 200) + plusnum
    imsi$ = imsi$ + "½êºìÀÏ°ö°³ "
   Case 8: pickupnum = pickupnum + (bettingnum * 400) + plusnum
    imsi$ = imsi$ + "½êºì¿©´ü°³ "
   Case 9: pickupnum = pickupnum + (bettingnum * 1000) + plusnum
    imsi$ = imsi$ + "½êºì¾ÆÈ©°³ "
  End Select
 boxprint all, &HFF
 End If
 
 
 
 
 
 Select Case all
  Case "999999999"  'all 3ºü
   pickupnum = pickupnum + (bettingnum * 500): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã¾²¸®ºü "
  Case "888888888"  'all 2ºü
   pickupnum = pickupnum + (bettingnum * 400): boxprint all, &H0
   imsi$ = imsi$ + "¿ÃÅõºü "
  Case "777777777"  'all 1ºü
   pickupnum = pickupnum + (bettingnum * 300): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã¿øºü "
  Case "111111111"  'µþ±â
   pickupnum = pickupnum + (bettingnum * 400): boxprint all, &H0
   imsi$ = imsi$ + "¿Ãµþ±â "
  Case "222222222"  'Ã¼¸®
   pickupnum = pickupnum + (bettingnum * 80): boxprint all, &H0
   imsi$ = imsi$ + "¿ÃÃ¼¸®"
  Case "333333333"  '·¹¸ó
   pickupnum = pickupnum + (bettingnum * 80): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã·¹¸ó"
  Case "666666666"  'ÇÜ¹ö°Å
   pickupnum = pickupnum + (bettingnum * 300): boxprint all, &H0
   imsi$ = imsi$ + "¿ÃÇÇ¸Á "
  Case "CCCCCCCCC"  '»ç°ú
   pickupnum = pickupnum + (bettingnum * 200): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã»ç°ú "
  Case "555555555"  'º¹¼þ¾Æ
   pickupnum = pickupnum + (bettingnum * 100): boxprint all, &H0
   imsi$ = imsi$ + "¿Ã¹Ù³ª³ª "
  Case "444444444"  'Æ÷µµ
   pickupnum = pickupnum + (bettingnum * 50): boxprint all, &H0
   imsi$ = imsi$ + "¿ÃÆ÷µµ "
  'Case "CCCCCCCCC"  'all ºü
 End Select
 
 If allfruitnum = 1 Then pickupnum = pickupnum + (bettingnum * 15): boxprint all, &H0: imsi$ = imsi$ + "¿Ã°úÀÏ "
 If allbar = 1 Then pickupnum = pickupnum + (bettingnum * 50): boxprint all, &H0: imsi$ = imsi$ + "¿Ãºü "
 If pickupnum = 0 Then buttontype 5 Else buttontype 3
 pickupprint 1
µþ±ânumprint
If soundonoff = 1 And pickupnum > 0 Then sndPlaySound App.Path + "\sound\pickup.wav", &H0
If soundonoff = 1 And bonusmode = 3 Then sndPlaySound App.Path + "\sound\f6up.wav", &H0
If soundonoff = 1 And bonusmode = 4 Then sndPlaySound App.Path + "\sound\1barbonus.wav", &H0

'Label1.Caption = imsi$
If creditnum = 0 And pickupnum = 0 Then gameover
If bonusmode = 3 Then pooltocredit
End Sub
Private Sub boxprint(all2 As String, boxcolor As String)
 Dim i%
 If boxcolor = &H0 Then
  For i = 1 To 9
   Picture2(i).Line (45 + 30, 45)-(87 * 15 + 15, 86 * 15), boxcolor, B
  Next
  Exit Sub
 End If
 If boxcolor = &H80FF Then
  For i = 1 To 9
    If Mid$(all2, i, 1) = "5" Or Mid$(all2, i, 1) = "6" Then Picture2(i).Line (45 + 30, 45)-(87 * 15 + 15, 86 * 15), boxcolor, B
  Next
  Exit Sub
 End If
 For i = 1 To 9
  Select Case bonusmode
   Case 1: If Mid$(all2, i, 1) = "A" Or Mid$(all2, i, 1) = "B" Then Picture2(i).Line (45 + 30, 45)-(87 * 15 + 15, 86 * 15), boxcolor, B
   Case 2: If Mid$(all2, i, 1) = "2" Or Mid$(all2, i, 1) = "3" Then Picture2(i).Line (45 + 30, 45)-(87 * 15 + 15, 86 * 15), boxcolor, B
   Case 3, 4: If Mid$(all2, i, 1) = "4" Then Picture2(i).Line (45 + 30, 45)-(87 * 15 + 15, 86 * 15), boxcolor, B
  End Select
 Next
End Sub
Private Sub lineprint(linenum As Integer)
 Dim i%
 For i = 1 To 3
  Select Case linenum
   Case 1: line1(i).DrawMode = 13: line1(i).Refresh
   Case 2: line2(i).DrawMode = 13: line2(i).Refresh
   Case 3: line3(i).DrawMode = 13: line3(i).Refresh
   Case 4: line4(i).DrawMode = 13: line4(i).Refresh
   Case 5: line5(i).DrawMode = 13: line5(i).Refresh
   Case 6: line6(i).DrawMode = 13: line6(i).Refresh
   Case 7: line7(i).DrawMode = 13: line7(i).Refresh
   Case 8: line8(i).DrawMode = 13: line8(i).Refresh
  End Select
 Next
End Sub
Private Sub chrprint(k As Integer)
 
End Sub
Private Sub piccls()
Set picArray(1) = LoadPicture("")
'Set picArray(1) = LoadPicture(App.Path & "\image\sb" & bld$ & ".jpg")
End Sub

Private Sub Command2_Click()
 Picture2(1).Line (45 + 30, 45)-(87 * 15 + 15, 86 * 15), &HFF, B
 'boxprint "AAAAAAAAA",
 
End Sub
Private Sub linebetnumcls()
 Dim i%, imsi2$
 For i = 1 To 8
 Select Case i
   Case 1 To 3: imsi2$ = "a"
   Case 4 To 6: imsi2$ = "b"
   Case 7: imsi2$ = "d"
   Case 8: imsi2$ = "c"
  End Select
    linebet(i - 1).Picture = LoadPicture(App.Path + "\image\" + imsi2$ + "-0" + ".jpg")
    linebet(i - 1).Refresh
 Next
End Sub
Private Sub linebetnumprint()
 Dim imsi$, imsi2$, i%
 
 'For i = 1 To bettingnum
  linebetnum = linebetnum + 1
  
  Select Case linebetnum
   Case 1 To 3: imsi2$ = "a"
   Case 4 To 6: imsi2$ = "b"
   Case 7: imsi2$ = "d"
   Case 8: imsi2$ = "c"
  End Select
  linebetting(linebetnum - 1) = linebetting(linebetnum - 1) + 1
  
    imsi = LTrim$(Str$(linebetting(linebetnum - 1)))
    
    
    
    
    
    linebet(linebetnum - 1).Picture = LoadPicture(App.Path + "\image\" + imsi2$ + "-" + imsi$ + ".jpg")
    linebet(linebetnum - 1).Refresh
'  linebet(linebetnum).picture
 If linebetnum > 7 Then linebetnum = 0
 'Next
End Sub

Private Sub Command7_Click()
Dim i%
If creditnum = 0 Then Exit Sub
If soundonoff = 1 Then sndPlaySound App.Path + "\sound\betting.wav", &H1
'If bettingnum = 50 And startbt = 1 Then Exit Sub
If startbt = 1 Then
 startbt = 0: bettingnum = 0: linebetnum = 0
 For i = 0 To 7
 linebetting(i) = 0
 Next
 linebetnumcls
End If

 
 bettingnum = bettingnum + 1
 If bettingnum > 50 Then bettingnum = 50: Exit Sub
 
 
  creditnum = creditnum - 1
  linebetnumprint
  'linebetnum = linebetnum + 1
 
 bettingnumprint 1
 creditnumprint 0
 buttontype 5
End Sub

Private Sub boxclear()
 Dim i%
 For i = 0 To 8
  Picture1(i).Line (15, 30)-(87 * 15, 87 * 15), vbWhite, B
 Next
End Sub
Private Sub lineclear()
Dim i%
For i = 1 To 3
 line1(i).DrawMode = 11 ': line1(i).Refresh
 line2(i).DrawMode = 11 ': line2(i).Refresh
 line3(i).DrawMode = 11 ': line3(i).Refresh
 line4(i).DrawMode = 11 ': line4(i).Refresh
 line5(i).DrawMode = 11 ': line5(i).Refresh
 line6(i).DrawMode = 11 ': line6(i).Refresh
 line7(i).DrawMode = 11 ': line7(i).Refresh
 line8(i).DrawMode = 11 ': line8(i).Refresh
Next

End Sub
Private Sub buttontype(btype As Integer)
 Select Case btype
  Case 1 '----ÃÊ±â»óÅÂ
   Command8.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
   Command6.Enabled = False
   Command7.Enabled = True
  Case 2 '----½½·Ô¸Ó½ÅÀÌ ÀÛµ¿ÁßÀÏ¶§
   Command8.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
   Command6.Enabled = False
   Command7.Enabled = False
  Case 3 '----µ·À» ½ÀµæÇßÀ»¶§
   Command8.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = True
   Command5.Enabled = True
   Command6.Enabled = False
   Command7.Enabled = False
  Case 4 '----´õºí¹èÆÃÀ» ÇßÀ»¶§
   Command8.Enabled = True
   Command3.Enabled = True
   Command4.Enabled = False
   Command5.Enabled = False
   Command6.Enabled = False
   Command7.Enabled = False
  Case 5 '----½ÃÀÛÀ» ±â´Ù¸±¶§
   Command8.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
   Command6.Enabled = True
   Command7.Enabled = True
  Case 6 '----ÇÇ¸Á, 1ºü º¸³Ê½ºÀÏ¶§
   Command8.Enabled = False
   Command3.Enabled = False
   Command4.Enabled = False
   Command5.Enabled = False
   Command6.Enabled = True
   Command7.Enabled = False
 End Select
End Sub
Private Sub gameover()
 If soundonoff = 1 Then sndPlaySound App.Path + "\sound\gameover.wav", &H0
End Sub


Private Sub Form_Load()
 Dim i%
 Randomize Timer
 bettingnum = 0: bettingnumprint 1
 creditnum = 1000: creditnumprint 1
 ÇÇ¸Ánum = 7: ÇÇ¸Ánumprint
 µþ±ânum = 12: µþ±ânumprint
 oddnum = 0: oddnumprint
 pickupnum = 0: pickupprint 1
 poolnum = 0: poolnumprint 1
 cardcho = 4: bonus2last = 7: bonusmode = 1
 soundonoff = 1
 buttontype 1
 bonuslast = 7
 linebetnum = 0
 lineclear
 cardbackprint
 bonuslastnum.Picture = LoadPicture("")
 betting(1).Picture = LoadPicture("")
 Image3.Picture = LoadPicture(App.Path & "\image\pickup.jpg")
 For i = 1 To 9
  Picture2(i).Picture = LoadPicture(App.Path & "\image\f1.jpg")
 Next
 For i = 0 To 7
  linebetting(i) = 0
 Next

 startbt = 0
End Sub

Private Sub filenewgame_Click()
 Form_Load
End Sub


Private Sub optionsound_Click()
 If optionsound.Checked = False Then
  optionsound.Checked = True
  soundonoff = 1
 Else
  optionsound.Checked = False
  soundonoff = 0
 End If

End Sub
