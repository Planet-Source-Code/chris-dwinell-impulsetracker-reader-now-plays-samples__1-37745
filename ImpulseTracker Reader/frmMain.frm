VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00D3CDD1&
   BorderStyle     =   0  'None
   Caption         =   "ImpulseTracker Reader"
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0BC2
   ScaleHeight     =   6450
   ScaleWidth      =   13500
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSampleConvertBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On: Samples stored as Delta Off: Samples stored as PCM"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   2
      Left            =   10680
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   4260
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleConvertBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On: Samples are signed       Off: Samples are unsigned"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   0
      Left            =   10680
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   3860
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "Sample associated with header"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   0
      Left            =   10680
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   660
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On = 16 bit                            Off = 8 bit"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   1
      Left            =   10680
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   1060
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On = Stereo                          Off = Mono"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   2
      Left            =   10680
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   1460
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "Compressed samples"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   3
      Left            =   10680
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   1860
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "Use loop"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   4
      Left            =   10680
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   2260
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "Use sustain loop"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   5
      Left            =   10680
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   2660
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On = Ping Pong Loop           Off = Forwards Loop"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   6
      Left            =   10680
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3060
      Width           =   2520
   End
   Begin VB.CheckBox chkSampleBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On = Ping Pong Sustain Loop Off = Forwards Sustain Loop"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   7
      Left            =   10680
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   3460
      Width           =   2520
   End
   Begin VB.TextBox txtSampleVibratoType 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   5820
      Width           =   855
   End
   Begin VB.TextBox txtSampleVibratoRate 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtSampleVibratoDepth 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   5220
      Width           =   855
   End
   Begin VB.TextBox txtSampleVibratoSpeed 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4920
      Width           =   855
   End
   Begin VB.TextBox txtSampleOffset 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   4620
      Width           =   855
   End
   Begin VB.TextBox txtSampleSustainLoopEnd 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4320
      Width           =   855
   End
   Begin VB.TextBox txtSampleSustainLoopBegin 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   4020
      Width           =   855
   End
   Begin VB.TextBox txtSampleC5Speed 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtSampleLoopEnd 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3420
      Width           =   855
   End
   Begin VB.TextBox txtSampleLoopBegin 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtSampleLength 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2820
      Width           =   855
   End
   Begin VB.TextBox txtSampleDefaultVolume 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtSampleGlobalVolume 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2220
      Width           =   855
   End
   Begin VB.ListBox lstSamples 
      Height          =   1425
      Left            =   7880
      TabIndex        =   38
      Top             =   680
      Width           =   2700
   End
   Begin VB.PictureBox picCloseArray 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   14400
      Picture         =   "frmMain.frx":11C32C
      ScaleHeight     =   330
      ScaleWidth      =   345
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picCloseArray 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   14040
      Picture         =   "frmMain.frx":11C99E
      ScaleHeight     =   330
      ScaleWidth      =   345
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox picClose 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   13100
      Picture         =   "frmMain.frx":11D010
      ScaleHeight     =   330
      ScaleWidth      =   345
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   60
      Width           =   345
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   14040
      Picture         =   "frmMain.frx":11D682
      ScaleHeight     =   435
      ScaleWidth      =   420
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.TextBox txtOrders 
      BackColor       =   &H00D3CDD1&
      Height          =   555
      Left            =   2190
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   5580
      Width           =   3015
   End
   Begin VB.TextBox txtIT 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4980
      Width           =   495
   End
   Begin VB.TextBox txtIS 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtMV 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4380
      Width           =   495
   End
   Begin VB.TextBox txtGV 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   4710
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtCwt 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox txtPatNum 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4980
      Width           =   495
   End
   Begin VB.TextBox txtSmpNum 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4680
      Width           =   495
   End
   Begin VB.TextBox txtInsNum 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4380
      Width           =   495
   End
   Begin VB.TextBox txtOrdNum 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4080
      Width           =   495
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00D3CDD1&
      Height          =   285
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3780
      Width           =   3015
   End
   Begin VB.CheckBox chkBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "Request embedded MIDI configuration"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   7
      Left            =   5610
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3460
      Width           =   1920
   End
   Begin VB.CheckBox chkBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "Use MIDI pitch controller"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   6
      Left            =   5610
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3060
      Width           =   1920
   End
   Begin VB.CheckBox chkBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "Link Effect G's memory with Effect E/F"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   5
      Left            =   5610
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2660
      Width           =   1920
   End
   Begin VB.CheckBox chkBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On = Old Effects      Off = IT Effects"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   4
      Left            =   5610
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2260
      Width           =   1920
   End
   Begin VB.CheckBox chkBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On = Linear slides    Off = Amiga slides"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   3
      Left            =   5610
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1860
      Width           =   1920
   End
   Begin VB.CheckBox chkBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On = Use instruments Off = Use samples"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   2
      Left            =   5610
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1460
      Width           =   1920
   End
   Begin VB.CheckBox chkBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "Vol0MixOptimizations"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   1
      Left            =   5610
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1060
      Width           =   1920
   End
   Begin VB.CheckBox chkBit 
      BackColor       =   &H00D3CDD1&
      Caption         =   "On = Stereo             Off = Mono"
      Enabled         =   0   'False
      ForeColor       =   &H005D4A4D&
      Height          =   400
      Index           =   0
      Left            =   5610
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   660
      Width           =   1920
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   3030
      Pattern         =   "*.it"
      TabIndex        =   2
      Top             =   870
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   150
      TabIndex        =   1
      Top             =   870
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   510
      Width           =   5175
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Vibrato Type:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   64
      Top             =   5850
      Width           =   1800
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Vibrato Rate:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   62
      Top             =   5550
      Width           =   1800
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Vibrato Depth:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   60
      Top             =   5250
      Width           =   1800
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Vibrato Speed:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   58
      Top             =   4950
      Width           =   1800
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Sample Offset:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   56
      Top             =   4650
      Width           =   1800
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Sustain Loop ends at:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   54
      Top             =   4350
      Width           =   1800
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Sustain Loop begins at:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   52
      Top             =   4050
      Width           =   1800
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "C-5 Speed:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   50
      Top             =   3750
      Width           =   1800
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Loop ends at:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   48
      Top             =   3450
      Width           =   1800
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Loop begins at:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   46
      Top             =   3165
      Width           =   1800
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Length:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   43
      Top             =   2865
      Width           =   1800
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Default Volume:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   41
      Top             =   2565
      Width           =   1800
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Global Volume:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   7800
      TabIndex        =   39
      Top             =   2265
      Width           =   1800
   End
   Begin VB.Label lblTitlebar 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Height          =   390
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   13095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Pattern Order:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   270
      TabIndex        =   32
      Top             =   5745
      Width           =   1785
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Initial Tempo:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   2790
      TabIndex        =   31
      Top             =   5010
      Width           =   1785
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Initial Speed:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   2790
      TabIndex        =   30
      Top             =   4725
      Width           =   1785
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Mix Volume:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   2790
      TabIndex        =   29
      Top             =   4425
      Width           =   1785
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Global Volume:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   2790
      TabIndex        =   28
      Top             =   4125
      Width           =   1785
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Created with tracker:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   270
      TabIndex        =   27
      Top             =   5310
      Width           =   1785
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Number of Patterns:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   270
      TabIndex        =   26
      Top             =   5010
      Width           =   1785
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Number of Samples:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   270
      TabIndex        =   25
      Top             =   4725
      Width           =   1785
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Number of Instruments:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   270
      TabIndex        =   24
      Top             =   4425
      Width           =   1785
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Number of Orders:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   270
      TabIndex        =   23
      Top             =   4125
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00D3CDD1&
      Caption         =   "Name:"
      ForeColor       =   &H005D4A4D&
      Height          =   195
      Left            =   270
      TabIndex        =   22
      Top             =   3825
      Width           =   1785
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Copyright 2000-2002, MTek Designs                              '''
'''                                                               '''
'''  This program was created by Chris Dwinell, MTek Designs.  It '''
'''was made solely for the purpose of learning how to read the    '''
'''ImpulseTracker music file format so that we may add it to our  '''
'''clone of Unreal Tournament which uses DirectX 8.  Source code  '''
'''for the clone is available at the Planet Source Code web site  '''
'''under the Visual Basic section.                                '''
'''Please enjoy and learn something from this!                    '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Public DirectX As New DirectX8
Public DirectSound As DirectSound8

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR = 2
Private lngRegion As Long
Private slide_Ctr As Integer
Private S_Out As Boolean, S_In As Boolean

Private mmflag As Boolean
Private sax As Integer
Private say As Integer

Private Bits(7) As Long
Private NumFlags As Integer

Private Type PatternAttributes
    HeaderOffset As Long
    Length As Integer
    Rows As Integer
End Type
Private Pattern() As PatternAttributes

Private Type SampleAttributes
    BitConvert(7) As Integer
    BitFlags(7) As Integer
    DefaultVolume As Byte
    DOSFilename As String
    DXSoundBuffer As DirectSoundSecondaryBuffer8
    DXSoundBufferReversed As DirectSoundSecondaryBuffer8
    C5Speed As Long
    GlobalVolume As Byte
    HeaderOffset As Long
    Length As Long
    LoopBegin As Long
    LoopEnd As Long
    Name As String
    SampleOffset As Long
    SustainLoopBegin As Long
    SustainLoopEnd As Long
    tmpBuffer() As Byte
    tmpBufferReversed() As Byte   'Used for Bidi looping
    VibratoSpeed As Byte
    VibratoDepth As Byte
    VibratoRate As Byte
    VibratoType As Byte
End Type
Private Sample() As SampleAttributes

Private Function ChangeMask(hWnd As Long, picMask As PictureBox)
    On Error Resume Next
    Dim lngRetr As Long
    lngRegion& = CreateRegionFromBitmap(picMask, vbWhite)
    lngRetr& = SetWindowRgn(hWnd, lngRegion&, True)
End Function

Private Function CreateRegionFromBitmap(picSource As PictureBox, Optional lngTransColor As Long) As Long
    Dim lngRetr As Long, lngHeight As Long, lngWidth As Long
    Dim lngRgnFinal As Long, lngRgnTmp As Long
    Dim lngStart As Long, lngRow As Long
    Dim lngCol As Long
    If lngTransColor& < 1 Then
        lngTransColor& = GetPixel(picSource.hDC, 0, 0)
    End If
    lngHeight& = picSource.Height / Screen.TwipsPerPixelY
    lngWidth& = picSource.Width / Screen.TwipsPerPixelX
    lngRgnFinal& = CreateRectRgn(0, 0, 0, 0)
    For lngRow& = 0 To lngHeight& - 1
        lngCol& = 0
        Do While lngCol& < lngWidth&
            Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) = lngTransColor&
                lngCol& = lngCol& + 1
            Loop
            If lngCol& < lngWidth& Then
                lngStart& = lngCol&
                Do While lngCol& < lngWidth& And GetPixel(picSource.hDC, lngCol&, lngRow&) <> lngTransColor&
                    lngCol& = lngCol& + 1
                Loop
                If lngCol& > lngWidth& Then lngCol& = lngWidth&
                lngRgnTmp& = CreateRectRgn(lngStart&, lngRow&, lngCol&, lngRow& + 1)
                lngRetr& = CombineRgn(lngRgnFinal&, lngRgnFinal&, lngRgnTmp&, RGN_OR)
                DeleteObject (lngRgnTmp&)
            End If
        Loop
    Next
    CreateRegionFromBitmap& = lngRgnFinal&
End Function

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    Call ResetItems
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
    Call ResetItems
End Sub

Private Sub File1_Click()
    Dim i As Integer
    Dim a As Long
    Dim strHex As String
    Dim strHeader As String
    Dim strName As String, intOrdNum As Integer
    Dim intInsNum As Integer, intSmpNum As Integer
    Dim intPatNum As Integer
    Dim strCwt As String, intFlags As Integer
    Dim bytGV As Byte, bytMV As Byte
    Dim bytIS As Byte, bytIT As Byte
    Dim strTempOrders1 As String, strTempOrders2 As String
    Dim strOrders As String
    Dim lngSampleHeadersOffset As Long
    Dim strSampleDOSName As String
    Dim bytSampleFlags As Byte
    Dim strSampleName As String
    Dim bytSampleConvertBits As Byte
    Dim lngPatternHeaderOffset As Long
    
    Call ResetItems
    
    DoEvents
    
    Open File1.Path & "\" & File1.FileName For Binary As #1
        Get #1, 1, strHeader
            strHeader = Input(4, 1)
            If strHeader <> "IMPM" Then
                Close #1
                MsgBox "The file you have selected either is not an ImpulseTracker" & vbNewLine & "file or it is corrupted and does not meet the standard file" & vbNewLine & "format for an ImpulseTracker file!", vbExclamation + vbOKOnly, "File Error!"
                Exit Sub
            End If
        Get #1, 5, strName
            strName = Input(26, 1)
            txtName.Text = strName
            DoEvents
        Get #1, 33, intOrdNum
            txtOrdNum.Text = intOrdNum
            DoEvents
        Get #1, 35, intInsNum
            txtInsNum.Text = intInsNum
            DoEvents
        Get #1, 37, intSmpNum
            txtSmpNum.Text = intSmpNum
            ReDim Sample(intSmpNum) As SampleAttributes
            DoEvents
        Get #1, 39, intPatNum
            txtPatNum.Text = intPatNum
            ReDim Pattern(intPatNum) As PatternAttributes
            DoEvents
        Get #1, 41, strCwt
            strCwt = Input(2, 1)
            strHex = ""
            For i = Len(strCwt) To 1 Step -1
                strHex = strHex & Format$(Hex(Asc(Mid$(strCwt, i, 1))), "00")
            Next i
            strCwt = "ImpulseTracker " & Int(Mid$(strHex, 1, 2)) & "." & Int(Mid$(strHex, 3, 2))
            txtCwt.Text = strCwt
            DoEvents
        Get #1, 45, intFlags
            NumFlags = 0
            For i = 0 To 7
                If intFlags And Bits(i) Then
                    NumFlags = NumFlags + 1
                    chkBit(i).Value = 1
                Else
                    chkBit(i).Value = 0
                End If
            Next i
            DoEvents
        Get #1, 49, bytGV
            txtGV.Text = bytGV
            DoEvents
        Get #1, 50, bytMV
            txtMV.Text = bytMV
            DoEvents
        Get #1, 51, bytIS
            txtIS.Text = bytIS
            DoEvents
        Get #1, 52, bytIT
            txtIT.Text = bytIT
            DoEvents
        Get #1, 193, strTempOrders1
            strTempOrders1 = Input(intOrdNum, 1)
            For i = 1 To intOrdNum
                strTempOrders2 = Mid$(strTempOrders1, i, 1)
                strTempOrders2 = CDec("&H" & Hex(Asc(strTempOrders2)))
                If strTempOrders2 = "255" Then
                    strOrders = Left$(strOrders, Len(strOrders) - 2)
                    Exit For
                End If
                strOrders = strOrders & strTempOrders2 & ", "
            Next i
            txtOrders.Text = strOrders
            DoEvents
        Get #1, 193 + intOrdNum + (intInsNum * 4), lngSampleHeadersOffset
            DoEvents
            lstSamples.Clear
            For i = 1 To intSmpNum
                If i = 1 Then Sample(i).HeaderOffset = lngSampleHeadersOffset + 1
                If i > 1 Then Sample(i).HeaderOffset = Sample(i - 1).HeaderOffset + 80
                strSampleDOSName = ""
                strSampleName = ""
                Get #1, Sample(i).HeaderOffset + 4, strSampleDOSName
                    strSampleDOSName = Input(12, 1)
                    Sample(i).DOSFilename = strSampleDOSName
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 17, Sample(i).GlobalVolume
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 18, bytSampleFlags
                    NumFlags = 0
                    For a = 0 To 7
                        If bytSampleFlags And Bits(a) Then
                            NumFlags = NumFlags + 1
                            Sample(i).BitFlags(a) = 1
                        Else
                            Sample(i).BitFlags(a) = 0
                        End If
                    Next a
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 19, Sample(i).DefaultVolume
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 20, strSampleName
                    strSampleName = Input(26, 1)
                    Sample(i).Name = strSampleName
                    lstSamples.AddItem Format(i, "00") & ": " & strSampleName
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 46, bytSampleConvertBits
                    NumFlags = 0
                    For a = 0 To 7
                        If bytSampleConvertBits And Bits(a) Then
                            NumFlags = NumFlags + 1
                            Sample(i).BitConvert(a) = 1
                        Else
                            Sample(i).BitConvert(a) = 0
                        End If
                    Next a
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 48, Sample(i).Length
                    If Sample(i).BitFlags(1) = 1 Then
                        Sample(i).Length = Sample(i).Length * 2
                    End If
                    ReDim Sample(i).tmpBuffer(Sample(i).Length) As Byte
                    ReDim Sample(i).tmpBufferReversed(Sample(i).Length) As Byte
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 52, Sample(i).LoopBegin
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 56, Sample(i).LoopEnd
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 60, Sample(i).C5Speed
                    If Sample(i).C5Speed > 100000 Then
                        Sample(i).C5Speed = 100000
                    ElseIf Sample(i).C5Speed < 100 Then
                        Sample(i).C5Speed = 100
                    End If
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 64, Sample(i).SustainLoopBegin
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 68, Sample(i).SustainLoopEnd
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 72, Sample(i).SampleOffset
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 76, Sample(i).VibratoSpeed
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 77, Sample(i).VibratoDepth
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 78, Sample(i).VibratoRate
                    DoEvents
                Get #1, Sample(i).HeaderOffset + 79, Sample(i).VibratoType
                    DoEvents
                Get #1, Sample(i).SampleOffset + 1, Sample(i).tmpBuffer
                    For a = 0 To UBound(Sample(i).tmpBufferReversed)
                        Sample(i).tmpBufferReversed(a) = Sample(i).tmpBuffer(UBound(Sample(i).tmpBufferReversed) - a)
                    Next a
                    DoEvents
                
                Call LoadSound(i)
                DoEvents
            Next i
            DoEvents
        Get #1, 193 + intOrdNum + (intInsNum * 4) + (intSmpNum * 4), lngPatternHeaderOffset
            DoEvents
            For i = 1 To intPatNum
                If i = 1 Then
                    Pattern(i).HeaderOffset = lngPatternHeaderOffset + 1
                End If
                If i > 1 Then
                    Pattern(i).HeaderOffset = Pattern(i - 1).HeaderOffset + Pattern(i - 1).Length + 8
                End If
                Get #1, Pattern(i).HeaderOffset, Pattern(i).Length
                Get #1, Pattern(i).HeaderOffset + 2, Pattern(i).Rows
            Next i
    Close #1
End Sub

Private Sub Form_Load()
    Dim i As Integer
    picMask.AutoSize = True
    Call ChangeMask(Me.hWnd, picMask)
    DoEvents
    For i = 0 To 7
        Bits(i) = (2 ^ i)
        DoEvents
    Next i
    DoEvents
    Set DirectSound = DirectX.DirectSoundCreate("")
    DirectSound.SetCooperativeLevel Me.hWnd, DSSCL_PRIORITY
    DoEvents
End Sub

Public Function LoadSound(SampleBuffer As Integer)
    On Error Resume Next
    Dim bufferDesc As DSBUFFERDESC
    Dim i As Integer
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_GLOBALFOCUS
    
    With bufferDesc.fxFormat
        .nFormatTag = WAVE_FORMAT_PCM
        If Sample(SampleBuffer).BitFlags(2) = 1 Then
            .nChannels = 2
        Else
            .nChannels = 1
        End If
        .lSamplesPerSec = Sample(SampleBuffer).C5Speed
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        '''For some reason the bit that is set that    '''
        '''tells us how many bits the sample is is     '''
        '''wrong.  If you set the nBitsPerSample = 8   '''
        '''the sample ends up sounding like crap.  So  '''
        '''just set the sample bit rate to 16.  They   '''
        '''sound correct this way.                     '''
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        'If Sample(SampleBuffer).BitFlags(1) = 1 Then
            .nBitsPerSample = 16
        'Else
        '    .nBitsPerSample = 8
        'End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''
        .nBlockAlign = (.nChannels * .nBitsPerSample) / 8
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
        .lExtra = 0
        .nSize = 0
    End With
    
    bufferDesc.lBufferBytes = Sample(SampleBuffer).Length
    
    If bufferDesc.lBufferBytes > 0 Then
        Set Sample(SampleBuffer).DXSoundBuffer = DirectSound.CreateSoundBuffer(bufferDesc)
        Sample(SampleBuffer).DXSoundBuffer.WriteBuffer 0, bufferDesc.lBufferBytes, Sample(SampleBuffer).tmpBuffer(0), DSBLOCK_DEFAULT
        Set Sample(SampleBuffer).DXSoundBufferReversed = DirectSound.CreateSoundBuffer(bufferDesc)
        Sample(SampleBuffer).DXSoundBufferReversed.WriteBuffer 0, bufferDesc.lBufferBytes, Sample(SampleBuffer).tmpBufferReversed(0), DSBLOCK_DEFAULT
    End If
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picClose.Picture = picCloseArray(0).Picture
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call ResetItems
    DoEvents
    Set DirectX = Nothing
    Set DirectSound = Nothing
    DoEvents
    End
End Sub

Private Sub lblTitlebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mmflag = False Then
        sax = X
        say = Y
        mmflag = True
    End If
End Sub

Private Sub lblTitlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim fml As Long
    Dim fmt As Long
    Dim a As Integer
    picClose.Picture = picCloseArray(0).Picture
    If mmflag = True Then
        fml = Me.Left: fmt = Me.Top
        If X > sax Then Me.Left = fml + (X - sax)
        If X < sax Then Me.Left = fml - (sax - X)
        If Y > say Then Me.Top = fmt + (Y - say)
        If Y < say Then Me.Top = fmt - (say - Y)
    End If
End Sub

Private Sub lblTitlebar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mmflag = False
End Sub

Private Sub lstSamples_Click()
    Dim i As Integer
    txtSampleGlobalVolume.Text = Sample(Left$(lstSamples.Text, 2)).GlobalVolume
    For i = 0 To 7
        chkSampleBit(i).Value = Sample(Left$(lstSamples.Text, 2)).BitFlags(i)
    Next i
    chkSampleConvertBit(0).Value = Sample(Left$(lstSamples.Text, 2)).BitConvert(0)
    chkSampleConvertBit(2).Value = Sample(Left$(lstSamples.Text, 2)).BitConvert(2)
    txtSampleDefaultVolume.Text = Sample(Left$(lstSamples.Text, 2)).DefaultVolume
    txtSampleLength.Text = Sample(Left$(lstSamples.Text, 2)).Length
    txtSampleLoopBegin.Text = Sample(Left$(lstSamples.Text, 2)).LoopBegin
    txtSampleLoopEnd.Text = Sample(Left$(lstSamples.Text, 2)).LoopEnd
    txtSampleC5Speed.Text = Sample(Left$(lstSamples.Text, 2)).C5Speed
    txtSampleSustainLoopBegin.Text = Sample(Left$(lstSamples.Text, 2)).SustainLoopBegin
    txtSampleSustainLoopEnd.Text = Sample(Left$(lstSamples.Text, 2)).SustainLoopEnd
    txtSampleOffset.Text = Sample(Left$(lstSamples.Text, 2)).SampleOffset
    txtSampleVibratoSpeed.Text = Sample(Left$(lstSamples.Text, 2)).VibratoSpeed
    txtSampleVibratoDepth.Text = Sample(Left$(lstSamples.Text, 2)).VibratoDepth
    txtSampleVibratoRate.Text = Sample(Left$(lstSamples.Text, 2)).VibratoRate
    txtSampleVibratoType.Text = Sample(Left$(lstSamples.Text, 2)).VibratoType
End Sub

Private Sub lstSamples_DblClick()
    On Error Resume Next
    Sample(Left$(lstSamples.Text, 2)).DXSoundBuffer.Play DSBPLAY_DEFAULT
End Sub

Private Sub lstSamples_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim Cursor As DSCURSORS, CursorReversed As DSCURSORS
    Dim CurrentBuffer As Integer
    
    If KeyAscii = vbKeyReturn Then
        CurrentBuffer = Left$(lstSamples.Text, 2)
        If Sample(CurrentBuffer).BitFlags(4) = 1 Then
            If Sample(CurrentBuffer).BitFlags(6) = 1 Then
                With Sample(CurrentBuffer)
                    .DXSoundBuffer.SetCurrentPosition 0
                    .DXSoundBuffer.Play DSBPLAY_DEFAULT
                    DoEvents
                '    Do
                '        DoEvents
                '        Do
                '            DoEvents
                '            .DXSoundBuffer.GetCurrentPosition Cursor
                '            DoEvents
                '            If .DXSoundBuffer.GetStatus = DSBSTATUS_PLAYING Then
                '                If Cursor.lPlay >= .LoopEnd Then
                '                    DoEvents
                '                    .DXSoundBuffer.Stop
                '                    DoEvents
                '                    .DXSoundBufferReversed.SetCurrentPosition .Length - .LoopEnd
                '                    .DXSoundBufferReversed.Play DSBPLAY_DEFAULT
                '                    DoEvents
                '                End If
                '            End If
                '            DoEvents
                '        Loop While .DXSoundBuffer.GetStatus = DSBSTATUS_PLAYING
                '        Do
                '            DoEvents
                '            .DXSoundBufferReversed.GetCurrentPosition CursorReversed
                '            DoEvents
                '            If .DXSoundBufferReversed.GetStatus = DSBSTATUS_PLAYING Then
                '                If CursorReversed.lPlay >= .Length - .LoopBegin Then
                '                    DoEvents
                '                    .DXSoundBufferReversed.Stop
                '                    DoEvents
                '                    .DXSoundBuffer.SetCurrentPosition .LoopBegin
                '                    .DXSoundBuffer.Play DSBPLAY_DEFAULT
                '                    DoEvents
                '                End If
                '            End If
                '            DoEvents
                '        Loop While .DXSoundBufferReversed.GetStatus = DSBSTATUS_PLAYING
                '        DoEvents
                '    Loop
                End With
            Else
                Do
                    Sample(CurrentBuffer).DXSoundBuffer.Play DSBPLAY_DEFAULT
                    DoEvents
                Loop Until KeyAscii = vbKeyEscape
            End If
        Else
            Sample(CurrentBuffer).DXSoundBuffer.Play DSBPLAY_DEFAULT
            DoEvents
        End If
    End If
End Sub

Private Sub picClose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picClose.Picture = picCloseArray(0).Picture
    End If
End Sub

Private Sub picClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picClose.Picture = picCloseArray(0).Picture
    Else
        picClose.Picture = picCloseArray(1).Picture
    End If
End Sub

Private Sub picClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If X >= 0 And X <= picClose.Width And Y >= 0 And Y <= picClose.Height Then
            Unload Me
        Else
            picClose.Picture = picCloseArray(0).Picture
        End If
    End If
End Sub

Public Function ResetItems()
    On Error Resume Next
    Dim i As Integer
    
    txtName.Text = ""
    txtOrdNum.Text = ""
    txtInsNum.Text = ""
    txtSmpNum.Text = ""
    txtPatNum.Text = ""
    txtGV.Text = ""
    txtMV.Text = ""
    txtIS.Text = ""
    txtIT.Text = ""
    txtCwt.Text = ""
    txtOrders.Text = ""
    For i = 0 To 7
        chkBit(i).Value = 0
        chkSampleBit(i).Value = 0
    Next i
    chkSampleConvertBit(0).Value = 0
    chkSampleConvertBit(2).Value = 0
    lstSamples.Clear
    txtSampleGlobalVolume.Text = ""
    txtSampleDefaultVolume.Text = ""
    txtSampleLength.Text = ""
    txtSampleLoopBegin.Text = ""
    txtSampleLoopEnd.Text = ""
    txtSampleC5Speed.Text = ""
    txtSampleSustainLoopBegin.Text = ""
    txtSampleSustainLoopEnd.Text = ""
    txtSampleOffset.Text = ""
    txtSampleVibratoSpeed.Text = ""
    txtSampleVibratoDepth.Text = ""
    txtSampleVibratoRate.Text = ""
    txtSampleVibratoType.Text = ""
    DoEvents
End Function

Public Function ReverseHex(strHex As String)
    Dim i As Integer
    For i = Len(strHex) - 1 To 1 Step -2
        ReverseHex = ReverseHex & Mid(strHex, i, 2)
    Next i
End Function
