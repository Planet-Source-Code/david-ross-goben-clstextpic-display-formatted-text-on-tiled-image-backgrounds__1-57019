VERSION 5.00
Begin VB.Form frmTestPic 
   Caption         =   "Text Picxture Test"
   ClientHeight    =   5310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   ScaleHeight     =   5310
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picFormTile 
      Height          =   1815
      Left            =   7920
      OLEDropMode     =   1  'Manual
      Picture         =   "frmTestPic.frx":0000
      ScaleHeight     =   1755
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   1980
      Width           =   435
   End
   Begin VB.PictureBox picTileBackground 
      Height          =   1575
      Left            =   7860
      Picture         =   "frmTestPic.frx":079E
      ScaleHeight     =   1515
      ScaleWidth      =   435
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      Height          =   315
      Left            =   7080
      TabIndex        =   1
      ToolTipText     =   "Select to test the picturebox text"
      Top             =   4320
      Width           =   855
   End
   Begin VB.PictureBox picContainer 
      Height          =   3615
      Left            =   240
      ScaleHeight     =   3555
      ScaleWidth      =   7455
      TabIndex        =   0
      Top             =   240
      Width           =   7515
      Begin VB.VScrollBar VScroll1 
         Height          =   3255
         Left            =   7020
         SmallChange     =   200
         TabIndex        =   4
         Top             =   60
         Width           =   240
      End
      Begin VB.PictureBox picText 
         BackColor       =   &H80000009&
         Height          =   3195
         Left            =   120
         ScaleHeight     =   3135
         ScaleWidth      =   6615
         TabIndex        =   3
         Top             =   120
         Width           =   6675
      End
   End
End
Attribute VB_Name = "frmTestPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Define our Text Picture class
'
Dim TextPic1 As clsTextPic

'*******************************************************************************
' Subroutine Name   : Form_Load
' Purpose           : Set up the new picture text class
'*******************************************************************************
Private Sub Form_Load()
  InitTileFormBackground Me.picFormTile           'init tiling for form itself
  
  Me.picText.BackColor = RGB(255, 235, 179)       'make background similar to tiled image
  Set TextPic1 = New clsTextPic                   'create new object
  Set TextPic1.Picture = Me.picText               'assing picture to use for text output
  Set TextPic1.TilePicture = Me.picTileBackground 'assign picture for tiling
  TextPic1.DoTiling = True                        'we will tile the image
  Set TextPic1.VScroll = Me.VScroll1              'assign a scrollbar
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Resize
' Purpose           : Resize form. Adjust container picture and command button,
'                   : then invoke the class Resize method
'*******************************************************************************
Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then Exit Sub 'do nothing if minimized
  
  Me.cmdTest.Top = Me.ScaleHeight - Me.cmdTest.Height - 60
  Me.picContainer.Width = Me.ScaleWidth - Me.picContainer.Left * 2
  Me.picContainer.Height = Me.cmdTest.Top - Me.picContainer.Top - 60
  Me.cmdTest.Left = Me.picContainer.Left + Me.picContainer.Width - Me.cmdTest.Width
  TextPic1.Resize                           'resize text picture and scrollbar
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Paint
' Purpose           : Tile the form's own background
'*******************************************************************************
Private Sub Form_Paint()
  PaintTileFormBackground Me, Me.picFormTile
End Sub

'*******************************************************************************
' Subroutine Name   : Form_Unload
' Purpose           : Unloading app. Release allocated resources
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
  Set TextPic1 = Nothing                    'release resources
End Sub

'*******************************************************************************
' Subroutine Name   : picText_Paint
' Purpose           : Invoke class' Repaint method on a Repaint Request
'*******************************************************************************
Private Sub picText_Paint()
  TextPic1.Repaint                          'repaint text picture image
End Sub

'*******************************************************************************
' Subroutine Name   : VScroll1_Change
' Purpose           : Scrolling changes
'*******************************************************************************
Private Sub VScroll1_Change()
  TextPic1.Scroll                           'handle scrolling
End Sub

Private Sub VScroll1_Scroll()
  TextPic1.Scroll                           'handle scrolling
End Sub

'*******************************************************************************
' Subroutine Name   : cmdTest_Click
' Purpose           : Test the class
'*******************************************************************************
Private Sub cmdTest_Click()
  Dim Fso As FileSystemObject
  Dim Ts As TextStream
  Dim Str As String
'
' disable test button
'
  Me.cmdTest.Enabled = False      'no need for this again
'
' read a preformatted file in
'
  Set Fso = New FileSystemObject
  Set Ts = Fso.OpenTextFile(App.Path & "\TmpIn.txt", ForReading, False)
  Str = Ts.ReadAll
  Ts.Close
  
  With TextPic1
    .InitText                                       'init new text
    .Indent = 240                                   'set indenting
    .HangingIndent = 480
    .AppendFormatted Str                            'append preformatted text (technically,
                                                    'this text could have been added as the
                                                    'optional parameter to InitText).
'"manual" append of pre-formatted text
    .AppendFormatted "<TBR><Font Symbol,14,B,,,FF0000><BR>Logon Tou TheoV" 'append preformatted
'normal text adding
    .AddText " " & Chr$(&HAC), "Symbol", 14         'Left Arrow from Symbol Character Set
    .AddText "  Blue Greek<BR><BR>", , 14           'add "normally" with vbCrLf's pre-converted
    .AddText "My Sample", "Arial", 12, , True, , vbRed 'display some text in Red
    .AddText " Text." & vbCrLf, "Arial", 12 'Display "My Sample Text." on a single line
    .Indent = 900   'set an indent for text (this is best done after a line has been
                    'added or before new lines).
    .AddText "A Final Sample Line.", "Webdings", 18, , , , vbBlue
'note: color constants do not HAVE to be one of the standard predefined "vb" colors.
'      You can bypass this with an RGB(r,g,b) setting instead, such as:
'    .AddText "<BR>A Final Sample Line.", "Webdings", 18, , , , RGB(128, 255, 179)
    .Repaint        'done adding, so display new stuff
'
' now write the stored preformatted and appended text data to a text file
'
    Set Ts = Fso.OpenTextFile(App.Path & "\TmpOut.txt", ForWriting, True)
    Ts.Write .Text
    Ts.Close
    MsgBox "Sample combined/compressed result file saved: " & App.Path & "\TmpOut.txt", _
            vbOKOnly Or vbInformation, "Result File Written"
  End With
  
  Set Fso = Nothing               'release instantiated object
End Sub
