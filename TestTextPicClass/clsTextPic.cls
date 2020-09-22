VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*******************************************************************************
' clsTextPic - Treat a picturebox like a textbox, with an optional tiled background
'
' This class assumes that the target "text" picturebox is contained within a bounding
' picture box.  Sizing witll default to the container picturebox height. This class
' assumes that the picturebox you will use for drawing text with contained within another
' picturebox (this allows smoother scrolling). Also, if you include a vertical scrollbar,
' it must also be a child of the container picturebox (simply create a containder picturebox
' and add a picturebox that will be used for textoutput and a scrollbar within it).
'
' I realize that this can be made into a custom control very easily, but this is just a
' quick example on accomplishing the task at hand.
'
' ASSUMPTIONS:
'  1. The picturebox that will be used for writing text is contained within another
'     picturebox as a bounding container.
'  2. If a vertical scrollbar will be used, it must also be placed within the
'     container picturebox.
'  3. The picturebox containing the image to be tiled in the background will not be
'     normally displayed on your form.
'
' HOWTO: Use the class:
''' In the Heading of the Form:
'  Dim TextPic1 As clsTextPic               'define a class reference variable
'
'''' in the Form_Load event:
'  Set TextPic1 = New clsTextPic            'instantialte class
'  Set TextPic1.Picture = Me.picText        'add picture to draw text upon. This must be
                                            'contained within another picturebox, unless you
                                            'want it to take up the entire form.
'  Set TextPic1.TilePicture = Me.Picture2   'optional picture to tile background image
'  TextPic1.DoTiling = True                 'we want the background image tiled, not single
'  Set TextPic1.VScroll = Me.VScroll1       'optional vertical scrollbar. This mus also share
                                            'the same parent as picText.
'
''' in the PicText_Repaint event:
'  TextPic1.Repaint                         'repaint as needed
'
''' In the Form_Resize Event, add:
'  TextPic1.Resize                          'shape to containder object
'
'''' in the Form_Unload event:
'  Set TextPic1 = Nothing                   'release allocated resources
'
''' If you have a vertical scrollbar, place this in the scrollbar's Change and Scroll events:
'  TextPic1.Scroll                          'handle scrolling events
'
'--------------------
' ADDING TEXT EXAMPLE
'--------------------
'  TextPic1.InitText       'start with completely fresh data. NOTE: if you have some
'                          'preformatted text (saved from a rpevious session; see below) then
'                          'you can insert it right here as the optional text parameter to
'                          'InitText().
'  TextPic1.Indent = 900   'set an indent for text (this is best done after a line has been
'                          'added or before new lines).
'  TextPic1.HangingIndent = 450  'hanging indent. Usually you set Indent or HangingIndent,
'                                'but seldome both, but this is still no problem.
'  TextPic1.AddText "My Sample","Arial",14
'  TextPic1.AddText " Text." & vbCrLf, "Arial",14 'Display "My Sample Text." on a single line
'  TextPic1.AddText "Another Sample Line.<br>" 'bypass adding vbCrLF to text by inserting
'                                              'linebreaks (<BR>) "manually". Also display
'                                              'in default 10-point Times New Roman.
'  TextPic1.Repaint        'done adding text, so update the display
'
'  Debug.Print TextPic1.Text 'retrieve preformatted text
'append some preformatted text (see below information)...
'  TextPic1.AppendFormatted "<Font Symbol,14,B,,,FF0000><BR><BR>Logon Tou TheoV"
'normal append of information...
'   TextPic1.AddText Chr$(&HAC), "Symbol", 14        'Left Arrow from Symbol Character Set
'  TextPic1.AddText " Blue Greek" & vbCrLf, "Arial", 14
'
'  TextPic1.Repaint        'done adding text, so update the display
'*******************************************************************************
' Formatted text information
'
' Though not sophisticated, this format works very well.  When adding text with
' the AddText method, the following text transformations take place:
' vbCrLf codes are changed to "<BR>".
' The font information is formatted thus, presenting the data being sent:
'   "<Font FontName,10,B,I,U,BBGGRR>"
'   Where:  'FontName' is the font name, such as Symbol, Arial, Times New Roman.
'           '10'       is the point size, such as 8, 10, 12, 14, 16, etc.
'           'B'        is any character to enable Bold (B is used). Not present means off.
'           'I'        is any Character to enable Italics (I is used). Not present means off.
'           'U'        is any Character to enable Underline (U is used). Not present means off.
'           'BBGGRR'   is a hex value of an RGB color. BB = Blue, GG = Green, RR = Red.
'              Normally, when you add colors manually, you use the system color values,
'              such as vbBlue, vbRed, vbBlack, etc., but you can also used the RGB() function, or
'              a Long variable containing a color value.
'
'  Indent format:         <IND value>   where 'value' is a decimal value in twips.
'  Hanging Indent format: <HIND value>  where 'value' is a decimal value in twips.
'
' Because the text is stored internally in a single string, a "cheat" has been added
' to make manually eding the data easier. In provided formatted text data, any encountered
' vbCrLf codes are internally transformed into "<TBR>", which is ignored during processing.
' If you retrieve the formatted text via the "Text" property, the internal <TBR> codes are
' converted back into vbCrLf.
'*******************************************************************************

'*******************************************************************************
' API goodies to grab text width from device context info.  Normally, I use a
'     non-visible autosizing label to do this, but this API method is much
'     faster (and also bypasses the need to ALSO define a label for the class
'     to use.
'*******************************************************************************
Private Type SIZE
  cx As Long
  cy As Long
End Type

Private Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" _
                (ByVal hdc As Long, _
                 ByVal lpszString As String, _
                 ByVal cbString As Long, _
                 lpSize As SIZE) As Long
'*******************************************************************************
' private, local storage for class
'*******************************************************************************
Private m_HasPic As Boolean         'true when picture can be processed
Private m_HasTile As Boolean        'when tiling image present
Private m_HasScroll As Boolean      'when scrollbar is present

Private m_Picture As PictureBox     'picturebox for text
Private m_PicTile As PictureBox     'picturebox containing a tiling pattern
Private m_DoTiling As Boolean       'true if we will want the background image tiled
Private m_Text As String            'text to draw
Private m_vScroll As VScrollBar     'optional vertical scrollbar
Private m_LabelWidth As Label       'label for figuring wraparound
Private m_Indent As Long            'paragraph indent in twips
Private m_HangingIndent As Long     'hanging indent in twips
'
' local-local stuff
'
Private LastFontName As String
Private LastFontSize As Long
Private LastFontBold As Boolean
Private LastFontItal As Boolean
Private LastFontUnd As Boolean
Private LastFontColor As ColorConstants
Private Tpp As Long                     'twips per pixel storage

Private Width  As Long                  'running line text width in twips accumulator

Private Const ScrollScale As Long = 100 'scrollbars are limit to -32768 to 32767,
                                        ' so we must scale its value

'*******************************************************************************
' Subroutine Name   : Class_Initialize
' Purpose           : Make sure text buffer is null
'*******************************************************************************
Private Sub Class_Initialize()
  m_Text = vbNullString                 'none of this really needed, but it makes
  LastFontName = vbNullString           'the intention clear
  m_Indent = 0
  m_HangingIndent = 0
  Width = 0
'
' keep 'puriests' happy
'
  Tpp = Screen.TwipsPerPixelX
End Sub

'*******************************************************************************
' Subroutine Name   : Class_Terminate
' Purpose           : Remove hooks to objects
'*******************************************************************************
Private Sub Class_Terminate()
  Set m_Picture = Nothing     'stomp it so that it cannot rise in the background
  Set m_PicTile = Nothing     'this is not usually necessary, but it makes the intention clear
  m_Text = vbNullString
  m_HasPic = False
  m_HasTile = False
  m_HasScroll = False
End Sub

'*******************************************************************************
' Set Name          : Picture
' Purpose           : Assign picture to store text
'*******************************************************************************
Public Property Set Picture(Pic As PictureBox)
  If m_HasPic Then Exit Property      '1-shot deal
  
  If Pic Is Nothing Then              'may be resetting...
    If m_HasPic Then m_Picture.Cls    'clear if so
    Set m_Picture = Nothing           'remove assignment
    m_HasPic = False                  'flag that we do not now have it
    Exit Property                     'exit
  End If
  
  Set m_Picture = Pic                 'assign picture
  m_HasPic = True                     'flag that we now have it
  With m_Picture
    .BorderStyle = 0                  'remove border
    .AutoRedraw = True                'force redraw (makes scrolling SMOOTH)
    .Top = .Container.ScaleTop        'ensure top/left in container is flush
    .Left = .Container.ScaleLeft
    .Height = .Container.ScaleHeight  'height is minimally full height of container
    .Width = .Container.ScaleWidth    'and full width of container
  End With
End Property

'*******************************************************************************
' Set Name          : TilePicture
' Purpose           : Assign tiling background
'*******************************************************************************
Public Property Set TilePicture(Pic As PictureBox)
  If Not m_HasPic Then Exit Property  'cannot do if picture not defined
  
  Set m_PicTile = Pic                 'assign tile picture
  If Pic Is Nothing Then
    m_HasTile = False                 'flag not having it
    m_DoTiling = False
    Call Me.Repaint                   'refresh screen
    Exit Property
  End If
  With Pic
    .AutoSize = True                  'make sure image is fully displayed
    .BorderStyle = 0                  'prevent copying border of picture
    .Visible = False                  'remove it form normal display
  End With
  m_HasTile = True                    'flag having it
  Call Me.Repaint                     'tile it in
End Property

'*******************************************************************************
' Let Name          : DoTiling
' Purpose           : Setting this property to True causes the Tile picture to
'                   : Actually Tile
'*******************************************************************************
Public Property Let DoTiling(Flag As Boolean)
  If Not m_HasTile Then Exit Property
  If m_DoTiling = Flag Then Exit Property  'pointless action
  m_DoTiling = Flag
  Call Me.Repaint                          'process it (on or off)
End Property

'*******************************************************************************
' Set Name          : VScroll
' Purpose           : Assign a companion vertical scroll bar
'*******************************************************************************
Public Property Set VScroll(ScrollBar As VScrollBar)
  If Not m_HasPic Then Exit Property        'cannot do if picture not defined
  
  Set m_vScroll = ScrollBar
  If m_vScroll Is Nothing Then              'we may be resetting
    m_HasScroll = False                     'flag no longer having it
    Call Me.Resize                          'resize stuff
    Exit Property                           'and leave
  End If
  
  m_HasScroll = True                        'flag having it
  With m_vScroll
    .Left = .Container.ScaleWidth - .Width  'set flush right
    .Top = .Container.ScaleTop              'force to top of container
    .Height = .Container.ScaleHeight        'and full height
    .SmallChange = 700 / ScrollScale        'small changes to about 1/2 inch
    m_Picture.Width = m_vScroll.Left        'force text picture left to accomodate scroller
  End With
  Call Me.Repaint                           'tile it in
End Property

'*******************************************************************************
' Let Name          : Indent
' Purpose           : Assign Left Indent Value
'
'This command is considered "live". It allows indenting and unindenting particular
'line of text.  Normally, you would invoke this after a line that contains a Newline
'code (vbCrLf or <BR>), and prior to the code you want indented.
'*******************************************************************************
Public Property Let Indent(ByVal IndentValue As Long)
  If Not m_HasPic Then Exit Property        'cannot do if picture not defined
  
  If IndentValue < 0 Then                           'do not allow negatives
    m_Indent = 0
  Else
    m_Indent = IndentValue
  End If
  m_Text = m_Text & "<IND " & CStr(m_Indent) & ">"  'add to text data
End Property

'*******************************************************************************
' Get Name          : Indent
' Purpose           : Grab Indent Value
'*******************************************************************************
Public Property Get Indent() As Long
  Indent = m_Indent
End Property

'*******************************************************************************
' Let Name          : HangingIndent
' Purpose           : Assign Hanging Indent Value
'
' this command is considered "live". It allows hanging indents and unindenting particular
'line of text.  Normally, you would invoke this after a line that contains a Newline
'code (vbCrLf), and prior to the code you want to having a hanging indent.
'*******************************************************************************
Public Property Let HangingIndent(ByVal IndentValue As Long)
  If Not m_HasPic Then Exit Property  'cannot do if picture not defined
  
  If IndentValue < 0 Then
    m_HangingIndent = 0
  Else
    m_HangingIndent = IndentValue
  End If
  m_Text = m_Text & "<HIND " & CStr(m_HangingIndent) & ">"
End Property

'*******************************************************************************
' Get Name          : HangingIndent
' Purpose           : Grab Hanging Indent Value
'*******************************************************************************
Public Property Get HangingIndent() As Long
  HangingIndent = m_HangingIndent
End Property

'*******************************************************************************
' Get Name          : Text
' Purpose           : Retrieve formatted text
'*******************************************************************************
Public Property Get Text() As String
  If Not m_HasPic Then Exit Property  'cannot do if picture not defined
  Text = FixFormat(m_Text)            'get data that can be saved to a file
End Function

'*******************************************************************************
' Sub Name          : AppendFormatted
' Purpose           : Append Pre-formatted Text
'*******************************************************************************
Public Sub AppendFormatted(Text As String)
  If Not m_HasPic Then Exit Sub                            'cannot do if picture not defined
  
  m_Text = m_Text & CompressTextFormat(UnFixFormat(Text))  'change any vbCrLf to <TBR>
End Sub

'*******************************************************************************
' Subroutine Name   : Repaint
' Purpose           : Repaint main picture
'
' Invoke this after you have added the text you want displayed
'*******************************************************************************
Public Sub Repaint()
  If Not m_HasPic Then Exit Sub                               'cannot do if picture not defined
  
  With m_Picture
    .Height = .Container.ScaleHeight                          'force height to minimum
    TileBackground                                            'tile background
    If Len(m_Text) <> 0 Then                                  'text present?
      DrawText                                                'draw it if so
      If .CurrentY > .ScaleHeight Then                        'data bigger than picture?
        .Height = .CurrentY                                   'yes, so readjust
        TileBackground                                        'and re-do
        DrawText
      End If
      '
      ' scroll bar support
      '
      If .Height > .Container.ScaleHeight Then                'height greater than container
        If m_HasScroll Then                                   'if we have a scroll bar
          m_vScroll.Min = 0
          m_vScroll.Max = (.Height - .Container.ScaleHeight) / ScrollScale 'set scaling (scale for short int)
          m_vScroll.LargeChange = m_vScroll.Height * 0.95 / ScrollScale   'large change (scale for short int)
          If m_vScroll.Value > m_vScroll.Max Then m_vScroll.Value = 0
          m_vScroll.Enabled = True                            'enable scroll
        End If
      Else
        If m_HasScroll Then
          m_vScroll.Value = 0                                 'set 0
          m_vScroll.Enabled = False                           'disable scrollbar
          m_Picture.Top = m_Picture.Container.ScaleTop        'not greater, set min height
        End If
      End If
    Else
      If m_HasScroll Then
          m_vScroll.Value = 0
          m_vScroll.Enabled = False                           'disable scrollbar
        m_Picture.Top = m_Picture.Container.ScaleTop
      End If
    End If
  End With
End Sub

'*******************************************************************************
' Subroutine Name   : Resize
' Purpose           : Handle resizing within the container control (form or pic)
'
' This should be invoked from your form's Resize event
'*******************************************************************************
Public Sub Resize()
  If Not m_HasPic Then Exit Sub         'cannot do if picture not defined

  If m_HasScroll Then
    With m_vScroll
      .Left = .Container.ScaleWidth - .Width
      .Height = .Container.ScaleHeight
      m_Picture.Height = .Height
      m_Picture.Width = .Left
    End With
  Else
    With m_Picture
      .Height = .Container.ScaleHeight
      .Width = .Container.ScaleWidth
    End With
  End If
  Call Repaint
End Sub

'*******************************************************************************
' Subroutine Name   : Scroll
' Purpose           : Handle scroll events
'
' This should be invoked from the scrollbar's Change and Scroll events
'*******************************************************************************
Public Sub Scroll()
  If m_vScroll Is Nothing Then Exit Sub
  m_Picture.Top = CLng(m_vScroll.Value) * -ScrollScale  'scrollbars are limited to short integers
End Sub

'*******************************************************************************
' Subroutine Name   : InitText
' Purpose           : Start new text
'
' Invoke this when you want to add completely new text to the display
'*******************************************************************************
Public Sub InitText(Optional Text As String = vbNullString)
  If Not m_HasPic Then Exit Sub                         'cannot do if picture not defined
  
  If Len(Text) <> 0 Then                                'much faster than checking for vbNullString
    m_Text = CompressTextFormat(UnFixFormat(Text))      'erase old text (add optional preformatted)
  Else
    m_Text = vbNullString
  End If
  LastFontName = vbNullString                           'reset last font
  If m_HasScroll Then                                   'if we have a scroll bar
    m_vScroll.Value = 0                                 'set minimum index
    m_vScroll.Enabled = False                           'disable it for now
  End If
  m_Indent = 0                                          'reset indents
  m_HangingIndent = 0
  Width = 0                                             'init width accumulator
  m_Picture.Top = m_Picture.Container.ScaleTop          'resize picture to minimum
  Call Repaint                                          'refresh display
End Sub

'*******************************************************************************
' Subroutine Name   : AddText
' Purpose           : Add some text to the data
'
' Rather than entering a ColorSonstant for the color (if you do not want Black), you
' can provide a Long variable that is set to a color (&HFF0000 = Blue), or provide
' an RGB() as data, such as RGB(128,255,128).
'*******************************************************************************
Public Sub AddText(Text As String, _
                   Optional FntName As String = "Times New Roman", _
                   Optional FntSize As Long = 10, _
                   Optional FntBold As Boolean = False, _
                   Optional FntItal As Boolean = False, _
                   Optional FntUnd As Boolean = False, _
                   Optional Color As ColorConstants = vbBlack)
  Dim S As String, Ary(5) As String
  Dim i As Long, j As Long
  
  If Not m_HasPic Then Exit Sub                         'cannot do if picture not defined
  S = Text                                              'grab text
  i = InStr(1, S, vbCrLf)                               'find terminators
  Do While i <> 0
    S = Left$(S, i - 1) & "<BR>" & Mid$(S, i + 2)       'convert CRLF to <BR>'
    i = InStr(i + 4, S, vbCrLf)
  Loop
'
' add a <TBR> to make text retreival beyond simply one line line of text with out "real"
' line breaks.  This is not neccessary, but makes later manual editing less of a headache
' if you save the result to a file for later fast-loading.
'
  If Right$(S, 4) = "<BR>" Then S = S & "<TBR>"
'
' if new font info matches current, do not bother keeping the header
'
  If StrComp(LastFontName, FntName, vbTextCompare) <> 0 Or _
     LastFontSize <> FntSize Or _
     LastFontBold <> FntBold Or _
     LastFontItal <> FntItal Or _
     LastFontUnd <> FntUnd Or _
     LastFontColor <> Color Then                        'if any do not match...
    LastFontName = FntName                              'save new last format
    LastFontSize = FntSize
    LastFontBold = FntBold
    LastFontItal = FntItal
    LastFontUnd = FntUnd
    LastFontColor = Color
     
    Ary(0) = FntName                                      'set font name
    Ary(1) = CStr(FntSize)                                'set point size
    If FntBold Then Ary(2) = "B"                          'bold flag
    If FntItal Then Ary(3) = "I"                          'italics flag
    If FntUnd Then Ary(4) = "U"                           'underlining
    Ary(5) = Hex$(Color)                                  'color
    m_Text = m_Text & "<Font " & Join(Ary, ",") & ">" & S 'append data
  Else
    m_Text = m_Text & S    'font info is the same as last, so just append data
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : DrawText
' Purpose           : Draw text on the main picture
'*******************************************************************************
Private Sub DrawText()
  Dim i As Long
  Dim S As String
  
  If Len(m_Text) = 0 Then Exit Sub          'nothing to do
  S = m_Text                                'else grab the text
  
  m_Indent = 0                              'reset indents
  m_HangingIndent = 0
  
  i = InStr(1, S, "<BR>", vbTextCompare)    'find a line terminator
  Do While i <> 0
    DrawLine Left$(S, i - 1)                'draw the line
    S = Mid$(S, i + 4)                      'trim printed stuff
    i = InStr(1, S, "<BR>", vbTextCompare)  'Find Next terminator
  Loop
  If Len(S) <> 0 Then DrawLine S            'print any remainders
End Sub

'*******************************************************************************
' Subroutine Name   : DrawLine
' Purpose           : Draw a line of text
'*******************************************************************************
Private Sub DrawLine(Text As String)
  Dim i As Long, j As Long
  Dim S As String, Ary() As String, T As String
  
  S = Text                                              'grab the text (allow blank text
  
  If Len(S) <> 0 Then
    i = InStr(1, S, "<Font ", vbTextCompare)              'find font information
    Do While i <> 0
      If i > 1 Then                                       'print previous data
        PaintLine Left$(S, i - 1)
      End If
      j = InStr(i + 1, S, ">")                            'find end of definitions
      T = Mid$(S, i + 6, j - i - 6)                       'grab definition text
      Ary = Split(T, ",")                                 'break up
      S = Mid$(S, j + 1)                                  'trim off font block
      With m_Picture
        With .Font
          .Name = Ary(0)                                    'set name of font
          .SIZE = CLng(Ary(1))                              'point size
          .Bold = Ary(2) <> vbNullString                    'bolding
          .Italic = Ary(3) <> vbNullString                  'italics
          .Underline = Ary(4) <> vbNullString               'underline
        End With
        .ForeColor = CLng("&H" & Ary(5))                    'text color
      End With
      i = InStr(1, S, "<Font ", vbTextCompare)              'find next font block
    Loop
  End If
'
' no more font block, process remainders
'
  If Len(S) <> 0 Then
    PaintLine S                                         'out text with newline
  End If
  m_Picture.Print                                       'nothing, so just newline
  Width = 0
  m_Picture.CurrentX = m_Indent                         'reset indent (if any)
End Sub

'*******************************************************************************
' Subroutine Name   : PaintLine
' Purpose           : Support line wrap
'*******************************************************************************
Private Sub PaintLine(Text As String)
  Dim Wd As Long, i As Long, j As Long
  Dim S As String, T As String
  Dim Sz As SIZE
'
' get a local copy of data
'
  S = Text
'
' strip temp break data
'
  If Len(S) = 0 Then Exit Sub                               'nothing to do
  i = InStr(1, S, "<TBR>", vbTextCompare)
  Do While i <> 0
    S = Left$(S, i - 1) & Mid$(S, i + 5)
    i = InStr(i, S, "<TBR>", vbTextCompare)
  Loop
'
' grab indenting info
'
  If Len(S) = 0 Then Exit Sub                               'nothing left to do
  i = InStr(1, S, "<IND ", vbTextCompare)
  Do While i <> 0
    j = InStr(i + 1, S, ">")
    T = Mid$(S, i + 5, j - i - 5)
    S = Left$(S, i - 1) & Mid$(S, j + 1)
    m_Indent = CLng(T)
    Width = m_Indent
    m_Picture.CurrentX = m_Indent
    i = InStr(1, S, "<IND ", vbTextCompare)
  Loop                                                      'loop for redundancies
'
' grab hanging indenting info
'
  If Len(S) = 0 Then Exit Sub                               'nothing left to do
  i = InStr(1, S, "<HIND ", vbTextCompare)
  Do While i <> 0
    j = InStr(i + 1, S, ">")
    T = Mid$(S, i + 6, j - i - 6)
    S = Left$(S, i - 1) & Mid$(S, j + 1)
    m_HangingIndent = CLng(T)
    i = InStr(1, S, "<HIND ", vbTextCompare)
  Loop                                                      'loop for redundancies
'
' now process actual text
'
  If Len(S) = 0 Then Exit Sub                               'nothing left to do
  Do
    T = vbNullString                                        'init bleed-over accumulator
    Call GetTextExtentPoint(m_Picture.hdc, S, Len(S), Sz)   'get size of string
    Wd = Width + Sz.cx * Tpp                                'compute width (pixels to twips)
    If Wd + 240 < m_Picture.ScaleWidth Then Exit Do         'we are still within the line
    Do While Wd + 240 > m_Picture.ScaleWidth                'else pull back until we are
      i = InStrRev(S, " ")                                  'find a space
      If i = 0 Then
        T = S & T
        S = vbNullString
        Exit Do
      Else
        If Len(T) <> 0 Then                                   'if T already has data
          T = Mid$(S, i) & T                                  'add text with a space sep
        Else
          T = Mid$(S, i)                                      'else no leading space
        End If
        S = Left$(S, i - 1)                                   'set new test string
        Call GetTextExtentPoint(m_Picture.hdc, S, Len(S), Sz) 'get size of string
        Wd = Width + Sz.cx * Tpp                              'grab new test width
      End If
    Loop
    m_Picture.Print S                                       'send out data
    S = LTrim$(T)                                           'grab remainder
    Width = m_HangingIndent                                 'init indenting
    m_Picture.CurrentX = Width                              'new start postion
  Loop
  If Len(S) <> 0 Then                                       'if data remains
    m_Picture.Print S;                                      'send it out
  End If
  Width = Wd                                                'width accumulator set
End Sub

'*******************************************************************************
' Subroutine Name   : TileBackground
' Purpose           : Tile the background image
'
' I revised this routine from another of my modules that tiled a picture to a form
' or a picture.
'*******************************************************************************
Private Sub TileBackground()
  Dim i As Long, j As Long
  
  If Not m_HasPic Then Exit Sub                         'cannot do if picture not defined
  m_Picture.Cls
  m_Picture.CurrentX = m_Indent                         'init with indent (if any)
  LastFontName = vbNullString
  Width = m_Indent
  If Not m_HasTile Then Exit Sub                        'cannot do if picture not defined
  With m_PicTile                                        'image to draw source
    If m_DoTiling Then
      For i = 0 To m_Picture.ScaleWidth Step .Width       'draw across top of m_Picture
        For j = 0 To m_Picture.ScaleHeight Step .Height   'draw across height of m_Picture
          m_Picture.PaintPicture .Picture, i, j           'draw a frame from m_PicTile
        Next j
      Next i
    Else
      m_Picture.PaintPicture .Picture, 0, 0               'draw a frame from m_PicTile
    End If
  End With
End Sub

'*******************************************************************************
' Function Name     : FixFormat
' Purpose           : Covert preformatted local data by changing any <TBR> codes
'                   : to CrLf codes
'*******************************************************************************
Private Function FixFormat(Text As String) As String
  Dim Ary() As String
  
  Ary = Split(Text, "<TBR>")    'replace temp breaks with vbCRLF
  FixFormat = Join(Ary, vbCrLf)
End Function

'*******************************************************************************
' Function Name     : UnFixFormat
' Purpose           : Covert preformatted input data by changing any CrLf codes
'                   : to <TBR> Temp BReaks
'*******************************************************************************
Private Function UnFixFormat(Text As String) As String
  Dim Ary() As String
  
  Ary = Split(Text, vbCrLf)
  UnFixFormat = Join(Ary, "<TBR>")        'change vbCrLf to Temp breaks
End Function

'*******************************************************************************
' Function Name     : CompressTextFormat
' Purpose           : Compress the font data by removing subsequent
'                   : duplicate definitions
'*******************************************************************************
Private Function CompressTextFormat(Text As String) As String
  Dim sIn As String, sOut As String, T As String, Ary() As String
  Dim LstNm As String, LstSiz As String
  Dim i As Long, j As Long, LstCol As Long
  Dim LstB As Boolean, LstI As Boolean, LstU As Boolean
  
  If Len(Text) = 0 Then Exit Function
  sIn = Text                                            'grab text to process
  i = InStr(1, sIn, "<Font ", vbTextCompare)            'find font information
  sOut = vbNullString
  Do While i <> 0
    If i > 1 Then                                       'save previous data
      sOut = sOut & Left$(sIn, i - 1)
    End If
    j = InStr(i + 1, sIn, ">")                          'find end of definitions
    T = Mid$(sIn, i + 6, j - i - 6)                     'grab definition text
    Ary = Split(T, ",")                                 'break up
    sIn = Mid$(sIn, j + 1)                              'trim off font block
    If StrComp(Ary(0), LstNm) <> 0 Or _
       LstSiz <> Ary(1) Or _
       LstB <> CBool(Len(Ary(2))) Or _
       LstI <> CBool(Len(Ary(3))) Or _
       LstU <> CBool(Len(Ary(4))) Or _
       LstCol <> CLng("&H" & Ary(5)) Then
      LstNm = Ary(0)                                    'no match, so update
      LstSiz = Ary(1)
      LstB = CBool(Len(Ary(2)))
      LstI = CBool(Len(Ary(3)))
      LstU = CBool(Len(Ary(4)))
      LstCol = CLng("&H" & Ary(5))
      sOut = sOut & "<Font " & T & ">"                  'and keep format
    End If
    i = InStr(1, sIn, "<Font ", vbTextCompare)          'find next font block
  Loop
  If Len(sIn) <> 0 Then sOut = sOut & sIn               'add any remainder
  CompressTextFormat = sOut                             'return finalized string
End Function

