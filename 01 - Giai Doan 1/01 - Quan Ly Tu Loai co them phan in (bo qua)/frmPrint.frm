VERSION 5.00
Object = "{1BCC7098-34C1-4749-B1A3-6C109878B38F}#1.0#0"; "vspdf8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form frmPrint 
   Caption         =   "Danh sach tu loai"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6210
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VSPrinter8LibCtl.VSPrinter VSP 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _cx             =   9763
      _cy             =   3836
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "In..."
      AbortTextButton =   "Dong"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   8.61742424242424
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
      AutoLinkNavigate=   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VSPDF8LibCtl.VSPDF8 VSPDF81 
      Left            =   480
      Top             =   2520
      Author          =   "Le Viet Man"
      Creator         =   ""
      Title           =   ""
      Subject         =   ""
      Keywords        =   ""
      Compress        =   3
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Show
    VSP.Height = Me.ScaleHeight
    VSP.Width = Me.ScaleWidth
    
    With VSP
        .Zoom = 100
        .PaperSize = pprLetter
        .Orientation = 0
        .MarginHeader = "1cm"
        .MarginFooter = "1cm"
        .MarginTop = "2cm"
        .MarginBottom = "2cm"
        .MarginLeft = "3.5cm"
        .MarginRight = "2.5cm"
        .Visible = True
    End With
    
    With VSP
        .Clear
        .StartDoc
        
        'Tao HeaderPage
        .HdrFontName = ".VnTime"
        .HdrFontSize = 13
        .Header = "Ch­¬ng tr×nh b¾t lçi chÝnh t¶||"
        
        'Tao FooterPage
        .Footer = Format(Now, "dd/mm/yyyy") & "||Trang %d"
        
        'Tao TitlePage
        .FontName = ".VntimeH"
        .FontSize = 20
        .FontBold = True
        .TextAlign = taCenterMiddle
        .Text = "danh s¸ch tõ lo¹i" & vbCrLf
        .TextAlign = taLeftMiddle
    
        .FontSize = 13
        
        'Tao Table
        .StartTable
        
        .FontBold = False
        .FontName = ".VnTime"
        .TableCell(tcCols) = 3
        .TableCell(tcRows) = 1
        .TableCell(tcColWidth, 1, 1) = "1.5cm"
        .TableCell(tcColWidth, 1, 2) = "3cm"
        .TableCell(tcColWidth, 1, 3) = "11cm"
                
        ' dien du lieu vao phan HeaderTable
        .TableCell(tcAlign) = taCenterMiddle
        .TableCell(tcText, 1, 1) = "Stt"
        .TableCell(tcText, 1, 2) = "M· tõ lo¹i"
        .TableCell(tcText, 1, 3) = "ý nghÜa"
        .EndTable
        
        bPrintingTable = True
        
        Call LoadWordClass
        
        bPrintingTable = False
        .EndDoc
    End With
End Sub

Private Sub LoadWordClass()
Dim command As String
Dim schehiername As String, schehiercode As String
Dim tHeader As String, tContent As String
Dim fHeader As String, fContent As String
Dim i As Integer
Dim ccom As Integer
Dim cd As Integer
    fContent = "<1.5cm|<3cm|<11cm;"
    With VSP
        .CurrentX = .CurrentX + 20
        For i = 1 To WC.Count
            If i = 1 Then
                tContent = i & "|" & WC.Sign(i) & "|" & UniStringToTCVNString(WC.Sense(i)) & ";"
            Else
                tContent = tContent & i & "|" & WC.Sign(i) & "|" & UniStringToTCVNString(WC.Sense(i)) & ";"
            End If
        Next i
        
        .FontBold = False
        .Table = fContent & tContent
        .CurrentX = .CurrentX + 20
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call TeminateWindows("About ComponentOne VSPrint 8")
    Call TeminateWindows("About ComponentOne VSPDF 8")
End Sub
