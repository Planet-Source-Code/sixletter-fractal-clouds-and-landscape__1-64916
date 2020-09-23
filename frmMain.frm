VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Landscape"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   319
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frm_Cloud 
      Height          =   6495
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4335
      Begin VB.CommandButton btn_Save 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   1095
         Left            =   240
         Picture         =   "frmMain.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5160
         Width           =   1815
      End
      Begin VB.PictureBox pic_Back2 
         Height          =   375
         Left            =   2280
         ScaleHeight     =   315
         ScaleWidth      =   1755
         TabIndex        =   10
         Top             =   720
         Width           =   1815
         Begin VB.PictureBox pic_Front2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   495
            TabIndex        =   12
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox pic_Back1 
         Height          =   375
         Left            =   2280
         ScaleHeight     =   315
         ScaleWidth      =   1755
         TabIndex        =   9
         Top             =   240
         Width           =   1815
         Begin VB.PictureBox pic_Front1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   600
            ScaleHeight     =   135
            ScaleWidth      =   495
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.PictureBox pic_Back 
         BackColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   240
         ScaleHeight     =   253
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   253
         TabIndex        =   7
         Top             =   1200
         Width           =   3855
         Begin VB.PictureBox pic_Frac 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   240
            ScaleHeight     =   21
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   21
            TabIndex        =   8
            Top             =   240
            Width           =   315
         End
      End
      Begin VB.ComboBox cmb_Style 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton btn_Go 
         Caption         =   "Create"
         Height          =   1095
         Left            =   2160
         Picture         =   "frmMain.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   5160
         Width           =   1935
      End
      Begin VB.ComboBox cmb_Size 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Style:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbl_Size 
         Caption         =   "Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog dlg_Save 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'this program starts with a square and iteratively
'subdivides it into the smallest square possible.
'As it subdivides it assigns colors to each point.
'this program can create random clouds and random
'landscapes using the same algorithm
Option Explicit
Private Type fPoint
    C1            As Single    'color of point
    OK            As Boolean   'has it already been colored
End Type
Private TP()  As fPoint
Private Type fSquare
    x1            As Long      'top left
    y1            As Long      'top left
    x2            As Long      'top right
    y2            As Long      'top right
    x3            As Long      'bottom right
    y3            As Long      'bottom right
    x4            As Long      'bottom left
    y4            As Long      'bottom left
    OK            As Boolean   'has it already been subdivided
    Level         As Single    'level of subdivision
End Type
Private SQ()  As fSquare
Private MaxH  As Single
Private MinH  As Single
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, _
                                               ByVal x As Long, _
                                               ByVal Y As Long, _
                                               ByVal crColor As Long) As Long

Private Sub Form_Load()
    Me.pic_Frac.Move 0, 0, 0, 0
    Fill_Combo
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End
End Sub

Private Sub btn_Go_Click()
    Me.btn_Go.Enabled = False
    Me.btn_Save.Enabled = False
    ReDim TP(0)
    ReDim SQ(0)
    Me.pic_Frac.Cls
    Create_Initial_Square
    SubDivide
    Draw_Fractal
    ReDim TP(0)
    ReDim SQ(0)
    Me.btn_Go.Enabled = True
    Me.btn_Save.Enabled = True
End Sub

Private Sub btn_Save_Click()
    Dim The_Path       As String
    Dim EnvString      As String
    Dim Indx           As Integer
    Dim FileInQuestion As String
    Dim Reply          As VbMsgBoxResult
    Dim SFile          As String
    Indx = 1
    The_Path = vbNullString
    The_Path = GetSetting("LandScape", "SaveAs", "SaveAs")
    If The_Path = vbNullString Then
        Do
            EnvString = Environ$(Indx)   ' Get environment
            If UCase$(Left$(EnvString, 12)) = "USERPROFILE=" Then   ' Check PATH entry.
                The_Path = Mid$(EnvString, 13, Len(EnvString) - 12)
                If Right$(The_Path, 1) = "\" Then
                    The_Path = The_Path & "My Documents"
                Else
                    The_Path = The_Path & "\My Documents"
                End If
                Exit Do
            Else
                Indx = Indx + 1
            End If
        Loop Until LenB(EnvString) = 0
        If LenB(The_Path) = 0 Then
            The_Path = App.Path
        End If
    End If
    With Me.dlg_Save
        .InitDir = The_Path
        .DialogTitle = "Save Picture"
        .CancelError = True
        .FileName = CStr(Me.cmb_Style.List(Me.cmb_Style.ListIndex) & " - (" & Format$(Date, "mm-dd-yyyy") & ")")
        .Filter = "Bitmap (*.bmp)|*.bmp"
        On Error Resume Next
        .ShowSave
        If Err.Number <> MSComDlg.cdlCancel Then
            The_Path = Mid$(.FileName, 1, Len(.FileName) - Len(.FileTitle))
            SaveSetting "LandScape", "SaveAs", "SaveAs", The_Path
            SFile = .FileName
            FileInQuestion = Dir(SFile)
            If LenB(FileInQuestion) Then
                Reply = MsgBox(FileInQuestion & " Already Exists." & vbLf & Chr$(10) & "Do You Want To Overwrite?", vbQuestion + vbYesNo, "Overwrite")
                If Reply = vbYes Then
                    Screen.MousePointer = vbHourglass
                    SavePicture Me.pic_Frac.Image, SFile
                    Screen.MousePointer = vbDefault
                End If
            Else
                Screen.MousePointer = vbHourglass
                SavePicture Me.pic_Frac.Image, SFile
                Screen.MousePointer = vbDefault
            End If
        End If
    End With
    On Error GoTo 0
End Sub

Private Sub cmb_Size_Click()
    With Me
        .btn_Save.Enabled = False
        .pic_Frac.Move .pic_Frac.Left, .pic_Frac.Top, CLng(.cmb_Size.List(.cmb_Size.ListIndex)), CLng(.cmb_Size.List(.cmb_Size.ListIndex))
        .pic_Frac.Cls
    End With 'Me
End Sub

Private Sub cmb_Style_Click()
    Me.pic_Frac.Cls
    Me.btn_Save.Enabled = False
End Sub

Private Sub Fill_Combo()
    'to make it easy, square sizes should be
    'of the order (2^n)+1
    With Me.cmb_Size
        .Clear
        .AddItem 5 '(2^2)+1
        .AddItem 9 '(2^3)+1
        .AddItem 17 '(2^4)+1
        .AddItem 33 '(2^5)+1
        .AddItem 65 '(2^6)+1
        .AddItem 129 '(2^7)+1
        .AddItem 257 '(2^8)+1
        .ListIndex = 6
    End With
    With Me.cmb_Style
        .Clear
        .AddItem "Cloud"
        .AddItem "Land"
        .ListIndex = 0
    End With
End Sub

Private Sub Create_Initial_Square()
    
    'set up initial square as the corners of the
    'picture.  Assign random color values to these corners
    'and mark those points as being set (OK=True).
    'I set the square as being not set (OK=False) becuase it
    'has not been subdivided yet.
    'I redim the points with all possible coordinates
    ReDim TP(Me.pic_Frac.Width - 1, Me.pic_Frac.Height - 1)
    ReDim SQ(1)
    With SQ(1)
    'square
        .x1 = 0
        .y1 = 0
        .x2 = Me.pic_Frac.Width - 1
        .y2 = 0
        .x3 = Me.pic_Frac.Width - 1
        .y3 = Me.pic_Frac.Height - 1
        .x4 = 0
        .y4 = Me.pic_Frac.Height - 1
        .OK = False
        .Level = 1
    'point
        MaxH = 0
        MinH = 0
        Randomize
        TP(.x1, .y1).OK = True
        TP(.x1, .y1).C1 = (2 * Rnd) - 1
        TP(.x2, .y2).OK = True
        TP(.x2, .y2).C1 = (2 * Rnd) - 1
        TP(.x3, .y3).OK = True
        TP(.x3, .y3).C1 = (2 * Rnd) - 1
        TP(.x4, .y4).OK = True
        TP(.x4, .y4).C1 = (2 * Rnd) - 1
        If TP(.x1, .y1).C1 > MaxH Then MaxH = TP(.x1, .y1).C1
        If TP(.x2, .y2).C1 > MaxH Then MaxH = TP(.x2, .y2).C1
        If TP(.x3, .y3).C1 > MaxH Then MaxH = TP(.x3, .y3).C1
        If TP(.x4, .y4).C1 > MaxH Then MaxH = TP(.x4, .y4).C1
        If TP(.x1, .y1).C1 < MinH Then MinH = TP(.x1, .y1).C1
        If TP(.x2, .y2).C1 < MinH Then MinH = TP(.x2, .y2).C1
        If TP(.x3, .y3).C1 < MinH Then MinH = TP(.x3, .y3).C1
        If TP(.x4, .y4).C1 < MinH Then MinH = TP(.x4, .y4).C1
    End With
End Sub

Private Sub SubDivide()
    Dim Its     As Long
    Dim ECycles As Long
    Dim SCycles As Long
    Dim A       As Long
    Dim B       As Long
    Dim C       As Long
    Dim FSize   As Long

    'I can determine how many Iterations I need to
    'do based on the size of the square.
    FSize = CLng(Me.cmb_Size.List(Me.cmb_Size.ListIndex)) - 1
    Its = Log(FSize) / Log(2#)
    'I cycle through all squares each iteration
    'If it has not been subdivided and it is bigger than
    'one pixel in size, I continue on with subdivision.
    'the point colors are determined by averaging points the corner colors
    'of the parent square and then adding a ranom value based on the
    'subdivision level.
    'this can probably be cleaned up.
    With Me
        .pic_Back1.Scale (0, 0)-(Its, 100)
        .pic_Front1.Move 0, 0, 0, 100
        .pic_Front1.Refresh
        .pic_Front1.Visible = True
    End With
    SCycles = 1
    For A = 1 To Its
        With Me
            .pic_Back2.Scale (0, 0)-(UBound(SQ), 100)
            .pic_Front2.Move 0, 0, 0, 100
            .pic_Front2.Refresh
            .pic_Front2.Visible = True
        End With 'Me
        ECycles = (4 ^ A - 1) / 3
        For B = SCycles To ECycles
            Randomize
            'split into 4 smaller squares
            ReDim Preserve SQ(UBound(SQ) + 4)
            For C = 1 To 4
                Select Case C
                Case 1 'top left square
                    With SQ(UBound(SQ) - 3)
                        .x1 = SQ(B).x1
                        .y1 = SQ(B).y1
                        .x2 = (SQ(B).x2 + SQ(B).x1) / 2
                        .y2 = SQ(B).y2
                        .x3 = (SQ(B).x3 + SQ(B).x1) / 2
                        .y3 = (SQ(B).y3 + SQ(B).y1) / 2
                        .x4 = SQ(B).x4
                        .y4 = (SQ(B).y4 + SQ(B).y1) / 2
                        .Level = SQ(B).Level / 2
                        'top right point
                    End With
                    With SQ(UBound(SQ) - 3)
                        If TP(.x2, .y2).OK = False Then
                            TP(.x2, .y2).C1 = (TP(SQ(B).x1, SQ(B).y1).C1 + TP(SQ(B).x2, SQ(B).y2).C1) / 2 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x2, .y2).OK = True
                            If TP(.x2, .y2).C1 > MaxH Then
                                MaxH = TP(.x2, .y2).C1
                            End If
                            If TP(.x2, .y2).C1 < MinH Then
                                MinH = TP(.x2, .y2).C1
                            End If
                        End If
                        'bottom right point
                        If TP(.x3, .y3).OK = False Then
                            TP(.x3, .y3).C1 = (TP(SQ(B).x1, SQ(B).y1).C1 + TP(SQ(B).x2, SQ(B).y2).C1 + TP(SQ(B).x3, SQ(B).y3).C1 + TP(SQ(B).x4, SQ(B).y4).C1) / 4 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x3, .y3).OK = True
                            If TP(.x3, .y3).C1 > MaxH Then
                                MaxH = TP(.x3, .y3).C1
                            End If
                            If TP(.x3, .y3).C1 < MinH Then
                                MinH = TP(.x3, .y3).C1
                            End If
                        End If
                        'bottom left point
                        If TP(.x4, .y4).OK = False Then
                            TP(.x4, .y4).C1 = (TP(SQ(B).x1, SQ(B).y1).C1 + TP(SQ(B).x4, SQ(B).y4).C1) / 2 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x4, .y4).OK = True
                            If TP(.x4, .y4).C1 > MaxH Then
                                MaxH = TP(.x4, .y4).C1
                            End If
                            If TP(.x4, .y4).C1 < MinH Then
                                MinH = TP(.x4, .y4).C1
                            End If
                        End If
                    End With
                Case 2 'top right square
                    With SQ(UBound(SQ) - 2)
                        .x1 = (SQ(B).x2 + SQ(B).x1) / 2
                        .y1 = SQ(B).y1
                        .x2 = SQ(B).x2
                        .y2 = SQ(B).y2
                        .x3 = SQ(B).x3
                        .y3 = (SQ(B).y3 + SQ(B).y1) / 2
                        .x4 = (SQ(B).x3 + SQ(B).x1) / 2
                        .y4 = (SQ(B).y4 + SQ(B).y1) / 2
                        .Level = SQ(B).Level / 2
                    End With
                    'top left point
                    With SQ(UBound(SQ) - 2)
                        If TP(.x1, .y1).OK = False Then
                            TP(.x1, .y1).C1 = (TP(SQ(B).x1, SQ(B).y1).C1 + TP(SQ(B).x2, SQ(B).y2).C1) / 2 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x1, .y1).OK = True
                            If TP(.x1, .y1).C1 > MaxH Then
                                MaxH = TP(.x1, .y1).C1
                            End If
                            If TP(.x1, .y1).C1 < MinH Then
                                MinH = TP(.x1, .y1).C1
                            End If
                        End If
                        'bottom right point
                        If TP(.x3, .y3).OK = False Then
                            TP(.x3, .y3).C1 = (TP(SQ(B).x2, SQ(B).y2).C1 + TP(SQ(B).x3, SQ(B).y3).C1) / 2 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x3, .y3).OK = True
                            If TP(.x3, .y3).C1 > MaxH Then
                                MaxH = TP(.x3, .y3).C1
                            End If
                            If TP(.x3, .y3).C1 < MinH Then
                                MinH = TP(.x3, .y3).C1
                            End If
                        End If
                        'bottom left point
                        If TP(.x4, .y4).OK = False Then
                            TP(.x4, .y4).C1 = (TP(SQ(B).x1, SQ(B).y1).C1 + TP(SQ(B).x2, SQ(B).y2).C1 + TP(SQ(B).x3, SQ(B).y3).C1 + TP(SQ(B).x4, SQ(B).y4).C1) / 4 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x4, .y4).OK = True
                            If TP(.x4, .y4).C1 > MaxH Then
                                MaxH = TP(.x4, .y4).C1
                            End If
                            If TP(.x4, .y4).C1 < MinH Then
                                MinH = TP(.x4, .y4).C1
                            End If
                        End If
                    End With
                Case 3 'bottom right square
                    With SQ(UBound(SQ) - 1)
                        .x1 = (SQ(B).x2 + SQ(B).x1) / 2
                        .y1 = (SQ(B).y3 + SQ(B).y1) / 2
                        .x2 = SQ(B).x2
                        .y2 = (SQ(B).y3 + SQ(B).y1) / 2
                        .x3 = SQ(B).x3
                        .y3 = SQ(B).y3
                        .x4 = (SQ(B).x3 + SQ(B).x1) / 2
                        .y4 = SQ(B).y4
                        .Level = SQ(B).Level / 2
        
                    End With
                    'top left point
                    With SQ(UBound(SQ) - 1)
                        If TP(.x1, .y1).OK = False Then
                            TP(.x1, .y1).C1 = (TP(SQ(B).x1, SQ(B).y1).C1 + TP(SQ(B).x2, SQ(B).y2).C1 + TP(SQ(B).x3, SQ(B).y3).C1 + TP(SQ(B).x4, SQ(B).y4).C1) / 4 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x1, .y1).OK = True
                            If TP(.x1, .y1).C1 > MaxH Then
                                MaxH = TP(.x1, .y1).C1
                            End If
                            If TP(.x1, .y1).C1 < MinH Then
                                MinH = TP(.x1, .y1).C1
                            End If
                        End If
                        'top right point
                        If TP(.x2, .y2).OK = False Then
                            TP(.x2, .y2).C1 = (TP(SQ(B).x2, SQ(B).y2).C1 + TP(SQ(B).x3, SQ(B).y3).C1) / 2 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x2, .y2).OK = True
                            If TP(.x2, .y2).C1 > MaxH Then
                                MaxH = TP(.x2, .y2).C1
                            End If
                            If TP(.x2, .y2).C1 < MinH Then
                                MinH = TP(.x2, .y2).C1
                            End If
                        End If
                        'bottom left point
                        If TP(.x4, .y4).OK = False Then
                            TP(.x4, .y4).C1 = (TP(SQ(B).x3, SQ(B).y3).C1 + TP(SQ(B).x4, SQ(B).y4).C1) / 2 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x4, .y4).OK = True
                            If TP(.x4, .y4).C1 > MaxH Then
                                MaxH = TP(.x4, .y4).C1
                            End If
                            If TP(.x4, .y4).C1 < MinH Then
                                MinH = TP(.x4, .y4).C1
                            End If
                        End If
                    End With
                Case 4 'bottom left square
                    With SQ(UBound(SQ))
                        .x1 = SQ(B).x1
                        .y1 = (SQ(B).y4 + SQ(B).y1) / 2
                        .x2 = (SQ(B).x3 + SQ(B).x1) / 2
                        .y2 = (SQ(B).y3 + SQ(B).y1) / 2
                        .x3 = (SQ(B).x3 + SQ(B).x1) / 2
                        .y3 = SQ(B).y4
                        .x4 = SQ(B).x4
                        .y4 = SQ(B).y4
                        .Level = SQ(B).Level / 2

                    End With
                    'top left point
                    With SQ(UBound(SQ))
                        If TP(.x1, .y1).OK = False Then
                            TP(.x1, .y1).C1 = (TP(SQ(B).x4, SQ(B).y4).C1 + TP(SQ(B).x1, SQ(B).y1).C1) / 2 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x1, .y1).OK = True
                            If TP(.x1, .y1).C1 > MaxH Then
                                MaxH = TP(.x1, .y1).C1
                            End If
                            If TP(.x1, .y1).C1 < MinH Then
                                MinH = TP(.x1, .y1).C1
                            End If
                        End If
                        'top right point
                        If TP(.x2, .y2).OK = False Then
                            TP(.x2, .y2).C1 = (TP(SQ(B).x1, SQ(B).y1).C1 + TP(SQ(B).x2, SQ(B).y2).C1 + TP(SQ(B).x3, SQ(B).y3).C1 + TP(SQ(B).x4, SQ(B).y4).C1) / 4 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x2, .y2).OK = True
                            If TP(.x2, .y2).C1 > MaxH Then
                                MaxH = TP(.x2, .y2).C1
                            End If
                            If TP(.x2, .y2).C1 < MinH Then
                                MinH = TP(.x2, .y2).C1
                            End If
                        End If
                        'bottom right point
                        If TP(.x3, .y3).OK = False Then
                            TP(.x3, .y3).C1 = (TP(SQ(B).x3, SQ(B).y3).C1 + TP(SQ(B).x4, SQ(B).y4).C1) / 2 + (2 * SQ(B).Level * Rnd) - SQ(B).Level
                            TP(.x3, .y3).OK = True
                            If TP(.x3, .y3).C1 > MaxH Then
                                MaxH = TP(.x3, .y3).C1
                            End If
                            If TP(.x3, .y3).C1 < MinH Then
                                MinH = TP(.x3, .y3).C1
                            End If
                        End If
                    End With
                End Select
            Next C
            Me.pic_Front2.Move 0, 0, B, 100
            Me.pic_Front2.Refresh
            DoEvents
        Next B
        Me.pic_Front1.Move 0, 0, A, 100
        Me.pic_Front1.Refresh
        SCycles = ECycles + 1
    Next A
    Me.pic_Front1.Visible = False
    Me.pic_Front2.Visible = False
End Sub

Private Sub Draw_Fractal()
    Dim A      As Long
    Dim B      As Long
    Dim CColor As Long
    Dim CPerc  As Single
    Select Case Me.cmb_Style.ListIndex
    Case 0 'Cloud
        For A = 0 To Me.pic_Frac.Height - 1
            For B = 0 To Me.pic_Frac.Width - 1
                'scale to maxH and minH to create shade of blue
                CPerc = (MaxH - TP(B, A).C1) / (MaxH - MinH)
                CColor = CPerc * 255
                SetPixel Me.pic_Frac.hdc, B, A, RGB(CColor, CColor, 255)
            Next B
        Next A
    Case 1 'land
        For A = 0 To Me.pic_Frac.Height - 1
            For B = 0 To Me.pic_Frac.Width - 1
                CPerc = (MaxH - TP(B, A).C1) / (MaxH - MinH)
                If CPerc <= 0.33 Then
                    CPerc = CPerc / 0.33
                    CColor = RGB(0, 0, CPerc * 255) 'blue
                ElseIf CPerc <= 0.75 Then
                    CPerc = CPerc / 0.75
                    CColor = RGB(0, CPerc * 255, 0) 'green shade
                ElseIf CPerc <= 0.95 Then
                    CPerc = CPerc / 0.95
                    CColor = RGB(CPerc * 139, CPerc * 69, CPerc * 19) 'brown shade
                Else
                    CColor = RGB(CPerc * 255, CPerc * 255, CPerc * 255) 'snow
                End If
                SetPixel Me.pic_Frac.hdc, B, A, CColor
            Next B
        Next A
    End Select
    Me.pic_Frac.Refresh
End Sub

