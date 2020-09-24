VERSION 5.00
Begin VB.Form ShapeCalc 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Area Calculator - Advanced"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      BackColor       =   &H80000001&
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   4560
      Width           =   3015
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Area = PIE * Radius²"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   305
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.Frame Frame6 
         BackColor       =   &H80000001&
         Height          =   615
         Left            =   0
         TabIndex        =   18
         Top             =   5400
         Width           =   8655
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Area Equals to: 0 units"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   8415
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000001&
         Height          =   5175
         Left            =   3120
         TabIndex        =   13
         Top             =   120
         Width           =   5535
         Begin VB.Timer Timer1 
            Left            =   360
            Top             =   1080
         End
         Begin VB.PictureBox PicBox 
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            Height          =   4335
            Left            =   600
            Picture         =   "ShapeCalc.frx":0000
            ScaleHeight     =   4335
            ScaleWidth      =   4335
            TabIndex        =   14
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "2) This is the shape which you chose.. yeah.. the graphics are crap, but what are you gona do? That's right, i'm challenging you!"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   15
            Top             =   240
            Width           =   5055
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000001&
         Height          =   2295
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   3015
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "1) Select the shape that you would like to calculate the area of."
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   12
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Circle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   840
            Width           =   2535
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Square"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   2535
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rectangle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1560
            Width           =   2535
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Triangle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1920
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000001&
         Height          =   2055
         Left            =   0
         TabIndex        =   1
         Top             =   2400
         Width           =   3015
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   1080
            TabIndex        =   3
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1080
            TabIndex        =   2
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "3) Enter the required data in order to process the area!"
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Radius"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Line B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   255
            Left            =   240
            TabIndex        =   4
            Top             =   1440
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "ShapeCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Shape As String
Private Sub Form_Load()
 Timer1.Interval = 50
 Timer1.Enabled = True
 Shape = "circle"
End Sub
Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label7.ForeColor = &HFF00&
    Label8.ForeColor = &HFF00&
    Label9.ForeColor = &HFF00&
    Label10.ForeColor = &HFF00&
End Sub
Private Sub Label10_Click()
    Shape = "triangle"
    PicBox.Picture = LoadPicture(App.Path & "\triangle.jpg")
    Label4.Caption = "Height"
    Label5.ForeColor = &H80000001
    Text2.BorderStyle = 0
    Text2.Appearance = 0
    Text2.BackColor = &H80000001
    Label6.Caption = "Area = ½Base * Height"
    Label5.Caption = "Base"
    Text2.Enabled = True
    Text2.BackColor = &H80000005
    Text2.BorderStyle = 1
    Text2.Appearance = 1
    Label5.ForeColor = &H80000005
End Sub
Private Sub Label7_Click()
    Shape = "circle"
    PicBox.Picture = LoadPicture(App.Path & "\circle.jpg")
    Label4.Caption = "Radius"
    Label5.ForeColor = &H80000001
    Text2.Appearance = 0
    Text2.BorderStyle = 0
    Text2.BackColor = &H80000001
    Text2.Enabled = False
    Label6.Caption = "Area = PIE * Radius²"
End Sub
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label7.ForeColor = &HFFFF&
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label8.ForeColor = &HFFFF&
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label9.ForeColor = &HFFFF&
End Sub
Private Sub Label10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label10.ForeColor = &HFFFF&
End Sub

Private Sub Label8_Click()
    Shape = "square"
    PicBox.Picture = LoadPicture(App.Path & "\square.jpg")
    Label4.Caption = "Side X"
    Label5.ForeColor = &H80000001
    Text2.BorderStyle = 0
    Text2.Appearance = 0
    Text2.BackColor = &H80000001
    Label6.Caption = "Area = Side X²"
    Text2.Text = ""
End Sub
Private Sub Label9_Click()
    Shape = "rectangle"
    PicBox.Picture = LoadPicture(App.Path & "\rectangle.jpg")
    Label4.Caption = "Side X"
    Label5.Caption = "Side Y"
    Text2.Enabled = True
    Text2.BackColor = &H80000005
    Text2.BorderStyle = 1
    Text2.Appearance = 1
    Label5.ForeColor = &H80000005
End Sub

Private Sub Text1_Change()
 Static strSaved As String
 If Text1.Text <> "" And Text1.Text <> "-" And Text1.Text <> "." And Text1.Text <> "-." Then
    If Not IsNumeric(Text1.Text) Then Text1.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text1.Text
End Sub

Private Sub Text2_Change()
 Static strSaved As String
 If Text2.Text <> "" And Text2.Text <> "-" And Text2.Text <> "." And Text2.Text <> "-." Then
    If Not IsNumeric(Text2.Text) Then Text2.Text = strSaved: SendKeys "{END}"
 End If
 strSaved = Text2.Text
End Sub

Private Sub Timer1_Timer()
    Dim TMP As String
    TMP = "Area Calculator - Advanced - " & UCase(Left(Shape, 1)) & LCase(Right(Shape, (Len(Shape) - 1)))
    If TMP <> ShapeCalc.Caption Then ShapeCalc.Caption = TMP
    Select Case Shape
        Case "circle"
            If Len(Text1.Text) > 0 Then
                Label11.Caption = "Area is Equal to: " & (3.14 * Val(Format(Text1.Text, "###,###,###.###########"))) & " units "
            Else
                Label11.Caption = vbNullString
            End If
        Case "square"
            If Len(Text1.Text) > 0 Then
                Label11.Caption = "Area is Equal to: " & (Val(Format(Text1.Text, "###,###,###.###########")) * Val(Format(Text1.Text, "###,###,###.###########"))) & " units"
            Else
                Label11.Caption = vbNullString
            End If
        Case "rectangle"
            If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 Then
               Label11.Caption = "Area is Equal to: " & (Val(Format(Text1.Text, "###,###,###.###########")) * Val(Format(Text2.Text, "###,###,###.###########"))) & " units"
            Else
                Label11.Caption = vbNullString
            End If
        Case "triangle"
             If Len(Text1.Text) > 0 And Len(Text2.Text) > 0 Then
                 Label11.Caption = "Area is Equal to: " & (Val(Format(Text2.Text, "###,###,###.###########")) / 2 * Val(Format(Text1.Text, "###,###,###.###########"))) & " units"
             Else
                Label11.Caption = vbNullString
             End If
        Case Else
             Label11.Caption = vbNullString
    End Select
End Sub
