VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "BitBlt  StretchBlt Example"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7305
   LinkTopic       =   "Form1"
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   487
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   6000
      TabIndex        =   4
      ToolTipText     =   " Clear Image "
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Same Size"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   " View image same size "
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy stretched "
      Height          =   495
      Left            =   360
      TabIndex        =   2
      ToolTipText     =   " View image stretched "
      Top             =   6240
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   5415
      Left            =   3000
      ScaleHeight     =   357
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5460
      Left            =   240
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   360
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Destination Image"
      Height          =   195
      Left            =   3120
      TabIndex        =   6
      Top             =   5640
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Source Image"
      Height          =   195
      Left            =   360
      TabIndex        =   5
      Top             =   5760
      Width           =   990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

        rtn = StretchBlt(Picture2.hdc, _
        0, _
        0, _
        300, _
        50, _
        Picture1.hdc, _
        0, _
        0, _
        173, _
        360, _
        SRCCOPY)
    
Picture2.Refresh


'''''''Below describes what the code above does''''''
'
'
'   The numbers below can be changed as needed for
'   your project
'
'
'
'rtn = StretchBlt(Picture2.hdc, _  'Destination image
'0, _         'set left of Picture2 (X pixel)
'0, _         'set top of Picture2 (Y pixel)
'300, _       'set Width of Picture2 (This is stretched)
'50, _        'set Height of Picture2 (This is stretched)
'Picture1.hdc, _                   'Source Image
'0, _         'selected left of Picture1
'0, _         'selected top of Picture1
'173, _       'Selected width from  Picture1
'360, _       'Selected height from  Picture1
'SRCCOPY)


'NOTE!!!!! When stretching an image, as shown above, the color pallete is reduced to
'256 color pallette. Something unavoidable in VB

End Sub

Private Sub Command2_Click()
    Suc% = BitBlt(Picture2.hdc, 0, 0, Picture1.Width, Picture1.Height, Picture1.hdc, 0, 0, &HCC0020)


''The code above just copies an image from one object to another''''
Picture2.Refresh
End Sub

Private Sub Command3_Click()
Picture2.Picture = LoadPicture()
End Sub

Public Sub Messag()
msg = "    Notice that Form1, Picture1 and Picture2 ScaleMode property is set at 3 (Pixel)." & vbCrLf
msg = msg & "Also, make sure that the AutoRedraw property for the picture boxes is set to TRUE." & vbCrLf & vbCrLf
msg = msg & "    Don't forget to use " & Chr(34) & "Picture1.Refresh" & Chr(34) & " and " & Chr(34) & "Picture2.Refresh" & Chr(34) & " after using the "
msg = msg & "StretchBlt and BitBlt codes, otherwise your image will not be displayed properly." & vbCrLf & vbCrLf & vbCrLf & vbCrLf
msg = msg & "    The VISIBLE property for Picture1 is set to FALSE. This is not neccessary, but just demonstrates that it can be done without error. "


MsgBox msg, vbCritical, "Visual Basic Programmer:  PLEASE READ!  "
End Sub

Private Sub Form_Load()
Messag
End Sub


