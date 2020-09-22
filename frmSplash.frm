VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   4440
   ClientTop       =   3720
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   223
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   343
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   480
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3000
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim H, W, col As Long
    
    Me.Width = Picture1.Width * 15 'Size the Splashscreenform to
    Me.Height = Picture1.Height * 15 ' fit the Picture

    Call GetDesktop(Me) 'Make Screenshot of Screen behind frmSplash
                        'and copy it to the form
    Me.Show

    For W = 0 To Picture1.ScaleWidth 'This is the effect itself
        For H = 0 To Picture1.ScaleHeight
            If Gerade(W) = Gerade(H) Then
            ' ^ Makes sure only every second Pixel is shown
                col = GetPixel(Picture1.hdc, W, H)
                SetPixel Me.hdc, destX + W, destY + H, col
            Else
                col = GetPixel(Picture1.hdc, Picture1.ScaleWidth - W, H)
                SetPixel Me.hdc, destX + Picture1.ScaleWidth - W, destY + H, col
            End If
        Next H
    
        Me.Refresh
        DoEvents
        Pause 1 'Here you can change the speed of the effect
    Next W

    Unload Me

End Sub
