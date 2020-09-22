VERSION 5.00
Begin VB.Form Level2 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer timer1 
      Interval        =   1000
      Left            =   3000
      Top             =   1320
   End
   Begin VB.Timer upgrade1timer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   6000
   End
   Begin VB.Timer kots2timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   5160
   End
   Begin VB.Timer alien2timer 
      Interval        =   15
      Left            =   240
      Top             =   4680
   End
   Begin VB.Timer kots1timer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   3480
   End
   Begin VB.Timer alien1timer 
      Interval        =   15
      Left            =   240
      Top             =   2880
   End
   Begin VB.Timer bullet2timer 
      Interval        =   100
      Left            =   240
      Top             =   960
   End
   Begin VB.Timer BulletTimer 
      Interval        =   100
      Left            =   240
      Top             =   360
   End
   Begin VB.Label lbllevel1clear 
      BackStyle       =   0  'Transparent
      Caption         =   "MISSION 2 ACCOMPLISHED"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1695
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label time 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TIME REMAINING:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Image upgrade1 
      Height          =   405
      Left            =   2760
      Picture         =   "level2.frx":0000
      Top             =   6720
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblAlien2power 
      BackStyle       =   0  'Transparent
      Caption         =   "ENEMY POWER: "
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   4095
   End
   Begin VB.Image kots2 
      Height          =   390
      Left            =   7440
      Picture         =   "level2.frx":026A
      Top             =   3480
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image alien2 
      Height          =   855
      Left            =   8520
      Picture         =   "level2.frx":0622
      Top             =   3000
      Width           =   1665
   End
   Begin VB.Label lblAlien1power 
      BackStyle       =   0  'Transparent
      Caption         =   "ENEMY POWER: "
      ForeColor       =   &H00FF8080&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lblInstr 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H8000000E&
      Height          =   1095
      Left            =   3840
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblGameOver 
      BackStyle       =   0  'Transparent
      Caption         =   "GAME OVER:   YOU LOSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   2415
      Left            =   2280
      TabIndex        =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.Image kots1 
      Height          =   390
      Left            =   3720
      Picture         =   "level2.frx":0FD9
      Top             =   1560
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Image alien1 
      Height          =   855
      Left            =   4680
      Picture         =   "level2.frx":1337
      Top             =   1080
      Width           =   1665
   End
   Begin VB.Image imgBullet2 
      Height          =   210
      Left            =   5040
      Picture         =   "level2.frx":1CFF
      Top             =   6240
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image imgBullet 
      Height          =   210
      Left            =   5400
      Picture         =   "level2.frx":1E2B
      Top             =   6240
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   5040
      Picture         =   "level2.frx":1F57
      Top             =   6480
      Width           =   885
   End
End
Attribute VB_Name = "Level2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim alien1modus As String
Dim alien2modus As Integer
Dim alien1power As Integer
Dim alien2power As Integer
Dim a As Integer
Dim game As Integer







Private Sub alien1timer_Timer()



If alien1.Left > Image1.Left - 300 And alien1.Left < Image1.Left + 300 And kots1timer.Enabled = False Then

    PlayPukeSound
 
    kots1.Left = alien1.Left
    kots1.Visible = True
    
    kots1timer.Enabled = True


End If


If alien1modus = 1 Then

    alien1.Left = alien1.Left - 250
    If alien1.Left < 200 Then alien1modus = 2
    
End If

If alien1modus = 2 Then
    alien1.Left = alien1.Left + 250
    If alien1.Left > 8000 Then alien1modus = 1
    
End If
    






End Sub



Private Sub alien2timer_Timer()
If alien2.Left > Image1.Left - 300 And alien2.Left < Image1.Left + 300 And kots2timer.Enabled = False Then

    PlayPukeSound
    
    kots2.Left = alien2.Left
    kots2.Visible = True
    
    kots2timer.Enabled = True


End If


If alien2modus = 1 Then

    alien2.Left = alien2.Left - 250
    If alien2.Left < 200 Then alien2modus = 2
    
End If

If alien2modus = 2 Then
    alien2.Left = alien2.Left + 250
    If alien2.Left > 8000 Then alien2modus = 1
    
End If
    
End Sub

Private Sub bullet2timer_Timer()
imgBullet2.Top = imgBullet2.Top - 650

' Wat te doen als een kogel de kots geraakt heeft...
If imgBullet2.Left > (kots1.Left) And imgBullet2.Left < (kots1.Left + 1000) Then
    If imgBullet2.Top > (kots1.Top - 300) And imgBullet2.Top < (kots1.Top + 300) Then

        imgBullet2.Top = 6240
        imgBullet2.Visible = False
        bullet2timer.Enabled = False
        
   kots1timer.Enabled = False
    kots1.Visible = False
    kots1.Top = 1560

End If
End If

' Wat te doen als een kogel de kots geraakt heeft...
If imgBullet.Left > (kots2.Left) And imgBullet.Left < (kots2.Left + 1000) And kots2.Visible = True Then
    If imgBullet.Top > (kots2.Top - 300) And imgBullet.Top < (kots2.Top + 300) Then

        imgBullet.Top = 6240
        imgBullet.Visible = False
        BulletTimer.Enabled = False
        
   kots2timer.Enabled = False
    kots2.Visible = False
    kots2.Top = 3480

End If
End If

' Kogel op kots van alien 2...
If imgBullet2.Left > (kots2.Left) And imgBullet2.Left < (kots2.Left + 1000) Then
    If imgBullet2.Top > (kots2.Top - 300) And imgBullet2.Top < (kots2.Top + 300) Then

        imgBullet2.Top = 6240
        imgBullet2.Visible = False
        bullet2timer.Enabled = False
        
   kots2timer.Enabled = False
    kots2.Visible = False
    kots2.Top = 3480

End If
End If





If imgBullet2.Top < 651 Then
    imgBullet2.Top = 6240
    imgBullet2.Visible = False
    
    bullet2timer.Enabled = False
End If

End Sub

Private Sub BulletTimer_Timer()
imgBullet.Top = imgBullet.Top - 650


If imgBullet.Left > alien1.Left And imgBullet.Left < (alien1.Left + 1000) Then
    If imgBullet.Top >= (alien1.Top) Then
    
    alien1power = alien1power - 1
    lblAlien1power.Caption = "ENEMY POWER: "
    
    For i = 1 To alien1power
    lblAlien1power.Caption = lblAlien1power.Caption & "|"
    Next
      
    If alien1power = 0 Then
        alien1.Visible = False
        Playlaser
        alien1timer.Enabled = False
        
        If alien2.Visible = True Then
            upgrade1.Left = alien1.Left
            upgrade1.Top = alien1.Top
            getupgrade1
        End If
        
        If alien2.Visible = False Then level1clear
        
            
    End If
End If

End If
If imgBullet.Left > alien2.Left And imgBullet.Left < (alien2.Left + 1000) Then
    If imgBullet.Top >= (alien2.Top) Then
    
    alien2power = alien2power - 1
    lblAlien2power.Caption = "ENEMY POWER: "
    
    For i = 1 To alien2power
    lblAlien2power.Caption = lblAlien2power.Caption & "|"
    Next
      
    If alien2power = 0 Then
        alien2.Visible = False
        Playlaser
        alien2timer.Enabled = False
        
        If alien1.Visible = True Then
        upgrade1.Left = alien2.Left
        upgrade1.Top = alien2.Top
         getupgrade1
        End If
        
        If alien1.Visible = False Then level1clear
        
        
        
    End If
    
    End If
End If


' Wat te doen als een kogel de kots geraakt heeft...
If imgBullet.Left > (kots1.Left) And imgBullet.Left < (kots1.Left + 1000) And kots1.Visible = True Then
    If imgBullet.Top > (kots1.Top - 300) And imgBullet.Top < (kots1.Top + 300) Then

        imgBullet.Top = 6240
        imgBullet.Visible = False
        BulletTimer.Enabled = False
        
   kots1timer.Enabled = False
    kots1.Visible = False
    kots1.Top = 1560

End If
End If


' Wat te doen als een kogel de kots geraakt heeft...
If imgBullet.Left > (kots2.Left) And imgBullet.Left < (kots2.Left + 1000) And kots2.Visible = True Then
    If imgBullet.Top > (kots2.Top - 300) And imgBullet.Top < (kots2.Top + 300) Then

        imgBullet.Top = 6240
        imgBullet.Visible = False
        BulletTimer.Enabled = False
        
   kots2timer.Enabled = False
    kots2.Visible = False
    kots2.Top = 3480

End If
End If



If imgBullet.Top < 651 Then
    imgBullet.Top = 6240
    imgBullet.Visible = False
    
    BulletTimer.Enabled = False
End If

End Sub

Private Sub Form_Activate()

BulletTimer.Interval = Level1.BulletTimer.Interval
bullet2timer.Interval = Level1.bullet2timer.Interval





game = 1
a = 0
Label1.Visible = False

Label1.Caption = "SHIP UPGRADED:" & vbCrLf & "YOU CAN NOW MOVE TWICE AS FAST!!!"

lblInstr.Caption = "ESC - Quit game"

alien1modus = 1
alien2modus = 1
alien1power = 70
alien2power = 70

lblAlien1power.Caption = "ENEMY POWER: "
lblAlien2power.Caption = "ENEMY POWER: "
    
    For i = 1 To alien1power
    lblAlien1power.Caption = lblAlien1power.Caption & "|"
    Next

For i = 1 To alien2power
    lblAlien2power.Caption = lblAlien2power.Caption & "|"
    Next

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)


If game = 0 Then

If KeyCode = vbKeyEscape Then End
End If


    

If game = 1 Then
If KeyCode = vbKeyLeft And Label1.Visible = False Then Image1.Left = Image1.Left - 100: Image1.Picture = LoadPicture(App.Path & "\goleft.jpg")
If KeyCode = vbKeyLeft And Label1.Visible = True Then Image1.Left = Image1.Left - 200: Image1.Picture = LoadPicture(App.Path & "\goleft.jpg")
If KeyCode = vbKeyRight And Label1.Visible = False Then Image1.Left = Image1.Left + 100: Image1.Picture = LoadPicture(App.Path & "\goright.jpg")
If KeyCode = vbKeyRight And Label1.Visible = True Then Image1.Left = Image1.Left + 200: Image1.Picture = LoadPicture(App.Path & "\goright.jpg")

If KeyCode = vbKeySpace Then


    If BulletTimer.Enabled = False Then

    imgBullet.Left = Image1.Left + 450
    imgBullet.Visible = True
    
    BulletTimer.Enabled = True
    
    Else
    
    imgBullet2.Left = Image1.Left + 450
    imgBullet2.Visible = True
    
     bullet2timer.Enabled = True
End If
   
End If
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If game = 1 Then Image1.Picture = LoadPicture(App.Path & "\char.jpg")

End Sub

Private Sub kots1timer_Timer()
kots1.Top = kots1.Top + 250

If kots1.Top > 6480 And kots1.Left > (Image1.Left - 800) And kots1.Left < (Image1.Left + 800) Then
gameover
End If


If kots1.Top > 7000 Then
    kots1timer.Enabled = False
    kots1.Visible = False
    kots1.Top = 1560
End If
    

End Sub

Private Sub Timer1_Timer()

a = a + 1

time.Caption = 200 - a

If a = 200 Then gameover
 

End Sub

Private Sub kots2timer_Timer()
kots2.Top = kots2.Top + 250

If kots2.Top > 6480 And kots2.Left > (Image1.Left - 800) And kots2.Left < (Image1.Left + 800) Then
gameover
End If


If kots2.Top > 7000 Then
    kots2timer.Enabled = False
    kots2.Visible = False
    kots2.Top = 3480
End If
    
End Sub
Private Sub getupgrade1()


upgrade1.Visible = True
upgrade1timer.Enabled = True



End Sub

Private Sub upgrade1timer_Timer()

upgrade1.Top = upgrade1.Top + 200


If upgrade1.Top > 6480 And upgrade1.Left > (Image1.Left - 800) And upgrade1.Left < (Image1.Left + 800) Then
    Label1.Visible = True
    
    
    
    
    upgrade1.Visible = False
    upgrade1timer.Enabled = False
    
    
End If


End Sub

Private Sub gameover()
Playlaser
Image1.Picture = LoadPicture(App.Path & "\dead.jpg")
kots1.Visible = False
kots2.Visible = False
alien1.Visible = False
alien2.Visible = False
alien1timer.Enabled = False
alien2timer.Enabled = False
kots1timer.Enabled = False
kots2timer.Enabled = False


Call mciSendString("Stop MM", 0, 0, 0)
    Call mciSendString("Close MM", 0, 0, 0)
filename = App.Path & "\ass.wav"

    Call mciSendString("Open " & filename & " Alias MM", 0, 0, 0)
    Call mciSendString("Play MM", 0, 0, 0)


lblGameOver.Visible = True
lblInstr.Visible = True
game = 0


End Sub

Private Sub PlayPukeSound()
Call mciSendString("Stop MM", 0, 0, 0)
    Call mciSendString("Close MM", 0, 0, 0)

filename = App.Path & "\puke.wav"


    Call mciSendString("Open " & filename & " Alias MM", 0, 0, 0)
    Call mciSendString("Play MM", 0, 0, 0)

End Sub

Private Sub Playlaser()
Call mciSendString("Stop MM", 0, 0, 0)
    Call mciSendString("Close MM", 0, 0, 0)

filename = App.Path & "\laser.wav"


    Call mciSendString("Open " & filename & " Alias MM", 0, 0, 0)
    Call mciSendString("Play MM", 0, 0, 0)

End Sub

Private Sub level1clear()


lbllevel1clear.Visible = True



End Sub
