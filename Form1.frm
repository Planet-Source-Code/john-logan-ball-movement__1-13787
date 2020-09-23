VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   180
   ClientTop       =   480
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   403
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape Didi 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   375
   End
   Begin VB.Menu Create 
      Caption         =   "&Create"
      Visible         =   0   'False
      Begin VB.Menu Ball 
         Caption         =   "&Ball"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Last step; enhancing the program:
'yes tere are few new variables, you'll see later on where we used them
Dim A As Integer, B As Integer, X0 As Integer, Y0 As Integer
Dim DeltaX(100) As Integer, DeltaY(100) As Integer
Dim Omega
Dim Alpha



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'we're saving the position of the mouse in these 2 variables
'so that we can make the ball start at that position
'later on
X0 = X
Y0 = Y
'so let's say your mouse was on position X=4857 and Y=3621
'X0 will become 4857 and Y0 = 3621

If Button = 2 Then
Form1.PopupMenu Create

End If

End Sub


Private Sub Ball_Click()

Randomize 'makes all the Rnd in the sub have a random number
'every time:
'here Rnd can give anything and even if you restart the program
'it'll give a new value
'without that, Rnd will follow a same sequence all the time
'like : 2-3-1-4-2-1-
'but with that it'll each time start something up randomely

Omega = Rnd * 5

'same here
B = B + 1
Load Didi(B)
Didi(B).Left = X0
Didi(B).Top = Y0

'and we can also make the new Didi be of a different color
'with RGB
'RGB(255,255,255) is white and RGB(0,0,0) is black
'we make it swicth randomely thrue all of these 3 values
'to get a new random colour.
Didi(B).BackColor = RGB(250 * Rnd + 1, 250 * Rnd + 1, 250 * Rnd + 1)


Didi(B).Visible = True


'don't forget you can only go up to 100 so you should
'make the program stop by then
If B >= 99 Then MsgBox ("To many balls"): End


Select Case Omega
'now, if Omega is below 1 the ball'll move down right
Case Is <= 1
DeltaX(A) = 5
DeltaY(A) = 5
'and, if Omega is between 1 and 2 the ball'll move up right
Case 1 To 2
DeltaX(A) = -5
DeltaY(A) = 5
Case 3 To 4
'or, if Omega is between 2 and 3 the ball'll move down left
DeltaX(A) = 5
DeltaY(A) = -5
Case Is >= 4
'finnaly, if Omega is higher then 4 the ball'll move up left.
DeltaX(A) = -5
DeltaY(A) = -5
End Select

End Sub

Private Sub Form_Load()
'same here
B = 0
DeltaX(0) = 5
DeltaY(0) = -5

End Sub
'same here
Private Sub Timer1_Timer()
For A = 0 To B
Didi(A).Move Didi(A).Left + DeltaX(A), Didi(A).Top + DeltaY(A)

If Didi(A).Left <= 0 Then DeltaX(A) = 5
If Didi(A).Top <= 0 Then DeltaY(A) = 5

If Didi(A).Left >= Me.ScaleWidth - Didi(A).Width Then DeltaX(A) = -5
If Didi(A).Top >= Me.ScaleHeight - Didi(A).Height Then DeltaY(A) = -5

Next

End Sub



Private Sub CmdOK_Click()
Unload Me

'by the way i just cared to say that this isn't how
'you should work typically when building a program,
'i'll make my next project the typical way, and i'll
'even make one just to show how you work.
'hope you enjoyed it;oh yes, my mail for any comments or any
'help you need or want to give bluespy_PJ@hotmail.com
'Bye :)
End Sub

