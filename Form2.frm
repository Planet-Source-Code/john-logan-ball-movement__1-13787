VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   Caption         =   "Form2"
   ClientHeight    =   6045
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   7350
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Shape Didi 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2760
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Third step, the hardest, now we have to make all the balls
'move thrue the screen and make them all bounce on the "walls".



Dim A As Integer, B As Integer
Dim DeltaX(100) As Integer, DeltaY(100) As Integer


'We start by making as many balls(only 100) as we want appear,
'we do that by making it load didi(B) which, on every click
'adds a new value to itself:

'First off course, the usuall step:
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
Form1.PopupMenu Create
End If

End Sub


Private Sub Ball_Click()
B = B + 1 'let's say B was 3, now i't ll become 4, and on
'an other click it'll become 5 then 6...

Load Didi(B) 'and then we'll be loading every time a new didi,
Didi(B).Left = 0 'setting the new didi on the top left corner,
Didi(B).Top = 0
Didi(B).Visible = True ' and making it visible.

DeltaX(B) = 5 'and we have to tell the ball to move
DeltaY(B) = 5 'down right

'the sub will start by adding a new value to B so B was 3
'becomes now 4 and then it'll do the usuall stuff with that
'new value, and every time the same.

End Sub

'Now, the bouncing part for all the balls:

Private Sub Timer1_Timer()
'we'll make all the balls move, we do so by making this timer
'loop every 1 milisecond thrue all the balls available
'and make them all move, therefore we tell it to go from
'the primary value (0) to the last added ball and make it move.
'the last added ball has the value B which is updated on
'every new addon.


For A = 0 To B

Didi(A).Move Didi(A).Left + DeltaX(A), Didi(A).Top + DeltaY(A)
'so, for A=0:
'didi(1).move didi(1).left + deltaX(1)(which is 5),didi(1).top+
'DeltaY(1)(which is also 5), u can c that in Ball_click, the
'last line

'As for the Bouncing; we have to check every ball if it is or
'not touching any border of the form; the 4 borders are:
'1-the left side which is X=0
'2-the top side which is Y=0
'3-the right side which is X=me.scalewidth
'4-the bottom side which is Y=me.scaleheight

'1-so if Didi's left goes below 0 then inverse it's side from
'-5 to 5:
If Didi(A).Left <= 0 Then DeltaX(A) = 5

'2-if Didi's top goes under 0 then inverse it's side from -5
'to 5
If Didi(A).Top <= 0 Then DeltaY(A) = 5

'3-if didi's left is over the scalewidth minus it's own width
'(because without that i'll go a little out of screen)
'inverse the size from 5 to -5
If Didi(A).Left >= Me.ScaleWidth - Didi(A).Width Then DeltaX(A) = -5

'4-if didi's top is over the scaleheight minus it's own height
'inverse the size from 5 to -5
If Didi(A).Top >= Me.ScaleHeight - Didi(A).Height Then DeltaY(A) = -5

Next

End Sub

'Off course we should give a starting value for the first Didi
'when the program starts B has to be 0 to make Didi(0) move
'and it's own deltaX and deltaY have to have a value or else
'it'll just stay on it's place

Private Sub Form_Load()
B = 0
DeltaX(0) = 5
DeltaY(0) = -5

End Sub


'Last step
Private Sub CmdNext_Click()
Form1.Show
Unload Me

End Sub

