VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   Caption         =   "Form3"
   ClientHeight    =   6045
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   7350
   LinkTopic       =   "Form3"
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
      Left            =   3360
      Shape           =   3  'Circle
      Top             =   3120
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
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the second step:

'here we will see all the variables needed and adapt
'them partly on the project.
'first we need to declare the new appearing balls,
'we'll call the variable B
Dim B As Integer

'then we need to know the X and Y position of every ball:
Dim DeltaX(100) As Integer, DeltaY(100) As Integer
'this type of variable can have up to 100 possiblities
'like deltax(1) or deltax(2), way up to 100, and if you would
'have put 1000 in stead of 100 like this: DeltaX(1000) there
'would be a 1000 possibilities.
'you're asking yourself why we didn't use (100) for the balls
'right?, well because once the balls are there we don't
'care about them anymore but we do care about their positions
'that's why we need a value for every ball so that we can loop
'between them to check on their positions.

'and we'll also need a variable to take the value of those 100
'possiblities:
Dim A As Integer


'it's a bit dark in your mind now, but it'll become clearer
'after you finished reading and you'll read it again.
'By the way, i won't explain the things i already did a second time.

'Now we'll work on the new shape appearance, therefore we go to
'the ball_click sub:

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'first we have to popup again:

If Button = 2 Then
Form1.PopupMenu Create
End If

End Sub


Private Sub Ball_Click()
'Now we make Didi be index value 0
'and we'll tell the project to:

Load Didi(1) 'load a new Didi, 0 is the only one available,
'now we just loaded the exact same Didi but for value 1
Didi(1).Left = 0 'we make it take position on the top
Didi(1).Top = 0 'left corner of the form

Didi(1).Visible = True 'and we make it visible, because a loaded
'object always starts visible=false

End Sub


'and now we'll show you how to make one ball move thrue
'the screen. Now let's go to timer1 sub:

Private Sub Timer1_Timer()

Didi(0).Move Didi(0).Left + 5, Didi(0).Top + 5
'this tells the shape to move right+5 and down+5 every 1 msec.


End Sub


'Next step
Private Sub CmdNext_Click()
Form2.Show
Unload Me
End Sub

