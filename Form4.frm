VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   Caption         =   "Form4"
   ClientHeight    =   6045
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   7350
   LinkTopic       =   "Form4"
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
      Left            =   2640
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   375
   End
   Begin VB.Menu Create 
      Caption         =   "Create"
      Visible         =   0   'False
      Begin VB.Menu Ball 
         Caption         =   "&Ball"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



'project for between beginners and intermediate
'we're assuming that you know the if...the...else
'and the loop statements and a little of indexes

'hi,it's me, nothing much to say about me, anyway...
'i always liked to understand many things at a time when
'i download my programs,
'or just have many usefull and easy things to understand
'in one project; therefore i'll be making many projects with
'2-3 things to learn in them, with a step by step "how to do",
'and off course allong the project i'll try to show you many
'other little nice things to know,
'so it'll be easy,helpfull and instructive.
'This is going to be my first project,in this i'll be teaching
'you how to mutilply dim with one variant;
'how to make the user create new objects;
'and a little about collision effects.

'the project is: letting the user be able to create new balls
'                on the form, and those balls have to rebounce
'                on the sides of the form.
'
'We'll divide the project in four steps
'
'the first step consists of taking care of the main idea
'of the project as well as the first few steps:

'*So we first add a circle shaped object on the form to have
'something to start with; then we think of the main idea:

'1-how will the user be adding a new ball?
'with a poping up menu with that option;so let's build it
'with the menu editor, and make it not visible(name:Didi)

'2-how will the ball know it's colliding?
'a timer set at 1 milisecond can keep on the track;
'so just add one(name:Timer1)

'*Now all that's left is the subs in which we will be working:

'We'll have to make the Ball menu available once you
'right click on the form:

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then 'button=1 means left mousebutton, and
'button=2 means right mouse button

Form1.PopupMenu Create 'here you give the name of the
'head menu to be loaded

End If
'here the Ball menu will popup and that will be our main sub:

End Sub

'and there is Timer1's sub to detect any collisions:

Private Sub Timer1_Timer()

End Sub

'and the next command to go to the next step

Private Sub CmdNext_Click()
Form3.Show
Unload Me

End Sub

