VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "One Scrollbar List Boxes"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Add Entry"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox MyListFront 
      Appearance      =   0  'Flat
      Height          =   2370
      Index           =   2
      Left            =   2280
      TabIndex        =   10
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox MyListFront 
      Appearance      =   0  'Flat
      Height          =   2370
      Index           =   1
      Left            =   1560
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
   Begin VB.ListBox MyListFront 
      Appearance      =   0  'Flat
      Height          =   2370
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   360
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Favorate Color"
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Weight"
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Height"
      Height          =   255
      Left            =   1560
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Age"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox MyList 
      Appearance      =   0  'Flat
      Height          =   2370
      Index           =   3
      ItemData        =   "1SBList.frx":0000
      Left            =   3000
      List            =   "1SBList.frx":0037
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
   Begin VB.ListBox MyList 
      Appearance      =   0  'Flat
      Height          =   420
      Index           =   2
      ItemData        =   "1SBList.frx":00BE
      Left            =   0
      List            =   "1SBList.frx":00F5
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox MyList 
      Appearance      =   0  'Flat
      Height          =   420
      Index           =   1
      ItemData        =   "1SBList.frx":014E
      Left            =   0
      List            =   "1SBList.frx":0185
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ListBox MyList 
      Appearance      =   0  'Flat
      Height          =   420
      Index           =   0
      ItemData        =   "1SBList.frx":01CD
      Left            =   0
      List            =   "1SBList.frx":0204
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight BG"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Height BG"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AgeBackground"
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command5_Click()
  MyList(0).AddItem "33", 0         'adds Items to the Lists and then
  MyList(1).AddItem "77", 0
  MyList(2).AddItem "270", 0
  MyList(3).AddItem "Moldy Green", 0
  PopulateLists                     'refills the front list boxes
End Sub

Private Sub Form_Load()
  Dim X As Long                 'Fills the Front Listboxes with Nothing
                                'so that they are easyer to deal with
  For X = 0 To 11
    MyListFront(0).AddItem ""
    MyListFront(1).AddItem ""
    MyListFront(2).AddItem ""
  Next X
  PopulateLists                 'then fills them with the Background data
End Sub

Private Sub MyList_Click(Index As Integer)
  If MyList(3).ListIndex - MyList(3).TopIndex >= 0 And MyList(3).ListIndex - MyList(3).TopIndex < 12 Then
    MyListFront(0).Selected(MyList(3).ListIndex - MyList(3).TopIndex) = True
    MyListFront(1).Selected(MyList(3).ListIndex - MyList(3).TopIndex) = True
    MyListFront(2).Selected(MyList(3).ListIndex - MyList(3).TopIndex) = True
  Else
    MyListFront(0).ListIndex = -1   'If the selected item in the right list is with in
    MyListFront(1).ListIndex = -1   'the scope of the Front Lists then it selectets them
    MyListFront(2).ListIndex = -1
  End If
End Sub

Private Sub MyList_Scroll(Index As Integer)
  MyList_Click 3  'Uses Current code to do scroll event
  PopulateLists
End Sub

Private Sub MyListFront_Click(Index As Integer)
  Dim X As Long
  
  If MyList(2).ListCount > 0 Then           'When Clicked selects all lists
    If MyListFront(Index).ListIndex >= 0 Then
      MyListFront(0).Selected(MyListFront(Index).ListIndex) = True
      MyListFront(1).Selected(MyListFront(0).ListIndex) = True
      MyListFront(2).Selected(MyListFront(0).ListIndex) = True
      MyList(3).Selected(MyListFront(Index).ListIndex + MyList(3).TopIndex) = True
    End If
  End If
End Sub

Sub MyLists_Click(Index As Integer)
  Dim X As Byte
  Dim TowerVar As String
  Dim LevelVar As String
    
  If MyList(Index).ListCount > 0 Then   'When Clicked selects all lists
    If MyList(Index).ListIndex - MyList(Index).TopIndex >= 0 And MyList(Index).ListIndex - MyList(Index).TopIndex < 12 Then
      MyListFront(0).Selected(MyList(3).ListIndex - MyList(3).TopIndex) = True
      MyListFront(1).Selected(MyList(3).ListIndex - MyList(3).TopIndex) = True
      MyListFront(2).Selected(MyList(3).ListIndex - MyList(3).TopIndex) = True
    Else
      MyListFront(0).ListIndex = -1
      MyListFront(1).ListIndex = -1
      MyListFront(2).ListIndex = -1
    End If
  End If
End Sub

Private Sub MyLists_Scroll(Index As Integer)
  Dim X As Byte
  
  For X = 0 To 2
    MyList(X).TopIndex = MyList(Index).TopIndex
  Next X
  If MyList(2).ListIndex - MyList(2).TopIndex < 0 Or MyList(2).ListIndex - MyList(2).TopIndex > 7 Then
    MyListFront(0).ListIndex = -1
    MyListFront(1).ListIndex = -1
  Else
    MyListFront(0).Selected(MyList(2).ListIndex - MyList(2).TopIndex) = True
    MyListFront(1).Selected(MyList(2).ListIndex - MyList(2).TopIndex) = True
  End If
  PopulateLists
End Sub

Sub PopulateLists()
  Dim X As Long

  For X = 0 To 11
    MyListFront(0).List(X) = MyList(0).List(X + MyList(3).TopIndex)
    MyListFront(1).List(X) = MyList(1).List(X + MyList(3).TopIndex)
    MyListFront(2).List(X) = MyList(2).List(X + MyList(3).TopIndex)
  Next X
  MyLists_Click 3
End Sub
