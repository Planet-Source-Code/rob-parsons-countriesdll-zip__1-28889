VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   4440
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6585
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1890
      Sorted          =   -1  'True
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1800
      Width           =   2835
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   1140
      Top             =   1680
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo1_Click()
Dim mycountries As Countries.Flags
Set mycountries = New Countries.Flags
With mycountries
    Me.Icon = .IconGet(Combo1.ItemData(Combo1.ListIndex))
    Image1.Picture = Me.Icon
    Me.Caption = .GetISOCode(Combo1.ItemData(Combo1.ListIndex)) & " - " & Combo1.Text
End With
If Not mycountries Is Nothing Then Set mycountries = Nothing

End Sub

Private Sub Form_Load()
Dim mycountries As Countries.Flags
Set mycountries = New Countries.Flags
With mycountries
    .FillCombo Combo1
    Combo1.ListIndex = 0
    Debug.Print .GetHTMLCombo(, Combo1.Text)
End With
If Not mycountries Is Nothing Then Set mycountries = Nothing
End Sub

