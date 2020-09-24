VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4515
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton btGetBag 
      Caption         =   "Get Bag Contents"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2977
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton btFillBag 
      Caption         =   "Fill Bag"
      Height          =   495
      Left            =   2962
      TabIndex        =   0
      Top             =   1305
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fest Einfach
      Height          =   525
      Left            =   3630
      Top             =   615
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Demonstrates the use of a property bag -

'  *  to store "things" in a file in any order
'  *  and to retrieve those things from that file in any order

Option Explicit

Private Sub btFillBag_Click()

  Dim objBag            As New PropertyBag
  Dim vntBagContents    As Variant

    With objBag

        'put things into the bag in any order
        .WriteProperty "Num2", -456                     'numerics...
        .WriteProperty "Num1", 4 * Atn(1)               'results (pi in this case)...
        .WriteProperty "String2", "This is StringTwo"   'strings...
        .WriteProperty "Bool2", (2 > 3)                 'booleans...
        .WriteProperty "String1", "This is String1"
        .WriteProperty "Bool1", (2 < 3)
        .WriteProperty "Font", Font                     'fonts...
        .WriteProperty "Picture", Icon                  'and even pictures can be put in the bag
        .WriteProperty "BackColor", vbGreen

        'alter font for now (to be able to restore it from the bag later)
        With Font
            .Name = "Courier New"
            .Size = 12
            .Italic = True
            .Bold = True
        End With 'FONT

        Print "Bag filled"

        'transfer contents into a variant
        vntBagContents = .Contents

    End With 'objBag

    Open App.Path & "\Things.Bag" For Binary As 1
    'write bag contents to output file
    Put #1, , vntBagContents
    Print "Bag written"
    Close 1

    'command buttons
    btFillBag.Enabled = False
    btGetBag.Enabled = True

End Sub

Private Sub btGetBag_Click()

  Dim objBag            As New PropertyBag
  Dim vntBagContents    As Variant

    Cls

    Open App.Path & "\Things.Bag" For Binary As 1
    'get bag contents into a variant
    Get #1, , vntBagContents
    Close 1

    With objBag
        'transfer contents into the bag
        .Contents = vntBagContents

        'take things out of the bag in any order
        BackColor = .ReadProperty("BackColor")
        Set Font = .ReadProperty("Font") 'restore font
        Print .ReadProperty("String1")
        Print .ReadProperty("String2")
        Print .ReadProperty("Bool1")
        Print .ReadProperty("Bool2")
        Print .ReadProperty("Num1")
        Print .ReadProperty("Num2")
        Image1 = .ReadProperty("Picture")

        'and an attempt to find a non-existent item in the bag
        Print .ReadProperty("XYZ", "Item XYZ Missing")

    End With 'objBag

End Sub

':) Ulli's VB Code Formatter V2.16.14 (2004-Mrz-15 01:52) 7 + 81 = 88 Lines
