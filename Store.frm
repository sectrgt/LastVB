VERSION 5.00
Begin VB.Form frmStore 
   Caption         =   "VERSION 5.00"
   ClientHeight    =   12120
   ClientLeft      =   -2265
   ClientTop       =   705
   ClientWidth     =   14715
   LinkTopic       =   "Form1"
   ScaleHeight     =   12120
   ScaleWidth      =   14715
   Begin VB.CommandButton cmdCategory 
      Caption         =   "Command1"
      Height          =   615
      Index           =   4
      Left            =   3840
      TabIndex        =   23
      Top             =   9120
      Width           =   3375
   End
   Begin VB.CommandButton cmdCategory 
      Caption         =   "Command1"
      Height          =   615
      Index           =   3
      Left            =   120
      TabIndex        =   22
      Top             =   9120
      Width           =   3375
   End
   Begin VB.CommandButton cmdCategory 
      Caption         =   "Command1"
      Height          =   615
      Index           =   2
      Left            =   7560
      TabIndex        =   21
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton cmdCategory 
      Caption         =   "Command1"
      Height          =   615
      Index           =   1
      Left            =   3960
      TabIndex        =   20
      Top             =   4800
      Width           =   3375
   End
   Begin VB.CommandButton cmdCategory 
      Caption         =   "Command1"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   4800
      Width           =   3495
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Command7"
      Height          =   735
      Left            =   3960
      TabIndex        =   17
      Top             =   10080
      Width           =   3375
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Frame1"
      Height          =   855
      Left            =   7920
      TabIndex        =   13
      Top             =   8640
      Width           =   4095
      Begin VB.OptionButton optQTN 
         Caption         =   "Option1"
         Height          =   495
         Left            =   2760
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optCost 
         Caption         =   "Option1"
         Height          =   495
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optName 
         Caption         =   "Option1"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Command8"
      Height          =   615
      Left            =   7920
      TabIndex        =   12
      Top             =   9720
      Width           =   4095
   End
   Begin VB.CommandButton cmdCheckout 
      Caption         =   "Command7"
      Height          =   735
      Left            =   12480
      TabIndex        =   10
      Top             =   7200
      Width           =   3975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   1095
      Left            =   14640
      TabIndex        =   8
      Top             =   9480
      Width           =   2535
   End
   Begin VB.ListBox lstList 
      Height          =   10590
      Left            =   17280
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Command7"
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   10080
      Width           =   3375
   End
   Begin VB.Label lblCounter 
      Caption         =   "Label1"
      Height          =   735
      Left            =   8160
      TabIndex        =   18
      Top             =   10920
      Width           =   2655
   End
   Begin VB.Label lblSearchTitle 
      Caption         =   "Label7"
      Height          =   735
      Left            =   7920
      TabIndex        =   11
      Top             =   7680
      Width           =   4095
   End
   Begin VB.Label lblInfo 
      Caption         =   "Label6"
      Height          =   2655
      Left            =   12480
      TabIndex        =   9
      Top             =   4320
      Width           =   3975
   End
   Begin VB.Image imgCheckout 
      Height          =   2535
      Left            =   12480
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Image imgBRT 
      Height          =   2415
      Left            =   3840
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label lblBRT 
      Caption         =   "Label1"
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Image imgTMT 
      Height          =   2415
      Left            =   3840
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label lblTMT 
      Caption         =   "Label1"
      Height          =   615
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Image imgBLT 
      Height          =   2415
      Left            =   120
      Top             =   6600
      Width           =   3495
   End
   Begin VB.Label lblBLT 
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Image imgTRT 
      Height          =   2415
      Left            =   7560
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label lblRMT 
      Caption         =   "Label1"
      Height          =   615
      Left            =   7560
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Image imgTLT 
      Height          =   2415
      Left            =   120
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Label lblTLT 
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label lbltitle 
      Caption         =   "Label1"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16815
   End
   Begin VB.Menu mnupsu 
      Caption         =   "Power Supplies"
      Index           =   5
   End
   Begin VB.Menu mnumobo 
      Caption         =   "Motherboards"
      Index           =   4
   End
   Begin VB.Menu mnuHDD 
      Caption         =   "Hard Drives"
      Index           =   3
   End
   Begin VB.Menu mnuRAM 
      Caption         =   "RAM"
      Index           =   2
   End
   Begin VB.Menu mnuCPU 
      Caption         =   "CPU's"
      Index           =   1
   End
End
Attribute VB_Name = "frmStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Type itemInfo
    itemName As String
    itemCost As Double
    itemQuantity As Integer
    itemPicturePath As String
End Type
Private udtItems(15) As itemInfo
Private intCategory As Integer
Private intIndex As Integer
Private Sub cmdBack_Click()
imgBLT.Visible = True
lblBLT.Visible = True
cmdBLT.Visible = True
imgBRT.Visible = True
lblBRT.Visible = True
cmdBRT.Visible = True
End Sub
Private Sub cmdBLT_Click()
intCategory = 4
Call imgDisplay
End Sub
Private Sub cmdBRT_Click()
intCategory = 5
Call imgDisplay
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdCheckout_Click()
Call itemCheckout
End Sub

Private Sub cmdSearch_Click()
If optName.Value = True Then
    Call searchName
ElseIf optCost.Value = True Then
    Call searchCost
ElseIf optQTN.Value = True Then
    Call searchQuantity
End If
End Sub
Sub searchCost()
Dim intX As Integer
Dim intCost As Integer
Dim intTemp As Integer
intCost = InputBox("Enter the cost of the product", "Kipplex") 'Gets search input from user
    For intX = 0 To UBound(udtItems())
    If intCost = udtItems(intX).itemCost Then  'Searches the item array for user input
        intTemp = 1
        intIndex = intX 'Puts place found in the array in a variable
    End If
    Next intX
If intTemp = 1 Then
    MsgBox "Item Found", , "Search"
    lblInfo.Caption = "Name: " & udtItems(intIndex).itemName & vbCrLf & "Cost: $" & udtItems(intIndex).itemCost & vbCrLf & "Quantity: " & udtItems(intIndex).itemQuantity
    imgCheckout.Picture = LoadPicture(udtItems(intIndex).itemPicturePath)
ElseIf intTemp <> 1 Then
        MsgBox "Item not found", , "Search"
End If
End Sub
Sub searchQuantity()
Dim intX As Integer
Dim intIndex As Integer
Dim intQuantity As Integer
Dim intTemp As Integer
intQuantity = InputBox("Enter the quantity of the product", "Kipplex") 'Gets search input from user
    For intX = 0 To UBound(udtItems())
    If intQuantity = udtItems(intX).itemQuantity Then  'Searches the item array for user input
        intTemp = 1
        intIndex = intX 'Puts place found in the array in a variable
    End If
    Next intX
If intTemp = 1 Then
    MsgBox "Item Found", , "Search"
    lblInfo.Caption = "Name: " & udtItems(intIndex).itemName & vbCrLf & "Cost: $" & udtItems(intIndex).itemCost & vbCrLf & "Quantity: " & udtItems(intIndex).itemQuantity
    imgCheckout.Picture = LoadPicture(udtItems(intIndex).itemPicturePath)
ElseIf intTemp <> 1 Then
        MsgBox "Item not found", , "Search"
End If
End Sub
Sub searchName()
Dim intX As Integer
Dim blnFound As Boolean
Dim intIndex As Integer
Dim strName As String
Dim intQuantity As Integer
Dim intCost As Integer
Dim intTemp As Integer
    strName = InputBox("Enter the name of the product", "Kipplex") 'Gets search input from user
    For intX = 0 To UBound(udtItems())
    If strName Like udtItems(intX).itemName = True Then 'Searches the item array for user input
        intTemp = 1
        intIndex = intX 'Puts place found in the array in a variable
    End If
    Next intX
If intTemp = 1 Then
    MsgBox "Item Found", , "Search"
    lblInfo.Caption = "Name: " & udtItems(intIndex).itemName & vbCrLf & "Cost: $" & udtItems(intIndex).itemCost & vbCrLf & "Quantity: " & udtItems(intIndex).itemQuantity
    imgCheckout.Picture = LoadPicture(udtItems(intIndex).itemPicturePath)
ElseIf intTemp <> 1 Then
        MsgBox "Item not found", , "Search"
End If
End Sub
Sub itemInfo()
'CPUs
udtItems(1).itemName = "Intel Core I7-6700K"
udtItems(1).itemCost = 420
udtItems(1).itemQuantity = 10
udtItems(1).itemPicturePath = "I7-6700K.jpg"

udtItems(2).itemName = "Intel Core I7-6700"
udtItems(2).itemCost = 400
udtItems(2).itemQuantity = 15
udtItems(2).itemPicturePath = "I7-6700.jpg"

udtItems(3).itemName = "Intel Core I5-6600K"
udtItems(3).itemPicturePath = "I5-6700.jpg"
udtItems(3).itemCost = 260
udtItems(3).itemQuantity = 25

'Motherboards

udtItems(4).itemName = "Gigabyte LGA1151"
udtItems(4).itemCost = 56
udtItems(4).itemQuantity = 30
udtItems(4).itemPicturePath = "Gigabyte.jpg"

udtItems(5).itemName = "ASUS Z170-A"
udtItems(5).itemCost = 170
udtItems(5).itemQuantity = 20
udtItems(5).itemPicturePath = "Asus.Jpg"

udtItems(6).itemCost = 100
udtItems(6).itemName = "MSI 970"
udtItems(6).itemPicturePath = "MSI.jpg"
udtItems(6).itemQuantity = 35

'Power Supplies
udtItems(7).itemCost = 40
udtItems(7).itemName = "EVGA 500"
udtItems(7).itemPicturePath = "EVGA.jpg"
udtItems(7).itemQuantity = 40

udtItems(8).itemCost = 79
udtItems(8).itemName = "Corsair CX Series"
udtItems(8).itemPicturePath = "Corsair.jpg"
udtItems(8).itemQuantity = 50

udtItems(9).itemCost = 180
udtItems(9).itemName = "EVGA Supernova"
udtItems(9).itemPicturePath = "EVGAS.jpg"
udtItems(9).itemQuantity = 80

'Hardrives
udtItems(10).itemCost = 54
udtItems(10).itemName = "Seagate 1TB HDD"
udtItems(10).itemPicturePath = "Seagate.Jpg"
udtItems(10).itemQuantity = 75

udtItems(11).itemCost = 96
udtItems(11).itemName = "WD Green 2TB HDD"
udtItems(11).itemPicturePath = "WD.Jpg"
udtItems(11).itemQuantity = 15

udtItems(12).itemCost = 600
udtItems(12).itemName = "HGST 8TB HDD"
udtItems(12).itemPicturePath = "HGST.Jpg"
udtItems(12).itemQuantity = 5

'RAM
udtItems(13).itemCost = 26
udtItems(13).itemName = "Crucial Ballistix Sport 4GB"
udtItems(13).itemPicturePath = "Crucial.Jpg"
udtItems(13).itemQuantity = 38

udtItems(14).itemName = "Vengeance LPX 64GB"
udtItems(14).itemCost = 594
udtItems(14).itemPicturePath = "CorsairV.Jpg"
udtItems(14).itemQuantity = 10

udtItems(15).itemCost = 41
udtItems(15).itemName = "Kingston HyperX FURY 8GB"
udtItems(15).itemPicturePath = "Kingston.jpg"
udtItems(15).itemQuantity = 23
End Sub
Private Sub cmdTMT_Click()
intCategory = 2
Call imgDisplay
End Sub
Private Sub cmdTRT_Click()
intCategory = 3
Call imgDisplay
End Sub

Private Sub Form_Load()
intCategory = 1
cmdBack.Caption = "Back"
frmStore.Caption = "Kipplex Hardware Store"
frmStore.WindowState = vbMaximized
lbltitle.Caption = "Welcome to the Kipplex Hardware Store. Select your item category below."
lblTLT.Caption = ""
lblTMT.Caption = ""
lblRMT.Caption = ""
lblBLT.Caption = ""
lblBRT.Caption = ""
cmdCancel.Caption = "Exit"
optName.Caption = "Name"
optQTN.Caption = "Quantity"
optCost.Caption = "Cost"
cmdCheckout.Caption = "Checkout"
lblSearchTitle.Caption = "Enter the name, cost, or quantity of the item to search"
cmdSearch.Caption = "Search"
lblTLT.Alignment = 2
lblTMT.Alignment = 2
lblRMT.Alignment = 2
lblBLT.Alignment = 2
lblBRT.Alignment = 2
lblSearchTitle.Alignment = 2
imgTLT.Stretch = True
imgTMT.Stretch = True
imgTRT.Stretch = True
imgBLT.Stretch = True
imgBRT.Stretch = True
imgCheckout.Stretch = True
fraOptions.Caption = "Options"
imgTLT.Picture = LoadPicture("Processor.jpg")
Call itemInfo
End Sub
Private Sub cmdTLT_Click()
'imgBLT.Visible = False
'lblBLT.Visible = False
'cmdBLT.Visible = False
'imgBRT.Visible = False
'lblBRT.Visible = False
'cmdBRT.Visible = False
'cmdTLT.Caption = "To View Info, Click on Photo"
intCategory = 1
Call imgDisplay
End Sub
Sub imgDisplay()
If intCategory = 1 Then
    imgTLT.Picture = LoadPicture(udtItems(7).itemPicturePath)
    imgTMT.Picture = LoadPicture(udtItems(8).itemPicturePath)
    imgTRT.Picture = LoadPicture(udtItems(9).itemPicturePath)
ElseIf intCategory = 2 Then
    imgTLT.Picture = LoadPicture(udtItems(4).itemPicturePath)
    imgTMT.Picture = LoadPicture(udtItems(5).itemPicturePath)
    imgTRT.Picture = LoadPicture(udtItems(6).itemPicturePath)
ElseIf intCategory = 3 Then
    imgTLT.Picture = LoadPicture(udtItems(10).itemPicturePath)
    imgTMT.Picture = LoadPicture(udtItems(11).itemPicturePath)
    imgTRT.Picture = LoadPicture(udtItems(12).itemPicturePath)
ElseIf intCategory = 4 Then
    imgTLT.Picture = LoadPicture(udtItems(13).itemPicturePath)
    imgTMT.Picture = LoadPicture(udtItems(14).itemPicturePath)
    imgTRT.Picture = LoadPicture(udtItems(15).itemPicturePath)
ElseIf intCategory = 5 Then
    imgTLT.Picture = LoadPicture(udtItems(1).itemPicturePath)
    imgTMT.Picture = LoadPicture(udtItems(2).itemPicturePath)
    imgTRT.Picture = LoadPicture(udtItems(3).itemPicturePath)
End If
End Sub

Private Sub imgTLT_Click()
Dim intIndex As Integer
If intCategory = 1 Then
    lblInfo.Caption = "Name: " & udtItems(6).itemName & vbCrLf & "Cost: $" & udtItems(6).itemCost & vbCrLf & "Quantity: " & udtItems(6).itemQuantity
ElseIf intCategory = 2 Then
    lblInfo.Caption = "Name: " & udtItems(3).itemName & vbCrLf & "Cost: $" & udtItems(3).itemCost & vbCrLf & "Quantity: " & udtItems(3).itemQuantity
ElseIf intCategory = 3 Then
    lblInfo.Caption = "Name: " & udtItems(6).itemName & vbCrLf & "Cost: $" & udtItems(6).itemCost & vbCrLf & "Quantity: " & udtItems(6).itemQuantity
ElseIf intCategory = 4 Then
    lblInfo.Caption = "Name: " & udtItems(9).itemName & vbCrLf & "Cost: $" & udtItems(9).itemCost & vbCrLf & "Quantity: " & udtItems(9).itemQuantity
ElseIf intCategory = 5 Then
    lblInfo.Caption = "Name: " & udtItems(12).itemName & vbCrLf & "Cost: $" & udtItems(12).itemCost & vbCrLf & "Quantity: " & udtItems(12).itemQuantity
End If
End Sub
Private Sub imgTMT_Click()
If intCategory = 1 Then
    lblInfo.Caption = "Name: " & udtItems(7).itemName & vbCrLf & "Cost: $" & udtItems(7).itemCost & vbCrLf & "Quantity: " & udtItems(7).itemQuantity
ElseIf intCategory = 2 Then
    lblInfo.Caption = "Name: " & udtItems(4).itemName & vbCrLf & "Cost: $" & udtItems(4).itemCost & vbCrLf & "Quantity: " & udtItems(4).itemQuantity
ElseIf intCategory = 3 Then
    lblInfo.Caption = "Name: " & udtItems(7).itemName & vbCrLf & "Cost: $" & udtItems(7).itemCost & vbCrLf & "Quantity: " & udtItems(7).itemQuantity
ElseIf intCategory = 4 Then
    lblInfo.Caption = "Name: " & udtItems(10).itemName & vbCrLf & "Cost: $" & udtItems(10).itemCost & vbCrLf & "Quantity: " & udtItems(10).itemQuantity
ElseIf intCategory = 5 Then
    lblInfo.Caption = "Name: " & udtItems(13).itemName & vbCrLf & "Cost: $" & udtItems(13).itemCost & vbCrLf & "Quantity: " & udtItems(13).itemQuantity
End If
End Sub
Private Sub imgTRT_Click()
Dim intIndex As Integer
If intCategory = 1 Then
    lblInfo.Caption = "Name: " & udtItems(8).itemName & vbCrLf & "Cost: $" & udtItems(8).itemCost & vbCrLf & "Quantity: " & udtItems(8).itemQuantity
ElseIf intCategory = 2 Then
    lblInfo.Caption = "Name: " & udtItems(5).itemName & vbCrLf & "Cost: $" & udtItems(5).itemCost & vbCrLf & "Quantity: " & udtItems(5).itemQuantity
ElseIf intCategory = 3 Then
    lblInfo.Caption = "Name: " & udtItems(8).itemName & vbCrLf & "Cost: $" & udtItems(8).itemCost & vbCrLf & "Quantity: " & udtItems(8).itemQuantity
ElseIf intCategory = 4 Then
    lblInfo.Caption = "Name: " & udtItems(11).itemName & vbCrLf & "Cost: $" & udtItems(11).itemCost & vbCrLf & "Quantity: " & udtItems(11).itemQuantity
ElseIf intCategory = 5 Then
    lblInfo.Caption = "Name: " & udtItems(14).itemName & vbCrLf & "Cost: $" & udtItems(14).itemCost & vbCrLf & "Quantity: " & udtItems(14).itemQuantity
End If
End Sub
Sub itemCheckout()
Dim dblCost As Double
Dim dblTotalCost As Double
Static intCounter As Integer
Const dblTax As Double = 0.08375
lstList.AddItem udtItems(intIndex).itemName
udtItems(intIndex).itemQuantity = udtItems(intIndex).itemQuantity - 1
dblCost = (dblCost + udtItems(intIndex).itemCost) * dblTax
intCounter = intCounter + 1
lblCounter.Caption = "Customer Number: " & intCounter
dblTotalCost = dblCost + udtItems(intIndex).itemCost
MsgBox "Sales Tax: " & dblTax & vbCrLf & "Total Cost: " & dblTotalCost, , "Receipt"
End Sub
Private Sub cmdCategory_Click(Index As Integer)
Index = (Index * intCategory) + 1
'lblInfo.Caption = "Name: " & udtItems(Index).itemName & vbCrLf & "Cost: $" & udtItems(Index).itemCost & vbCrLf & "Quantity: " & udtItems(Index).itemQuantity

MsgBox Index
'MsgBox udtItems(Index).itemName
End Sub
Private Sub mnuCPU_Click(Index As Integer)
intCategory = 5 'Check category
Call imgDisplay
End Sub
Private Sub mnumobo_Click(Index As Integer)
intCategory = 2
Call imgDisplay
End Sub
Private Sub mnupsu_Click(Index As Integer)
intCategory = 1
Call imgDisplay
End Sub
Private Sub mnuHDD_Click(Index As Integer)
intCategory = 3
Call imgDisplay
End Sub
Private Sub mnuRAM_Click(Index As Integer)
intCategory = 4
Call imgDisplay
End Sub

