VERSION 5.00
Begin VB.Form frmStore 
   ClientHeight    =   12195
   ClientLeft      =   -2265
   ClientTop       =   705
   ClientWidth     =   21360
   LinkTopic       =   "Form1"
   ScaleHeight     =   12195
   ScaleWidth      =   21360
   Begin VB.Frame fraProducts 
      Caption         =   "Frame1"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   4800
      Width           =   11055
      Begin VB.OptionButton optThree 
         Caption         =   "Option1"
         Height          =   375
         Left            =   8160
         TabIndex        =   21
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optTwo 
         Caption         =   "Option1"
         Height          =   375
         Left            =   4680
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton optOne 
         Caption         =   "Option1"
         Height          =   375
         Left            =   960
         TabIndex        =   19
         Top             =   240
         Width           =   1215
      End
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
   Begin VB.CommandButton cmdAddToCart 
      Caption         =   "Command7"
      Height          =   735
      Left            =   12480
      TabIndex        =   10
      Top             =   7200
      Width           =   3975
   End
   Begin VB.CommandButton cmdCheckout 
      Caption         =   "Command6"
      Height          =   1095
      Left            =   14520
      TabIndex        =   8
      Top             =   9600
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
      Left            =   18000
      TabIndex        =   17
      Top             =   11160
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
   Begin VB.Label lblBRT 
      Caption         =   "Label1"
      Height          =   615
      Left            =   3840
      TabIndex        =   5
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Image imgTwo 
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
   Begin VB.Label lblBLT 
      Caption         =   "Label1"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   5760
      Width           =   3495
   End
   Begin VB.Image imgThree 
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
   Begin VB.Image imgOne 
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
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Index           =   2
      End
   End
   Begin VB.Menu mnupsu 
      Caption         =   "Power Supplies"
      Index           =   1
   End
   Begin VB.Menu mnumobo 
      Caption         =   "Motherboards"
      Index           =   2
   End
   Begin VB.Menu mnuHDD 
      Caption         =   "Hard Drives"
      Index           =   3
   End
   Begin VB.Menu mnuRAM 
      Caption         =   "RAM"
      Index           =   4
   End
   Begin VB.Menu mnuCPU 
      Caption         =   "CPU's"
      Index           =   5
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
      Index           =   6
      Begin VB.Menu mnuQtnSearh 
         Caption         =   "Quantity"
         Index           =   7
      End
      Begin VB.Menu mnuSearchCost 
         Caption         =   "Cost"
         Index           =   8
      End
      Begin VB.Menu mnuSeachName 
         Caption         =   "Name"
         Index           =   9
      End
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
Private udtItems(14) As itemInfo
Private Type Receipt
    totalCost As Double
    salesTax  As Double
End Type
Private udtReceipt As Receipt
Private intIndex As Integer
Private intItemOne As Integer
Private intItemTwo As Integer
Private intItemThree As Integer
Private intCounter As Integer
Private dblCost As Double
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmdCheckout_Click()
Call itemCheckout
End Sub
Sub itemCheckout()
udtReceipt.totalCost = dblCost
udtReceipt.salesTax = 0.08375
intCounter = intCounter + 1
lblCounter.Caption = "Customer number: " & intCounter
udtReceipt.totalCost = udtReceipt.totalCost + (udtReceipt.totalCost * udtReceipt.salesTax)
MsgBox "Cost before tax: $" & dblCost & vbCrLf & "Sales Tax: " & udtReceipt.salesTax & vbCrLf & "Total Cost: $" & udtReceipt.totalCost, , "Receipt"
MsgBox "Cost: $" & udtReceipt.totalCost, , "Receipt"
lstList.Clear
udtReceipt.totalCost = 0
dblCost = 0
imgOne.Picture = LoadPicture("")
imgTwo.Picture = LoadPicture("")
imgThree.Picture = LoadPicture("")
lblInfo.Caption = ""
optOne.Value = False
optTwo.Value = False
optThree.Value = False
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
udtItems(0).itemName = "Intel Core I7-6700K"
udtItems(0).itemCost = 420
udtItems(0).itemQuantity = 10
udtItems(0).itemPicturePath = "I7-6700K.jpg"

udtItems(1).itemName = "Intel Core I7-6700"
udtItems(1).itemCost = 400
udtItems(1).itemQuantity = 15
udtItems(1).itemPicturePath = "I7-6700.jpg"

udtItems(2).itemName = "Intel Core I5-6600K"
udtItems(2).itemPicturePath = "I5-6700.jpg"
udtItems(2).itemCost = 260
udtItems(2).itemQuantity = 25

'Motherboards

udtItems(3).itemName = "Gigabyte LGA1151 Intell H110 Micro ATX DDR4"
udtItems(3).itemCost = 56
udtItems(3).itemQuantity = 30
udtItems(3).itemPicturePath = "Gigabyte.jpg"

udtItems(4).itemName = "ASUS Z170-A"
udtItems(4).itemCost = 170
udtItems(4).itemQuantity = 20
udtItems(4).itemPicturePath = "Asus.Jpg"

udtItems(5).itemCost = 100
udtItems(5).itemName = "MSI 970"
udtItems(5).itemPicturePath = "MSI.jpg"
udtItems(5).itemQuantity = 35

'Power Supplies
udtItems(6).itemCost = 40
udtItems(6).itemName = "EVGA 500"
udtItems(6).itemPicturePath = "EVGA.jpg"
udtItems(6).itemQuantity = 40

udtItems(7).itemCost = 79
udtItems(7).itemName = "Corsair CX Series"
udtItems(7).itemPicturePath = "Corsair.jpg"
udtItems(7).itemQuantity = 50

udtItems(8).itemCost = 180
udtItems(8).itemName = "EVGA Supernova"
udtItems(8).itemPicturePath = "EVGAS.jpg"
udtItems(8).itemQuantity = 80

'Hardrives
udtItems(9).itemCost = 54
udtItems(9).itemName = "Seagate 1TB HDD"
udtItems(9).itemPicturePath = "Seagate.Jpg"
udtItems(9).itemQuantity = 75

udtItems(10).itemCost = 96
udtItems(10).itemName = "WD Green 2TB HDD"
udtItems(10).itemPicturePath = "WD.Jpg"
udtItems(10).itemQuantity = 15

udtItems(11).itemCost = 600
udtItems(11).itemName = "HGST 8TB HDD"
udtItems(11).itemPicturePath = "HGST.Jpg"
udtItems(11).itemQuantity = 5

'RAM
udtItems(12).itemCost = 26
udtItems(12).itemName = "Crucial Ballistix Sport 4GB"
udtItems(12).itemPicturePath = "Crucial.Jpg"
udtItems(12).itemQuantity = 38

udtItems(13).itemName = "Vengeance LPX 64GB"
udtItems(13).itemCost = 594
udtItems(13).itemPicturePath = "CorsairV.Jpg"
udtItems(13).itemQuantity = 10

udtItems(14).itemCost = 41
udtItems(14).itemName = "Kingston HyperX FURY 8GB"
udtItems(14).itemPicturePath = "Kingston.jpg"
udtItems(14).itemQuantity = 23
End Sub
Private Sub Form_Load()
intCounter = 1
lblCounter.Caption = "Customer: " & intCounter
cmdCheckout.Caption = "Checkout"
lblInfo.Caption = ""
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
optOne.Caption = "Item One"
optTwo.Caption = "Item Two"
optThree.Caption = "Item Three"
fraProducts.Caption = "Products"
cmdCheckout.Caption = "Checkout"
cmdAddToCart.Caption = "Add to Cart"
lblSearchTitle.Caption = "Enter the name, cost, or quantity of the item to search"
cmdSearch.Caption = "Search"
lblTLT.Alignment = 2
lblTMT.Alignment = 2
lblRMT.Alignment = 2
lblBLT.Alignment = 2
lblBRT.Alignment = 2
lblSearchTitle.Alignment = 2
imgOne.Stretch = True
imgTwo.Stretch = True
imgThree.Stretch = True
imgCheckout.Stretch = True
fraOptions.Caption = "Options"
Call itemInfo
End Sub
Sub imgDisplay(ByVal intOne As Integer, ByVal intTwo As Integer, ByVal intThree As Integer)
imgOne.Picture = LoadPicture(udtItems(intOne).itemPicturePath)
imgTwo.Picture = LoadPicture(udtItems(intTwo).itemPicturePath)
imgThree.Picture = LoadPicture(udtItems(intThree).itemPicturePath)
End Sub
Private Sub cmdAddToCart_Click()
Call AddToCart
End Sub
Sub AddToCart()
Dim dblTotalCost As Double
Static intCounter As Integer
Const dblTax As Double = 0.08375
Dim intQuantity As Integer
Dim strTemp As String
intQuantity = udtItems(intIndex).itemQuantity
strTemp = MsgBox("Add To Cart?", vbYesNo, "Cart")
If strTemp = vbYes Then
    If intQuantity >= 1 Then
    udtItems(intIndex).itemQuantity = udtItems(intIndex).itemQuantity - 1
    lstList.AddItem udtItems(intIndex).itemName
    dblCost = dblCost + udtItems(intIndex).itemCost
    ElseIf intQuantity >= 0 Then
        MsgBox "Out of stock", , "Stock"
    End If
ElseIf strTemp = vbNo Then
    MsgBox "Canceled", , "Cart"
End If
lblInfo.Caption = "Name: " & udtItems(intIndex).itemName & vbCrLf & "Cost: $" & udtItems(intIndex).itemCost & vbCrLf & "Quantity: " & udtItems(intIndex).itemQuantity
End Sub

Private Sub mnuAbout_Click(Index As Integer)
MsgBox ("Created by Tarek Elkheir and Jonathan Bubloski" & vbCrLf & "Mr. Nickels period 2" & vbCrLf & "Used with permission from Kipplex, visit our real store at kipplex.com")
End Sub

Private Sub mnuCPU_Click(Index As Integer)
Call imgDisplay(0, 1, 2)
intItemOne = 0
intItemTwo = 1
intItemThree = 2
End Sub
Private Sub mnumobo_Click(Index As Integer)
Call imgDisplay(3, 4, 5)
intItemOne = 3
intItemTwo = 4
intItemThree = 5
End Sub
Private Sub mnupsu_Click(Index As Integer)
Call imgDisplay(6, 7, 8)
intItemOne = 6
intItemTwo = 7
intItemThree = 8
End Sub
Private Sub mnuHDD_Click(Index As Integer)
Call imgDisplay(9, 10, 11)
intItemOne = 9
intItemTwo = 10
intItemThree = 11
End Sub

Private Sub mnuQtnSearh_Click(Index As Integer)
Call searchQuantity
End Sub

Private Sub mnuRAM_Click(Index As Integer)
Call imgDisplay(12, 13, 14)
intItemOne = 12
intItemTwo = 13
intItemThree = 14
End Sub

Private Sub mnuSeachName_Click(Index As Integer)
Call searchName
End Sub

Private Sub mnuSearchCost_Click(Index As Integer)
Call searchCost
End Sub

Private Sub optOne_Click()
lblInfo.Caption = "Name: " & udtItems(intItemOne).itemName & vbCrLf & "Cost: $" & udtItems(intItemOne).itemCost & vbCrLf & "Quantity: " & udtItems(intItemOne).itemQuantity
intIndex = intItemOne
End Sub

Private Sub optThree_Click()
lblInfo.Caption = "Name: " & udtItems(intItemThree).itemName & vbCrLf & "Cost: $" & udtItems(intItemThree).itemCost & vbCrLf & "Quantity: " & udtItems(intItemThree).itemQuantity
intIndex = intItemThree
End Sub

Private Sub optTwo_Click()
lblInfo.Caption = "Name: " & udtItems(intItemTwo).itemName & vbCrLf & "Cost: $" & udtItems(intItemTwo).itemCost & vbCrLf & "Quantity: " & udtItems(intItemTwo).itemQuantity
intIndex = intItemTwo
End Sub
