VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LIBRARY MANAGERS...."
   ClientHeight    =   6060
   ClientLeft      =   225
   ClientTop       =   1125
   ClientWidth     =   10980
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":030A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10980
      _ExtentX        =   19368
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "1a"
            ImageIndex      =   1
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "rt"
                  Object.Tag             =   "234"
                  Text            =   "amit"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "2a"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "3a"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "4a"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "5a"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "6a"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "7a"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "8a"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "9a"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "10a"
            ImageIndex      =   16
         EndProperty
      EndProperty
      BorderStyle     =   1
      MousePointer    =   1
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   1320
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":7F95
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":83E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8839
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":8C8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":90DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":952F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9981
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":9DD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A225
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":A677
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AAC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":AF1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B375
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":B7C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":BC19
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":C06B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2280
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   17
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":C4BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":C7D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":CAF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":CE0B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":D125
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":D43F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":D759
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":DA73
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":E0ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":E767
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":EDE1
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":F45B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":FAD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":1014F
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":107C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":10E43
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":114BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuLibStructure 
      Caption         =   "&Library Structure"
      Begin VB.Menu submnuAuthor 
         Caption         =   "Author"
      End
      Begin VB.Menu submnuLanguage 
         Caption         =   "Language"
      End
      Begin VB.Menu submnuSubject 
         Caption         =   "Subject"
      End
      Begin VB.Menu submnuCourse 
         Caption         =   "Course"
      End
      Begin VB.Menu submnuLibCardType 
         Caption         =   "Library Card"
      End
      Begin VB.Menu A1 
         Caption         =   "-"
      End
      Begin VB.Menu submnuSupplier 
         Caption         =   "Supplier"
      End
      Begin VB.Menu submnuPublisher 
         Caption         =   "Publisher"
      End
      Begin VB.Menu A2 
         Caption         =   "-"
      End
      Begin VB.Menu submnuMaxAssets 
         Caption         =   "Asset Planner"
      End
      Begin VB.Menu submnuRates 
         Caption         =   "Fine "
      End
      Begin VB.Menu A3 
         Caption         =   "-"
      End
      Begin VB.Menu submnuFloorSectionAlmira 
         Caption         =   "Floor/Section/Almira"
      End
   End
   Begin VB.Menu mnuAsset 
      Caption         =   "&Asset"
      Begin VB.Menu submnuCatalogItems 
         Caption         =   "Catalogue Items"
      End
      Begin VB.Menu submnuGenAssets 
         Caption         =   "General Items"
      End
      Begin VB.Menu qwer 
         Caption         =   "-"
      End
      Begin VB.Menu submnuStationary 
         Caption         =   "Stationary"
      End
   End
   Begin VB.Menu mnuIssueReturn 
      Caption         =   "&Issue/Return"
      Begin VB.Menu submnuIssueReturn 
         Caption         =   "Issue/Return"
         Begin VB.Menu submnuAccessionWiseISSUERETURN 
            Caption         =   "Accession Number Wise"
         End
         Begin VB.Menu submnuBookTitleWsieISSUERETURN 
            Caption         =   "Book Title Wise"
         End
         Begin VB.Menu submnuCDIssue 
            Caption         =   "CD"
         End
      End
      Begin VB.Menu submnuMembers 
         Caption         =   "Members"
      End
      Begin VB.Menu submnuLibCard 
         Caption         =   "Library Card"
      End
      Begin VB.Menu bvbv 
         Caption         =   "-"
      End
      Begin VB.Menu submnuShelfArrangement 
         Caption         =   "Shelf Arrangement"
      End
   End
   Begin VB.Menu mnuCatalog 
      Caption         =   "&Cataloguing"
      Begin VB.Menu submnuGeneralCatalog 
         Caption         =   "General Catalog"
      End
      Begin VB.Menu submnuNewAdditionList 
         Caption         =   "New Addition List"
      End
   End
   Begin VB.Menu mnuAcquisition 
      Caption         =   "&Acquisition"
      Enabled         =   0   'False
      Begin VB.Menu submnuOrderEntry 
         Caption         =   "Order Entry"
      End
      Begin VB.Menu submnuOrderPlace 
         Caption         =   "Order Place"
      End
      Begin VB.Menu aaa 
         Caption         =   "-"
      End
      Begin VB.Menu submnuReceive 
         Caption         =   "Order Receive"
      End
      Begin VB.Menu submnuOrderPayment 
         Caption         =   "Order Payment"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu subMnuBookSearch 
         Caption         =   "Book"
      End
      Begin VB.Menu submnuPubBookSEarch 
         Caption         =   "Publisher Wise Book Search"
      End
      Begin VB.Menu subMnuCDSearch 
         Caption         =   "CD"
      End
      Begin VB.Menu subMnuJournalSearch 
         Caption         =   "Jounrnal"
      End
   End
   Begin VB.Menu mnuQueries 
      Caption         =   "&Queries"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuRepot 
      Caption         =   "&Report"
      Begin VB.Menu mnuBookRep 
         Caption         =   "Book"
         Begin VB.Menu mnuAccessRep 
            Caption         =   "Accession Wise"
         End
         Begin VB.Menu mnusubRep 
            Caption         =   "Subject Wise"
         End
         Begin VB.Menu PublisherAutherTitle 
            Caption         =   "PublisherAutherTitle"
         End
         Begin VB.Menu submnuISSUELOST 
            Caption         =   "Issue/Lost"
         End
         Begin VB.Menu submnuIssuedBookReport 
            Caption         =   "Issued Book Report"
         End
         Begin VB.Menu submnuIssuedBookReportMemberWise 
            Caption         =   "Issued Book Report Member Wise"
         End
         Begin VB.Menu submnuReturnedBookReportMemberWise 
            Caption         =   "Returned Book Report Member Wise"
         End
         Begin VB.Menu mnuRefRep 
            Caption         =   "Reference Book"
         End
         Begin VB.Menu submnuJournal 
            Caption         =   "Journal"
            Begin VB.Menu submnuJournalAccessionWise 
               Caption         =   "Accession Wise"
            End
            Begin VB.Menu submnuJournalSubject 
               Caption         =   "Subject Wise"
            End
         End
      End
      Begin VB.Menu mnuRepMem 
         Caption         =   "Member"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu submnuLogin 
         Caption         =   "Login"
         Shortcut        =   ^I
      End
      Begin VB.Menu submnuLogOut 
         Caption         =   "Log Out"
         Shortcut        =   ^U
      End
      Begin VB.Menu asdasd 
         Caption         =   "-"
      End
      Begin VB.Menu submnuBackUp 
         Caption         =   "Back Up"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnustatus 
         Caption         =   "Library Status"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help?"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu UPDATE 
      Caption         =   "&UPDATE"
   End
   Begin VB.Menu mnuLoginName 
      Caption         =   ""
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c1 As New ClassMaster

Private Sub MDIForm_Load()
Load MDIForm1
MDIForm1.Show
'frmAuthorEntry.MDIChild = False
'Load frmAuthorEntry

End Sub
Private Sub mnuArrangement_Click()
Load frmShelfManagement
frmShelfManagement.Show
End Sub

Private Sub mnuAssets_Click()
Load frmRequiredAssets
frmRequiredAssets.Show
End Sub

Private Sub mnuAuthor_Click()
Load frmAuthorEntry
frmAuthorEntry.Show
End Sub


Private Sub MDIForm_Terminate()
End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuAccessRep_Click()
Load frmbookReportAccessionNo
frmbookReportAccessionNo.Show
End Sub

Private Sub mnuExit_Click()
a = MsgBox("Do You Want to Exit", vbQuestion + vbYesNo, "Exit")
If a = vbYes Then
'Unload Me
End
End If
End Sub

Private Sub mnuHelp_Click()
Load FrmHelp
FrmHelp.Show
End Sub

Private Sub mnuRefRep_Click()
Load frmReferenceReport
frmReferenceReport.Show
End Sub

Private Sub mnuRepMem_Click()
Load frmMemberReport
frmMemberReport.Show
End Sub

Private Sub mnustatus_Click()
Load frmstatus
frmstatus.Image4.Visible = False
frmstatus.Show
End Sub

Private Sub mnusubRep_Click()
Load frmBookReport
frmBookReport.Show
End Sub
Private Sub PublisherAutherTitle_Click()
'Load frmWS
'frmWS.Show
Load Form3
Form3.Show
End Sub

Private Sub submnuAccessionWiseISSUERETURN_Click()
Load Form2
Form2.Show
End Sub

Private Sub submnuBookIssue_Click()

End Sub

Private Sub submnuBookTitleWsieISSUERETURN_Click()
Load frmIssueReturn
frmIssueReturn.Show
End Sub

Private Sub submnuCDIssue_Click()
Load frmIssuereturnCD
frmIssuereturnCD.Show

End Sub

Private Sub subMnuCDSearch_Click()
Load frmCDSearch
frmCDSearch.Show
End Sub

Private Sub submnuLibAssets_Click()

End Sub



Private Sub submnuLogin_Click()
frmLogin_Lib.Show
End Sub

Private Sub submnuOrder_Click()

End Sub

Private Sub submnuPlacement_Click()
End Sub

Private Sub submnuLogOut_Click()
Call c1.proLogOut
End Sub
Private Sub submnuNewAdditionList_Click()
'Load frmNewItemList
'frmNewItemList.Show
End Sub

Private Sub submnuShipping_Click()
End Sub

Private Sub submnuTenderPlace_Click()
End Sub









Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Toolbar1(BackColor) = vbRed
End Sub

Private Sub UPDATE_Click()
'Load frmupdateEntry
'frmupdateEntry.Show
End Sub

