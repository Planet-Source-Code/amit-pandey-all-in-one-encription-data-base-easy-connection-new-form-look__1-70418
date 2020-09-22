Attribute VB_Name = "Module_Institute"
 'Option Explicit
 Public Global_CompanyId As Integer, Global_financialYear, Global_Institute_Name As String
 Public con As Connection, SortResult As Boolean
 Public AccessionEnd, AccessionStart As Integer
 Dim ascii As Integer
 Public FlagLostGlobal As Boolean
 Public FlagDamageGlobal As Boolean
 Public FlagRepairGlobal As Boolean
 
 Public Sub Main()
  Set con = New Connection
'  con.CursorLocation = adUseClient
'  con.Open "PROVIDER=MSDASQL;dsn=Library;uid=;pwd=uilibstarui;database=Library;"

'  Set con = New Connection
'    con.CursorLocation = adUseClient
'    con.Open "PROVIDER=MSDASQL;dsn=Libstar;uid=;pwd=uilibstarui;database=Library;"

  con.Provider = "microsoft.jet.oledb.4.0; data source =" & App.Path & "\library.mdb;jet OLEDB:Database Password=uilibstarui"
  con.Open
    'Global_financialYear = 2004 - 2005
    Global_CompanyId = 1
    FlagLostGlobal = False
    FlagDamageGlobal = False
    FlagRepairGlobal = False
 'Call TrailVersion
 Load frmstatus
 'Load frmSplash
frmstatus.Show
' Load Form2
' Form2.Show
 'frmSplash.Show
 'Load Form3
 'Form3.Show
 'Load MDIForm1
 'MDIForm1.Show
' Load Form5
' Form5.Show
'Load Form6
' Form6.Show
 End Sub
 
 Sub TrailVersion()
 On Error GoTo Er
Dim rsbook, rsBookAuthor As Recordset
Set rsbook = New Recordset
rsbook.Open "SELECT * from book_master ", con, adOpenKeyset, adLockOptimistic
If rsbook.RecordCount > 500 Then
   MsgBox "Trail Version Over Limit Is Completed"
   End
End If

'-----------------------------------Error Handling-----------------------------
Exit Sub
Er:
MsgBox Err.Description
'-----------------------------------------------------------------------------------------

 End Sub

Public Function funTextLock(aKeyAscii As Integer) As Integer
   If aKeyAscii >= 48 And aKeyAscii <= 57 Then
     funTextLock = aKeyAscii
   Else
     aKeyAscii = 0
     funTextLock = aKeyAscii
End If
  
 End Function
Public Function funNumberLock(aKeyAscii As Integer) As Integer
   If aKeyAscii >= 48 And aKeyAscii <= 57 Then
     funNumberLock = 0
   Else
     funNumberLock = aKeyAscii
 End If
End Function
 
 Public Sub address_display(lvw As ListView, frm As Form, ind As Integer)
  ''MsgBox ("hello")
  Dim str, store, ch As String
  str = Trim(lvw.SelectedItem.SubItems(ind))
  Dim strlen As Integer
  strlen = Len(str)
  Dim i, co As Integer
  co = 1
  store = ""
  For i = 1 To strlen
      ch = Mid(str, i, 1)
      If Not ch = "," Then
             store = store + ch
      End If
      If ch = "," Then
         If co = 1 Then
            frm.txtHouse.Text = store
            store = ""
         End If
         If co = 2 Then
            frm.txtStreet.Text = store
            store = ""
         End If
         If co = 3 Then
            frm.txtLocality.Text = store
            store = ""
         End If
         If co = 4 Then
            frm.txtpin.Text = store
            store = ""
         End If
      co = co + 1
      End If
      ''co = co + 1
  Next i
End Sub
Public Sub city(frm As Form)
 Dim rs1, rs2 As Recordset
 Set rs1 = New Recordset
 rs1.Open "select state from state_master where state_id in (select state_id from city_master where city='" & frm.Combocity.Text & "')", con, adLockOptimistic, adOpenKeyset
 If rs1.RecordCount = 0 Then
   Exit Sub
 Else
 frm.Combostate.Text = rs1.Fields(0).Value
 End If
 Set rs2 = New Recordset
 rs2.Open "select country from country_master where country_id in (select country_id from state_master where state='" & frm.Combostate.Text & "')", con, adLockOptimistic, adOpenKeyset
 frm.Combocountry.Text = rs2.Fields(0).Value
End Sub
 Public Function StringFormat(ByVal stringIn As String) As String
  Dim S, c As String, i, l As Integer
S = Trim(stringIn)
l = Len(S)
For i = 1 To l
    ascii = Asc(Mid(S, i, 1))
   If i = 1 Then
          If ascii >= 97 And ascii <= 122 Then
              ascii = ascii - 32
          End If
               Mid(S, i, 1) = Chr(ascii)
   Else
 'SDFSDFSFSDF
          If ascii = 32 Then
              'Mid(S, i, 1) = Chr(ascii)
              i = i + 1
              Dim aa As Integer
              aa = Asc(Mid(S, i, 1))
                  If aa >= 97 And aa <= 122 Then
                    aa = aa - 32
                    Mid(S, i, 1) = Chr(aa)
                  End If
           Else
        
             If ascii >= 65 And ascii <= 91 Then
             ascii = ascii + 32
             End If
             Mid(S, i, 1) = Chr(ascii)
          End If
          
   End If
   
Next
StringFormat = Trim(S)
 End Function
 Public Function DuplicateCheck1(ByVal tablename As String, ByVal SearchField As String, ByVal searchvalue As String) As Boolean
Dim rsopen As Recordset
Set rsopen = New Recordset
rsopen.Open "select * from " & tablename & " where " & SearchField & "= " & Trim(StringFormat(searchvalue)) & "  and company_Id=" & Global_CompanyId & "  And del_Flag = false", con, adOpenKeyset, adLockOptimistic
If rsopen.RecordCount = 0 Then
 'MsgBox ("record not found")
  DuplicateCheck1 = False
Else
 'MsgBox ("record exist")
 DuplicateCheck1 = True
End If
End Function
Public Sub additems(ByVal Id_Field As String, ByVal fieldname As String, ByVal tablename As String, cBox As ComboBox)
Dim i As Integer
Dim Rs As Recordset
Set Rs = New Recordset
Rs.Open "select distinct(" & fieldname & ")," & Id_Field & " from " & tablename & "  where del_flag=0 and company_Id=" & Global_CompanyId & " order by " & fieldname & " asc", con, adOpenKeyset, adLockOptimistic
If Rs.RecordCount = 0 Then
Exit Sub
Else
 Rs.MoveFirst
 While Not Rs.EOF
  cBox.List(i) = Rs.Fields(0).Value
  cBox.ItemData(i) = Rs.Fields(1).Value
  i = i + 1
  Rs.MoveNext
 Wend
End If
End Sub

Public Function DuplicateCheck(ByVal tablename As String, ByVal SearchField As String, ByVal searchvalue As String) As Boolean
Dim rsopen As Recordset
Set rsopen = New Recordset
rsopen.Open "select * from " & tablename & " where " & SearchField & "= '" & Trim(StringFormat(searchvalue)) & "'  and company_Id=" & Global_CompanyId & "  And del_Flag = false", con, adOpenKeyset, adLockOptimistic
If rsopen.RecordCount = 0 Then
 'MsgBox ("record not found")
  DuplicateCheck = False
Else
 'MsgBox ("record exist")
 DuplicateCheck = True
End If
End Function

Public Function Country_State_City_Entry(country1 As ComboBox, state1 As ComboBox, city1 As ComboBox) As Integer
 Dim strCountry, strState, strCity As String
 Dim rs1, rsSearch1, rs2, rsSearch2, rs_2, rs3, rsSearch3, rs_31, rs_32 As Recordset
 Set rs1 = New Recordset
 Set rsSearch1 = New Recordset
 Set rs2 = New Recordset
 Set rsSearch2 = New Recordset
 Set rs3 = New Recordset
 Set rsSearch3 = New Recordset
 Set rs_2 = New Recordset
 Set rs_31 = New Recordset
 Set rs_32 = New Recordset
 
''------------Insertion into Country_master---------------
 rsSearch1.Open "select country from country_master where country='" & StringFormat(country1.Text) & "'", con, adLockOptimistic, adOpenKeyset
 If rsSearch1.RecordCount = 0 Then
    rs1.Open "select Max(country_id) from country_master", con, adLockOptimistic, adOpenKeyset
    Dim countryID As Integer
    countryID = rs1.Fields(0).Value
    If countryID = 0 Then
       countryID = 1
    Else
    countryID = countryID + 1
    End If
    strCountry = "insert into country_master values(" & country_id & ",'" & StringFormat(country1.Text) & "')"
    con.Execute (strCountry)
 End If

 
''---------Insertion Into State_Master--------------------
 rsSearch2.Open "select state from state_master where state='" & StringFormat(state1.Text) & "' and country_id in (select country_id from country_master where country='" & Trim(country1.Text) & "')", con, adLockOptimistic, adOpenKeyset
 If rsSearch2.RecordCount = 0 Then
    rs2.Open "select Max(state_id) from state_master", con, adLockOptimistic, adOpenKeyset
    Dim stateID As Integer
    stateID = rs2.Fields(0).Value
    If stateID = 0 Then
       stateID = 1
    Else
       stateID = stateID + 1
    End If
    rs_2.Open "select country_id from country_master where country='" & StringFormat(country1) & "'", con, adLockOptimistic, adOpenKeyset
    Dim con_id As Integer
    con_id = rs_2.Fields(0).Value
    strState = "insert into state_master values(" & stateID & "," & con_id & ",'" & StringFormat(state1.Text) & "')"
    con.Execute (strState)
 End If
 
 
''--------------Insertion into City_master----------------
rsSearch3.Open "select city from city_master where city='" & StringFormat(city1.Text) & "' and state_id in(select state_id from state_master where state='" & StringFormat(state1.Text) & "' and country_id in(select country_id from country_master where country='" & StringFormat(country1.Text) & "'))", con, adLockOptimistic, adOpenKeyset
If rsSearch3.RecordCount = 0 Then
    Dim stateID2, cityID As Integer
    rs3.Open "select state_id from state_master where state='" & StringFormat(state1.Text) & "' and country_id in(select country_id from country_master where country='" & StringFormat(country1.Text) & "')", con, adLockOptimistic, adOpenKeyset
    stateID2 = rs3.Fields(0).Value
    rs_31.Open "select Max(city_id) from city_master ", con, adLockOptimistic, adOpenKeyset
    cityID = rs_31.Fields(0).Value
    If cityID = 0 Then
       cityID = 1
    Else
       cityID = cityID + 1
    End If
    strCity = "insert into city_master values(" & cityID & "," & stateID2 & ",'" & StringFormat(city1.Text) & "')"
    con.Execute (strCity)
 End If
 
 '----------------selection of city id'------------------------------------------------------------
Dim rsSearchcityID As Recordset
Set rsSearchcityID = New Recordset
 rsSearchcityID.Open "select city_id from city_master where city='" & StringFormat(city1.Text) & "' and state_id in(select state_id from state_master where state='" & StringFormat(state1.Text) & "' and country_id in(select country_id from country_master where country='" & StringFormat(country1.Text) & "'))", con, adLockOptimistic, adOpenKeyset
 cityID = rsSearchcityID.Fields(0).Value
 
 Country_State_City_Entry = city_Id 'return city_Id
End Function
 
'---for setting of country state and city in a single click of CityCombo-------------
Public Sub select_City(City_C As ComboBox, State_C As ComboBox, Country_C As ComboBox)
Dim i As Integer
Dim rsState, rsCountry As Recordset
 Set rsState = New Recordset
 rsState.Open "select state from state_master where state_id in (select state_id from city_master where city='" & StringFormat(City_C.Text) & "')", con, adLockOptimistic, adOpenKeyset
 If rsState.RecordCount = 0 Then
 Else
   State_C.Clear
 If rsState.RecordCount > 1 Then
    While Not rsState.EOF
     State_C.List(i) = rsState.Fields(0).Value
     If i = 0 Then 'for selecting first record
        Value = rsState.Fields(0).Value
     End If
       i = i + 1
      rsState.MoveNext
     Wend
     State_C.Text = Value
 Else
    State_C.Text = rsState.Fields(0).Value
 
 End If
 End If
 '-------------------select country-----------------------------------------------------
 Set rsCountry = New Recordset
 rsCountry.Open "select country from country_master where country_id in (select country_id from state_master where state='" & StringFormat(State_C.Text) & "')", con, adLockOptimistic, adOpenKeyset
 If rsCountry.RecordCount = 0 Then
 Else
    Country_C.Clear
   If rsCountry.RecordCount > 1 Then
      While Not rsCountry.EOF
      Combostate.List(i) = rsCountry.Fields(0).Value
      If i = 0 Then 'for selecting first record
      Value = rsCountry.Fields(0).Value
      End If
       i = i + 1
       rsCountry.MoveNext
      Wend
      Country_C.Text = Value
   Else
      Country_C.Text = rsCountry.Fields(0).Value
  End If
 End If
End Sub

Public Function ID_Generator(ByVal Table As String, ByVal fieldname As String, ByVal prefix As String) As String
Dim rs1 As Recordset
Set rs1 = New Recordset
rs1.Open "SElect * from " & Table & " where company_Id=" & Global_CompanyId & "", con, adOpenKeyset, adLockOptimistic
If rs1.RecordCount <= 0 Then
   ID_Generator = 1
   Exit Function
End If
  Dim Rs As Recordset
  Set Rs = New Recordset
  Rs.Open "select max(" & fieldname & ") from " & Table & "", con, adLockOptimistic, adOpenKeyset
  Dim ID As String
  If Rs.RecordCount = 0 Then
     ID = prefix + 1
  ElseIf prefix = Trim("") Then
  ID = Rs.Fields(0).Value
  ID = ID + 1
  Else
  ID = prefix + str(Rs.Fields(0).Value)
  End If
  ID_Generator = Trim(ID)
  
End Function
Public Sub SetCompany_YEAR(Comp_name As String, financial_Year As String)
Dim Rs As Recordset
Set Rs = New Recordset
Rs.Open "select company_id from company_master where company_name='" & Trim(Comp_name) & "' and del_flag=0", con, adOpenKeyset, adLockOptimistic
If Rs.RecordCount = 0 Then
 MsgBox "Company Does Not Exist", vbInformation + vbOKOnly
 Else
Global_CompanyId = Rs.Fields(0).Value
Global_financialYear = financial_Year
End If
End Sub
Public Sub ListViewSort(lstview As ListView, ColumnHeader As MSComctlLib.ColumnHeader)
'Dim columnHeader As ColumnHeaders
lstview.SortKey = ColumnHeader.SubItemIndex
Load frmSplashsort1
frmSplashsort1.Show vbModal
    If SortResult = True Then
    lstview.SortOrder = lvwAscending
    Else
    lstview.SortOrder = lvwDescending
    End If
lstview.Sorted = True

lstview.Sorted = False

End Sub

Public Function select_ID(ByVal Id_Field As String, ByVal tablename As String, ByVal SearchField As String, ByVal searchvalue As String) As Integer
   Dim Rs As Recordset
   Set Rs = New Recordset
   Rs.Open "select " & Id_Field & " from " & tablename & " where " & SearchField & " ='" & searchvalue & "'", con, adOpenKeyset, adLockOptimistic
   If Rs.RecordCount = 0 Then
     select_ID = 0
     
   Else
   select_ID = Rs.Fields(0).Value
   End If
End Function
Private Function funIsText(txt As TextBox) As Boolean
 S = txt.Text
 i = Mid(S, 1, 1)
 If Asc(i) < 65 And Asc(i) > 91 Then
    MsgBox "input Texts", vbInformation + vbOKOnly, "Appraisal System"
    txt.SetFocus
    txt = Clear
    
    funIsText = False
 ElseIf Asc(i) > 97 And Asc(i) < 122 Then
    MsgBox "input Texts", vbInformation + vbOKOnly, "Appraisal System"
    txt.SetFocus
    txt = Clear
    funIsText = False
  Else
    funIsText = True
  End If
End Function
 Public Function city_select(cit As Integer) As String
        Dim Rs As Recordset
        Set Rs = New Recordset
        Rs.Open "select  city from city_master where city_id=" & cit & "", con, adLockOptimistic, adOpenKeyset
        If Rs.RecordCount = 0 Then
           Exit Function
        End If
        city_select = Rs.Fields(0).Value
 End Function
        
  Public Function course_select(br As Integer) As String
        Dim Rs As Recordset
        Set Rs = New Recordset
        Rs.Open "select course_name from course_master where course_id in(select course_id from branch_master where branch_id=" & br & ")", con, adLockOptimistic, adOpenKeyset
        If Rs.RecordCount = 0 Then
           Exit Function
        End If
        course_select = Rs.Fields(0).Value
 End Function
Function proitemType_or_ItemID(Optional ByVal varitem_id As String, Optional item_type As String) As String
    
    Dim rsitemtype_or_itemid As Recordset
    Set rsitemtype_or_itemid = New Recordset
    If item_type = "" Then
        rsitemtype_or_itemid.Open "select Item_type from Required_Asset_Master where item_id=" & Val(varitem_id) & "", con, adOpenKeyset, adLockOptimistic
   Else
       rsitemtype_or_itemid.Open "select Item_id from Required_Asset_Master where Item_type='" & item_type & "'", con, adOpenKeyset, adLockOptimistic
   End If
    If rsitemtype_or_itemid.RecordCount = 0 Then
        MsgBox "NO ITEM TYPW IN REQUIRED_ASSET_MASTER TABLE CORRESPONDING TO ITEM_ID"
    Else
        proitemType_or_ItemID = rsitemtype_or_itemid.Fields(0).Value
    End If

End Function

             
