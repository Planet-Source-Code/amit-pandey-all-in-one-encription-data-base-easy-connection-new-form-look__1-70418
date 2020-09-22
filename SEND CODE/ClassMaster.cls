VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Public Sub AddItemCityStateCountry(ByVal fieldname As String, ByVal tablename As String, cBox As ComboBox)
Dim i As Integer
Dim Rs As Recordset
Set Rs = New Recordset
Rs.Open "select distinct(" & fieldname & ") from " & tablename & "  order by " & fieldname & " asc", con, adOpenKeyset, adLockOptimistic
If Rs.RecordCount = 0 Then
Exit Sub
Else
 Rs.MoveFirst
 While Not Rs.EOF
  cBox.List(i) = Rs.Fields(0).Value
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
Sub proLogOut()
MDIForm1.mnuLibStructure.Enabled = False
MDIForm1.mnuIssueReturn.Enabled = False
MDIForm1.submnuShelfArrangement.Enabled = False 'vikjk
MDIForm1.mnuAsset.Enabled = False
MDIForm1.mnuCatalog.Enabled = False
MDIForm1.mnuAcquisition.Enabled = False
MDIForm1.mnuSearch.Enabled = False
MDIForm1.mnuQueries.Enabled = False
MDIForm1.mnuRepot.Enabled = False
MDIForm1.submnuBackUp.Enabled = False
End Sub
Sub proUserLogin()
MDIForm1.mnuLibStructure.Enabled = False
MDIForm1.mnuIssueReturn.Enabled = True
MDIForm1.submnuShelfArrangement.Enabled = False 'vikjk
MDIForm1.mnuAsset.Enabled = True
MDIForm1.mnuCatalog.Enabled = True
   MDIForm1.submnuOrderEntry.Enabled = False
   MDIForm1.submnuOrderPayment.Enabled = False
   MDIForm1.submnuOrderPlace.Enabled = False
   MDIForm1.submnuReceive.Enabled = True
MDIForm1.mnuAcquisition.Enabled = False
MDIForm1.mnuSearch.Enabled = True
MDIForm1.mnuQueries.Enabled = True
MDIForm1.mnuRepot.Enabled = True
MDIForm1.submnuBackUp.Enabled = False
End Sub
Sub proAdminLogin()
MDIForm1.mnuLibStructure.Enabled = True
MDIForm1.mnuIssueReturn.Enabled = True
MDIForm1.submnuShelfArrangement.Enabled = True 'vikjk
MDIForm1.mnuAsset.Enabled = True
MDIForm1.submnuStationary.Enabled = True
MDIForm1.mnuCatalog.Enabled = True
   MDIForm1.submnuOrderEntry.Enabled = True
   MDIForm1.submnuOrderPayment.Enabled = True
   MDIForm1.submnuOrderPlace.Enabled = True
   MDIForm1.submnuReceive.Enabled = True
MDIForm1.mnuAcquisition.Enabled = False
MDIForm1.mnuSearch.Enabled = True
MDIForm1.mnuQueries.Enabled = True
MDIForm1.mnuRepot.Enabled = True
MDIForm1.submnuGenAssets.Enabled = True
MDIForm1.submnuBackUp.Enabled = True
End Sub
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
  Rs.Open "select max(" & fieldname & ") from " & Table & " ", con, adLockOptimistic, adOpenKeyset
  Dim ID As String
  If Rs.RecordCount = 0 And prefix = "" Then
     ID = 1
     ID_Generator = Trim(ID)
     Exit Function
  End If
   If Not Rs.RecordCount = 0 And prefix = "" Then
     ID_Generator = Trim(str(Rs.Fields(0).Value + 1))
     Exit Function
  End If
  If Rs.RecordCount = 0 And Not prefix = "" Then
     ID = 1
     ID_Generator = Trim(Trim(prefix) + ID)
     Exit Function
  End If
  If Not Rs.RecordCount = 0 And Not Trim(prefix) = "" Then
     ID_Generator = Trim(Trim(prefix) + str(Rs.Fields(0).Value + 1))
     Exit Function
  End If
  Rs.Close
 End Function
Public Function Date_Check(dt As Date) As Date
  Dim c_Date As Date
  c_Date = Date
  If dt > c_Date Then
   MsgBox "Entry Date should not be More than Current Date", vbInformation + vbOKOnly, "EasyLib"
   dt = Date
  End If
    Date_Check = dt
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
'-------------- For searching Id of that perticular rearch value--------
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
Public Function funSelectMonth(ByVal dd As Date) As String
 Dim i As Integer
 i = Month(dd)
 If i = 1 Then
    funSelectMonth = "Jan"
    Exit Function
 ElseIf i = 2 Then
    funSelectMonth = "Feb"
    Exit Function
 ElseIf i = 3 Then
    funSelectMonth = "Mar"
    Exit Function
 ElseIf i = 4 Then
    funSelectMonth = "Apr"
    Exit Function
ElseIf i = 5 Then
    funSelectMonth = "May"
    Exit Function
ElseIf i = 6 Then
    funSelectMonth = "Jun"
    Exit Function
ElseIf i = 7 Then
    funSelectMonth = "Jul"
    Exit Function
ElseIf i = 8 Then
    funSelectMonth = "Aug"
    Exit Function
ElseIf i = 9 Then
    funSelectMonth = "Sep"
    Exit Function
ElseIf i = 10 Then
    funSelectMonth = "Oct"
    Exit Function
ElseIf i = 11 Then
    funSelectMonth = "Nov"
    Exit Function
Else
    funSelectMonth = "Dec"
    Exit Function
End If
End Function
