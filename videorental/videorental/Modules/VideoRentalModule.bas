Attribute VB_Name = "VideoRentalModule"
Option Explicit

Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
Public Const HELPC = &H3&
Public strImgN As String
Public BImg() As Byte
Public strvalid As String
Public SQLTEXT As String
Public cnn As ADODB.Connection
Public Const hl = "{HOME}+{END}"
Public adomembership As ADODB.Recordset
Public adoautonum As ADODB.Recordset
Public adoitemlist As ADODB.Recordset
Public adorr As ADODB.Recordset
Public adorent As ADODB.Recordset
Public adorrstatus As ADODB.Recordset
Public adopanerio As ADODB.Recordset
Public rslog As ADODB.Recordset
Public selek As Boolean, wfine As Boolean
Public useradd As Boolean, useredit As Boolean, LogIn As Boolean
Public AllUsers As Boolean, Admin As Boolean, Emp As Boolean

Public Sub main()
  If App.PrevInstance = True Then
    MsgBox "Video Rental System is already open.", vbOKOnly + vbInformation, App.Title
    End
  End If
  'help files
  App.HelpFile = App.Path & "\VideoRental.HLP"
  Load frmSplash
  frmSplash.Show
End Sub

'database connection
Public Sub getconnected()
  Set cnn = New ADODB.Connection
  cnn.CursorLocation = adUseClient
  cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\videorental.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=incubus"
  cnn.Open
End Sub

'display items information on main menu (frmMain.frm)
Public Sub Due()
   'info on items due today
   Set adopanerio = New ADODB.Recordset
   adopanerio.Open "SELECT * FROM rentreturn where duedate = #" & Date & "# And rentreturnstatus = 'UnReturned'", cnn, adOpenStatic, adLockReadOnly
   frmMain.lblDueItemsToday.Caption = adopanerio.RecordCount
   Set adopanerio = Nothing
   'overdue items
   Set adopanerio = New ADODB.Recordset
   adopanerio.Open "SELECT * FROM rentreturn where duedate < #" & Date & "# And rentreturnstatus = 'UnReturned'", cnn, adOpenStatic, adLockReadOnly
   frmMain.lblOverdueItems.Caption = adopanerio.RecordCount
   Set adopanerio = Nothing
   'items in
   Set adopanerio = New ADODB.Recordset
   adopanerio.Open "SELECT * FROM itemlist where status = 'IN'", cnn, adOpenStatic, adLockReadOnly
   frmMain.lblItemsIn.Caption = adopanerio.RecordCount
   Set adopanerio = Nothing
   'items out
   Set adopanerio = New ADODB.Recordset
   adopanerio.Open "SELECT * FROM rentreturn where duedate > #" & Date & "# And rentreturnstatus = 'UnReturned'", cnn, adOpenStatic, adLockReadOnly
   frmMain.lblItemsOut.Caption = adopanerio.RecordCount
   Set adopanerio = Nothing
   'total items
   Set adopanerio = New ADODB.Recordset
   adopanerio.Open "SELECT * FROM itemlist", cnn, adOpenStatic, adLockReadOnly
   frmMain.lblTotalItems.Caption = adopanerio.RecordCount
   Set adopanerio = Nothing
End Sub
 
'connect to recordset used for displaying Name and Address
Public Sub setup_connected()
  Set adopanerio = New ADODB.Recordset
  adopanerio.Open "Select * from videorentalsetup", cnn, adOpenStatic, adLockPessimistic
End Sub

'used to display Name and Address of firm in Datareport
Public Sub InOut()
  Call setup_connected
  If wfine = True And selek = True Then
    With DataReport6.Sections("Section2").Controls
      .Item("lblName").Caption = adopanerio!nname
      .Item("lblAddr").Caption = adopanerio!address
    End With
  ElseIf wfine = False And selek = False Then
    With DataReport9.Sections("Section2").Controls
      .Item("lblName").Caption = adopanerio!nname
      .Item("lblAddr").Caption = adopanerio!address
    End With
  End If
    Set adopanerio = Nothing
End Sub

'code used to load flash file on VideoRental.res file
Public Sub LoadDataIntoFile(DataName As Integer, FileName As String)
  Dim myArray() As Byte
  Dim myFile As Long
    If Dir(FileName) = "" Then
      myArray = LoadResData(DataName, "CUSTOM")
      myFile = FreeFile
      Open FileName For Binary Access Write As #myFile
      Put #myFile, , myArray
      Close #myFile
    End If
End Sub

'used to display record number on frmUserConfig.frm
Public Sub User_recno()
  On Error Resume Next
  If adopanerio.AbsolutePosition < 1 Then
    frmUserConfig.lblRecordNo.Caption = " Record " & adopanerio.AbsolutePosition + 1 & " of " & adopanerio.RecordCount
  Else
    frmUserConfig.lblRecordNo.Caption = " Record " & adopanerio.AbsolutePosition & " of " & adopanerio.RecordCount
  End If
End Sub

'Used for Users LogIn time and date
Public Sub User_LogIn()
  Set rslog = New ADODB.Recordset
  rslog.Open "SELECT * FROM userslog", cnn, adOpenStatic, adLockPessimistic
  rslog.AddNew
  rslog!Level = frmMain.StatusBar1.Panels(2).Text
  rslog!UserName = frmMain.StatusBar1.Panels(5).Text
  rslog!logintime = frmMain.StatusBar1.Panels(8).Text
  rslog!logindate = frmMain.StatusBar1.Panels(11).Text
  rslog.Update
End Sub

'Used for Users LogOut time and date
Public Sub User_LogOut()
  rslog!logouttime = Time
  rslog!logoutdate = Date
  rslog.UpdateBatch adAffectCurrent
  Set rslog = Nothing
End Sub

'display item information found on frmedititem.frm
Public Sub edit_display_item_data()
  With frmedititem
    .txtlitemid.Text = adoitemlist.Fields("itemid").value & ""
    .cmblformat.Text = adoitemlist("format").value & ""
    .txtFormatRentPrice.Text = adoitemlist!amount & ""
    .txtltitle.Text = adoitemlist!Title & ""
    .txtlcategory.Text = adoitemlist!category & ""
    .txtlmaincast.Text = adoitemlist!maincast & ""
    .txtlsecondcast.Text = adoitemlist!secondcast & ""
    .txtlprice.Text = adoitemlist!price & ""
    .txtlnoofdays.Text = adoitemlist!noofdays & ""
    .txtldate.Text = adoitemlist!datepurchase & ""
    .txtlnoofcd.Text = adoitemlist!noofcd & ""
    .txtlstatus.Text = adoitemlist!Status & ""
  End With
End Sub

'save changes made on frmedititem.frm
Public Sub edit_save_item_data()
  With frmedititem
    adoitemlist!itemid = .txtlitemid.Text
    adoitemlist!Format = .cmblformat.Text
    adoitemlist!amount = .txtFormatRentPrice.Text
    adoitemlist!Title = .txtltitle.Text
    adoitemlist!category = .txtlcategory.Text
    adoitemlist!maincast = .txtlmaincast.Text
    adoitemlist!secondcast = .txtlsecondcast.Text
    adoitemlist!price = .txtlprice.Text
    adoitemlist!noofdays = .txtlnoofdays.Text
    adoitemlist!datepurchase = .txtldate.Text
    adoitemlist!noofcd = .txtlnoofcd.Text
    adoitemlist!Status = .txtlstatus.Text
    adoitemlist.UpdateBatch adAffectCurrent
  End With
End Sub

'connect to recordset used in frmadditem.frm, frmedititem.frm and frmformat.frm
Public Sub connectcombo()
  Set adopanerio = New ADODB.Recordset
  adopanerio.Open "Select * from format", cnn, adOpenStatic, adLockPessimistic
End Sub
  
'used to fill combo box found on frmformat.frm
Public Sub fillcombo()
  On Error Resume Next
  Dim z
  For z = 1 To adopanerio.RecordCount
    frmformat.Combo1.AddItem adopanerio!itemformat
    adopanerio.MoveNext
  Next z
End Sub

'used to fill combo box found on frmadditem.frm
Public Sub fillformcombo()
  Dim z
  For z = 1 To adopanerio.RecordCount
    frmadditem.cmblformat.AddItem adopanerio.Fields(0)
    adopanerio.MoveNext
    Next z
End Sub

'connect to recordset used in frmItemList.frm
Public Sub rs_act_itemlist()
  Set adoitemlist = New ADODB.Recordset
  adoitemlist.Open "Select * from itemlist", cnn, adOpenStatic, adLockPessimistic
End Sub

'code used for searching itemid found on frmItemList.frm
Public Sub SearchItemID(reqtext As String)
  On Error Resume Next
  adoitemlist.Close
  Dim SQLTEXT As String
  SQLTEXT = "SELECT * FROM itemlist WHERE Left(itemid," & Len(reqtext) & ")='" & reqtext & "';"
  Set adoitemlist = New ADODB.Recordset
  adoitemlist.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  adoitemlist.Find "itemid  = " & "'" & reqtext & "'", , adSearchForward, 1
  Set frmItemList.DataGrid1.DataSource = adoitemlist
  Call recnoIL
  If adoitemlist.RecordCount < 1 Then
    MsgBox "Item ID number does not exist", vbInformation, "Find"
    frmItemList.txtISearchItem.Text = ""
  End If
End Sub

'code used for searching movie title found on frmItemList.frm
Public Sub SearchMovieTitle(reqtext As String)
    On Error Resume Next
    adoitemlist.Close
    Dim SQLTEXT As String
    SQLTEXT = "SELECT * FROM itemlist WHERE Left(title," & Len(reqtext) & ")='" & reqtext & "';"
    Set adoitemlist = New ADODB.Recordset
    adoitemlist.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
    adoitemlist.Find "title  = " & "'" & reqtext & "'", , adSearchForward, 1
    Set frmItemList.DataGrid1.DataSource = adoitemlist
    Call recnoIL
    If adoitemlist.RecordCount < 1 Then
     MsgBox "Movie Title does not exist", vbInformation, "Find"
     frmItemList.txtISearchMovie.Text = ""
    End If
End Sub

'code used for searching category title found on frmItemList.frm
Public Sub SearchCategory(reqtext As String)
  On Error Resume Next
  adoitemlist.Close
  Dim SQLTEXT As String
  SQLTEXT = "SELECT * FROM itemlist WHERE Left(category," & Len(reqtext) & ")='" & reqtext & "';"
  Set adoitemlist = New ADODB.Recordset
  adoitemlist.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  adoitemlist.Find "category  = " & "'" & reqtext & "'", , adSearchForward, 1
  Set frmItemList.DataGrid1.DataSource = adoitemlist
  Call recnoIL
  If adoitemlist.RecordCount < 1 Then
    MsgBox "Category name does not exist", vbInformation, "Find"
    frmItemList.txtISearchCategory.Text = ""
  End If
End Sub

'code used to sort fields found on frmIteList.frm
Public Sub SortIL()
  On Error Resume Next
  Dim SQLTEXT As String
  adoitemlist.Close
  SQLTEXT = "SELECT * FROM itemlist ORDER BY" + " " & frmItemList.cmbISort.Text + " " & frmItemList.cmbISortOrder.Text
  adoitemlist.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  Set frmItemList.DataGrid1.DataSource = adoitemlist
End Sub

'used to display record number on frmItemList.frm
Public Sub recnoIL()
  On Error Resume Next
  If adoitemlist.AbsolutePosition < 1 Then
    frmItemList.lblRecordNoA.Caption = " Record " & adoitemlist.AbsolutePosition + 1 & " of " & adoitemlist.RecordCount
  Else
    frmItemList.lblRecordNoA.Caption = " Record " & adoitemlist.AbsolutePosition & " of " & adoitemlist.RecordCount
  End If
End Sub

'clears textbox fields found on frmItemList.frm
Public Sub clear_opt_txtbox_itemlist()
  With frmItemList
    .txtISearchItem.Text = ""
    .txtISearchMovie.Text = ""
    .txtISearchCategory.Text = ""
    .cmbISort.Text = ""
    .cmbISortOrder.Text = ""
  End With
End Sub

'write data from controls. found on frmaddmembership.frm
Public Sub WriteDataFromControls()
  On Error Resume Next
  With frmaddmembership
    If .txtPictureName.Text = "" Then
      Call nopicture
      GoTo jessiepanerio:
    Else
jessiepanerio:
      Call Image
      adomembership!picfilename = .txtPictureName
      adomembership.Fields("picblob").AppendChunk BImg
      adoautonum!autonum = .txtautonum.Text
      adomembership!membershipid = .txtMID.Text
      adomembership!lastname = .txtLName.Text
      adomembership!firstname = .txtFName.Text
      adomembership!middlename = .txtMName.Text
      adomembership!Date = .dateMem.Text
      adomembership!birthdate = .txtBDate.Text
      adomembership!gender = .cmbGender.Text
      adomembership!address = .txtAddress.Text
      adomembership!landline = .txtLandLine.Text
      adomembership!mobile = .txtMobile.Text
    End If
  End With
End Sub

'code used if borrower has no picture. found on frmaddmembership.frm
Public Sub nopicture()
  With frmaddmembership
    .cd2.InitDir = App.Path & "\image"
    .cd2.FileName = App.Path & "\image\temp.jpg"
      If .cd2.FileName <> "" Then
        strImgN = .cd2.FileName
        .txtPictureName.Text = "No Picture"
        .imgpic.Picture = LoadPicture(.cd2.FileName)
      End If
  End With
End Sub

'code for image
Public Sub Image()
  On Error Resume Next
  Dim IntNum As Integer
  IntNum = FreeFile
  Open strImgN For Binary As #IntNum
  ReDim BImg(FileLen(strImgN))
  Get #IntNum, , BImg
  Close #1
End Sub

'code for loading image on frmmembership.frm
Public Sub LoadImage()
  On Error Resume Next
  Dim ImgS As Long
  Dim OS As Long
  Dim TmpPic As String
  Const conCS = 100
  TmpPic = App.Path & "\tmpPic.bmp"
    If Len(Dir(TmpPic)) > 0 Then
      Kill TmpPic
    End If
      Dim F As Integer
      F = FreeFile
      Open App.Path & "\tmpPic.bmp" For Binary As #F
      ImgS = adomembership.Fields("picblob").ActualSize
      Do While OS < ImgS
        BImg() = adomembership _
        ("picblob").GetChunk(conCS)
        Put #F, , BImg
        OS = OS + conCS
      Loop
        Close #F
        frmMembership.imgpic.Picture = LoadPicture(App.Path & "\tmpPic.bmp")
        Kill App.Path & "\tmpPic.bmp"
End Sub

'connect to recordset. used on frmaddmembership.frm, frmrent.frm and frmadditem.frm
Public Sub rs_act_autonum()
  Set adoautonum = New ADODB.Recordset
  adoautonum.Open "Select * from autonum", cnn, adOpenStatic, adLockPessimistic
End Sub

'connect to recordset. used on frmmembership.frm
Public Sub rs_act_membership()
  Set adomembership = New ADODB.Recordset
  adomembership.Open "Select * from membership", cnn, adOpenStatic, adLockPessimistic
End Sub

'code for searching Lastname field. found on fmMembership.frm
Public Sub SearchLastName(reqtext As String)
  On Error Resume Next
  adomembership.Close
  Dim SQLTEXT As String
  SQLTEXT = "SELECT * FROM membership WHERE Left(lastname," & Len(reqtext) & ")='" & reqtext & "';"
  Set adomembership = New ADODB.Recordset
  adomembership.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  adomembership.Find "lastname  = " & "'" & reqtext & "'", , adSearchForward, 1
  Set frmMembership.DataGrid1.DataSource = adomembership
  recno
  If adomembership.RecordCount < 1 Then
    Set frmMembership.imgpic.Picture = Nothing
    MsgBox "Customer name does not exist", vbInformation, "Find"
    frmMembership.txtMSearchLN.Text = ""
  Else
    Call LoadImage
  End If
End Sub

'code for searching Borrowers ID# field. found on fmMembership.frm
Public Sub SearchMemID(reqtext As String)
  On Error Resume Next
  adomembership.Close
  Dim SQLTEXT As String
  SQLTEXT = "SELECT * FROM membership WHERE Left(membershipid," & Len(reqtext) & ")='" & reqtext & "';"
  Set adomembership = New ADODB.Recordset
  adomembership.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  adomembership.Find "membershipid  = " & "'" & reqtext & "'", , adSearchForward, 1
  Set frmMembership.DataGrid1.DataSource = adomembership
  recno
  If adomembership.RecordCount < 1 Then
    Set frmMembership.imgpic.Picture = Nothing
    MsgBox "ID number does not exist", vbInformation, "Find"
    frmMembership.txtMSearchMem.Text = ""
  Else
    Call LoadImage
  End If
End Sub

'code for sorting fields. found on fmMembership.frm
Public Sub Sort()
  On Error Resume Next
  Dim SQLTEXT As String
  adomembership.Close
  SQLTEXT = "SELECT * FROM membership ORDER BY" + " " & frmMembership.cmbMSort.Text + " " & frmMembership.cmbMSortOrder.Text
  adomembership.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  Set frmMembership.DataGrid1.DataSource = adomembership
  If adomembership.RecordCount < 1 Then
    Set frmMembership.imgpic.Picture = Nothing
  Else
    Call LoadImage
  End If
End Sub

'used to display record number on frmMembership.frm
Public Sub recno()
  On Error Resume Next
  If adomembership.AbsolutePosition < 1 Then
    frmMembership.lblRecordNoA.Caption = " Record " & adomembership.AbsolutePosition + 1 & " of " & adomembership.RecordCount
  Else
    frmMembership.lblRecordNoA.Caption = " Record " & adomembership.AbsolutePosition & " of " & adomembership.RecordCount
  End If
End Sub

'used to clear textbox on frmMembership.frm
Public Sub clear_opt_txtbox_members()
  With frmMembership
    .txtMSearchMem.Text = ""
    .txtMSearchLN.Text = ""
    .cmbMSort.Text = ""
    .cmbMSortOrder.Text = ""
  End With
End Sub

'load data for viewing. found on frmeditmembership.frm
Public Sub edit_display_data()
  On Error Resume Next
    With frmeditmembership
      .txtMID.Text = adomembership!membershipid
      .txtLName.Text = adomembership!lastname
      .txtFName.Text = adomembership!firstname
      .txtMName.Text = adomembership!middlename
      .txtDate.Text = adomembership!Date
      .txtBDate.Text = adomembership!birthdate
      .cmbGender.Text = adomembership!gender
      .txtAddress.Text = adomembership!address
      .txtLandLine.Text = adomembership!landline
      .txtMobile.Text = adomembership!mobile
      .txtPictureName = adomembership!picfilename
      Set .imgpic.Picture = frmMembership.imgpic.Picture
    End With
End Sub

'write data from controls found on frmeditmembership.frm
Public Sub edit_save_data()
  On Error Resume Next
  With frmeditmembership
    If .txtPictureName.Text = adomembership!picfilename Then
      GoTo jessiepanerio:
    Else
      Call Image
      adomembership!picfilename = .txtPictureName
      adomembership.Fields("picblob").AppendChunk BImg
jessiepanerio:
      adomembership!membershipid = .txtMID.Text
      adomembership!lastname = .txtLName.Text
      adomembership!firstname = .txtFName.Text
      adomembership!middlename = .txtMName.Text
      adomembership!Date = .txtDate.Text
      adomembership!birthdate = .txtBDate.Text
      adomembership!gender = .cmbGender.Text
      adomembership!address = .txtAddress.Text
      adomembership!landline = .txtLandLine.Text
      adomembership!mobile = .txtMobile.Text
      adomembership.UpdateBatch adAffectCurrent
    End If
  End With
End Sub

'write data from controls. found on frmadditem.frm
Public Sub WriteDataFromControlslist()
  On Error Resume Next
  With frmadditem
    adoautonum!itemnum = .txtlitemid.Text
    adoitemlist!itemid = .txtlitemid.Text
    adoitemlist!Format = .cmblformat.Text
    adoitemlist!amount = .txtFormatRentPrice.Text
    adoitemlist!Title = .txtltitle.Text
    adoitemlist!category = .txtlcategory.Text
    adoitemlist!maincast = .txtlmaincast.Text
    adoitemlist!secondcast = .txtlsecondcast.Text
    adoitemlist!price = .txtlprice.Text
    adoitemlist!noofdays = .txtlnoofdays.Text
    adoitemlist!datepurchase = .txtDate.Text
    adoitemlist!noofcd = .txtnoofcd.Text
    adoitemlist!Status = .txtstatus.Text
  End With
End Sub

'code for searching Borrowers ID #. used in returning items. found on frmRent.frm
Public Sub MemID(search As String)
  On Error Resume Next
  adorent.Close
  SQLTEXT = "SELECT * FROM membership WHERE Left(membershipid," & Len(search) & ")='" & search & "' ORDER BY membershipid;"
  Set adorent = New ADODB.Recordset
  adorent.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  frmRent.lstMemID.clear
  If adorent.RecordCount = 0 Then
    MsgBox "Membership ID number does not exist.", vbInformation, "Validation"
    frmRent.txtMemID.Text = ""
    Exit Sub
  End If
  Do Until adorent.EOF
    frmRent.lstMemID.AddItem adorent.Fields(0)
    adorent.MoveNext
  Loop
  If frmRent.lstMemID.ListCount = 1 Then
    adorent.MoveFirst
    frmRent.txtMemID.Text = frmRent.lstMemID.List(0)
    frmRent.txtMemID.SelLength = Len(frmRent.txtMemID.Text)
  End If
  frmRent.lblMemName.Caption = adorent!firstname & "  " & adorent!lastname
  frmRent.txt5.Text = adorent!membershipid
  frmRent.txt6.Text = adorent!lastname
  frmRent.txt7.Text = adorent!firstname
End Sub

'code for searching Borrowers ID #, Title, format etc. found on frmRent.frm
Public Sub Item(search As String)
  On Error Resume Next
  adorent.Close
  SQLTEXT = "SELECT * FROM itemlist WHERE Left(itemid," & Len(search) & ")='" & search & "' And status = 'IN' ORDER BY itemid;"
  Set adorent = New ADODB.Recordset
  adorent.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  frmRent.lstItemID.clear
  If adorent.RecordCount = 0 Then
    MsgBox "Item ID number does not exist.", vbInformation, "Validation"
    frmRent.txtItemID.Text = ""
    Exit Sub
  End If
  Do Until adorent.EOF
    frmRent.lstItemID.AddItem adorent!itemid
    adorent.MoveNext
  Loop
  If frmRent.lstItemID.ListCount = 1 Then
    adorent.MoveFirst
    frmRent.txtItemID.Text = frmRent.lstItemID.List(0)
    frmRent.txtItemID.SelLength = Len(frmRent.txtItemID.Text)
  End If
    frmRent.lblMovieTitle.Caption = adorent!Title
    frmRent.txtstatus.Text = adorent!Status
    frmRent.txt1.Text = adorent!itemid
    frmRent.txt2.Text = adorent!Title
    frmRent.txt3.Text = adorent!Format
    frmRent.txt4.Text = adorent!amount
    frmRent.txt8.Text = Date
    frmRent.txt9.Text = CDate(frmRent.txt8.Text) + CDate(adorent!noofdays)
End Sub

'code used to fill list box with Borrowers ID #. found on frmRent.frm
Public Sub MID(search As String)
    On Error Resume Next
    adorent.Close
    SQLTEXT = "SELECT * FROM membership WHERE membershipid='" & search & "';"
    Set adorent = New ADODB.Recordset
    adorent.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
    If adorent.RecordCount = 0 Then Exit Sub
    frmRent.lstMemID.clear
    frmRent.lstMemID.AddItem adorent!membershipid
    frmRent.txtMemID.SetFocus
    SendKeys hl
End Sub

'code used to fill list box with Item ID #. found on frmRent.frm
Public Sub IID(search As String)
  On Error Resume Next
  adorent.Close
  SQLTEXT = "SELECT * FROM itemlist WHERE itemid='" & search & "';"
  Set adorent = New ADODB.Recordset
  adorent.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  If adorent.RecordCount = 0 Then Exit Sub
  frmRent.lstItemID.clear
  frmRent.lstItemID.AddItem adorent!itemid
  frmRent.txtItemID.SetFocus
  SendKeys hl
End Sub

'connect to recordset. used on frmRent.frm
Public Sub rentreturnconnect()
  Set adorr = New ADODB.Recordset
  adorr.Open "Select * from rentreturn", cnn, adOpenStatic, adLockPessimistic
End Sub

'used to display info on datareport. found on frmcashier.frm
Public Sub print_or(reqtext As String)
  Dim SQLTEXT As String
  SQLTEXT = "SELECT * FROM rentreturn WHERE rrtransno >= '" & reqtext & "';"
  Set adopanerio = New ADODB.Recordset
  adopanerio.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  Set DataReport1.DataSource = adopanerio
  DataReport1.Show
  Set adopanerio = Nothing
End Sub

'clears textbox field. found on frmRent.frm
Public Sub clearall()
  With frmRent
    .Text1.Text = ""
    .Text2.Text = ""
    .Text3.Text = ""
    .Text4.Text = ""
    .Text5.Text = ""
    .Text6.Text = ""
    .Text7.Text = ""
    .Text8.Text = ""
    .Text9.Text = ""
    .Text10.Text = ""
    .txtA.Text = ""
    .txtM.Text = ""
    .txtTitl.Text = ""
    .txtAmount.Text = ""
    .txtItemCount.Text = ""
  End With
End Sub

'code for searching Item ID # used for returning items. found on frmReturn.frm
Public Sub sItem(search As String)
  On Error Resume Next
  adorr.Close
  Dim SQLTEXT As String
  SQLTEXT = "SELECT * FROM rentreturn WHERE Left(itemidnumber," & Len(search) & ")='" & search & "' And rentreturnstatus = 'UnReturned';"
  Set adorr = New ADODB.Recordset
  adorr.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  frmReturn.lstItemID.clear
    If adorr.RecordCount = 0 Then
      frmReturn.txtItemID.Text = ""
      Exit Sub
    End If
    Do Until adorr.EOF
      frmReturn.lstItemID.AddItem adorr!itemidnumber
      adorr.MoveNext
    Loop
    If frmReturn.lstItemID.ListCount = 1 Then
      adorr.MoveFirst
      frmReturn.txtItemID.Text = frmReturn.lstItemID.List(0)
      frmReturn.txtItemID.SelLength = Len(frmReturn.txtItemID.Text)
    End If
    frmReturn.lblMemName.Caption = adorr!firstname & " " & adorr!lastname
    frmReturn.lblMovieTitle.Caption = adorr!Title
    frmReturn.lblMemID.Caption = adorr!membershipid
    frmReturn.lblFormat.Caption = adorr!Format
    frmReturn.lblItemID.Caption = adorr!itemidnumber
    frmReturn.lblAmount.Caption = adorr!amount
    frmReturn.lblDateBorrowed.Caption = adorr!datebor
    frmReturn.lblDueDate.Caption = adorr!duedate
    frmReturn.txtrrtransno.Text = adorr!rrtransno
End Sub

'used to fill list box with Item ID #. found on frmReturn.frm
Public Sub sID(search As String)
  On Error Resume Next
  adorr.Close
  Dim SQLTEXT As String
  SQLTEXT = "SELECT * FROM rentreturn WHERE itemidnumber='" & search & "';"
  Set adorr = New ADODB.Recordset
  adorr.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  If adorr.RecordCount = 0 Then Exit Sub
  frmReturn.lstItemID.clear
  frmReturn.lstItemID.AddItem adorr!itemidnumber
  frmReturn.txtItemID.SelStart = 0
  frmReturn.txtItemID.SelLength = Len(frmReturn.txtItemID.Text)
End Sub

'used to get and display Item Status (In or Out). found on frmReturn.frm
Public Sub itemstatus(search As String)
  On Error Resume Next
  Dim SQLTEXT As String
  SQLTEXT = "SELECT * FROM itemlist WHERE itemid='" & search & "';"
  Set adorrstatus = New ADODB.Recordset
  adorrstatus.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  frmReturn.lblItemStatus.Caption = adorrstatus!Status
End Sub

'used to display penalty information. found on frmReturn.frm
Public Sub penalty()
  With frmReturn
  Set adopanerio = New ADODB.Recordset
  adopanerio.Open "Select * from penaltyrateperday", cnn, adOpenStatic, adLockPessimistic
  .lblnoofdayspenalty.Caption = CDate(.DateReturned.Caption) - CDate(.lblDueDate.Caption)
  If .lblnoofdayspenalty.Caption < 1 Then
    .lblnoofdayspenalty.Caption = 0
    .lblStatus.Caption = "Returned"
    .lblTotalAmountPenalty.Caption = 0
  Else
    .lblStatus.Caption = "Returned"
    .lblTotalAmountPenalty.Caption = Val(.lblnoofdayspenalty.Caption) * Val(adopanerio!penaltyrateperday)
    .lblTotalAmountPenalty.Caption = Format(.lblTotalAmountPenalty.Caption, "####.00")
  End If
  Set adopanerio = Nothing
  End With
End Sub

'used to display items (out) on datagrid
Public Sub unreturned(search As String)
  On Error Resume Next
  SQLTEXT = "SELECT * FROM rentreturn WHERE Left(membershipid," & Len(search) & ")='" & search & "' And rentreturnstatus = 'UnReturned' ORDER BY itemidnumber;"
  Set adopanerio = New ADODB.Recordset
  adopanerio.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  Set frmReturn.DataGrid1.DataSource = adopanerio
  Set adopanerio = Nothing
End Sub
   
'prints official receipt for penalty
Public Sub printpenalty(reqtext As String)
  On Error Resume Next
  SQLTEXT = "SELECT * FROM rentreturn WHERE rrtransno = '" & reqtext & "';"
  Set adopanerio = New ADODB.Recordset
  adopanerio.Open SQLTEXT, cnn, adOpenStatic, adLockPessimistic
  Set DataReport2.DataSource = adopanerio
  With DataReport2.Sections("Section2").Controls
    .Item("Label10").Caption = adopanerio!firstname & " " & adopanerio!lastname
    .Item("Label12").Caption = adopanerio!membershipid
    .Item("Label13").Caption = adopanerio!totalpenaltyamount
    .Item("Label20").Caption = adopanerio!DateReturned
    .Item("Label21").Caption = adopanerio!noofdayspenalty
    .Item("Label3").Caption = Val(.Item("Label13").Caption) / Val(.Item("Label21").Caption)
  End With
  DataReport2.Show
  Set adopanerio = Nothing
End Sub
