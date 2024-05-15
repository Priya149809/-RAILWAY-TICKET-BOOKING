If Text3.Text = "admin" And Text4.Text = "admin" Then
Unload Me
MDIForm1.Show
Else
MsgBox ("Invalid Username/Password")
End If
End Sub
Private Sub CLEAR_Click()
End
End Sub
Private Sub Form_Load()
Text1.Text = y
Text2.Text = z
End Sub

WELCOME PAGE
Private Sub Form_Load()
Timer1.Interval = 50
End Sub
Private Sub Picture2_Click()
End Sub
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 1
Select Case ProgressBar1.Value
Case "10"
Label1.Caption = "loading..."
Case "20"
Label1.Caption = "opening database..."
Case "30"
Label1.Caption = "checking connectivity..."
Case "80"
Label1.Caption = "Welcome to RRS"
Case "100"
Unload Me
Form1.Show
End Select
y = "gopal"
z = "krishna"
End SuB

Private Sub Command1_Click()
Form12.Show
End Sub
Private Sub mnuAbt_Click()
frmAbout.Show
End Sub
Private Sub mnuCan_Click()
Form6.Show
End Sub
Private Sub mnuCascade_Click()
MDIForm1.Arrange vbCascade
End Sub
Private Sub mnuFD_Click()
Form13.Show
End Sub
Private Sub mnuFL_Click()
Form11.Show
End Sub
Private Sub mnuRepRes_Click()
Form8.Show
End Sub
MAIN PAGE

Private Sub mnuRes_Click()
Form4.Show
End Sub
Private Sub mnuSearch_Click()
Form4.Show
End Sub
Private Sub mnuSL_Click()
Form10.Show
End Sub
Private Sub mnuTH_Click()
MDIForm1.Arrange vbTileHorizontal
End Sub
Private Sub mnuTL_Click()
Form9.Show
End Sub
Private Sub mnuTV_Click()
MDIForm1.Arrange vbTileVertical
End Sub
Private Sub mnuxit_Click()
If MsgBox("Are you sure you want to exit?", vbYesNo, "RRS") = vbYes Then
End
End If
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Date
Label4.Caption = Time
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Caption
Case "Search"
Form4.Show
Form4.Command1.Visible = False
Form4.Command2.Visible = True
Form4.Label1.Caption = "Search Train"
Case "Reservation"
Form4.Show
Form4.Caption = "Select Train Number"
Form4.Command2.Visible = False
Case "Cancellation"
Form6.Show
Case "Exit"
If MsgBox("Are you sure you want to exit?", vbYesNo, "RRS") = vbYes Then
End
End If
Case "About"
frmAbout.Show
End Select
End Sub

Private Sub Command1_Click()
If Text1.Text <> "" And Combo7.Text <> "" Then
Label14.Caption = Text1.Text
Label10.Caption = Combo7.Text
s = "select * from fares where train_no = " & Label14.Caption & " "
connect (s)
Set Label2.DataSource = rs
Label2.DataField = "train_name"
Select Case Label14.Caption
Case "1", "3", "5", "7", "9"
Select Case Label10.Caption
Case "General"
Label5.Caption = "200"
Label6.Caption = "100"
Label7.Caption = "150"
Case "II class"
Label5.Caption = "350"
Label6.Caption = "200"
Label7.Caption = "300"
Case "II sitting"
Label5.Caption = "220"
Label6.Caption = "120"
Label7.Caption = "200"
End Select
FAIR DETAILS

Case "2", "4", "6", "8"
Select Case Label10.Caption
Case "General"
Label5.Caption = "175"
Label6.Caption = "75"
Label7.Caption = "150"
Case "II class"
Label5.Caption = "330"
Label6.Caption = "190"
Label7.Caption = "290"
Case "II sitting"
Label5.Caption = "200"
Label6.Caption = "100"
Label7.Caption = "150"
Case "II sleeper"
Label5.Caption = "375"
Label6.Caption = "240"
Label7.Caption = "330"
Case "I class"
Label5.Caption = "550"
Label6.Caption = "300"
Label7.Caption = "475"

Case "III tier AC"
Label5.Caption = "730"
Label6.Caption = "450"
Label7.Caption = "580"
Case "II Tier AC"
Label5.Caption = "1100"
Label6.Caption = "625"
Label7.Caption = "775"
Case "I AC"
Label5.Caption = "2500"
Label6.Caption = "1300"
Label7.Caption = "1975"
End Select
End Select
Else
MsgBox ("Please do not leave any field blank")
End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Dim cn1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset
Dim cn2 As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim cn3 As New ADODB.Connection
Dim rs3 As New ADODB.Recordset
Dim cn4 As New ADODB.Connection
Dim rs4 As New ADODB.Recordset
Dim cn5 As New ADODB.Connection
Dim rs5 As New ADODB.Recordset
Dim cn6 As New ADODB.Connection
Dim rs6 As New ADODB.Recordset
Private Sub Combo7_Click()
s = "select * from seats where train_no = " & Text1.Text & " AND class = '" & Combo7.Text & "' "
connect (s)
Set Text66.DataSource = rs
Text66.DataField = "available_seats"
If Text66.Text = "0" Then
MsgBox ("No Seats Available in" & Combo7.Text)
Combo7.Text = ""
End If
End Sub

If Text6.Text <> "" And Text12.Text <> "" And Combo2.Text <> "" Then
If Check2.Value = True Then
Text28.Text = "Yes"
Else
Text28.Text = "No"
End If
Text29.Text = DTPicker1.Value
Text30.Text = Combo7.Text
Text68.Text = Text66.Text
Text66.Text = Text66.Text - 1
rs.Update
rs.MoveNext
rs.MovePrevious
rs2.Update
rs2.MoveNext
rs2.MovePrevious
If Text7.Text <> "" And Text13.Text <> "" And Combo3.Text <> "" Then
If Check3.Value = True Then
Text35.Text = "Yes"
Else
Text35.Text = "No"
End If
End If
Page 40 of 56
Text36.Text = DTPicker1.Value
Text37.Text = Combo7.Text
Text69.Text = Text66.Text
Text66.Text = Text66.Text - 1
rs.Update
rs.MoveNext
rs.MovePrevious
rs3.Update
rs3.MoveNext
rs3.MovePrevious
End If
If Text8.Text <> "" And Text14.Text <> "" And Combo4.Text <> "" Then
If Check4.Value = True Then
Text42.Text = "Yes"
Else
Text42.Text = "No"
End If
Text43.Text = DTPicker1.Value
Text44.Text = Combo7.Text
Adodc2.Refresh
Adodc2.Recordset.MoveLast
Text63.Text = Text59 + 1
Text70.Text = Text66.Text
Text66.Text = Text66.Text - 1
Page 41 of 56
If Text10.Text <> "" And Text16.Text <> "" And Combo6.Text <> "" Then
If Check6.Value = True Then
Text56.Text = "Yes"
Else
Text56.Text = "No"
End If
Text57.Text = DTPicker1.Value
Text58.Text = Combo7.Text
Text72.Text = Text66.Text
Text66.Text = Text66.Text - 1
rs.Update
rs.MoveNext
rs.MovePrevious
rs6.Update
rs6.MoveNext
rs6.MovePrevious
End If
Unload Me
Load Form3
Form3.Show
End Sub
Unload Me
Load Form3
Form3.Show
End Sub
Page 42 of 56
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
Combo5.Text = ""
Combo6.Text = ""
Combo7.Text = ""
Check1.Value = False
Check2.Value = False
Check3.Value = False
Check4.Value = False
Check5.Value = False
Option6.Value = False
End Sub
Private Sub Form_Load()
DTPicker1.Value = Date$
Text1.Text = Temp1
s1 = "select * from reservation"
s2 = "select * from reservation"
s3 = "select * from reservation"
s4 = "select * from reservation"
s5 = "select * from reservation"
s6 = "select * from reservation"

Set cn1 = New ADODB.Connection
cn1.CursorLocation = adUseClient
cn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"
cn1.Open
Set rs1 = New ADODB.Recordset
rs1.CursorType = adOpenDynamic
rs1.LockType = adLockOptimistic
rs1.ActiveConnection = cn1
rs1.Open s1
Set cn2 = New ADODB.Connection
cn2.CursorLocation = adUseClient
cn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"
cn2.Open
Set rs2 = New ADODB.Recordset
rs2.CursorType = adOpenDynamic
rs2.LockType = adLockOptimistic
rs2.ActiveConnection = cn2
rs2.Open s2
Set cn3 = New ADODB.Connection
cn3.CursorLocation = adUseClient
cn3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"

cn3.Open
Set rs3 = New ADODB.Recordset
rs3.CursorType = adOpenDynamic
rs3.LockType = adLockOptimistic
rs3.ActiveConnection = cn3
rs3.Open s3
Set cn4 = New ADODB.Connection
cn4.CursorLocation = adUseClient
cn4.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"
cn4.Open
Set rs4 = New ADODB.Recordset
rs4.CursorType = adOpenDynamic
rs4.LockType = adLockOptimistic
rs4.ActiveConnection = cn4
rs4.Open s4
Set cn5 = New ADODB.Connection
cn5.CursorLocation = adUseClient
cn5.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"
cn5.Open
Set rs5 = New ADODB.Recordset
rs5.CursorType = adOpenDynamic
rs5.LockType = adLockOptimistic
rs5.ActiveConnection = cn5
rs5.Open s5

Set cn6 = New ADODB.Connection
cn6.CursorLocation = adUseClient
cn6.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Railway\Railway Reservation.mdb;Persist Security Info=False"
cn6.Open
Set rs6 = New ADODB.Recordset
rs6.CursorType = adOpenDynamic
rs6.LockType = adLockOptimistic
rs6.ActiveConnection = cn6
rs6.Open s6
Set Text5.DataSource = rs1
Text5.DataField = "Passenger_name"
Set Text11.DataSource = rs1
Text11.DataField = "Age"
Set Combo1.DataSource = rs1
Combo1.DataField = "Sex"
Set Text2.DataSource = rs1
Text2.DataField = "Train_No"
Set Text3.DataSource = rs1
Text3.DataField = "Train_Name"
Set Text4.DataSource = rs1
Text4.DataField = "From"
Set Text17.DataSource = rs1
Text17.DataField = "To"
Set Text21.DataSource = rs1
Text21.DataField = "Senior_Citizen"

Text49.DataField = "Senior_Citizen"
Set Text50.DataSource = rs5
Text50.DataField = "Date_Travel"
Set Text51.DataSource = rs5
Text51.DataField = "Class"
Set Text64.DataSource = rs5
Text64.DataField = "PNR_NO"
Set Text71.DataSource = rs5
Text71.DataField = "Seat_no"
Set Text10.DataSource = rs6
Text10.DataField = "Passenger_name"
Set Text16.DataSource = rs6
Text16.DataField = "Age"
Set Combo6.DataSource = rs6
Combo6.DataField = "Sex"
Set Text52.DataSource = rs6
Text52.DataField = "Train_No"
Set Text53.DataSource = rs6
Text53.DataField = "Train_Name"
Set Text54.DataSource = rs6
Text54.DataField = "From"
Set Text55.DataSource = rs6
Text55.DataField = "To"
Set Text56.DataSource = rs6
Text56.DataField = "Senior_Citizen"
Set Text57.DataSource = rs6

Private Sub check4_Click()
If Text14.Text < 60 Then
MsgBox ("Age Should Be More Than 60")
Check4.Value = False
End If
End Sub
Private Sub check5_Click()
If Text15.Text < 60 Then
MsgBox ("Age Should Be More Than 60")
Check5.Value = False
End If
End Sub
Private Sub Option6_Click()
If Text16.Text < 60 Then
MsgBox ("Age Should Be More Than 60")
Option6.Value = False
End If
End Sub

Private Sub Command1_Click()
Label12.Caption = (Val(Label5.Caption) * Val(Text1.Text)) + (Val(Label6.Caption) * Val(Text2.Text)) + (Val(Label7.Caption) * Val(Text3.Text))
End Sub
Private Sub Command2_Click()
temp3 = Label12.Caption
Form5.Label11.Caption = Text1.Text
Form5.Label13.Caption = Text2.Text
Form5.Label15.Caption = Text3.Text
Unload Me
Load Form5
Form5.Show
End Sub
Private Sub Form_Load()
Text4.Text = n1
Text5.Text = n2
Text6.Text = n3
Text7.Text = n4
Text8.Text = n5
Text9.Text = n6
Label10.Caption = Temp4
Label14.Caption = Temp6
FORM NO: 3

End Select
End Select
End Sub
Private Sub Text3_Change()
If Val(Text10.Text) > 0 And Val(Text10.Text) < 18 Then
Text2.Text = Text2.Text + 1
ElseIf Val(Text10.Text) > 18 And Val(Text10.Text) < 60 Then
Text1.Text = Text1.Text + 1
ElseIf Val(Text10.Text) > 60 Then
Text3.Text = Text3.Text + 1
End If
End Sub
Private Sub Text4_Change()
If Val(Text11.Text) > 0 And Val(Text11.Text) < 18 Then
Text2.Text = Text2.Text + 1
ElseIf Val(Text11.Text) > 18 And Val(Text11.Text) < 60 Then
Text1.Text = Text1.Text + 1
ElseIf Val(Text11.Text) > 60 Then
Text3.Text = Text3.Text + 1
End If
End Sub
Private Sub Text12_Change()
If Val(Text12.Text) > 0 And Val(Text12.Text) < 18 Then

Text2.Text = Text2.Text + 1
ElseIf Val(Text12.Text) > 18 And Val(Text12.Text) < 60 Then
Text1.Text = Text1.Text + 1
ElseIf Val(Text12.Text) > 60 Then
Text3.Text = Text3.Text + 1
End If
End Sub
Private Sub Text5_Change()
If Val(Text13.Text) > 0 And Val(Text13.Text) < 18 Then
Text2.Text = Text2.Text + 1
ElseIf Val(Text13.Text) > 18 And Val(Text13.Text) < 60 Then
Text1.Text = Text1.Text + 1
ElseIf Val(Text13.Text) > 60 Then
Text3.Text = Text3.Text + 1
End If
End Sub
Private Sub Text6_Change()
If Val(Text14.Text) > 0 And Val(Text14.Text) < 18 Then
Text2.Text = Text2.Text + 1
ElseIf Val(Text14.Text) > 18 And Val(Text14.Text) < 60 Then
Text1.Text = Text1.Text + 1
ElseIf Val(Text14.Text) > 60 Then
Text3.Text = Text3.Text + 1
End If
End Sub

Private Sub Combo1_Click()
Adodc1.Refresh
Adodc1.Recordset.Find "Train_No =" & Combo1.Text, 0, adSearchForward
If Adodc1.Recordset.EOF = True Then
MsgBox ("Train not Available")
End If
End Sub
Private Sub Command1_Click()
Temp1 = Combo1.Text
Unload Me
Load Form2
Form2.Show
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
FORM NO :4 SEARCH TRAIN
Page 52 of 56
GENRAL FORM
Dim rs1 As New ADODB.Recordset
Dim cn1 As New ADODB.Connection
Private Sub Command1_Click()
Command1.Visible = False
Dim Beginpage, EndPage, NumCopies, orientation, i
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.ShowPrinter
Beginpage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies
orientation = CommonDialog1.orientation
For i = 1 To NumCopies
Form5.PrintForm
Next
Exit Sub
ErrHandler:
Exit Sub
End Sub
FORM NO :5 TICKET DETAILS
Page 53 of 56
FORM NO :6 CHECK PNR
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()
If MsgBox(" Are you sure you want to cancel this ticket?", vbYesNo, "RRS") = vbYes Then
s1 = "delete from reservation where PNR_NO = " & Text1.Text & " "
connect (s1)
Temp5 = Text1.Text
n7 = Text2.Text
Unload Me
Load Form7
Form7.Show
End If
End Sub
Private Sub Text1_Change()
s = "select * from reservation where PNR_NO = " & Text1.Text & " "
connect (s)
Set DataGrid1.DataSource = rs
Text2.Text = rs.RecordCount
End Sub
Page 54 of 56
FORM NO :7 TICKET PRINTING
Private Sub Command1_Click()
Command1.Visible = False
Dim Beginpage, EndPage, NumCopies, orientation, i
CommonDialog1.CancelError = True
On Error GoTo ErrHandler
CommonDialog1.ShowPrinter
Beginpage = CommonDialog1.FromPage
EndPage = CommonDialog1.ToPage
NumCopies = CommonDialog1.Copies
orientation = CommonDialog1.orientation
For i = 1 To NumCopies
Form7.PrintForm
Next
Exit Sub
ErrHandler:
Exit Sub
End Sub
Private Sub Form_Load()
Text6.Text = n7
Text3.Text = Temp5
Text5.Text = Temp5
End Sub
Private Sub Text6_Change()
Label9.Caption = "Rs." & Val(Text6.Text) * 20
End Sub
Page 55 of 56
FORM NO :8 RESERVATION LIST
Private Sub Command1_Click()
Unload Me
End Sub
Private Sub Command2_Click()
s = "select * from reservation where date_travel = ' " + Combo1.Text + " ' "
connect (s)
Set DataGrid1.DataSource = rs
End Sub
FORM NO :9 TRAIN LIST
Private Sub Command1_Click()
Unload Me
End Sub
FORM NO :10 SEAR AVAILIBILTY
Private Sub Command1_Click()
Unload Me
End Sub
Page 56 of 56
IMPORTANT NOTICE Smoking is banned in public place in India. We all know that. In long distance trains, we have to spent a long duration of time in the train itself. For the smokers it is really a pain to keep a control on their smoking habits. I feel that their should be a separate Smoking Zones in every trains, so that the passengers who smoke can have some relief. Please note that I am not promoting smoking. I am trying to quit this harmful habit. But as long as I am a smoker, it is really a pain to travel in long distance trains.I think every smoker have problems with long distance trains. And i also think that most of the Indian people don`t like the smell of cigarette.
THE
