
'Student Registration System
'Starting with Splash Form

Private Sub Form_Load()

End Sub
Private Sub Timer1_Timer()
Splash.Hide
frmLogin.Show
Timer1.Enabled = False
End Sub


'Login Form

Option Explicit
Public LoginSucceeded As Boolean
Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    'check for correct password
    If txtPassword = "admin" Then
        'place code to here to pass the
        'success to the calling sub
        'setting a global var is the easiest
        LoginSucceeded = True
        Me.Hide
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If
     If txtUserName = "admin" Then       
        LoginSucceeded = True
        Me.Hide
        Form1.Show
           End If
          End Sub
		  
		  

Private Sub Register_Click()
Register.Show
Login.Hide

End Sub

Private Sub Form_Load()
End Sub


'Register Form


Private Sub Already_Click()
Login.Show

End Sub


Private Sub Cancel_Click()
End
End Sub
Private Sub Form_Load()
Adodc1.Recordset.Update
Text1.Text = ""
Text2.Text = ""


End Sub

Private Sub Register_Click()
Adodc1.Recordset.Addnew

End Sub

'Student Registration Form

Private Sub Last_Click()
Adodc1.Recordset.MoveLast
End Sub

Private Sub Exit_Click()
Unload Form1

End Sub

Private Sub AddNew_Click()
Adodc1.Recordset.Addnew
clear
End Sub


Sub clear()
Text6.Text = "12/01/1947"
Combo1.Text = "Select Class"
Combo2.Text = "Select Class"
Combo3.Text = "Select Religion"
Combo4.Text = "Select Gender"

End Sub

Private Sub Update_Click()
Adodc1.Recordset.Update
MsgBox "Record Saved Successfully"

End Sub

Private Sub Delete_Click()
confirmation = MsgBox("Do you want to delete this record", vbYesNo + vbCritical, "Delete Record Confirmation")
If confirmation = vbYes Then
Adodc1.Recordset.Delete
MsgBox "Record has been Deleted Successfully", vbInformation, "Message"
Else
MsgBox "Record Not Deleted...!!", vbInformation, "Message"

End If

End Sub

Private Sub ViewRecord_Click()
Form2.Show
End Sub

Private Sub Text15_Change()
Picture1.Picture = LoadPicture(Text15.Text)
End Sub

Private Sub SaveBtn_Click()
Adodc1.Recordset.Update
MsgBox "Record Saved Successfully"
End Sub

Private Sub Previous_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Next_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub First_Click()
Adodc1.Recordset.MoveFirst
End Sub


Private Sub Form_Load()
Dim str As String

'Adding items in combo1
Combo1.AddItem "One"
Combo1.AddItem "Two"
Combo1.AddItem "Three"
Combo1.AddItem "Four"
Combo1.AddItem "Five"
Combo1.AddItem "Sixth"
Combo1.AddItem "Seventh"
Combo1.AddItem "Eighth"
Combo1.AddItem "Nineth"
Combo1.AddItem "Matric"

'Adding options in combo4

Combo4.AddItem "Male"
Combo4.AddItem "Female"

'Adding options in combo2

Combo2.AddItem "One"
Combo2.AddItem "Two"
Combo2.AddItem "Three"
Combo2.AddItem "Four"
Combo2.AddItem "Five"
Combo2.AddItem "Sixth"
Combo2.AddItem "Seventh"
Combo2.AddItem "Eighth"
Combo2.AddItem "Nineth"
Combo2.AddItem "Matric"

'Adding options in combo3

Combo3.AddItem " Islam"
Combo3.AddItem " Hindu"
Combo3.AddItem " Chiristian"
Combo3.AddItem " Jews"



End Sub

Private Sub UploadImage_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "Jpeg|*jpg"
Text15.Text = CommonDialog1.FileName

End Sub



'Records Form


Private Sub Form_Load()
End Sub

Private Sub GO_Click()
Adodc1.RecordSource = "Select * from Table1 where Name of Student = '" + Text1.Text + "' or Caste = '" + Text1.Text + "'"
If Adodc1.Recordset.EOF Then
MsgBox "Record Not Found", vbCritical, "Message"
Else
Adodc1.Caption = Adodc1.RecordSource
End If

End Sub
