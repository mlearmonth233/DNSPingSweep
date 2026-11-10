' NetToolsLogin - the SSH login box (username + masked password).
' build.ps1 adds an empty UserForm called NetToolsLogin and puts this code
' in it; the controls are created at run time, so no binary .frx is needed.
Option Explicit

Private WithEvents btnOK As MSForms.CommandButton
Private WithEvents btnCancel As MSForms.CommandButton
Private txtUser As MSForms.TextBox
Private txtPass As MSForms.TextBox
Private mOK As Boolean

Private Sub UserForm_Initialize()
    Dim lbl As MSForms.Label
    Me.Caption = "NetTools - SSH login"
    Me.Width = 280
    Me.Height = 178

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblUser")
    lbl.Caption = "Username:"
    lbl.Left = 12: lbl.Top = 16: lbl.Width = 60

    Set txtUser = Me.Controls.Add("Forms.TextBox.1", "txtUser")
    txtUser.Left = 80: txtUser.Top = 12: txtUser.Width = 175

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblPass")
    lbl.Caption = "Password:"
    lbl.Left = 12: lbl.Top = 46: lbl.Width = 60

    Set txtPass = Me.Controls.Add("Forms.TextBox.1", "txtPass")
    txtPass.Left = 80: txtPass.Top = 42: txtPass.Width = 175
    txtPass.PasswordChar = "*"

    Set lbl = Me.Controls.Add("Forms.Label.1", "lblNote")
    lbl.Caption = "Used for this run only - the password is never saved."
    lbl.Left = 12: lbl.Top = 74: lbl.Width = 245
    lbl.ForeColor = RGB(96, 96, 96)

    Set btnOK = Me.Controls.Add("Forms.CommandButton.1", "btnOK")
    btnOK.Caption = "OK"
    btnOK.Left = 105: btnOK.Top = 104: btnOK.Width = 70: btnOK.Height = 24
    btnOK.Default = True

    Set btnCancel = Me.Controls.Add("Forms.CommandButton.1", "btnCancel")
    btnCancel.Caption = "Cancel"
    btnCancel.Left = 185: btnCancel.Top = 104: btnCancel.Width = 70: btnCancel.Height = 24
    btnCancel.Cancel = True
End Sub

' Shows the box (modal). user is pre-filled and returned; pw is returned.
Public Function Ask(ByRef user As String, ByRef pw As String) As Boolean
    txtUser.Text = user
    txtPass.Text = ""
    mOK = False
    Me.Show
    If mOK Then
        user = Trim$(txtUser.Text)
        pw = txtPass.Text
        Ask = True
    End If
    txtPass.Text = ""
End Function

Private Sub UserForm_Activate()
    If Len(txtUser.Text) > 0 Then txtPass.SetFocus Else txtUser.SetFocus
End Sub

Private Sub btnOK_Click()
    mOK = True
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    mOK = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then     ' the X button
        Cancel = True
        mOK = False
        Me.Hide
    End If
End Sub
