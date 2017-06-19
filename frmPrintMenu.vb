Option Strict Off
Option Explicit On
Friend Class frmPrintMenu
	Inherits System.Windows.Forms.Form
	
	Private Sub comPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comPrint.Click
		Dim x As Short
		For x = 0 To 6
			If chkPrintMenu(x).CheckState = 1 Then
				PrintJob(x) = True
			Else
				PrintJob(x) = False
			End If
		Next x
		If chkPrintMenu(7).CheckState = 1 Then
			Me.Close()
		Else
			frmOutPrint.Show()
			Me.Close()
		End If
	End Sub
End Class