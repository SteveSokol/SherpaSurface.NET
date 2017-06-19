Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmOutPrint
	Inherits System.Windows.Forms.Form
	Dim LastCell As Short
	'UPGRADE_WARNING: Event chkPrintTitle.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkPrintTitle_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPrintTitle.CheckStateChanged
		Dim Index As Short = chkPrintTitle.GetIndex(eventSender)
		Select Case Index
			Case 0
				chkPrintTitle(1).Focus()
				lstFontList.Visible = False
				Label1(3).Visible = False
			Case 1
				chkPrintTitle(2).Focus()
			Case 2
				txtPrintOutItem(2).Focus()
				lstValueFontSize.Visible = True
				Label1(5).Visible = True
		End Select
	End Sub
	Private Sub chkPrintTitle_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPrintTitle.Enter
		Dim Index As Short = chkPrintTitle.GetIndex(eventSender)
		Label1(3).Visible = False
		lstFontList.Visible = False
		Label1(4).Visible = False
		lstFontSize.Visible = False
	End Sub
	'UPGRADE_WARNING: Event chkPrintToFile.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub chkPrintToFile_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkPrintToFile.CheckStateChanged
		If chkPrintToFile.CheckState = 1 Then
			txtPrintOutItem(5).Enabled = True
			labPrintOutItems(5).ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
			txtPrintOutItem(5).BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
            txtPrintOutItem(5).Text = frmSurfaceFileMaker.fileSurfaceFile.Path
        Else
            labPrintOutItems(5).ForeColor = System.Drawing.ColorTranslator.FromOle(&HE0E0E0)
            txtPrintOutItem(5).BackColor = System.Drawing.ColorTranslator.FromOle(&HE0E0E0)
            txtPrintOutItem(5).Enabled = False
        End If
    End Sub
    Private Sub comLeavePrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comLeavePrint.Click
        Me.Close()
    End Sub
    Private Sub comPrintItOut_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles comPrintItOut.Click
        If chkPrintToFile.CheckState = 1 Then
            PrintToName = LTrim(RTrim(txtPrintOutItem(5).Text))
            Call FilePrinter()
        Else
            Call PrintOutEnglish()
        End If
        Me.Close()
    End Sub
    'UPGRADE_WARNING: Form event frmOutPrint.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmOutPrint_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Dim Printer As New Printer
        LastCell = 0
        'Dim i As Short
        'For i = 0 To Printer.FontCount - 1
        'lstFontList.Items.Add(Printer.Fonts(i))
        'Next i
    End Sub
    Private Sub frmOutPrint_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - (VB6.PixelsToTwipsY(Me.Height) + 350)) / 2)
        Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
        If VB6.PixelsToTwipsY(Me.Top) < 0 Then Me.Top = 0
        If VB6.PixelsToTwipsX(Me.Left) < 0 Then Me.Left = 0
    End Sub
    'UPGRADE_WARNING: Event lstFontList.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub lstFontList_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFontList.SelectedIndexChanged
        If LastCell = 0 Then
            txtPrintOutItem(0).Font = VB6.FontChangeName(txtPrintOutItem(0).Font, VB6.GetItemString(lstFontList, lstFontList.SelectedIndex))
            txtPrintOutItem(0).Text = VB6.GetItemString(lstFontList, lstFontList.SelectedIndex)
            Label1(3).Visible = False
            lstFontList.Visible = False
            txtPrintOutItem(1).Focus()
        ElseIf LastCell = 2 Then
            txtPrintOutItem(2).Font = VB6.FontChangeName(txtPrintOutItem(2).Font, VB6.GetItemString(lstFontList, lstFontList.SelectedIndex))
            txtPrintOutItem(2).Text = VB6.GetItemString(lstFontList, lstFontList.SelectedIndex)
            Label1(3).Visible = False
            lstFontList.Visible = False
            txtPrintOutItem(3).Focus()
        End If
    End Sub
    'UPGRADE_WARNING: Event lstFontSize.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub lstFontSize_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstFontSize.SelectedIndexChanged
        txtPrintOutItem(1).Text = VB6.GetItemString(lstFontSize, lstFontSize.SelectedIndex)
        Label1(3).Visible = False
        Label1(4).Visible = False
        lstFontList.Visible = False
        lstFontSize.Visible = False
        chkPrintTitle(0).Focus()
    End Sub
    'UPGRADE_WARNING: Event lstValueFontSize.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub lstValueFontSize_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstValueFontSize.SelectedIndexChanged
        txtPrintOutItem(3).Text = VB6.GetItemString(lstValueFontSize, lstValueFontSize.SelectedIndex)
        Label1(5).Visible = False
        lstValueFontSize.Visible = False
    End Sub
    Private Sub txtPrintOutItem_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPrintOutItem.Enter
        Dim Index As Short = txtPrintOutItem.GetIndex(eventSender)
        LastCell = Index
        If Index = 0 Or Index = 2 Then
            Label1(3).Visible = True
            lstFontList.Visible = True
            Label1(4).Visible = False
            lstFontSize.Visible = False
            Label1(5).Visible = False
            lstValueFontSize.Visible = False
        ElseIf Index = 1 Then
            Label1(3).Visible = False
            lstFontList.Visible = False
            Label1(4).Visible = True
            lstFontSize.Visible = True
            Label1(5).Visible = False
            lstValueFontSize.Visible = False
        ElseIf Index = 3 Then
            Label1(3).Visible = False
            lstFontList.Visible = False
            Label1(4).Visible = False
            lstFontSize.Visible = False
            Label1(5).Visible = True
            lstValueFontSize.Visible = True
        Else
            Label1(3).Visible = False
            lstFontList.Visible = False
            Label1(4).Visible = False
            lstFontSize.Visible = False
            Label1(5).Visible = False
            lstValueFontSize.Visible = False
        End If
    End Sub
    Private Sub FilePrinter()
        Dim saveinput As Short = Nothing
        Dim i As Short
        Dim wordlength As Short
        Dim theword As String = Nothing
        Dim spreadtheword As String
        Dim addext As Short
        Dim tempdir As String = Nothing

        JumpShip = False
        wordlength = Len(PrintToName)

        For i = 1 To wordlength
            If Mid(PrintToName, i, 1) = "\" Then
                theword = ""
            Else
                theword = theword & Mid(PrintToName, i, 1)
            End If
        Next i

        wordlength = Len(theword)

        addext = True
        For i = 1 To wordlength
            If Mid(theword, i, 1) = "." Then
                addext = False
            End If
        Next i

        If addext = True Then
            spreadtheword = PrintToName & ".prn"
            theword = PrintToName & ".txt"
            addext = False
        Else
            spreadtheword = PrintToName
            theword = PrintToName
            addext = False
        End If

        'For i = 0 To frmSurfaceFileMaker.fileSurfaceFile.Items.Count
        'If VB.Left(theword, 8) = VB.Left(frmSurfaceFileMaker.fileSurfaceFile.Items(i), 8) Then
        'addext = True
        'End If
        'Next i
        addext = False

        If addext = True Then
            WarnNumber = 9
            frmWarning.Show()
        Else
            PrintFileName = theword
            LotusFileName = spreadtheword
            '    GetPrintFileName = False
        End If

        If PrintFileName = "" Then
            WarnNumber = 9
            frmWarning.Show()
        End If

        PrintFileNumber = FreeFile()
        FileOpen(PrintFileNumber, PrintFileName, OpenMode.Output)

        Call PrintOutEngFile(PrintFileNumber, LotusFileNumber)

        FileClose(PrintFileNumber)
    End Sub
End Class