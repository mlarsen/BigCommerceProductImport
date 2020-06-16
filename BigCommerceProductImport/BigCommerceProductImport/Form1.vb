Public Class Form1
    Dim stFileName As String
    'Dim stPath As String = "C:\CoreTech\47Photo\Data"




    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If stChannelAccount = "" Then
            MessageBox.Show("Channel Accounts have not been selected.")

        Else
            OpenFileDialog1.ShowDialog()
            stFileName = OpenFileDialog1.FileName
        End If


    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        TextBox1.Clear()
        TextBox1.AppendText(OpenFileDialog1.FileName)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Module1.Main(stFileName, stPath)
    End Sub

    Private Sub FormLoad(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Count As Integer = 1

        'Fix the data path
        If Strings.Right(stPath, 1) <> "\" Then
            stPath = stPath + "\"
        End If

        'Get the connection string
        CoreFunctions.stODBCString = CoreFunctions.GetDSN(stPath)

        If CoreFunctions.stODBCString = "" Then
            MessageBox.Show("This utility is not installed in a valid Channergy data folder.", "Error", MessageBoxButtons.OK)
            Close()
        Else
            If CoreFunctions.IsLicensed("BC Product Import") = True Then

                'Get the version infomration
                Me.txtVersion.Text = "Version:" + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Major.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Minor.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Build.ToString + "." + Reflection.Assembly.GetExecutingAssembly.GetName.Version.Revision.ToString

                Me.txtDataPath.AppendText(stPath)
                LoadForm()
            Else
                frmNotLicensed.ShowDialog()
                Me.Close()
            End If

        End If


        ''Check of ODBC driver installed
        'If My.Computer.FileSystem.FileExists("C:\dbisam\odbc\std\ver4\lib\dbodbc\dbodbc.dll") = True Then
        '    'Check of the Mailware DSN has been created
        '    If DSNExists() = False Then
        '        MessageBox.Show("The Mailware ODBC connection has not been created. Set up the ODBC connection from the Setup Menu.")
        '    End If
        'Else
        '    MessageBox.Show("The Mailware ODBC drivers have not been installed.  Please download the drivers from Setup->ODBC->Download Drivers")
        'End If

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        stChannelAccount = ComboBox1.SelectedItem
    End Sub




    Private Sub cboFileType_SelectedIndexChanged(sender As Object, e As EventArgs)


        'Clear out existing combo boxes
        ComboBox1.Items.Clear()
        ComboBox1.Refresh()
 

        'Clear out file text box
        TextBox1.Clear()
        TextBox2.Clear()

        LoadForm()
    End Sub

    Private Sub UseCurrentRetailPrice_CheckedChanged(sender As Object, e As EventArgs) Handles UseCurrentRetailPrice.CheckedChanged
        If UseCurrentRetailPrice.Checked = True Then
            stIsUseRetailPrice = "True"
        Else
            stIsUseRetailPrice = "False"
        End If

    End Sub

    Private Sub UseCurrentInStock_CheckedChanged(sender As Object, e As EventArgs) Handles UseCurrentInStock.CheckedChanged
        If UseCurrentInStock.Checked = True Then
            stIsUseCurrentInStock = "True"
        Else
            stIsUseCurrentInStock = "False"
        End If
    End Sub

    Private Sub UseProductNoAsSKU_CheckedChanged(sender As Object, e As EventArgs) Handles UseProductNoAsSKU.CheckedChanged
        If UseProductNoAsSKU.Checked = True Then
            stUseProductNoAsSKU = "True"
        Else
            stUseProductNoAsSKU = "False"
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles UpdateCurrentProd.CheckedChanged
        Dim Ans As Object

        If UpdateCurrentProd.Checked = True Then
            Ans = MessageBox.Show("Warning, this will update all of your existing products in Mailware with this imported data.  Is this what you want to do?", "Warning", MessageBoxButtons.YesNo)
            If Ans = vbYes Then
                bolUpdateExistingProd = True
            Else
                bolUpdateExistingProd = False
                UpdateCurrentProd.Checked = False
            End If
        Else
            bolUpdateExistingProd = False
        End If


    End Sub

    Private Sub CheckBox1_CheckedChanged_1(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            bolAddNewProducts = True
        Else
            bolAddNewProducts = False
        End If
    End Sub
End Class
