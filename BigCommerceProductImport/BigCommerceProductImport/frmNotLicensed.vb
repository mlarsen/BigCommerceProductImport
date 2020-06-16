Public Class frmNotLicensed
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        System.Diagnostics.Process.Start(stBuyNowURL)
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class