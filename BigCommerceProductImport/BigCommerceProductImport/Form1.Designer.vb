<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.UseCurrentRetailPrice = New System.Windows.Forms.CheckBox()
        Me.UseCurrentInStock = New System.Windows.Forms.CheckBox()
        Me.UseProductNoAsSKU = New System.Windows.Forms.CheckBox()
        Me.UpdateCurrentProd = New System.Windows.Forms.CheckBox()
        Me.txtDataPath = New System.Windows.Forms.TextBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(443, 150)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(115, 23)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Select File"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.DefaultExt = "*.txt"
        Me.OpenFileDialog1.FileName = "*.txt;*.csv"
        Me.OpenFileDialog1.Filter = "Text Files|*.txt|All files|*.*"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(45, 152)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(380, 20)
        Me.TextBox1.TabIndex = 1
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(45, 324)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(196, 28)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Process Active Listings"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(42, 136)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(175, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Select the active listings report here"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(45, 368)
        Me.TextBox2.Multiline = True
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox2.Size = New System.Drawing.Size(452, 99)
        Me.TextBox2.TabIndex = 4
        Me.TextBox2.TabStop = False
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.SelectedPath = "C:\Mailware\Data"
        '
        'ComboBox1
        '
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Location = New System.Drawing.Point(209, 70)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(216, 21)
        Me.ComboBox1.TabIndex = 7
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(42, 77)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(122, 13)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "Select Channel Account"
        '
        'UseCurrentRetailPrice
        '
        Me.UseCurrentRetailPrice.AutoSize = True
        Me.UseCurrentRetailPrice.Location = New System.Drawing.Point(49, 191)
        Me.UseCurrentRetailPrice.Name = "UseCurrentRetailPrice"
        Me.UseCurrentRetailPrice.Size = New System.Drawing.Size(139, 17)
        Me.UseCurrentRetailPrice.TabIndex = 14
        Me.UseCurrentRetailPrice.Text = "Use Current Retail Price"
        Me.UseCurrentRetailPrice.UseVisualStyleBackColor = True
        '
        'UseCurrentInStock
        '
        Me.UseCurrentInStock.AutoSize = True
        Me.UseCurrentInStock.Location = New System.Drawing.Point(49, 214)
        Me.UseCurrentInStock.Name = "UseCurrentInStock"
        Me.UseCurrentInStock.Size = New System.Drawing.Size(125, 17)
        Me.UseCurrentInStock.TabIndex = 15
        Me.UseCurrentInStock.Text = "Use Current In Stock"
        Me.UseCurrentInStock.UseVisualStyleBackColor = True
        '
        'UseProductNoAsSKU
        '
        Me.UseProductNoAsSKU.AutoSize = True
        Me.UseProductNoAsSKU.Location = New System.Drawing.Point(49, 237)
        Me.UseProductNoAsSKU.Name = "UseProductNoAsSKU"
        Me.UseProductNoAsSKU.Size = New System.Drawing.Size(164, 17)
        Me.UseProductNoAsSKU.TabIndex = 16
        Me.UseProductNoAsSKU.Text = "Use Product Number as SKU"
        Me.UseProductNoAsSKU.UseVisualStyleBackColor = True
        '
        'UpdateCurrentProd
        '
        Me.UpdateCurrentProd.AutoSize = True
        Me.UpdateCurrentProd.Location = New System.Drawing.Point(49, 260)
        Me.UpdateCurrentProd.Name = "UpdateCurrentProd"
        Me.UpdateCurrentProd.Size = New System.Drawing.Size(151, 17)
        Me.UpdateCurrentProd.TabIndex = 17
        Me.UpdateCurrentProd.Text = "Update Existing Products?"
        Me.UpdateCurrentProd.UseVisualStyleBackColor = True
        '
        'txtDataPath
        '
        Me.txtDataPath.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDataPath.HideSelection = False
        Me.txtDataPath.Location = New System.Drawing.Point(330, 489)
        Me.txtDataPath.Name = "txtDataPath"
        Me.txtDataPath.ReadOnly = True
        Me.txtDataPath.Size = New System.Drawing.Size(282, 20)
        Me.txtDataPath.TabIndex = 18
        Me.txtDataPath.TabStop = False
        Me.txtDataPath.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtDataPath.WordWrap = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Location = New System.Drawing.Point(49, 283)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(115, 17)
        Me.CheckBox1.TabIndex = 19
        Me.CheckBox1.Text = "Add New Products"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(-3, 483)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(615, 26)
        Me.ProgressBar1.TabIndex = 20
        Me.ProgressBar1.Visible = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(609, 506)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.txtDataPath)
        Me.Controls.Add(Me.UpdateCurrentProd)
        Me.Controls.Add(Me.UseProductNoAsSKU)
        Me.Controls.Add(Me.UseCurrentInStock)
        Me.Controls.Add(Me.UseCurrentRetailPrice)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "Big Commerce Product Import"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents UseCurrentRetailPrice As System.Windows.Forms.CheckBox
    Friend WithEvents UseCurrentInStock As System.Windows.Forms.CheckBox
    Friend WithEvents UseProductNoAsSKU As System.Windows.Forms.CheckBox
    Friend WithEvents UpdateCurrentProd As System.Windows.Forms.CheckBox
    Friend WithEvents txtDataPath As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar

End Class
