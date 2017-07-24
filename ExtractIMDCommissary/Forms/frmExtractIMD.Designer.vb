<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmExtractIMD
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lvIMD = New System.Windows.Forms.ListView()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtQuery = New System.Windows.Forms.TextBox()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.pbProgress = New System.Windows.Forms.ProgressBar()
        Me.btnExtract = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lvIMD)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 158)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(801, 275)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Data List"
        '
        'lvIMD
        '
        Me.lvIMD.FullRowSelect = True
        Me.lvIMD.GridLines = True
        Me.lvIMD.Location = New System.Drawing.Point(6, 19)
        Me.lvIMD.Name = "lvIMD"
        Me.lvIMD.Size = New System.Drawing.Size(789, 245)
        Me.lvIMD.TabIndex = 0
        Me.lvIMD.UseCompatibleStateImageBehavior = False
        Me.lvIMD.View = System.Windows.Forms.View.Details
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.btnCancel)
        Me.GroupBox2.Controls.Add(Me.btnOK)
        Me.GroupBox2.Controls.Add(Me.txtQuery)
        Me.GroupBox2.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(801, 140)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Query"
        '
        'txtQuery
        '
        Me.txtQuery.Location = New System.Drawing.Point(6, 19)
        Me.txtQuery.Multiline = True
        Me.txtQuery.Name = "txtQuery"
        Me.txtQuery.Size = New System.Drawing.Size(680, 111)
        Me.txtQuery.TabIndex = 0
        '
        'btnOK
        '
        Me.btnOK.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOK.Location = New System.Drawing.Point(692, 19)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(103, 48)
        Me.btnOK.TabIndex = 1
        Me.btnOK.Text = "&OK"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(692, 84)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(103, 46)
        Me.btnCancel.TabIndex = 2
        Me.btnCancel.Text = "&Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'pbProgress
        '
        Me.pbProgress.Location = New System.Drawing.Point(12, 439)
        Me.pbProgress.Name = "pbProgress"
        Me.pbProgress.Size = New System.Drawing.Size(713, 23)
        Me.pbProgress.TabIndex = 2
        '
        'btnExtract
        '
        Me.btnExtract.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExtract.Location = New System.Drawing.Point(731, 439)
        Me.btnExtract.Name = "btnExtract"
        Me.btnExtract.Size = New System.Drawing.Size(82, 23)
        Me.btnExtract.TabIndex = 3
        Me.btnExtract.Text = "&Extract"
        Me.btnExtract.UseVisualStyleBackColor = True
        '
        'frmExtractIMD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(820, 470)
        Me.Controls.Add(Me.btnExtract)
        Me.Controls.Add(Me.pbProgress)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmExtractIMD"
        Me.Text = "Extract Commisary Data"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lvIMD As System.Windows.Forms.ListView
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents txtQuery As System.Windows.Forms.TextBox
    Friend WithEvents pbProgress As System.Windows.Forms.ProgressBar
    Friend WithEvents btnExtract As System.Windows.Forms.Button

End Class
