<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSwitchFamilyHead
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSwitchFamilyHead))
        Me.dgChosen = New System.Windows.Forms.DataGridView()
        Me.btSwitch = New System.Windows.Forms.Button()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.lblMsg = New System.Windows.Forms.Label()
        CType(Me.dgChosen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dgChosen
        '
        Me.dgChosen.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgChosen.Location = New System.Drawing.Point(14, 46)
        Me.dgChosen.Name = "dgChosen"
        Me.dgChosen.ReadOnly = True
        Me.dgChosen.Size = New System.Drawing.Size(515, 224)
        Me.dgChosen.TabIndex = 3
        '
        'btSwitch
        '
        Me.btSwitch.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btSwitch.Location = New System.Drawing.Point(294, 288)
        Me.btSwitch.Name = "btSwitch"
        Me.btSwitch.Size = New System.Drawing.Size(84, 32)
        Me.btSwitch.TabIndex = 15
        Me.btSwitch.Text = "바꾸기"
        Me.btSwitch.UseVisualStyleBackColor = True
        '
        'btCancel
        '
        Me.btCancel.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCancel.Location = New System.Drawing.Point(164, 288)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(84, 32)
        Me.btCancel.TabIndex = 16
        Me.btCancel.Text = "닫기"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'lblMsg
        '
        Me.lblMsg.AutoSize = True
        Me.lblMsg.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMsg.Location = New System.Drawing.Point(15, 14)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(46, 17)
        Me.lblMsg.TabIndex = 17
        Me.lblMsg.Text = "Label1"
        '
        'frmSwitchFamilyHead
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(543, 340)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.btSwitch)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.dgChosen)
        Me.Font = New System.Drawing.Font("Malgun Gothic", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSwitchFamilyHead"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "가족 세대주 변경"
        CType(Me.dgChosen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dgChosen As DataGridView
    Friend WithEvents btSwitch As Button
    Friend WithEvents btCancel As Button
    Friend WithEvents lblMsg As Label
End Class
