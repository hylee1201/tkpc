<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmSwitchFamilyChild
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSwitchFamilyChild))
        Me.btCancel = New System.Windows.Forms.Button()
        Me.dgChosen = New System.Windows.Forms.DataGridView()
        Me.lblMsg = New System.Windows.Forms.Label()
        CType(Me.dgChosen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btCancel
        '
        Me.btCancel.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCancel.Location = New System.Drawing.Point(229, 285)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(84, 32)
        Me.btCancel.TabIndex = 19
        Me.btCancel.Text = "닫기"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'dgChosen
        '
        Me.dgChosen.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgChosen.Location = New System.Drawing.Point(14, 43)
        Me.dgChosen.Name = "dgChosen"
        Me.dgChosen.ReadOnly = True
        Me.dgChosen.Size = New System.Drawing.Size(515, 224)
        Me.dgChosen.TabIndex = 17
        '
        'lblMsg
        '
        Me.lblMsg.AutoSize = True
        Me.lblMsg.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMsg.Location = New System.Drawing.Point(16, 18)
        Me.lblMsg.Name = "lblMsg"
        Me.lblMsg.Size = New System.Drawing.Size(46, 17)
        Me.lblMsg.TabIndex = 20
        Me.lblMsg.Text = "Label1"
        '
        'frmSwitchFamilyChild
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(543, 340)
        Me.Controls.Add(Me.lblMsg)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.dgChosen)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSwitchFamilyChild"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "가족 타입 child를 Parent로 변경"
        CType(Me.dgChosen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btCancel As Button
    Friend WithEvents dgChosen As DataGridView
    Friend WithEvents lblMsg As Label
End Class
