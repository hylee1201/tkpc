<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmReceiptInfo
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmReceiptInfo))
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.txtAddr1 = New System.Windows.Forms.TextBox()
        Me.txtAddr2 = New System.Windows.Forms.TextBox()
        Me.txtPhone = New System.Windows.Forms.TextBox()
        Me.txtRegNo = New System.Windows.Forms.TextBox()
        Me.txtTreasurer = New System.Windows.Forms.TextBox()
        Me.btOK = New System.Windows.Forms.Button()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.dtpReceiptIssue = New System.Windows.Forms.DateTimePicker()
        Me.picSignature = New System.Windows.Forms.PictureBox()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.pbReceipt = New System.Windows.Forms.ProgressBar()
        CType(Me.picSignature, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.Location = New System.Drawing.Point(91, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 17)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "교회명:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.Location = New System.Drawing.Point(78, 69)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(63, 17)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "교회주소:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.Location = New System.Drawing.Point(47, 134)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(94, 17)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "교회 전화번호:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label4.Location = New System.Drawing.Point(47, 167)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(94, 17)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "교회 등록번호:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label5.Location = New System.Drawing.Point(65, 201)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(76, 17)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "재정담당자:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label6.Location = New System.Drawing.Point(34, 270)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(107, 17)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "재정담당자 싸인:"
        '
        'txtName
        '
        Me.txtName.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.Location = New System.Drawing.Point(149, 32)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(216, 23)
        Me.txtName.TabIndex = 6
        '
        'txtAddr1
        '
        Me.txtAddr1.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddr1.Location = New System.Drawing.Point(149, 65)
        Me.txtAddr1.Name = "txtAddr1"
        Me.txtAddr1.Size = New System.Drawing.Size(216, 23)
        Me.txtAddr1.TabIndex = 7
        '
        'txtAddr2
        '
        Me.txtAddr2.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtAddr2.Location = New System.Drawing.Point(149, 98)
        Me.txtAddr2.Name = "txtAddr2"
        Me.txtAddr2.Size = New System.Drawing.Size(216, 23)
        Me.txtAddr2.TabIndex = 8
        '
        'txtPhone
        '
        Me.txtPhone.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPhone.Location = New System.Drawing.Point(148, 131)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(217, 23)
        Me.txtPhone.TabIndex = 9
        '
        'txtRegNo
        '
        Me.txtRegNo.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegNo.Location = New System.Drawing.Point(148, 164)
        Me.txtRegNo.Name = "txtRegNo"
        Me.txtRegNo.Size = New System.Drawing.Size(217, 23)
        Me.txtRegNo.TabIndex = 10
        '
        'txtTreasurer
        '
        Me.txtTreasurer.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTreasurer.Location = New System.Drawing.Point(149, 197)
        Me.txtTreasurer.Name = "txtTreasurer"
        Me.txtTreasurer.Size = New System.Drawing.Size(216, 23)
        Me.txtTreasurer.TabIndex = 11
        '
        'btOK
        '
        Me.btOK.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btOK.Location = New System.Drawing.Point(229, 366)
        Me.btOK.Name = "btOK"
        Me.btOK.Size = New System.Drawing.Size(75, 29)
        Me.btOK.TabIndex = 13
        Me.btOK.Text = "실행"
        Me.btOK.UseVisualStyleBackColor = True
        '
        'btCancel
        '
        Me.btCancel.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btCancel.Location = New System.Drawing.Point(112, 366)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(75, 29)
        Me.btCancel.TabIndex = 14
        Me.btCancel.Text = "취소"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label7.Location = New System.Drawing.Point(78, 234)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(63, 17)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "발행일자:"
        '
        'dtpReceiptIssue
        '
        Me.dtpReceiptIssue.Font = New System.Drawing.Font("맑은 고딕", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.dtpReceiptIssue.Location = New System.Drawing.Point(149, 231)
        Me.dtpReceiptIssue.Name = "dtpReceiptIssue"
        Me.dtpReceiptIssue.Size = New System.Drawing.Size(216, 23)
        Me.dtpReceiptIssue.TabIndex = 16
        '
        'picSignature
        '
        Me.picSignature.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picSignature.Location = New System.Drawing.Point(149, 268)
        Me.picSignature.Name = "picSignature"
        Me.picSignature.Size = New System.Drawing.Size(216, 62)
        Me.picSignature.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.picSignature.TabIndex = 12
        Me.picSignature.TabStop = False
        '
        'lblPath
        '
        Me.lblPath.AutoSize = True
        Me.lblPath.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPath.Location = New System.Drawing.Point(98, 333)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(91, 15)
        Me.lblPath.TabIndex = 17
        Me.lblPath.Text = "C:\TKPC_ROOT\"
        '
        'pbReceipt
        '
        Me.pbReceipt.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pbReceipt.Location = New System.Drawing.Point(0, 429)
        Me.pbReceipt.Name = "pbReceipt"
        Me.pbReceipt.Size = New System.Drawing.Size(417, 10)
        Me.pbReceipt.TabIndex = 18
        Me.pbReceipt.Visible = False
        '
        'frmReceiptInfo
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(417, 439)
        Me.Controls.Add(Me.pbReceipt)
        Me.Controls.Add(Me.lblPath)
        Me.Controls.Add(Me.dtpReceiptIssue)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btOK)
        Me.Controls.Add(Me.picSignature)
        Me.Controls.Add(Me.txtTreasurer)
        Me.Controls.Add(Me.txtRegNo)
        Me.Controls.Add(Me.txtPhone)
        Me.Controls.Add(Me.txtAddr2)
        Me.Controls.Add(Me.txtAddr1)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmReceiptInfo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "헌금영수증 발행자 입력창"
        CType(Me.picSignature, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents txtName As TextBox
    Friend WithEvents txtAddr1 As TextBox
    Friend WithEvents txtAddr2 As TextBox
    Friend WithEvents txtPhone As TextBox
    Friend WithEvents txtRegNo As TextBox
    Friend WithEvents txtTreasurer As TextBox
    Friend WithEvents picSignature As PictureBox
    Friend WithEvents btOK As Button
    Friend WithEvents btCancel As Button
    Friend WithEvents Label7 As Label
    Friend WithEvents dtpReceiptIssue As DateTimePicker
    Friend WithEvents lblPath As Label
    Friend WithEvents pbReceipt As ProgressBar
End Class
