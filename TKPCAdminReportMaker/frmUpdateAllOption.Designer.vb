<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUpdateAllOption
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUpdateAllOption))
        Me.dtpOption = New System.Windows.Forms.DateTimePicker()
        Me.cbOption = New System.Windows.Forms.ComboBox()
        Me.dgChosen = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.cbOption2 = New System.Windows.Forms.ComboBox()
        Me.btOK = New System.Windows.Forms.Button()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.한글명 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.person_id = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.등록 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.나이 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.성별 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.pbAll = New System.Windows.Forms.ProgressBar()
        CType(Me.dgChosen, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'dtpOption
        '
        Me.dtpOption.CalendarFont = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpOption.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.dtpOption.Location = New System.Drawing.Point(118, 216)
        Me.dtpOption.Name = "dtpOption"
        Me.dtpOption.Size = New System.Drawing.Size(200, 25)
        Me.dtpOption.TabIndex = 0
        '
        'cbOption
        '
        Me.cbOption.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbOption.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cbOption.FormattingEnabled = True
        Me.cbOption.Items.AddRange(New Object() {"목사", "전도사", "평신도", "서리집사", "안수집사", "시무장로", "사역장로", "은퇴장로", "증경장로", "시무권사", "사역권사", "명예권사"})
        Me.cbOption.Location = New System.Drawing.Point(118, 180)
        Me.cbOption.Name = "cbOption"
        Me.cbOption.Size = New System.Drawing.Size(154, 25)
        Me.cbOption.TabIndex = 1
        '
        'dgChosen
        '
        Me.dgChosen.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgChosen.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.한글명, Me.person_id, Me.등록, Me.나이, Me.성별})
        Me.dgChosen.Location = New System.Drawing.Point(26, 15)
        Me.dgChosen.Name = "dgChosen"
        Me.dgChosen.Size = New System.Drawing.Size(396, 150)
        Me.dgChosen.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.Location = New System.Drawing.Point(24, 183)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(94, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "일괄처리 옵션:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label2.Location = New System.Drawing.Point(24, 220)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(94, 17)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "일괄처리 날짜:"
        '
        'cbOption2
        '
        Me.cbOption2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cbOption2.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.cbOption2.FormattingEnabled = True
        Me.cbOption2.Items.AddRange(New Object() {"infant_baptism", "confirmation", "baptism"})
        Me.cbOption2.Location = New System.Drawing.Point(118, 180)
        Me.cbOption2.Name = "cbOption2"
        Me.cbOption2.Size = New System.Drawing.Size(154, 25)
        Me.cbOption2.TabIndex = 5
        '
        'btOK
        '
        Me.btOK.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btOK.Location = New System.Drawing.Point(243, 288)
        Me.btOK.Name = "btOK"
        Me.btOK.Size = New System.Drawing.Size(75, 29)
        Me.btOK.TabIndex = 6
        Me.btOK.Text = "저장"
        Me.btOK.UseVisualStyleBackColor = True
        '
        'btCancel
        '
        Me.btCancel.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btCancel.Location = New System.Drawing.Point(130, 288)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(75, 29)
        Me.btCancel.TabIndex = 7
        Me.btCancel.Text = "취소"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label3.Location = New System.Drawing.Point(22, 251)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(361, 17)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "(일괄처리 날짜: 직분일 경우 임명일, 신급일 경우 받은 날짜)"
        '
        '한글명
        '
        Me.한글명.HeaderText = "한글명"
        Me.한글명.Name = "한글명"
        '
        'person_id
        '
        Me.person_id.HeaderText = "person_id"
        Me.person_id.Name = "person_id"
        Me.person_id.Visible = False
        '
        '등록
        '
        Me.등록.HeaderText = "등록"
        Me.등록.Name = "등록"
        Me.등록.Width = 70
        '
        '나이
        '
        Me.나이.HeaderText = "나이"
        Me.나이.Name = "나이"
        Me.나이.Width = 60
        '
        '성별
        '
        Me.성별.HeaderText = "성별"
        Me.성별.Name = "성별"
        Me.성별.Width = 60
        '
        'pbAll
        '
        Me.pbAll.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pbAll.Location = New System.Drawing.Point(0, 335)
        Me.pbAll.Name = "pbAll"
        Me.pbAll.Size = New System.Drawing.Size(449, 13)
        Me.pbAll.TabIndex = 9
        Me.pbAll.Visible = False
        '
        'frmUpdateAllOption
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(449, 348)
        Me.Controls.Add(Me.pbAll)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btOK)
        Me.Controls.Add(Me.cbOption2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgChosen)
        Me.Controls.Add(Me.cbOption)
        Me.Controls.Add(Me.dtpOption)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmUpdateAllOption"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "일괄처리 옵션창"
        CType(Me.dgChosen, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents dtpOption As DateTimePicker
    Friend WithEvents cbOption As ComboBox
    Friend WithEvents dgChosen As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents cbOption2 As ComboBox
    Friend WithEvents btOK As Button
    Friend WithEvents btCancel As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents 한글명 As DataGridViewTextBoxColumn
    Friend WithEvents person_id As DataGridViewTextBoxColumn
    Friend WithEvents 등록 As DataGridViewTextBoxColumn
    Friend WithEvents 나이 As DataGridViewTextBoxColumn
    Friend WithEvents 성별 As DataGridViewTextBoxColumn
    Friend WithEvents pbAll As ProgressBar
End Class
