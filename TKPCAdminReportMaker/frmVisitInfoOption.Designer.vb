<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmVisitInfoOption
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
        Me.components = New System.ComponentModel.Container()
        Dim Appearance25 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance26 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance27 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance28 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance29 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance30 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance31 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance32 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance33 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance34 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance35 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance36 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmVisitInfoOption))
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ugResult = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.btOK = New System.Windows.Forms.Button()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.bnList = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.ToolStripButton1 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripLabel1 = New System.Windows.Forms.ToolStripLabel()
        Me.ToolStripButton2 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton3 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton4 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripTextBox1 = New System.Windows.Forms.ToolStripTextBox()
        Me.ToolStripSeparator4 = New System.Windows.Forms.ToolStripSeparator()
        Me.ToolStripButton5 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripButton6 = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbtCheckAll = New System.Windows.Forms.ToolStripButton()
        Me.tsbtUncheckAll = New System.Windows.Forms.ToolStripButton()
        Me.tsbtFindExit = New System.Windows.Forms.ToolStripButton()
        Me.spVisitOption = New System.Windows.Forms.SplitContainer()
        Me.dgVisitInfo = New System.Windows.Forms.DataGridView()
        Me.col1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.col2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.col3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.ugResult, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bnList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.bnList.SuspendLayout()
        CType(Me.spVisitOption, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.spVisitOption.Panel1.SuspendLayout()
        Me.spVisitOption.Panel2.SuspendLayout()
        Me.spVisitOption.SuspendLayout()
        CType(Me.dgVisitInfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ugResult
        '
        Appearance25.BackColor = System.Drawing.SystemColors.Window
        Appearance25.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.ugResult.DisplayLayout.Appearance = Appearance25
        Me.ugResult.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ExtendLastColumn
        Me.ugResult.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.ugResult.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance26.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance26.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance26.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance26.BorderColor = System.Drawing.SystemColors.Window
        Me.ugResult.DisplayLayout.GroupByBox.Appearance = Appearance26
        Appearance27.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ugResult.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance27
        Me.ugResult.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance28.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance28.BackColor2 = System.Drawing.SystemColors.Control
        Appearance28.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance28.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ugResult.DisplayLayout.GroupByBox.PromptAppearance = Appearance28
        Appearance29.BackColor = System.Drawing.SystemColors.Window
        Appearance29.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ugResult.DisplayLayout.Override.ActiveCellAppearance = Appearance29
        Appearance30.BackColor = System.Drawing.SystemColors.Highlight
        Appearance30.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.ugResult.DisplayLayout.Override.ActiveRowAppearance = Appearance30
        Me.ugResult.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Synchronized
        Me.ugResult.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.[False]
        Me.ugResult.DisplayLayout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.[False]
        Me.ugResult.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.ugResult.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance31.BackColor = System.Drawing.SystemColors.Window
        Me.ugResult.DisplayLayout.Override.CardAreaAppearance = Appearance31
        Appearance32.BorderColor = System.Drawing.Color.Silver
        Appearance32.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.ugResult.DisplayLayout.Override.CellAppearance = Appearance32
        Me.ugResult.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ugResult.DisplayLayout.Override.CellPadding = 0
        Me.ugResult.DisplayLayout.Override.ColumnSizingArea = Infragistics.Win.UltraWinGrid.ColumnSizingArea.EntireColumn
        Me.ugResult.DisplayLayout.Override.FilterUIType = Infragistics.Win.UltraWinGrid.FilterUIType.FilterRow
        Appearance33.BackColor = System.Drawing.SystemColors.Control
        Appearance33.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance33.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance33.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance33.BorderColor = System.Drawing.SystemColors.Window
        Me.ugResult.DisplayLayout.Override.GroupByRowAppearance = Appearance33
        Me.ugResult.DisplayLayout.Override.GroupBySummaryDisplayStyle = Infragistics.Win.UltraWinGrid.GroupBySummaryDisplayStyle.SummaryCellsAlwaysBelowDescription
        Appearance34.TextHAlignAsString = "Left"
        Me.ugResult.DisplayLayout.Override.HeaderAppearance = Appearance34
        Me.ugResult.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugResult.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance35.BackColor = System.Drawing.SystemColors.Window
        Appearance35.BorderColor = System.Drawing.Color.Silver
        Me.ugResult.DisplayLayout.Override.RowAppearance = Appearance35
        Me.ugResult.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand
        Me.ugResult.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.ColumnChooserButton
        Me.ugResult.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[True]
        Me.ugResult.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
        Me.ugResult.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.RowSelectorsOnly
        Me.ugResult.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.ExtendedAutoDrag
        Me.ugResult.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.ExtendedAutoDrag
        Me.ugResult.DisplayLayout.Override.SummaryDisplayArea = Infragistics.Win.UltraWinGrid.SummaryDisplayAreas.None
        Appearance36.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ugResult.DisplayLayout.Override.TemplateAddRowAppearance = Appearance36
        Me.ugResult.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugResult.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ugResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugResult.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugResult.Location = New System.Drawing.Point(0, 0)
        Me.ugResult.Name = "ugResult"
        Me.ugResult.Size = New System.Drawing.Size(309, 225)
        Me.ugResult.TabIndex = 22
        '
        'btOK
        '
        Me.btOK.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btOK.Location = New System.Drawing.Point(450, 273)
        Me.btOK.Name = "btOK"
        Me.btOK.Size = New System.Drawing.Size(112, 28)
        Me.btOK.TabIndex = 23
        Me.btOK.Text = "실행"
        Me.btOK.UseVisualStyleBackColor = True
        '
        'btCancel
        '
        Me.btCancel.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCancel.Location = New System.Drawing.Point(259, 273)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(112, 28)
        Me.btCancel.TabIndex = 24
        Me.btCancel.Text = "취소"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'bnList
        '
        Me.bnList.AddNewItem = Me.ToolStripButton1
        Me.bnList.CountItem = Me.ToolStripLabel1
        Me.bnList.DeleteItem = Me.ToolStripButton2
        Me.bnList.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripButton3, Me.ToolStripButton4, Me.ToolStripSeparator3, Me.ToolStripTextBox1, Me.ToolStripLabel1, Me.ToolStripSeparator4, Me.ToolStripButton5, Me.ToolStripButton6, Me.ToolStripSeparator5, Me.ToolStripButton1, Me.ToolStripButton2, Me.tsbtCheckAll, Me.tsbtUncheckAll, Me.tsbtFindExit})
        Me.bnList.Location = New System.Drawing.Point(0, 0)
        Me.bnList.MoveFirstItem = Me.ToolStripButton3
        Me.bnList.MoveLastItem = Me.ToolStripButton6
        Me.bnList.MoveNextItem = Me.ToolStripButton5
        Me.bnList.MovePreviousItem = Me.ToolStripButton4
        Me.bnList.Name = "bnList"
        Me.bnList.PositionItem = Me.ToolStripTextBox1
        Me.bnList.Size = New System.Drawing.Size(822, 25)
        Me.bnList.TabIndex = 25
        '
        'ToolStripButton1
        '
        Me.ToolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton1.Image = CType(resources.GetObject("ToolStripButton1.Image"), System.Drawing.Image)
        Me.ToolStripButton1.Name = "ToolStripButton1"
        Me.ToolStripButton1.RightToLeftAutoMirrorImage = True
        Me.ToolStripButton1.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton1.Text = "Add new"
        Me.ToolStripButton1.Visible = False
        '
        'ToolStripLabel1
        '
        Me.ToolStripLabel1.Name = "ToolStripLabel1"
        Me.ToolStripLabel1.Size = New System.Drawing.Size(27, 22)
        Me.ToolStripLabel1.Text = "/{0}"
        Me.ToolStripLabel1.ToolTipText = "Total number of items"
        '
        'ToolStripButton2
        '
        Me.ToolStripButton2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton2.Image = CType(resources.GetObject("ToolStripButton2.Image"), System.Drawing.Image)
        Me.ToolStripButton2.Name = "ToolStripButton2"
        Me.ToolStripButton2.RightToLeftAutoMirrorImage = True
        Me.ToolStripButton2.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton2.Text = "Delete"
        Me.ToolStripButton2.Visible = False
        '
        'ToolStripButton3
        '
        Me.ToolStripButton3.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton3.Image = CType(resources.GetObject("ToolStripButton3.Image"), System.Drawing.Image)
        Me.ToolStripButton3.Name = "ToolStripButton3"
        Me.ToolStripButton3.RightToLeftAutoMirrorImage = True
        Me.ToolStripButton3.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton3.Text = "Move first"
        '
        'ToolStripButton4
        '
        Me.ToolStripButton4.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton4.Image = CType(resources.GetObject("ToolStripButton4.Image"), System.Drawing.Image)
        Me.ToolStripButton4.Name = "ToolStripButton4"
        Me.ToolStripButton4.RightToLeftAutoMirrorImage = True
        Me.ToolStripButton4.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton4.Text = "Move previous"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripTextBox1
        '
        Me.ToolStripTextBox1.AccessibleName = "Position"
        Me.ToolStripTextBox1.AutoSize = False
        Me.ToolStripTextBox1.Name = "ToolStripTextBox1"
        Me.ToolStripTextBox1.Size = New System.Drawing.Size(58, 23)
        Me.ToolStripTextBox1.Text = "0"
        Me.ToolStripTextBox1.ToolTipText = "Current position"
        '
        'ToolStripSeparator4
        '
        Me.ToolStripSeparator4.Name = "ToolStripSeparator4"
        Me.ToolStripSeparator4.Size = New System.Drawing.Size(6, 25)
        '
        'ToolStripButton5
        '
        Me.ToolStripButton5.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton5.Image = CType(resources.GetObject("ToolStripButton5.Image"), System.Drawing.Image)
        Me.ToolStripButton5.Name = "ToolStripButton5"
        Me.ToolStripButton5.RightToLeftAutoMirrorImage = True
        Me.ToolStripButton5.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton5.Text = "Move next"
        '
        'ToolStripButton6
        '
        Me.ToolStripButton6.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.ToolStripButton6.Image = CType(resources.GetObject("ToolStripButton6.Image"), System.Drawing.Image)
        Me.ToolStripButton6.Name = "ToolStripButton6"
        Me.ToolStripButton6.RightToLeftAutoMirrorImage = True
        Me.ToolStripButton6.Size = New System.Drawing.Size(23, 22)
        Me.ToolStripButton6.Text = "Move last"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(6, 25)
        '
        'tsbtCheckAll
        '
        Me.tsbtCheckAll.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.tsbtCheckAll.Image = Global.TKPC.My.Resources.Resources.checkAll
        Me.tsbtCheckAll.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtCheckAll.Name = "tsbtCheckAll"
        Me.tsbtCheckAll.Size = New System.Drawing.Size(85, 22)
        Me.tsbtCheckAll.Text = "모두 선택"
        '
        'tsbtUncheckAll
        '
        Me.tsbtUncheckAll.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.tsbtUncheckAll.Image = Global.TKPC.My.Resources.Resources.uncheckAll
        Me.tsbtUncheckAll.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtUncheckAll.Name = "tsbtUncheckAll"
        Me.tsbtUncheckAll.Size = New System.Drawing.Size(85, 22)
        Me.tsbtUncheckAll.Text = "모두 해제"
        '
        'tsbtFindExit
        '
        Me.tsbtFindExit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbtFindExit.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtFindExit.Image = Global.TKPC.My.Resources.Resources.exit1
        Me.tsbtFindExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtFindExit.Name = "tsbtFindExit"
        Me.tsbtFindExit.Size = New System.Drawing.Size(54, 22)
        Me.tsbtFindExit.Text = "종료"
        '
        'spVisitOption
        '
        Me.spVisitOption.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.spVisitOption.Location = New System.Drawing.Point(0, 27)
        Me.spVisitOption.Name = "spVisitOption"
        '
        'spVisitOption.Panel1
        '
        Me.spVisitOption.Panel1.Controls.Add(Me.ugResult)
        '
        'spVisitOption.Panel2
        '
        Me.spVisitOption.Panel2.Controls.Add(Me.dgVisitInfo)
        Me.spVisitOption.Size = New System.Drawing.Size(820, 227)
        Me.spVisitOption.SplitterDistance = 311
        Me.spVisitOption.SplitterWidth = 5
        Me.spVisitOption.TabIndex = 26
        '
        'dgVisitInfo
        '
        Me.dgVisitInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgVisitInfo.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.col1, Me.col2, Me.col3})
        Me.dgVisitInfo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgVisitInfo.Location = New System.Drawing.Point(0, 0)
        Me.dgVisitInfo.Name = "dgVisitInfo"
        Me.dgVisitInfo.Size = New System.Drawing.Size(502, 225)
        Me.dgVisitInfo.TabIndex = 0
        '
        'col1
        '
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.col1.DefaultCellStyle = DataGridViewCellStyle1
        Me.col1.HeaderText = "Date"
        Me.col1.Name = "col1"
        Me.col1.Width = 90
        '
        'col2
        '
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.col2.DefaultCellStyle = DataGridViewCellStyle2
        Me.col2.HeaderText = "Korean Name"
        Me.col2.Name = "col2"
        Me.col2.Width = 70
        '
        'col3
        '
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.col3.DefaultCellStyle = DataGridViewCellStyle3
        Me.col3.HeaderText = "Note"
        Me.col3.Name = "col3"
        Me.col3.Width = 500
        '
        'frmVisitInfoOption
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(822, 322)
        Me.Controls.Add(Me.spVisitOption)
        Me.Controls.Add(Me.bnList)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.btOK)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmVisitInfoOption"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "심방기록 옵션"
        CType(Me.ugResult, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bnList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.bnList.ResumeLayout(False)
        Me.bnList.PerformLayout()
        Me.spVisitOption.Panel1.ResumeLayout(False)
        Me.spVisitOption.Panel2.ResumeLayout(False)
        CType(Me.spVisitOption, System.ComponentModel.ISupportInitialize).EndInit()
        Me.spVisitOption.ResumeLayout(False)
        CType(Me.dgVisitInfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ugResult As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btOK As Button
    Friend WithEvents btCancel As Button
    Friend WithEvents bnList As BindingNavigator
    Friend WithEvents ToolStripButton1 As ToolStripButton
    Friend WithEvents ToolStripLabel1 As ToolStripLabel
    Friend WithEvents ToolStripButton2 As ToolStripButton
    Friend WithEvents ToolStripButton3 As ToolStripButton
    Friend WithEvents ToolStripButton4 As ToolStripButton
    Friend WithEvents ToolStripSeparator3 As ToolStripSeparator
    Friend WithEvents ToolStripTextBox1 As ToolStripTextBox
    Friend WithEvents ToolStripSeparator4 As ToolStripSeparator
    Friend WithEvents ToolStripButton5 As ToolStripButton
    Friend WithEvents ToolStripButton6 As ToolStripButton
    Friend WithEvents ToolStripSeparator5 As ToolStripSeparator
    Friend WithEvents tsbtCheckAll As ToolStripButton
    Friend WithEvents tsbtUncheckAll As ToolStripButton
    Friend WithEvents tsbtFindExit As ToolStripButton
    Friend WithEvents spVisitOption As SplitContainer
    Friend WithEvents dgVisitInfo As DataGridView
    Friend WithEvents col1 As DataGridViewTextBoxColumn
    Friend WithEvents col2 As DataGridViewTextBoxColumn
    Friend WithEvents col3 As DataGridViewTextBoxColumn
End Class
