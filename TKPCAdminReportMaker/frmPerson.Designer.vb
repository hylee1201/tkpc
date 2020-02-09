<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPerson
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPerson))
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance11 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Dim Appearance12 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance()
        Me.bnList = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.tsbtCheckAll = New System.Windows.Forms.ToolStripButton()
        Me.tsbtUncheckAll = New System.Windows.Forms.ToolStripButton()
        Me.tsbtExcel = New System.Windows.Forms.ToolStripDropDownButton()
        Me.모든데이터ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.선택한데이터만ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.선택하지않은데이터만ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsbtExit = New System.Windows.Forms.ToolStripButton()
        Me.tscbDataOption = New System.Windows.Forms.ToolStripComboBox()
        Me.tsbtFamily = New System.Windows.Forms.ToolStripDropDownButton()
        Me.가족모드로ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.교적카드로ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.설문지양식으로ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.교인주소록으로ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.개인모드로ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ugResult = New Infragistics.Win.UltraWinGrid.UltraGrid()
        Me.ugexcel = New Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tspbTKPC = New System.Windows.Forms.ToolStripProgressBar()
        CType(Me.bnList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.bnList.SuspendLayout()
        CType(Me.ugResult, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'bnList
        '
        Me.bnList.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.bnList.CountItem = Me.BindingNavigatorCountItem
        Me.bnList.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.bnList.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem, Me.tsbtCheckAll, Me.tsbtUncheckAll, Me.tsbtExcel, Me.tsbtExit, Me.tscbDataOption, Me.tsbtFamily})
        Me.bnList.Location = New System.Drawing.Point(0, 0)
        Me.bnList.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.bnList.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.bnList.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.bnList.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.bnList.Name = "bnList"
        Me.bnList.PositionItem = Me.BindingNavigatorPositionItem
        Me.bnList.Size = New System.Drawing.Size(1013, 30)
        Me.bnList.TabIndex = 18
        Me.bnList.Text = "BindingNavigator1"
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 27)
        Me.BindingNavigatorAddNewItem.Text = "Add new"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(35, 27)
        Me.BindingNavigatorCountItem.Text = "of {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Total number of items"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 27)
        Me.BindingNavigatorDeleteItem.Text = "Delete"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 27)
        Me.BindingNavigatorMoveFirstItem.Text = "Move first"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 27)
        Me.BindingNavigatorMovePreviousItem.Text = "Move previous"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 30)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Current position"
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator1"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 30)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 27)
        Me.BindingNavigatorMoveNextItem.Text = "Move next"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 27)
        Me.BindingNavigatorMoveLastItem.Text = "Move last"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 30)
        '
        'tsbtCheckAll
        '
        Me.tsbtCheckAll.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtCheckAll.Image = Global.TKPC.My.Resources.Resources.checkAll
        Me.tsbtCheckAll.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtCheckAll.Name = "tsbtCheckAll"
        Me.tsbtCheckAll.Size = New System.Drawing.Size(81, 27)
        Me.tsbtCheckAll.Text = "모두 선택"
        Me.tsbtCheckAll.Visible = False
        '
        'tsbtUncheckAll
        '
        Me.tsbtUncheckAll.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtUncheckAll.Image = Global.TKPC.My.Resources.Resources.uncheckAll
        Me.tsbtUncheckAll.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtUncheckAll.Name = "tsbtUncheckAll"
        Me.tsbtUncheckAll.Size = New System.Drawing.Size(81, 27)
        Me.tsbtUncheckAll.Text = "모두 해제"
        Me.tsbtUncheckAll.Visible = False
        '
        'tsbtExcel
        '
        Me.tsbtExcel.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.모든데이터ToolStripMenuItem, Me.선택한데이터만ToolStripMenuItem, Me.선택하지않은데이터만ToolStripMenuItem})
        Me.tsbtExcel.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtExcel.Image = Global.TKPC.My.Resources.Resources.excel1
        Me.tsbtExcel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtExcel.Name = "tsbtExcel"
        Me.tsbtExcel.Size = New System.Drawing.Size(102, 27)
        Me.tsbtExcel.Text = "엑셀로 출력"
        Me.tsbtExcel.Visible = False
        '
        '모든데이터ToolStripMenuItem
        '
        Me.모든데이터ToolStripMenuItem.Name = "모든데이터ToolStripMenuItem"
        Me.모든데이터ToolStripMenuItem.Size = New System.Drawing.Size(206, 24)
        Me.모든데이터ToolStripMenuItem.Text = "모든 데이터"
        '
        '선택한데이터만ToolStripMenuItem
        '
        Me.선택한데이터만ToolStripMenuItem.Name = "선택한데이터만ToolStripMenuItem"
        Me.선택한데이터만ToolStripMenuItem.Size = New System.Drawing.Size(206, 24)
        Me.선택한데이터만ToolStripMenuItem.Text = "선택한 데이터만"
        '
        '선택하지않은데이터만ToolStripMenuItem
        '
        Me.선택하지않은데이터만ToolStripMenuItem.Name = "선택하지않은데이터만ToolStripMenuItem"
        Me.선택하지않은데이터만ToolStripMenuItem.Size = New System.Drawing.Size(206, 24)
        Me.선택하지않은데이터만ToolStripMenuItem.Text = "선택하지 않은 데이터만"
        '
        'tsbtExit
        '
        Me.tsbtExit.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbtExit.Font = New System.Drawing.Font("Calibri", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtExit.Image = Global.TKPC.My.Resources.Resources.exit1
        Me.tsbtExit.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtExit.Name = "tsbtExit"
        Me.tsbtExit.Size = New System.Drawing.Size(60, 27)
        Me.tsbtExit.Text = "종료"
        '
        'tscbDataOption
        '
        Me.tscbDataOption.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.tscbDataOption.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tscbDataOption.Items.AddRange(New Object() {"모든 데이터 (고인포함)", "등록교인만 (고인제외)", "이적교인만 (고인제외)", "새교우만 (고인제외)", "직분자만 (고인제외한 등록교인중)", "유아세례만 (고인제외한 등록교인중)", "세례교인만 (고인제외한 등록교인중)", "입교교인만 (고인제외한 등록교인중)"})
        Me.tscbDataOption.Name = "tscbDataOption"
        Me.tscbDataOption.Size = New System.Drawing.Size(191, 30)
        '
        'tsbtFamily
        '
        Me.tsbtFamily.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.가족모드로ToolStripMenuItem, Me.교적카드로ToolStripMenuItem, Me.설문지양식으로ToolStripMenuItem, Me.교인주소록으로ToolStripMenuItem, Me.개인모드로ToolStripMenuItem})
        Me.tsbtFamily.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtFamily.Image = Global.TKPC.My.Resources.Resources.people
        Me.tsbtFamily.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtFamily.Name = "tsbtFamily"
        Me.tsbtFamily.Size = New System.Drawing.Size(86, 27)
        Me.tsbtFamily.Text = "모드변경"
        Me.tsbtFamily.Visible = False
        '
        '가족모드로ToolStripMenuItem
        '
        Me.가족모드로ToolStripMenuItem.Name = "가족모드로ToolStripMenuItem"
        Me.가족모드로ToolStripMenuItem.Size = New System.Drawing.Size(166, 24)
        Me.가족모드로ToolStripMenuItem.Text = "가족모드로"
        '
        '교적카드로ToolStripMenuItem
        '
        Me.교적카드로ToolStripMenuItem.Name = "교적카드로ToolStripMenuItem"
        Me.교적카드로ToolStripMenuItem.Size = New System.Drawing.Size(166, 24)
        Me.교적카드로ToolStripMenuItem.Text = "교적카드로"
        '
        '설문지양식으로ToolStripMenuItem
        '
        Me.설문지양식으로ToolStripMenuItem.Name = "설문지양식으로ToolStripMenuItem"
        Me.설문지양식으로ToolStripMenuItem.Size = New System.Drawing.Size(166, 24)
        Me.설문지양식으로ToolStripMenuItem.Text = "설문지 양식으로"
        '
        '교인주소록으로ToolStripMenuItem
        '
        Me.교인주소록으로ToolStripMenuItem.Name = "교인주소록으로ToolStripMenuItem"
        Me.교인주소록으로ToolStripMenuItem.Size = New System.Drawing.Size(166, 24)
        Me.교인주소록으로ToolStripMenuItem.Text = "교인 주소록으로"
        '
        '개인모드로ToolStripMenuItem
        '
        Me.개인모드로ToolStripMenuItem.Name = "개인모드로ToolStripMenuItem"
        Me.개인모드로ToolStripMenuItem.Size = New System.Drawing.Size(166, 24)
        Me.개인모드로ToolStripMenuItem.Text = "개인모드로"
        '
        'ugResult
        '
        Appearance1.BackColor = System.Drawing.SystemColors.Window
        Appearance1.BorderColor = System.Drawing.SystemColors.InactiveCaption
        Me.ugResult.DisplayLayout.Appearance = Appearance1
        Me.ugResult.DisplayLayout.AutoFitStyle = Infragistics.Win.UltraWinGrid.AutoFitStyle.ExtendLastColumn
        Me.ugResult.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Me.ugResult.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.[False]
        Appearance2.BackColor = System.Drawing.SystemColors.ActiveBorder
        Appearance2.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance2.BackGradientStyle = Infragistics.Win.GradientStyle.Vertical
        Appearance2.BorderColor = System.Drawing.SystemColors.Window
        Me.ugResult.DisplayLayout.GroupByBox.Appearance = Appearance2
        Appearance3.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ugResult.DisplayLayout.GroupByBox.BandLabelAppearance = Appearance3
        Me.ugResult.DisplayLayout.GroupByBox.BorderStyle = Infragistics.Win.UIElementBorderStyle.Solid
        Appearance4.BackColor = System.Drawing.SystemColors.ControlLightLight
        Appearance4.BackColor2 = System.Drawing.SystemColors.Control
        Appearance4.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance4.ForeColor = System.Drawing.SystemColors.GrayText
        Me.ugResult.DisplayLayout.GroupByBox.PromptAppearance = Appearance4
        Appearance5.BackColor = System.Drawing.SystemColors.Window
        Appearance5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ugResult.DisplayLayout.Override.ActiveCellAppearance = Appearance5
        Appearance6.BackColor = System.Drawing.SystemColors.Highlight
        Appearance6.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.ugResult.DisplayLayout.Override.ActiveRowAppearance = Appearance6
        Me.ugResult.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Synchronized
        Me.ugResult.DisplayLayout.Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.[True]
        Me.ugResult.DisplayLayout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.[True]
        Me.ugResult.DisplayLayout.Override.BorderStyleCell = Infragistics.Win.UIElementBorderStyle.Dotted
        Me.ugResult.DisplayLayout.Override.BorderStyleRow = Infragistics.Win.UIElementBorderStyle.Dotted
        Appearance7.BackColor = System.Drawing.SystemColors.Window
        Me.ugResult.DisplayLayout.Override.CardAreaAppearance = Appearance7
        Appearance8.BorderColor = System.Drawing.Color.Silver
        Appearance8.TextTrimming = Infragistics.Win.TextTrimming.EllipsisCharacter
        Me.ugResult.DisplayLayout.Override.CellAppearance = Appearance8
        Me.ugResult.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.EditAndSelectText
        Me.ugResult.DisplayLayout.Override.CellPadding = 0
        Me.ugResult.DisplayLayout.Override.ColumnSizingArea = Infragistics.Win.UltraWinGrid.ColumnSizingArea.EntireColumn
        Me.ugResult.DisplayLayout.Override.FilterUIType = Infragistics.Win.UltraWinGrid.FilterUIType.FilterRow
        Appearance9.BackColor = System.Drawing.SystemColors.Control
        Appearance9.BackColor2 = System.Drawing.SystemColors.ControlDark
        Appearance9.BackGradientAlignment = Infragistics.Win.GradientAlignment.Element
        Appearance9.BackGradientStyle = Infragistics.Win.GradientStyle.Horizontal
        Appearance9.BorderColor = System.Drawing.SystemColors.Window
        Me.ugResult.DisplayLayout.Override.GroupByRowAppearance = Appearance9
        Me.ugResult.DisplayLayout.Override.GroupBySummaryDisplayStyle = Infragistics.Win.UltraWinGrid.GroupBySummaryDisplayStyle.SummaryCellsAlwaysBelowDescription
        Appearance10.TextHAlignAsString = "Left"
        Me.ugResult.DisplayLayout.Override.HeaderAppearance = Appearance10
        Me.ugResult.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugResult.DisplayLayout.Override.HeaderStyle = Infragistics.Win.HeaderStyle.WindowsXPCommand
        Appearance11.BackColor = System.Drawing.SystemColors.Window
        Appearance11.BorderColor = System.Drawing.Color.Silver
        Me.ugResult.DisplayLayout.Override.RowAppearance = Appearance11
        Me.ugResult.DisplayLayout.Override.RowFilterMode = Infragistics.Win.UltraWinGrid.RowFilterMode.AllRowsInBand
        Me.ugResult.DisplayLayout.Override.RowSelectorHeaderStyle = Infragistics.Win.UltraWinGrid.RowSelectorHeaderStyle.ColumnChooserButton
        Me.ugResult.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.[True]
        Me.ugResult.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
        Me.ugResult.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.RowSelectorsOnly
        Me.ugResult.DisplayLayout.Override.SelectTypeCol = Infragistics.Win.UltraWinGrid.SelectType.ExtendedAutoDrag
        Me.ugResult.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.ExtendedAutoDrag
        Me.ugResult.DisplayLayout.Override.SummaryDisplayArea = Infragistics.Win.UltraWinGrid.SummaryDisplayAreas.Top
        Appearance12.BackColor = System.Drawing.SystemColors.ControlLight
        Me.ugResult.DisplayLayout.Override.TemplateAddRowAppearance = Appearance12
        Me.ugResult.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugResult.DisplayLayout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate
        Me.ugResult.DisplayLayout.ViewStyleBand = Infragistics.Win.UltraWinGrid.ViewStyleBand.OutlookGroupBy
        Me.ugResult.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugResult.Font = New System.Drawing.Font("Calibri", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugResult.Location = New System.Drawing.Point(0, 30)
        Me.ugResult.Name = "ugResult"
        Me.ugResult.Size = New System.Drawing.Size(1013, 458)
        Me.ugResult.TabIndex = 21
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tspbTKPC})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 488)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1013, 22)
        Me.StatusStrip1.TabIndex = 22
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tspbTKPC
        '
        Me.tspbTKPC.Name = "tspbTKPC"
        Me.tspbTKPC.Size = New System.Drawing.Size(200, 16)
        Me.tspbTKPC.Visible = False
        '
        'frmPerson
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.ClientSize = New System.Drawing.Size(1013, 510)
        Me.Controls.Add(Me.ugResult)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.bnList)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmPerson"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "교우 리스트"
        CType(Me.bnList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.bnList.ResumeLayout(False)
        Me.bnList.PerformLayout()
        CType(Me.ugResult, System.ComponentModel.ISupportInitialize).EndInit()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents bnList As BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As ToolStripSeparator
    Friend WithEvents tsbtExit As ToolStripButton
    Friend WithEvents ugResult As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugexcel As Infragistics.Win.UltraWinGrid.ExcelExport.UltraGridExcelExporter
    Friend WithEvents tsbtCheckAll As ToolStripButton
    Friend WithEvents tsbtUncheckAll As ToolStripButton
    Friend WithEvents tsbtExcel As ToolStripDropDownButton
    Friend WithEvents 모든데이터ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 선택한데이터만ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 선택하지않은데이터만ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents tspbTKPC As ToolStripProgressBar
    Friend WithEvents tscbDataOption As ToolStripComboBox
    Friend WithEvents tsbtFamily As ToolStripDropDownButton
    Friend WithEvents 교적카드로ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 설문지양식으로ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 교인주소록으로ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 가족모드로ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents 개인모드로ToolStripMenuItem As ToolStripMenuItem
End Class
