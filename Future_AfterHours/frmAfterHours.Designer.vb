<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAfterHours
    Inherits System.Windows.Forms.Form

    'Form 覆寫 Dispose 以清除元件清單。
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

    '為 Windows Form 設計工具的必要項
    Private components As System.ComponentModel.IContainer

    '注意: 以下為 Windows Form 設計工具所需的程序
    '可以使用 Windows Form 設計工具進行修改。
    '請不要使用程式碼編輯器進行修改。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle4 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle5 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle6 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim ChartArea2 As System.Windows.Forms.DataVisualization.Charting.ChartArea = New System.Windows.Forms.DataVisualization.Charting.ChartArea()
        Me.btnShow = New System.Windows.Forms.Button()
        Me.lblTimes = New System.Windows.Forms.Label()
        Me.dgvDisplay = New System.Windows.Forms.DataGridView()
        Me.chtDisplay = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.btnSavePic = New System.Windows.Forms.Button()
        Me.cboType = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.gupChartItems = New System.Windows.Forms.GroupBox()
        Me.lblNote = New System.Windows.Forms.Label()
        CType(Me.dgvDisplay, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.chtDisplay, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnShow
        '
        Me.btnShow.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnShow.Location = New System.Drawing.Point(330, 16)
        Me.btnShow.Name = "btnShow"
        Me.btnShow.Size = New System.Drawing.Size(205, 30)
        Me.btnShow.TabIndex = 4
        Me.btnShow.Text = "取得所選報表"
        Me.btnShow.UseVisualStyleBackColor = True
        '
        'lblTimes
        '
        Me.lblTimes.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblTimes.Location = New System.Drawing.Point(12, 71)
        Me.lblTimes.Name = "lblTimes"
        Me.lblTimes.Size = New System.Drawing.Size(269, 20)
        Me.lblTimes.TabIndex = 5
        '
        'dgvDisplay
        '
        Me.dgvDisplay.AllowUserToAddRows = False
        Me.dgvDisplay.AllowUserToDeleteRows = False
        Me.dgvDisplay.AllowUserToResizeColumns = False
        Me.dgvDisplay.AllowUserToResizeRows = False
        Me.dgvDisplay.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.dgvDisplay.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvDisplay.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.[Single]
        DataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle4.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle4.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        DataGridViewCellStyle4.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDisplay.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle4
        Me.dgvDisplay.ColumnHeadersHeight = 24
        Me.dgvDisplay.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvDisplay.ColumnHeadersVisible = False
        Me.dgvDisplay.Location = New System.Drawing.Point(24, 150)
        Me.dgvDisplay.Margin = New System.Windows.Forms.Padding(0)
        Me.dgvDisplay.Name = "dgvDisplay"
        Me.dgvDisplay.ReadOnly = True
        DataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle5.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle5.Font = New System.Drawing.Font("新細明體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        DataGridViewCellStyle5.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle5.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dgvDisplay.RowHeadersDefaultCellStyle = DataGridViewCellStyle5
        Me.dgvDisplay.RowHeadersVisible = False
        Me.dgvDisplay.RowHeadersWidth = 100
        Me.dgvDisplay.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        DataGridViewCellStyle6.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle6.SelectionBackColor = System.Drawing.Color.White
        DataGridViewCellStyle6.SelectionForeColor = System.Drawing.Color.White
        Me.dgvDisplay.RowsDefaultCellStyle = DataGridViewCellStyle6
        Me.dgvDisplay.RowTemplate.DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        Me.dgvDisplay.RowTemplate.DefaultCellStyle.SelectionBackColor = System.Drawing.Color.White
        Me.dgvDisplay.RowTemplate.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.Black
        Me.dgvDisplay.RowTemplate.Height = 24
        Me.dgvDisplay.RowTemplate.ReadOnly = True
        Me.dgvDisplay.RowTemplate.Resizable = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvDisplay.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.dgvDisplay.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect
        Me.dgvDisplay.Size = New System.Drawing.Size(122, 78)
        Me.dgvDisplay.TabIndex = 8
        Me.dgvDisplay.Visible = False
        '
        'chtDisplay
        '
        ChartArea2.Name = "ChartArea1"
        Me.chtDisplay.ChartAreas.Add(ChartArea2)
        Me.chtDisplay.Location = New System.Drawing.Point(189, 125)
        Me.chtDisplay.Name = "chtDisplay"
        Me.chtDisplay.Size = New System.Drawing.Size(134, 82)
        Me.chtDisplay.TabIndex = 9
        Me.chtDisplay.Text = "Chart1"
        Me.chtDisplay.Visible = False
        '
        'btnSavePic
        '
        Me.btnSavePic.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.btnSavePic.Location = New System.Drawing.Point(570, 16)
        Me.btnSavePic.Name = "btnSavePic"
        Me.btnSavePic.Size = New System.Drawing.Size(233, 30)
        Me.btnSavePic.TabIndex = 10
        Me.btnSavePic.Text = "另存為圖片"
        Me.btnSavePic.UseVisualStyleBackColor = True
        '
        'cboType
        '
        Me.cboType.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.cboType.FormattingEnabled = True
        Me.cboType.Items.AddRange(New Object() {"大台期貨", "小台期貨", "大小台期貨合計", "選擇權合計(口數)", "選擇權合計(契約金額)", "選擇權CALL合計", "選擇權PUT合計"})
        Me.cboType.Location = New System.Drawing.Point(106, 17)
        Me.cboType.Name = "cboType"
        Me.cboType.Size = New System.Drawing.Size(185, 28)
        Me.cboType.TabIndex = 11
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.Label1.Location = New System.Drawing.Point(10, 21)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(89, 20)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "報表種類︰"
        '
        'gupChartItems
        '
        Me.gupChartItems.Font = New System.Drawing.Font("微軟正黑體", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.gupChartItems.Location = New System.Drawing.Point(287, 52)
        Me.gupChartItems.Name = "gupChartItems"
        Me.gupChartItems.Size = New System.Drawing.Size(530, 55)
        Me.gupChartItems.TabIndex = 19
        Me.gupChartItems.TabStop = False
        Me.gupChartItems.Text = "圖表顯示項目"
        '
        'lblNote
        '
        Me.lblNote.AutoSize = True
        Me.lblNote.Font = New System.Drawing.Font("微軟正黑體", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.lblNote.Location = New System.Drawing.Point(10, 585)
        Me.lblNote.Name = "lblNote"
        Me.lblNote.Size = New System.Drawing.Size(89, 20)
        Me.lblNote.TabIndex = 20
        Me.lblNote.Text = "報表種類︰"
        Me.lblNote.Visible = False
        '
        'frmAfterHours
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(825, 658)
        Me.Controls.Add(Me.lblNote)
        Me.Controls.Add(Me.gupChartItems)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.cboType)
        Me.Controls.Add(Me.btnSavePic)
        Me.Controls.Add(Me.chtDisplay)
        Me.Controls.Add(Me.dgvDisplay)
        Me.Controls.Add(Me.lblTimes)
        Me.Controls.Add(Me.btnShow)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.Name = "frmAfterHours"
        Me.RightToLeftLayout = True
        Me.Text = "台灣期貨盤後資訊"
        CType(Me.dgvDisplay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.chtDisplay, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnShow As System.Windows.Forms.Button
    Friend WithEvents lblTimes As System.Windows.Forms.Label
    Friend WithEvents dgvDisplay As System.Windows.Forms.DataGridView
    Friend WithEvents chtDisplay As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents btnSavePic As System.Windows.Forms.Button
    Friend WithEvents cboType As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents gupChartItems As System.Windows.Forms.GroupBox
    Friend WithEvents lblNote As System.Windows.Forms.Label

End Class
