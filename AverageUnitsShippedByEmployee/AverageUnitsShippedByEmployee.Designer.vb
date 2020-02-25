<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAverageUnitsShippedByEmployee
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
        Me.lblAverageOutputEmployee1 = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.btnEnter = New System.Windows.Forms.Button()
        Me.txtEnteredUnitsEmployee1 = New System.Windows.Forms.TextBox()
        Me.lblDay = New System.Windows.Forms.Label()
        Me.lblUnitsDaily = New System.Windows.Forms.Label()
        Me.txtUnitsInput = New System.Windows.Forms.TextBox()
        Me.ttpAverageUnitsShipped = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblAverageOutputEmployee2 = New System.Windows.Forms.Label()
        Me.txtEnteredUnitsEmployee2 = New System.Windows.Forms.TextBox()
        Me.lblAverageOutputEmployee3 = New System.Windows.Forms.Label()
        Me.txtEnteredUnitsEmployee3 = New System.Windows.Forms.TextBox()
        Me.lblEmployee1 = New System.Windows.Forms.Label()
        Me.lblEmployee2 = New System.Windows.Forms.Label()
        Me.lblEmployee3 = New System.Windows.Forms.Label()
        Me.lblOverallAverageOutput = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblAverageOutputEmployee1
        '
        Me.lblAverageOutputEmployee1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAverageOutputEmployee1.Location = New System.Drawing.Point(24, 275)
        Me.lblAverageOutputEmployee1.Name = "lblAverageOutputEmployee1"
        Me.lblAverageOutputEmployee1.Size = New System.Drawing.Size(140, 33)
        Me.lblAverageOutputEmployee1.TabIndex = 9
        Me.lblAverageOutputEmployee1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ttpAverageUnitsShipped.SetToolTip(Me.lblAverageOutputEmployee1, "Once 7 values are entered above, this field will output the average number of uni" &
        "ts shipped per day for Employee 1.")
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnExit.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnExit.Location = New System.Drawing.Point(338, 363)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(140, 23)
        Me.btnExit.TabIndex = 14
        Me.btnExit.Text = "E&xit"
        Me.ttpAverageUnitsShipped.SetToolTip(Me.btnExit, "Press this button to exit the application.")
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'btnReset
        '
        Me.btnReset.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnReset.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnReset.Location = New System.Drawing.Point(180, 363)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(140, 23)
        Me.btnReset.TabIndex = 13
        Me.btnReset.Text = "&Reset"
        Me.ttpAverageUnitsShipped.SetToolTip(Me.btnReset, "Press this button to clear all values and reset the application to it's initial s" &
        "tate.")
        Me.btnReset.UseVisualStyleBackColor = False
        '
        'btnEnter
        '
        Me.btnEnter.BackColor = System.Drawing.SystemColors.ControlLight
        Me.btnEnter.Location = New System.Drawing.Point(24, 363)
        Me.btnEnter.Name = "btnEnter"
        Me.btnEnter.Size = New System.Drawing.Size(140, 23)
        Me.btnEnter.TabIndex = 12
        Me.btnEnter.Text = "&Enter"
        Me.ttpAverageUnitsShipped.SetToolTip(Me.btnEnter, "Press this button to submit the number of units shipped for a single day.")
        Me.btnEnter.UseVisualStyleBackColor = False
        '
        'txtEnteredUnitsEmployee1
        '
        Me.txtEnteredUnitsEmployee1.AcceptsReturn = True
        Me.txtEnteredUnitsEmployee1.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtEnteredUnitsEmployee1.Location = New System.Drawing.Point(24, 96)
        Me.txtEnteredUnitsEmployee1.Multiline = True
        Me.txtEnteredUnitsEmployee1.Name = "txtEnteredUnitsEmployee1"
        Me.txtEnteredUnitsEmployee1.ReadOnly = True
        Me.txtEnteredUnitsEmployee1.Size = New System.Drawing.Size(140, 168)
        Me.txtEnteredUnitsEmployee1.TabIndex = 6
        Me.ttpAverageUnitsShipped.SetToolTip(Me.txtEnteredUnitsEmployee1, "This field displays all daily units shipped values entered so far for this week f" &
        "or Employee 1.")
        '
        'lblDay
        '
        Me.lblDay.AutoSize = True
        Me.lblDay.Location = New System.Drawing.Point(23, 16)
        Me.lblDay.Name = "lblDay"
        Me.lblDay.Size = New System.Drawing.Size(26, 13)
        Me.lblDay.TabIndex = 0
        Me.lblDay.Text = "&Day"
        '
        'lblUnitsDaily
        '
        Me.lblUnitsDaily.AutoSize = True
        Me.lblUnitsDaily.Location = New System.Drawing.Point(21, 45)
        Me.lblUnitsDaily.Name = "lblUnitsDaily"
        Me.lblUnitsDaily.Size = New System.Drawing.Size(37, 13)
        Me.lblUnitsDaily.TabIndex = 1
        Me.lblUnitsDaily.Text = "&Units: "
        '
        'txtUnitsInput
        '
        Me.txtUnitsInput.Location = New System.Drawing.Point(69, 42)
        Me.txtUnitsInput.Name = "txtUnitsInput"
        Me.txtUnitsInput.Size = New System.Drawing.Size(64, 20)
        Me.txtUnitsInput.TabIndex = 2
        Me.ttpAverageUnitsShipped.SetToolTip(Me.txtUnitsInput, "Enter the  daily number of units shipped in this field.")
        '
        'lblAverageOutputEmployee2
        '
        Me.lblAverageOutputEmployee2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAverageOutputEmployee2.Location = New System.Drawing.Point(180, 275)
        Me.lblAverageOutputEmployee2.Name = "lblAverageOutputEmployee2"
        Me.lblAverageOutputEmployee2.Size = New System.Drawing.Size(140, 33)
        Me.lblAverageOutputEmployee2.TabIndex = 10
        Me.lblAverageOutputEmployee2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ttpAverageUnitsShipped.SetToolTip(Me.lblAverageOutputEmployee2, "Once 7 values are entered above, this field will output the average number of uni" &
        "ts shipped per day for Employee 2.")
        '
        'txtEnteredUnitsEmployee2
        '
        Me.txtEnteredUnitsEmployee2.AcceptsReturn = True
        Me.txtEnteredUnitsEmployee2.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtEnteredUnitsEmployee2.Location = New System.Drawing.Point(180, 96)
        Me.txtEnteredUnitsEmployee2.Multiline = True
        Me.txtEnteredUnitsEmployee2.Name = "txtEnteredUnitsEmployee2"
        Me.txtEnteredUnitsEmployee2.ReadOnly = True
        Me.txtEnteredUnitsEmployee2.Size = New System.Drawing.Size(140, 168)
        Me.txtEnteredUnitsEmployee2.TabIndex = 7
        Me.ttpAverageUnitsShipped.SetToolTip(Me.txtEnteredUnitsEmployee2, "This field displays all daily units shipped values entered so far for this week f" &
        "or Employee 2.")
        '
        'lblAverageOutputEmployee3
        '
        Me.lblAverageOutputEmployee3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAverageOutputEmployee3.Location = New System.Drawing.Point(338, 275)
        Me.lblAverageOutputEmployee3.Name = "lblAverageOutputEmployee3"
        Me.lblAverageOutputEmployee3.Size = New System.Drawing.Size(140, 33)
        Me.lblAverageOutputEmployee3.TabIndex = 11
        Me.lblAverageOutputEmployee3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ttpAverageUnitsShipped.SetToolTip(Me.lblAverageOutputEmployee3, "Once 7 values are entered above, this field will output the average number of uni" &
        "ts shipped per day for Employee 3.")
        '
        'txtEnteredUnitsEmployee3
        '
        Me.txtEnteredUnitsEmployee3.AcceptsReturn = True
        Me.txtEnteredUnitsEmployee3.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.txtEnteredUnitsEmployee3.Location = New System.Drawing.Point(338, 96)
        Me.txtEnteredUnitsEmployee3.Multiline = True
        Me.txtEnteredUnitsEmployee3.Name = "txtEnteredUnitsEmployee3"
        Me.txtEnteredUnitsEmployee3.ReadOnly = True
        Me.txtEnteredUnitsEmployee3.Size = New System.Drawing.Size(140, 168)
        Me.txtEnteredUnitsEmployee3.TabIndex = 8
        Me.ttpAverageUnitsShipped.SetToolTip(Me.txtEnteredUnitsEmployee3, "This field displays all daily units shipped values entered so far for this week f" &
        "or Employee 3.")
        '
        'lblEmployee1
        '
        Me.lblEmployee1.AutoSize = True
        Me.lblEmployee1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEmployee1.Location = New System.Drawing.Point(61, 76)
        Me.lblEmployee1.Name = "lblEmployee1"
        Me.lblEmployee1.Size = New System.Drawing.Size(72, 13)
        Me.lblEmployee1.TabIndex = 3
        Me.lblEmployee1.Text = "Employee &1"
        '
        'lblEmployee2
        '
        Me.lblEmployee2.AutoSize = True
        Me.lblEmployee2.Location = New System.Drawing.Point(217, 76)
        Me.lblEmployee2.Name = "lblEmployee2"
        Me.lblEmployee2.Size = New System.Drawing.Size(62, 13)
        Me.lblEmployee2.TabIndex = 4
        Me.lblEmployee2.Text = "Employee &2"
        '
        'lblEmployee3
        '
        Me.lblEmployee3.AutoSize = True
        Me.lblEmployee3.Location = New System.Drawing.Point(375, 76)
        Me.lblEmployee3.Name = "lblEmployee3"
        Me.lblEmployee3.Size = New System.Drawing.Size(62, 13)
        Me.lblEmployee3.TabIndex = 5
        Me.lblEmployee3.Text = "Employee &3"
        '
        'lblOverallAverageOutput
        '
        Me.lblOverallAverageOutput.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOverallAverageOutput.Location = New System.Drawing.Point(24, 321)
        Me.lblOverallAverageOutput.Name = "lblOverallAverageOutput"
        Me.lblOverallAverageOutput.Size = New System.Drawing.Size(454, 33)
        Me.lblOverallAverageOutput.TabIndex = 15
        Me.lblOverallAverageOutput.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.ttpAverageUnitsShipped.SetToolTip(Me.lblOverallAverageOutput, "Once 7 values are entered above, this field will output the average number of uni" &
        "ts shipped per day for Employee 1.")
        '
        'frmAverageUnitsShippedByEmployee
        '
        Me.AcceptButton = Me.btnEnter
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnReset
        Me.ClientSize = New System.Drawing.Size(501, 393)
        Me.Controls.Add(Me.lblOverallAverageOutput)
        Me.Controls.Add(Me.lblEmployee3)
        Me.Controls.Add(Me.lblEmployee2)
        Me.Controls.Add(Me.lblEmployee1)
        Me.Controls.Add(Me.lblAverageOutputEmployee3)
        Me.Controls.Add(Me.txtEnteredUnitsEmployee3)
        Me.Controls.Add(Me.lblAverageOutputEmployee2)
        Me.Controls.Add(Me.txtEnteredUnitsEmployee2)
        Me.Controls.Add(Me.lblAverageOutputEmployee1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnReset)
        Me.Controls.Add(Me.btnEnter)
        Me.Controls.Add(Me.txtEnteredUnitsEmployee1)
        Me.Controls.Add(Me.lblDay)
        Me.Controls.Add(Me.lblUnitsDaily)
        Me.Controls.Add(Me.txtUnitsInput)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmAverageUnitsShippedByEmployee"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Average Units Shipped By Employee"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblAverageOutputEmployee1 As Label
    Friend WithEvents ttpAverageUnitsShipped As ToolTip
    Friend WithEvents btnExit As Button
    Friend WithEvents btnReset As Button
    Friend WithEvents btnEnter As Button
    Friend WithEvents txtEnteredUnitsEmployee1 As TextBox
    Friend WithEvents lblDay As Label
    Friend WithEvents lblUnitsDaily As Label
    Friend WithEvents txtUnitsInput As TextBox
    Friend WithEvents lblAverageOutputEmployee2 As Label
    Friend WithEvents txtEnteredUnitsEmployee2 As TextBox
    Friend WithEvents lblAverageOutputEmployee3 As Label
    Friend WithEvents txtEnteredUnitsEmployee3 As TextBox
    Friend WithEvents lblEmployee1 As Label
    Friend WithEvents lblEmployee2 As Label
    Friend WithEvents lblEmployee3 As Label
    Friend WithEvents lblOverallAverageOutput As Label
End Class
