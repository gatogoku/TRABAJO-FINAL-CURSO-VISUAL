<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmReservas
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
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Button14 = New System.Windows.Forms.Button()
        Me.Button13 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.DataGridView1.BackgroundColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(234, 171)
        Me.DataGridView1.Margin = New System.Windows.Forms.Padding(4)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(921, 414)
        Me.DataGridView1.TabIndex = 3
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.GroupBox3.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox3.Controls.Add(Me.Button14)
        Me.GroupBox3.Controls.Add(Me.Button13)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox3.Location = New System.Drawing.Point(1233, 253)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4)
        Me.GroupBox3.Size = New System.Drawing.Size(165, 239)
        Me.GroupBox3.TabIndex = 4
        Me.GroupBox3.TabStop = False
        '
        'Button14
        '
        Me.Button14.BackColor = System.Drawing.Color.Black
        Me.Button14.Font = New System.Drawing.Font("Brandish", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button14.ForeColor = System.Drawing.Color.Transparent
        Me.Button14.Location = New System.Drawing.Point(26, 126)
        Me.Button14.Margin = New System.Windows.Forms.Padding(4)
        Me.Button14.Name = "Button14"
        Me.Button14.Size = New System.Drawing.Size(115, 88)
        Me.Button14.TabIndex = 1
        Me.Button14.Text = "QUITAR RESERVA"
        Me.Button14.UseVisualStyleBackColor = False
        '
        'Button13
        '
        Me.Button13.BackColor = System.Drawing.Color.Black
        Me.Button13.BackgroundImage = Global.ProyectoFinCurso.My.Resources.Resources.volver
        Me.Button13.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Button13.ForeColor = System.Drawing.Color.Transparent
        Me.Button13.Location = New System.Drawing.Point(26, 23)
        Me.Button13.Margin = New System.Windows.Forms.Padding(4)
        Me.Button13.Name = "Button13"
        Me.Button13.Size = New System.Drawing.Size(115, 88)
        Me.Button13.TabIndex = 0
        Me.Button13.UseVisualStyleBackColor = False
        '
        'FrmReservas
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.ProyectoFinCurso.My.Resources.Resources._74959
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1463, 741)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.DataGridView1)
        Me.DoubleBuffered = True
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "FrmReservas"
        Me.Text = "FrmReservas"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents GroupBox3 As GroupBox
    Friend WithEvents Button14 As Button
    Friend WithEvents Button13 As Button
End Class
