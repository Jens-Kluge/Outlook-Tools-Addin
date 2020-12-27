<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmSearch
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSearch))
        Me.tvFolders = New System.Windows.Forms.TreeView()
        Me.lstResults = New System.Windows.Forms.ListView()
        Me.btnLoadFolders = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'tvFolders
        '
        Me.tvFolders.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.tvFolders.Location = New System.Drawing.Point(28, 80)
        Me.tvFolders.Name = "tvFolders"
        Me.tvFolders.Size = New System.Drawing.Size(232, 427)
        Me.tvFolders.TabIndex = 0
        '
        'lstResults
        '
        Me.lstResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstResults.FullRowSelect = True
        Me.lstResults.GridLines = True
        Me.lstResults.HideSelection = False
        Me.lstResults.Location = New System.Drawing.Point(310, 80)
        Me.lstResults.Name = "lstResults"
        Me.lstResults.Size = New System.Drawing.Size(685, 427)
        Me.lstResults.TabIndex = 1
        Me.lstResults.UseCompatibleStateImageBehavior = False
        Me.lstResults.View = System.Windows.Forms.View.Details
        '
        'btnLoadFolders
        '
        Me.btnLoadFolders.Location = New System.Drawing.Point(28, 22)
        Me.btnLoadFolders.Name = "btnLoadFolders"
        Me.btnLoadFolders.Size = New System.Drawing.Size(115, 35)
        Me.btnLoadFolders.TabIndex = 2
        Me.btnLoadFolders.Text = "Load Folders"
        Me.btnLoadFolders.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(307, 31)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Search for:"
        '
        'txtSearch
        '
        Me.txtSearch.Location = New System.Drawing.Point(392, 31)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(227, 22)
        Me.txtSearch.TabIndex = 4
        '
        'frmSearch
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1067, 535)
        Me.Controls.Add(Me.txtSearch)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnLoadFolders)
        Me.Controls.Add(Me.lstResults)
        Me.Controls.Add(Me.tvFolders)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmSearch"
        Me.Text = "frmSearch"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tvFolders As Windows.Forms.TreeView
    Friend WithEvents lstResults As Windows.Forms.ListView
    Friend WithEvents btnLoadFolders As Windows.Forms.Button
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents txtSearch As Windows.Forms.TextBox
End Class
