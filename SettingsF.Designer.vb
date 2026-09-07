<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class SettingsF
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SettingsF))
        APIKTabControl = New TabControl()
        OllamaTabPage = New TabPage()
        GoogleCheckBox = New CheckBox()
        AnthropicCheckBox = New CheckBox()
        OpenAICheckBox = New CheckBox()
        OllamaCheckBox = New CheckBox()
        MNameListBox = New ListBox()
        HostALabel = New Label()
        ModelNLabel = New Label()
        HAddressTextBox = New TextBox()
        ServerSButton = New Button()
        ServerCButton = New Button()
        BootTabPage = New TabPage()
        SaveBootButton = New Button()
        ClearBootButton = New Button()
        BootRichTextBox = New RichTextBox()
        OSTabPage = New TabPage()
        SaveOSButton = New Button()
        LoadOSButton = New Button()
        OSRichTextBox = New RichTextBox()
        ClearOSButton = New Button()
        PictureBox1 = New PictureBox()
        Label1 = New Label()
        APIKTabControl.SuspendLayout()
        OllamaTabPage.SuspendLayout()
        BootTabPage.SuspendLayout()
        OSTabPage.SuspendLayout()
        CType(PictureBox1, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' APIKTabControl
        ' 
        APIKTabControl.Controls.Add(OllamaTabPage)
        APIKTabControl.Controls.Add(BootTabPage)
        APIKTabControl.Controls.Add(OSTabPage)
        APIKTabControl.Location = New Point(12, 12)
        APIKTabControl.Name = "APIKTabControl"
        APIKTabControl.SelectedIndex = 0
        APIKTabControl.Size = New Size(750, 529)
        APIKTabControl.TabIndex = 0
        ' 
        ' OllamaTabPage
        ' 
        OllamaTabPage.BackColor = Color.Black
        OllamaTabPage.Controls.Add(GoogleCheckBox)
        OllamaTabPage.Controls.Add(AnthropicCheckBox)
        OllamaTabPage.Controls.Add(OpenAICheckBox)
        OllamaTabPage.Controls.Add(OllamaCheckBox)
        OllamaTabPage.Controls.Add(MNameListBox)
        OllamaTabPage.Controls.Add(HostALabel)
        OllamaTabPage.Controls.Add(ModelNLabel)
        OllamaTabPage.Controls.Add(HAddressTextBox)
        OllamaTabPage.Controls.Add(ServerSButton)
        OllamaTabPage.Controls.Add(ServerCButton)
        OllamaTabPage.ForeColor = Color.Gray
        OllamaTabPage.Location = New Point(4, 24)
        OllamaTabPage.Name = "OllamaTabPage"
        OllamaTabPage.Padding = New Padding(3)
        OllamaTabPage.Size = New Size(742, 501)
        OllamaTabPage.TabIndex = 2
        OllamaTabPage.Text = "Ollama"
        ' 
        ' GoogleCheckBox
        ' 
        GoogleCheckBox.AutoSize = True
        GoogleCheckBox.Location = New Point(397, 195)
        GoogleCheckBox.Name = "GoogleCheckBox"
        GoogleCheckBox.Size = New Size(64, 19)
        GoogleCheckBox.TabIndex = 3
        GoogleCheckBox.Text = "Google"
        GoogleCheckBox.UseVisualStyleBackColor = True
        ' 
        ' AnthropicCheckBox
        ' 
        AnthropicCheckBox.AutoSize = True
        AnthropicCheckBox.Location = New Point(312, 195)
        AnthropicCheckBox.Name = "AnthropicCheckBox"
        AnthropicCheckBox.Size = New Size(79, 19)
        AnthropicCheckBox.TabIndex = 2
        AnthropicCheckBox.Text = "Anthropic"
        AnthropicCheckBox.UseVisualStyleBackColor = True
        ' 
        ' OpenAICheckBox
        ' 
        OpenAICheckBox.AutoSize = True
        OpenAICheckBox.Location = New Point(240, 195)
        OpenAICheckBox.Name = "OpenAICheckBox"
        OpenAICheckBox.Size = New Size(66, 19)
        OpenAICheckBox.TabIndex = 1
        OpenAICheckBox.Text = "OpenAI"
        OpenAICheckBox.UseVisualStyleBackColor = True
        ' 
        ' OllamaCheckBox
        ' 
        OllamaCheckBox.AutoSize = True
        OllamaCheckBox.Location = New Point(170, 195)
        OllamaCheckBox.Name = "OllamaCheckBox"
        OllamaCheckBox.Size = New Size(64, 19)
        OllamaCheckBox.TabIndex = 0
        OllamaCheckBox.Text = "Ollama"
        OllamaCheckBox.UseVisualStyleBackColor = True
        ' 
        ' MNameListBox
        ' 
        MNameListBox.BackColor = Color.Black
        MNameListBox.BorderStyle = BorderStyle.FixedSingle
        MNameListBox.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        MNameListBox.ForeColor = Color.Gray
        MNameListBox.FormattingEnabled = True
        MNameListBox.ItemHeight = 15
        MNameListBox.Location = New Point(170, 249)
        MNameListBox.Name = "MNameListBox"
        MNameListBox.Size = New Size(350, 32)
        MNameListBox.TabIndex = 6
        ' 
        ' HostALabel
        ' 
        HostALabel.AutoSize = True
        HostALabel.ForeColor = Color.RoyalBlue
        HostALabel.Location = New Point(535, 224)
        HostALabel.Name = "HostALabel"
        HostALabel.Size = New Size(77, 15)
        HostALabel.TabIndex = 5
        HostALabel.Text = "Host/API Key"
        ' 
        ' ModelNLabel
        ' 
        ModelNLabel.AutoSize = True
        ModelNLabel.ForeColor = Color.RoyalBlue
        ModelNLabel.Location = New Point(535, 258)
        ModelNLabel.Name = "ModelNLabel"
        ModelNLabel.Size = New Size(41, 15)
        ModelNLabel.TabIndex = 7
        ModelNLabel.Text = "Model"
        ' 
        ' HAddressTextBox
        ' 
        HAddressTextBox.BackColor = Color.Black
        HAddressTextBox.BorderStyle = BorderStyle.FixedSingle
        HAddressTextBox.Font = New Font("Segoe UI", 9F, FontStyle.Bold)
        HAddressTextBox.ForeColor = Color.Gray
        HAddressTextBox.Location = New Point(170, 220)
        HAddressTextBox.Name = "HAddressTextBox"
        HAddressTextBox.Size = New Size(350, 23)
        HAddressTextBox.TabIndex = 4
        ' 
        ' ServerSButton
        ' 
        ServerSButton.BackColor = Color.Black
        ServerSButton.FlatAppearance.BorderColor = Color.Black
        ServerSButton.FlatAppearance.BorderSize = 0
        ServerSButton.FlatAppearance.CheckedBackColor = Color.Transparent
        ServerSButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        ServerSButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        ServerSButton.FlatStyle = FlatStyle.Flat
        ServerSButton.Font = New Font("Webdings", 12F)
        ServerSButton.ForeColor = Color.RoyalBlue
        ServerSButton.Location = New Point(680, 470)
        ServerSButton.Name = "ServerSButton"
        ServerSButton.Size = New Size(25, 25)
        ServerSButton.TabIndex = 8
        ServerSButton.Text = "Í"
        ServerSButton.UseVisualStyleBackColor = False
        ' 
        ' ServerCButton
        ' 
        ServerCButton.BackColor = Color.Black
        ServerCButton.FlatAppearance.BorderColor = Color.Black
        ServerCButton.FlatAppearance.BorderSize = 0
        ServerCButton.FlatAppearance.CheckedBackColor = Color.Transparent
        ServerCButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        ServerCButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        ServerCButton.FlatStyle = FlatStyle.Flat
        ServerCButton.Font = New Font("Webdings", 12F)
        ServerCButton.ForeColor = Color.RoyalBlue
        ServerCButton.Location = New Point(711, 469)
        ServerCButton.Name = "ServerCButton"
        ServerCButton.Size = New Size(25, 25)
        ServerCButton.TabIndex = 9
        ServerCButton.Text = "r"
        ServerCButton.UseVisualStyleBackColor = False
        ' 
        ' BootTabPage
        ' 
        BootTabPage.BackColor = Color.Black
        BootTabPage.Controls.Add(SaveBootButton)
        BootTabPage.Controls.Add(ClearBootButton)
        BootTabPage.Controls.Add(BootRichTextBox)
        BootTabPage.Location = New Point(4, 24)
        BootTabPage.Name = "BootTabPage"
        BootTabPage.Padding = New Padding(3)
        BootTabPage.Size = New Size(742, 501)
        BootTabPage.TabIndex = 3
        BootTabPage.Text = "Boot"
        ' 
        ' SaveBootButton
        ' 
        SaveBootButton.BackColor = Color.Black
        SaveBootButton.FlatAppearance.BorderColor = Color.Black
        SaveBootButton.FlatAppearance.BorderSize = 0
        SaveBootButton.FlatAppearance.CheckedBackColor = Color.Transparent
        SaveBootButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        SaveBootButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        SaveBootButton.FlatStyle = FlatStyle.Flat
        SaveBootButton.Font = New Font("Webdings", 12F)
        SaveBootButton.ForeColor = Color.RoyalBlue
        SaveBootButton.Location = New Point(680, 469)
        SaveBootButton.Name = "SaveBootButton"
        SaveBootButton.Size = New Size(25, 25)
        SaveBootButton.TabIndex = 5
        SaveBootButton.Text = "Í"
        SaveBootButton.UseVisualStyleBackColor = False
        ' 
        ' ClearBootButton
        ' 
        ClearBootButton.BackColor = Color.Black
        ClearBootButton.FlatAppearance.BorderColor = Color.Black
        ClearBootButton.FlatAppearance.BorderSize = 0
        ClearBootButton.FlatAppearance.CheckedBackColor = Color.Transparent
        ClearBootButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        ClearBootButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        ClearBootButton.FlatStyle = FlatStyle.Flat
        ClearBootButton.Font = New Font("Webdings", 12F)
        ClearBootButton.ForeColor = Color.RoyalBlue
        ClearBootButton.Location = New Point(711, 468)
        ClearBootButton.Name = "ClearBootButton"
        ClearBootButton.Size = New Size(25, 25)
        ClearBootButton.TabIndex = 6
        ClearBootButton.Text = "r"
        ClearBootButton.UseVisualStyleBackColor = False
        ' 
        ' BootRichTextBox
        ' 
        BootRichTextBox.BackColor = Color.Black
        BootRichTextBox.BorderStyle = BorderStyle.None
        BootRichTextBox.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        BootRichTextBox.ForeColor = Color.Gray
        BootRichTextBox.Location = New Point(6, 6)
        BootRichTextBox.Name = "BootRichTextBox"
        BootRichTextBox.Size = New Size(730, 458)
        BootRichTextBox.TabIndex = 1
        BootRichTextBox.Text = ""
        ' 
        ' OSTabPage
        ' 
        OSTabPage.BackColor = Color.Black
        OSTabPage.Controls.Add(SaveOSButton)
        OSTabPage.Controls.Add(LoadOSButton)
        OSTabPage.Controls.Add(OSRichTextBox)
        OSTabPage.Controls.Add(ClearOSButton)
        OSTabPage.ForeColor = Color.Gray
        OSTabPage.Location = New Point(4, 24)
        OSTabPage.Name = "OSTabPage"
        OSTabPage.Padding = New Padding(3)
        OSTabPage.Size = New Size(742, 501)
        OSTabPage.TabIndex = 1
        OSTabPage.Text = "OS"
        ' 
        ' SaveOSButton
        ' 
        SaveOSButton.BackColor = Color.Black
        SaveOSButton.FlatAppearance.BorderColor = Color.Black
        SaveOSButton.FlatAppearance.BorderSize = 0
        SaveOSButton.FlatAppearance.CheckedBackColor = Color.Transparent
        SaveOSButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        SaveOSButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        SaveOSButton.FlatStyle = FlatStyle.Flat
        SaveOSButton.Font = New Font("Webdings", 12F)
        SaveOSButton.ForeColor = Color.RoyalBlue
        SaveOSButton.Location = New Point(680, 469)
        SaveOSButton.Name = "SaveOSButton"
        SaveOSButton.Size = New Size(25, 25)
        SaveOSButton.TabIndex = 6
        SaveOSButton.Text = "Í"
        SaveOSButton.UseVisualStyleBackColor = False
        ' 
        ' LoadOSButton
        ' 
        LoadOSButton.BackColor = Color.Black
        LoadOSButton.FlatAppearance.BorderColor = Color.Black
        LoadOSButton.FlatAppearance.BorderSize = 0
        LoadOSButton.FlatAppearance.CheckedBackColor = Color.Transparent
        LoadOSButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        LoadOSButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        LoadOSButton.FlatStyle = FlatStyle.Flat
        LoadOSButton.Font = New Font("Webdings", 12F)
        LoadOSButton.ForeColor = Color.RoyalBlue
        LoadOSButton.Location = New Point(649, 469)
        LoadOSButton.Name = "LoadOSButton"
        LoadOSButton.Size = New Size(25, 25)
        LoadOSButton.TabIndex = 1
        LoadOSButton.Text = "Ì"
        LoadOSButton.UseVisualStyleBackColor = False
        ' 
        ' OSRichTextBox
        ' 
        OSRichTextBox.BackColor = Color.Black
        OSRichTextBox.BorderStyle = BorderStyle.None
        OSRichTextBox.Font = New Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        OSRichTextBox.ForeColor = Color.Gray
        OSRichTextBox.Location = New Point(6, 6)
        OSRichTextBox.Name = "OSRichTextBox"
        OSRichTextBox.Size = New Size(730, 458)
        OSRichTextBox.TabIndex = 0
        OSRichTextBox.Text = ""
        ' 
        ' ClearOSButton
        ' 
        ClearOSButton.BackColor = Color.Black
        ClearOSButton.FlatAppearance.BorderColor = Color.Black
        ClearOSButton.FlatAppearance.BorderSize = 0
        ClearOSButton.FlatAppearance.CheckedBackColor = Color.Transparent
        ClearOSButton.FlatAppearance.MouseDownBackColor = Color.Transparent
        ClearOSButton.FlatAppearance.MouseOverBackColor = Color.Transparent
        ClearOSButton.FlatStyle = FlatStyle.Flat
        ClearOSButton.Font = New Font("Webdings", 12F)
        ClearOSButton.ForeColor = Color.RoyalBlue
        ClearOSButton.Location = New Point(711, 468)
        ClearOSButton.Name = "ClearOSButton"
        ClearOSButton.Size = New Size(25, 25)
        ClearOSButton.TabIndex = 3
        ClearOSButton.Text = "r"
        ClearOSButton.UseVisualStyleBackColor = False
        ' 
        ' PictureBox1
        ' 
        PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), Image)
        PictureBox1.Location = New Point(607, 546)
        PictureBox1.Name = "PictureBox1"
        PictureBox1.Size = New Size(166, 94)
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom
        PictureBox1.TabIndex = 15
        PictureBox1.TabStop = False
        ' 
        ' Label1
        ' 
        Label1.AutoSize = True
        Label1.ForeColor = Color.RoyalBlue
        Label1.Location = New Point(12, 610)
        Label1.Name = "Label1"
        Label1.Size = New Size(266, 30)
        Label1.TabIndex = 1
        Label1.Text = "© Copyright 2026 Oranyx Labs/Elliot Monteverde" & vbCrLf & "'GNU General Public License v3.0"
        ' 
        ' SettingsF
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.FromArgb(CByte(64), CByte(64), CByte(64))
        ClientSize = New Size(774, 656)
        Controls.Add(Label1)
        Controls.Add(PictureBox1)
        Controls.Add(APIKTabControl)
        ForeColor = Color.Black
        FormBorderStyle = FormBorderStyle.FixedSingle
        Icon = CType(resources.GetObject("$this.Icon"), Icon)
        MaximizeBox = False
        Name = "SettingsF"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Settings"
        APIKTabControl.ResumeLayout(False)
        OllamaTabPage.ResumeLayout(False)
        OllamaTabPage.PerformLayout()
        BootTabPage.ResumeLayout(False)
        OSTabPage.ResumeLayout(False)
        CType(PictureBox1, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents APIKTabControl As TabControl
    Friend WithEvents OSTabPage As TabPage
    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents SaveBootButton As Button
    Friend WithEvents ClearOSButton As Button
    Friend WithEvents OllamaTabPage As TabPage
    Friend WithEvents ServerSButton As Button
    Friend WithEvents ServerCButton As Button
    Friend WithEvents OSRichTextBox As RichTextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents LoadOSButton As Button
    Friend WithEvents HAddressTextBox As TextBox
    Friend WithEvents HostALabel As Label
    Friend WithEvents ModelNLabel As Label
    Friend WithEvents BootTabPage As TabPage
    Friend WithEvents BootRichTextBox As RichTextBox
    Friend WithEvents SaveOSButton As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents ClearBootButton As Button
    Friend WithEvents MNameListBox As ListBox
    Friend WithEvents AnthropicCheckBox As CheckBox
    Friend WithEvents OpenAICheckBox As CheckBox
    Friend WithEvents OllamaCheckBox As CheckBox
    Friend WithEvents GoogleCheckBox As CheckBox
End Class
