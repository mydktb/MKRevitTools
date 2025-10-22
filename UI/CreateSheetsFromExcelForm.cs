using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MKRevitTools.UI
{
    public partial class CreateSheetsFromExcelForm : System.Windows.Forms.Form
    {
        // Public properties to pass data back to the command
        public string ExcelFilePath { get; private set; }
        public bool CreateSheets { get; private set; }
        public string SheetNumberColumn { get; private set; }
        public string SheetNameColumn { get; private set; }

        // Declare all controls for designer support
        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Label titleLabel;
        private System.Windows.Forms.Label instructionLabel;
        private System.Windows.Forms.Label filePathLabel;
        private System.Windows.Forms.TextBox filePathTextBox;
        private System.Windows.Forms.Button browseButton;
        private System.Windows.Forms.Label sheetNumberLabel;
        private System.Windows.Forms.TextBox sheetNumberTextBox;
        private System.Windows.Forms.Label sheetNameLabel;
        private System.Windows.Forms.TextBox sheetNameTextBox;
        private System.Windows.Forms.Button createButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label statusLabel;

        public CreateSheetsFromExcelForm()
        {
            InitializeComponent();
            ApplyThemeAtRuntime();
            InitializeDefaults();
        }

        private void InitializeComponent()
        {
            // Initialize all components
            this.mainPanel = new System.Windows.Forms.Panel();
            this.titleLabel = new System.Windows.Forms.Label();
            this.instructionLabel = new System.Windows.Forms.Label();
            this.filePathLabel = new System.Windows.Forms.Label();
            this.filePathTextBox = new System.Windows.Forms.TextBox();
            this.browseButton = new System.Windows.Forms.Button();
            this.sheetNumberLabel = new System.Windows.Forms.Label();
            this.sheetNumberTextBox = new System.Windows.Forms.TextBox();
            this.sheetNameLabel = new System.Windows.Forms.Label();
            this.sheetNameTextBox = new System.Windows.Forms.TextBox();
            this.createButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.statusLabel = new System.Windows.Forms.Label();

            this.mainPanel.SuspendLayout();
            this.SuspendLayout();

            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.statusLabel);
            this.mainPanel.Controls.Add(this.cancelButton);
            this.mainPanel.Controls.Add(this.createButton);
            this.mainPanel.Controls.Add(this.sheetNameTextBox);
            this.mainPanel.Controls.Add(this.sheetNameLabel);
            this.mainPanel.Controls.Add(this.sheetNumberTextBox);
            this.mainPanel.Controls.Add(this.sheetNumberLabel);
            this.mainPanel.Controls.Add(this.browseButton);
            this.mainPanel.Controls.Add(this.filePathTextBox);
            this.mainPanel.Controls.Add(this.filePathLabel);
            this.mainPanel.Controls.Add(this.instructionLabel);
            this.mainPanel.Controls.Add(this.titleLabel);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Padding = new System.Windows.Forms.Padding(20);
            this.mainPanel.Size = new System.Drawing.Size(500, 400);
            this.mainPanel.TabIndex = 0;

            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold);
            this.titleLabel.ForeColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.titleLabel.Location = new System.Drawing.Point(20, 20);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(202, 25);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "Create Sheets from Excel";

            // 
            // instructionLabel
            // 
            this.instructionLabel.AutoSize = true;
            this.instructionLabel.ForeColor = System.Drawing.Color.LightGray;
            this.instructionLabel.Location = new System.Drawing.Point(20, 55);
            this.instructionLabel.Name = "instructionLabel";
            this.instructionLabel.Size = new System.Drawing.Size(350, 15);
            this.instructionLabel.TabIndex = 1;
            this.instructionLabel.Text = "Select an Excel file and specify the column names for sheet data.";

            // 
            // filePathLabel
            // 
            this.filePathLabel.AutoSize = true;
            this.filePathLabel.ForeColor = System.Drawing.Color.White;
            this.filePathLabel.Location = new System.Drawing.Point(20, 90);
            this.filePathLabel.Name = "filePathLabel";
            this.filePathLabel.Size = new System.Drawing.Size(79, 15);
            this.filePathLabel.TabIndex = 2;
            this.filePathLabel.Text = "Excel File Path:";

            // 
            // filePathTextBox
            // 
            this.filePathTextBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.filePathTextBox.ForeColor = System.Drawing.Color.White;
            this.filePathTextBox.Location = new System.Drawing.Point(120, 88);
            this.filePathTextBox.Name = "filePathTextBox";
            this.filePathTextBox.ReadOnly = true;
            this.filePathTextBox.Size = new System.Drawing.Size(250, 23);
            this.filePathTextBox.TabIndex = 3;

            // 
            // browseButton
            // 
            this.browseButton.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.browseButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.browseButton.ForeColor = System.Drawing.Color.White;
            this.browseButton.Location = new System.Drawing.Point(380, 88);
            this.browseButton.Name = "browseButton";
            this.browseButton.Size = new System.Drawing.Size(80, 23);
            this.browseButton.TabIndex = 4;
            this.browseButton.Text = "Browse";
            this.browseButton.UseVisualStyleBackColor = false;
            this.browseButton.Click += new System.EventHandler(this.BrowseButton_Click);

            // 
            // sheetNumberLabel
            // 
            this.sheetNumberLabel.AutoSize = true;
            this.sheetNumberLabel.ForeColor = System.Drawing.Color.White;
            this.sheetNumberLabel.Location = new System.Drawing.Point(20, 130);
            this.sheetNumberLabel.Name = "sheetNumberLabel";
            this.sheetNumberLabel.Size = new System.Drawing.Size(94, 15);
            this.sheetNumberLabel.TabIndex = 5;
            this.sheetNumberLabel.Text = "Sheet Number Col:";

            // 
            // sheetNumberTextBox
            // 
            this.sheetNumberTextBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.sheetNumberTextBox.ForeColor = System.Drawing.Color.White;
            this.sheetNumberTextBox.Location = new System.Drawing.Point(120, 128);
            this.sheetNumberTextBox.Name = "sheetNumberTextBox";
            this.sheetNumberTextBox.Size = new System.Drawing.Size(150, 23);
            this.sheetNumberTextBox.TabIndex = 6;
            this.sheetNumberTextBox.Text = "Sheet Number";

            // 
            // sheetNameLabel
            // 
            this.sheetNameLabel.AutoSize = true;
            this.sheetNameLabel.ForeColor = System.Drawing.Color.White;
            this.sheetNameLabel.Location = new System.Drawing.Point(20, 165);
            this.sheetNameLabel.Name = "sheetNameLabel";
            this.sheetNameLabel.Size = new System.Drawing.Size(78, 15);
            this.sheetNameLabel.TabIndex = 7;
            this.sheetNameLabel.Text = "Sheet Name Col:";

            // 
            // sheetNameTextBox
            // 
            this.sheetNameTextBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.sheetNameTextBox.ForeColor = System.Drawing.Color.White;
            this.sheetNameTextBox.Location = new System.Drawing.Point(120, 163);
            this.sheetNameTextBox.Name = "sheetNameTextBox";
            this.sheetNameTextBox.Size = new System.Drawing.Size(150, 23);
            this.sheetNameTextBox.TabIndex = 8;
            this.sheetNameTextBox.Text = "Sheet Name";

            // 
            // createButton
            // 
            this.createButton.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.createButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.createButton.ForeColor = System.Drawing.Color.White;
            this.createButton.Location = new System.Drawing.Point(20, 220);
            this.createButton.Name = "createButton";
            this.createButton.Size = new System.Drawing.Size(150, 35);
            this.createButton.TabIndex = 9;
            this.createButton.Text = "Create Sheets";
            this.createButton.UseVisualStyleBackColor = false;
            this.createButton.Click += new System.EventHandler(this.CreateButton_Click);

            // 
            // cancelButton
            // 
            this.cancelButton.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.cancelButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancelButton.ForeColor = System.Drawing.Color.White;
            this.cancelButton.Location = new System.Drawing.Point(180, 220);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(150, 35);
            this.cancelButton.TabIndex = 10;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = false;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);

            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.ForeColor = System.Drawing.Color.LightGray;
            this.statusLabel.Location = new System.Drawing.Point(20, 270);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(175, 15);
            this.statusLabel.TabIndex = 11;
            this.statusLabel.Text = "Select an Excel file to get started.";

            // 
            // CreateSheetsFromExcelForm
            // 
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.ClientSize = new System.Drawing.Size(500, 400);
            this.Controls.Add(this.mainPanel);
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CreateSheetsFromExcelForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create Sheets from Excel";

            this.mainPanel.ResumeLayout(false);
            this.mainPanel.PerformLayout();
            this.ResumeLayout(false);
        }

        private void InitializeDefaults()
        {
            ExcelFilePath = string.Empty;
            CreateSheets = false;
            SheetNumberColumn = "Sheet Number";
            SheetNameColumn = "Sheet Name";
        }

        private void ApplyThemeAtRuntime()
        {
            // Apply theme at runtime
            this.BackColor = AppTheme.BackgroundColor;
            this.ForeColor = AppTheme.TextColor;
            this.Font = AppTheme.NormalFont;

            // Apply theme to specific controls
            titleLabel.ForeColor = AppTheme.AccentColor;
            titleLabel.Font = AppTheme.HeaderFont;

            instructionLabel.ForeColor = AppTheme.SecondaryTextColor;
            statusLabel.ForeColor = AppTheme.SecondaryTextColor;

            filePathTextBox.BackColor = AppTheme.PanelColor;
            filePathTextBox.ForeColor = AppTheme.TextColor;

            sheetNumberTextBox.BackColor = AppTheme.PanelColor;
            sheetNumberTextBox.ForeColor = AppTheme.TextColor;

            sheetNameTextBox.BackColor = AppTheme.PanelColor;
            sheetNameTextBox.ForeColor = AppTheme.TextColor;

            browseButton.BackColor = AppTheme.AccentColor;
            browseButton.ForeColor = AppTheme.TextColor;

            createButton.BackColor = AppTheme.AccentColor;
            createButton.ForeColor = AppTheme.TextColor;

            cancelButton.BackColor = AppTheme.PanelColor;
            cancelButton.ForeColor = AppTheme.TextColor;
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ExcelFilePath = openFileDialog.FileName;
                    filePathTextBox.Text = ExcelFilePath;
                    statusLabel.Text = $"Selected: {Path.GetFileName(ExcelFilePath)}";
                }
            }
        }

        private void CreateButton_Click(object sender, EventArgs e)
        {
            // Validate inputs
            if (string.IsNullOrEmpty(ExcelFilePath))
            {
                MessageBox.Show("Please select an Excel file.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!File.Exists(ExcelFilePath))
            {
                MessageBox.Show("The selected Excel file does not exist.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(sheetNumberTextBox.Text.Trim()))
            {
                MessageBox.Show("Please enter a sheet number column name.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrEmpty(sheetNameTextBox.Text.Trim()))
            {
                MessageBox.Show("Please enter a sheet name column name.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Set properties for the command to use
            SheetNumberColumn = sheetNumberTextBox.Text.Trim();
            SheetNameColumn = sheetNameTextBox.Text.Trim();
            CreateSheets = true;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            CreateSheets = false;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }
    }
}