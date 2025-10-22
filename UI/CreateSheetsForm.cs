using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace MKRevitTools.UI
{
    public partial class CreateSheetsForm : System.Windows.Forms.Form
    {
        // Public properties to pass data back to the command
        public bool CreateSheets { get; private set; }
        public int SheetCount { get; private set; }
        public int StartNumber { get; private set; }
        public string NumberPrefix { get; private set; }
        public string NumberSuffix { get; private set; }

        // Declare all controls for designer support
        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Label titleLabel;
        private System.Windows.Forms.Label sheetCountLabel;
        private System.Windows.Forms.NumericUpDown sheetCountNumeric;
        private System.Windows.Forms.Label startNumberLabel;
        private System.Windows.Forms.NumericUpDown startNumberNumeric;
        private System.Windows.Forms.Label prefixLabel;
        private System.Windows.Forms.TextBox prefixTextBox;
        private System.Windows.Forms.Label suffixLabel;
        private System.Windows.Forms.TextBox suffixTextBox;
        private System.Windows.Forms.Button createButton;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Label instructionsLabel;

        public CreateSheetsForm()
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
            this.instructionsLabel = new System.Windows.Forms.Label();
            this.sheetCountLabel = new System.Windows.Forms.Label();
            this.sheetCountNumeric = new System.Windows.Forms.NumericUpDown();
            this.startNumberLabel = new System.Windows.Forms.Label();
            this.startNumberNumeric = new System.Windows.Forms.NumericUpDown();
            this.prefixLabel = new System.Windows.Forms.Label();
            this.prefixTextBox = new System.Windows.Forms.TextBox();
            this.suffixLabel = new System.Windows.Forms.Label();
            this.suffixTextBox = new System.Windows.Forms.TextBox();
            this.createButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();

            this.mainPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sheetCountNumeric)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.startNumberNumeric)).BeginInit();
            this.SuspendLayout();

            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.cancelButton);
            this.mainPanel.Controls.Add(this.createButton);
            this.mainPanel.Controls.Add(this.suffixTextBox);
            this.mainPanel.Controls.Add(this.suffixLabel);
            this.mainPanel.Controls.Add(this.prefixTextBox);
            this.mainPanel.Controls.Add(this.prefixLabel);
            this.mainPanel.Controls.Add(this.startNumberNumeric);
            this.mainPanel.Controls.Add(this.startNumberLabel);
            this.mainPanel.Controls.Add(this.sheetCountNumeric);
            this.mainPanel.Controls.Add(this.sheetCountLabel);
            this.mainPanel.Controls.Add(this.instructionsLabel);
            this.mainPanel.Controls.Add(this.titleLabel);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Padding = new System.Windows.Forms.Padding(20);
            this.mainPanel.Size = new System.Drawing.Size(450, 400);
            this.mainPanel.TabIndex = 0;

            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold);
            this.titleLabel.ForeColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.titleLabel.Location = new System.Drawing.Point(20, 20);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(137, 25);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "Create Sheets";

            // 
            // instructionsLabel
            // 
            this.instructionsLabel.AutoSize = true;
            this.instructionsLabel.ForeColor = System.Drawing.Color.LightGray;
            this.instructionsLabel.Location = new System.Drawing.Point(20, 55);
            this.instructionsLabel.Name = "instructionsLabel";
            this.instructionsLabel.Size = new System.Drawing.Size(300, 15);
            this.instructionsLabel.TabIndex = 1;
            this.instructionsLabel.Text = "Configure the sheet creation settings below:";

            // 
            // sheetCountLabel
            // 
            this.sheetCountLabel.AutoSize = true;
            this.sheetCountLabel.ForeColor = System.Drawing.Color.White;
            this.sheetCountLabel.Location = new System.Drawing.Point(20, 90);
            this.sheetCountLabel.Name = "sheetCountLabel";
            this.sheetCountLabel.Size = new System.Drawing.Size(74, 15);
            this.sheetCountLabel.TabIndex = 2;
            this.sheetCountLabel.Text = "Sheet Count:";

            // 
            // sheetCountNumeric
            // 
            this.sheetCountNumeric.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.sheetCountNumeric.ForeColor = System.Drawing.Color.White;
            this.sheetCountNumeric.Location = new System.Drawing.Point(150, 88);
            this.sheetCountNumeric.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.sheetCountNumeric.Name = "sheetCountNumeric";
            this.sheetCountNumeric.Size = new System.Drawing.Size(80, 23);
            this.sheetCountNumeric.TabIndex = 3;
            this.sheetCountNumeric.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});

            // 
            // startNumberLabel
            // 
            this.startNumberLabel.AutoSize = true;
            this.startNumberLabel.ForeColor = System.Drawing.Color.White;
            this.startNumberLabel.Location = new System.Drawing.Point(20, 125);
            this.startNumberLabel.Name = "startNumberLabel";
            this.startNumberLabel.Size = new System.Drawing.Size(124, 15);
            this.startNumberLabel.TabIndex = 4;
            this.startNumberLabel.Text = "Starting Sheet Number:";

            // 
            // startNumberNumeric
            // 
            this.startNumberNumeric.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.startNumberNumeric.ForeColor = System.Drawing.Color.White;
            this.startNumberNumeric.Location = new System.Drawing.Point(150, 123);
            this.startNumberNumeric.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.startNumberNumeric.Name = "startNumberNumeric";
            this.startNumberNumeric.Size = new System.Drawing.Size(80, 23);
            this.startNumberNumeric.TabIndex = 5;
            this.startNumberNumeric.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});

            // 
            // prefixLabel
            // 
            this.prefixLabel.AutoSize = true;
            this.prefixLabel.ForeColor = System.Drawing.Color.White;
            this.prefixLabel.Location = new System.Drawing.Point(20, 160);
            this.prefixLabel.Name = "prefixLabel";
            this.prefixLabel.Size = new System.Drawing.Size(79, 15);
            this.prefixLabel.TabIndex = 6;
            this.prefixLabel.Text = "Number Prefix:";

            // 
            // prefixTextBox
            // 
            this.prefixTextBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.prefixTextBox.ForeColor = System.Drawing.Color.White;
            this.prefixTextBox.Location = new System.Drawing.Point(150, 158);
            this.prefixTextBox.Name = "prefixTextBox";
            this.prefixTextBox.Size = new System.Drawing.Size(80, 23);
            this.prefixTextBox.TabIndex = 7;

            // 
            // suffixLabel
            // 
            this.suffixLabel.AutoSize = true;
            this.suffixLabel.ForeColor = System.Drawing.Color.White;
            this.suffixLabel.Location = new System.Drawing.Point(20, 195);
            this.suffixLabel.Name = "suffixLabel";
            this.suffixLabel.Size = new System.Drawing.Size(79, 15);
            this.suffixLabel.TabIndex = 8;
            this.suffixLabel.Text = "Number Suffix:";

            // 
            // suffixTextBox
            // 
            this.suffixTextBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.suffixTextBox.ForeColor = System.Drawing.Color.White;
            this.suffixTextBox.Location = new System.Drawing.Point(150, 193);
            this.suffixTextBox.Name = "suffixTextBox";
            this.suffixTextBox.Size = new System.Drawing.Size(80, 23);
            this.suffixTextBox.TabIndex = 9;

            // 
            // createButton
            // 
            this.createButton.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.createButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.createButton.ForeColor = System.Drawing.Color.White;
            this.createButton.Location = new System.Drawing.Point(20, 250);
            this.createButton.Name = "createButton";
            this.createButton.Size = new System.Drawing.Size(150, 35);
            this.createButton.TabIndex = 10;
            this.createButton.Text = "Create Sheets";
            this.createButton.UseVisualStyleBackColor = false;
            this.createButton.Click += new System.EventHandler(this.CreateButton_Click);

            // 
            // cancelButton
            // 
            this.cancelButton.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.cancelButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.cancelButton.ForeColor = System.Drawing.Color.White;
            this.cancelButton.Location = new System.Drawing.Point(180, 250);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(150, 35);
            this.cancelButton.TabIndex = 11;
            this.cancelButton.Text = "Cancel";
            this.cancelButton.UseVisualStyleBackColor = false;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);

            // 
            // CreateSheetsForm
            // 
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.ClientSize = new System.Drawing.Size(450, 400);
            this.Controls.Add(this.mainPanel);
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CreateSheetsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create Multiple Sheets";

            this.mainPanel.ResumeLayout(false);
            this.mainPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sheetCountNumeric)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.startNumberNumeric)).EndInit();
            this.ResumeLayout(false);
        }

        private void InitializeDefaults()
        {
            // Set default values
            sheetCountNumeric.Value = 1;
            startNumberNumeric.Value = 1;
            prefixTextBox.Text = "";
            suffixTextBox.Text = "";
            CreateSheets = false;
        }

        private void ApplyThemeAtRuntime()
        {
            // Apply theme colors and fonts at runtime
            this.BackColor = AppTheme.BackgroundColor;
            this.ForeColor = AppTheme.TextColor;
            this.Font = AppTheme.NormalFont;

            // Apply theme to controls
            titleLabel.ForeColor = AppTheme.AccentColor;
            titleLabel.Font = AppTheme.HeaderFont;

            // Style buttons
            createButton.BackColor = AppTheme.AccentColor;
            createButton.ForeColor = AppTheme.TextColor;
            createButton.FlatStyle = FlatStyle.Flat;
            createButton.Font = AppTheme.NormalFont;

            cancelButton.BackColor = AppTheme.PanelColor;
            cancelButton.ForeColor = AppTheme.TextColor;
            cancelButton.FlatStyle = FlatStyle.Flat;
            cancelButton.Font = AppTheme.NormalFont;

            // Style textboxes and numeric controls
            prefixTextBox.BackColor = AppTheme.PanelColor;
            prefixTextBox.ForeColor = AppTheme.TextColor;
            prefixTextBox.BorderStyle = BorderStyle.FixedSingle;

            suffixTextBox.BackColor = AppTheme.PanelColor;
            suffixTextBox.ForeColor = AppTheme.TextColor;
            suffixTextBox.BorderStyle = BorderStyle.FixedSingle;

            sheetCountNumeric.BackColor = AppTheme.PanelColor;
            sheetCountNumeric.ForeColor = AppTheme.TextColor;

            startNumberNumeric.BackColor = AppTheme.PanelColor;
            startNumberNumeric.ForeColor = AppTheme.TextColor;
        }

        private void CreateButton_Click(object sender, EventArgs e)
        {
            // Validate inputs
            if (sheetCountNumeric.Value < 1)
            {
                MessageBox.Show("Please enter a valid sheet count (at least 1).", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Set properties for the command to use
            SheetCount = (int)sheetCountNumeric.Value;
            StartNumber = (int)startNumberNumeric.Value;
            NumberPrefix = prefixTextBox.Text.Trim();
            NumberSuffix = suffixTextBox.Text.Trim();
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