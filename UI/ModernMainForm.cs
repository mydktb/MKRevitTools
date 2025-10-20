using System;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;

namespace MKRevitTools.UI
{
    public partial class ModernMainForm : Form
    {
        private Panel mainPanel;
        private Label headerLabel;
        private Button createSheetsFromExcelBtn;
        private Button createMultipleSheetsBtn;
        private Button settingsBtn;
        private Label versionLabel;

        public ModernMainForm()
        {
            InitializeComponent();
            ApplyModernStyling();
        }

        private void InitializeComponent()
        {
            this.mainPanel = new System.Windows.Forms.Panel();
            this.versionLabel = new System.Windows.Forms.Label();
            this.settingsBtn = new System.Windows.Forms.Button();
            this.createMultipleSheetsBtn = new System.Windows.Forms.Button();
            this.createSheetsFromExcelBtn = new System.Windows.Forms.Button();
            this.headerLabel = new System.Windows.Forms.Label();
            this.mainPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.versionLabel);
            this.mainPanel.Controls.Add(this.settingsBtn);
            this.mainPanel.Controls.Add(this.createMultipleSheetsBtn);
            this.mainPanel.Controls.Add(this.createSheetsFromExcelBtn);
            this.mainPanel.Controls.Add(this.headerLabel);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Padding = new System.Windows.Forms.Padding(20);
            this.mainPanel.Size = new System.Drawing.Size(450, 350);
            this.mainPanel.TabIndex = 0;
            this.mainPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.mainPanel_Paint);
            // 
            // versionLabel
            // 
            this.versionLabel.AutoSize = true;
            this.versionLabel.ForeColor = System.Drawing.Color.LightGray;
            this.versionLabel.Location = new System.Drawing.Point(20, 220);
            this.versionLabel.Name = "versionLabel";
            this.versionLabel.Size = new System.Drawing.Size(102, 20);
            this.versionLabel.TabIndex = 4;
            this.versionLabel.Text = "Version 1.0.0";
            // 
            // settingsBtn
            // 
            this.settingsBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.settingsBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.settingsBtn.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.settingsBtn.ForeColor = System.Drawing.Color.White;
            this.settingsBtn.Location = new System.Drawing.Point(20, 160);
            this.settingsBtn.Name = "settingsBtn";
            this.settingsBtn.Size = new System.Drawing.Size(280, 35);
            this.settingsBtn.TabIndex = 3;
            this.settingsBtn.Text = "Settings";
            this.settingsBtn.UseVisualStyleBackColor = false;
            this.settingsBtn.Click += new System.EventHandler(this.SettingsBtn_Click);
            // 
            // createMultipleSheetsBtn
            // 
            this.createMultipleSheetsBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.createMultipleSheetsBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.createMultipleSheetsBtn.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.createMultipleSheetsBtn.ForeColor = System.Drawing.Color.White;
            this.createMultipleSheetsBtn.Location = new System.Drawing.Point(20, 115);
            this.createMultipleSheetsBtn.Name = "createMultipleSheetsBtn";
            this.createMultipleSheetsBtn.Size = new System.Drawing.Size(280, 35);
            this.createMultipleSheetsBtn.TabIndex = 2;
            this.createMultipleSheetsBtn.Text = "Create Multiple Sheets";
            this.createMultipleSheetsBtn.UseVisualStyleBackColor = false;
            this.createMultipleSheetsBtn.Click += new System.EventHandler(this.CreateMultipleSheetsBtn_Click);
            // 
            // createSheetsFromExcelBtn
            // 
            this.createSheetsFromExcelBtn.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(62)))), ((int)(((byte)(62)))), ((int)(((byte)(66)))));
            this.createSheetsFromExcelBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.createSheetsFromExcelBtn.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.createSheetsFromExcelBtn.ForeColor = System.Drawing.Color.White;
            this.createSheetsFromExcelBtn.Location = new System.Drawing.Point(20, 70);
            this.createSheetsFromExcelBtn.Name = "createSheetsFromExcelBtn";
            this.createSheetsFromExcelBtn.Size = new System.Drawing.Size(280, 35);
            this.createSheetsFromExcelBtn.TabIndex = 1;
            this.createSheetsFromExcelBtn.Text = "Create Sheets from Excel";
            this.createSheetsFromExcelBtn.UseVisualStyleBackColor = false;
            this.createSheetsFromExcelBtn.Click += new System.EventHandler(this.CreateSheetsFromExcelBtn_Click);
            // 
            // headerLabel
            // 
            this.headerLabel.AutoSize = true;
            this.headerLabel.Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold);
            this.headerLabel.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(122)))), ((int)(((byte)(204)))));
            this.headerLabel.Location = new System.Drawing.Point(20, 20);
            this.headerLabel.Name = "headerLabel";
            this.headerLabel.Size = new System.Drawing.Size(244, 45);
            this.headerLabel.TabIndex = 0;
            this.headerLabel.Text = "MK Revit Tools";
            // 
            // ModernMainForm
            // 
            this.ClientSize = new System.Drawing.Size(450, 350);
            this.Controls.Add(this.mainPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ModernMainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MK Revit Tools - Dashboard";
            this.mainPanel.ResumeLayout(false);
            this.mainPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        private void ApplyModernStyling()
        {
            this.BackColor = Color.FromArgb(45, 45, 48);
            this.ForeColor = Color.White;
        }

        private void CreateSheetsFromExcelBtn_Click(object sender, EventArgs e)
        {
            // Launch the Excel-based sheet creation
            using (var form = new CreateSheetsForm())
            {
                form.ShowDialog();
            }
        }

        private void CreateMultipleSheetsBtn_Click(object sender, EventArgs e)
        {
            // Launch the basic multiple sheets creation
            // This will use your existing CreateSheetsCommand
            try
            {
                // You might need to create a form for this or use the existing command directly
                // For now, we'll show a message and you can implement the form tomorrow
                MessageBox.Show("Create Multiple Sheets feature will be implemented.\n\nThis will open the basic sheet creation interface.",
                    "Coming Soon",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                // TODO: Add your CreateMultipleSheets form here when ready
                // using (var form = new CreateMultipleSheetsForm())
                // {
                //     form.ShowDialog();
                // }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SettingsBtn_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Settings coming soon!", "Info",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void mainPanel_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}