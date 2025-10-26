using System;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using MKRevitTools.UI;

namespace MKRevitTools.Commands
{
    public class CreateSheetsFromExcelForm : Form
    {
        private TextBox filePathTextBox;
        private Button browseButton;
        private Button okButton;
        private Button cancelButton;
        private Label instructionLabel;

        public string FilePath => filePathTextBox.Text;

        public CreateSheetsFromExcelForm()
        {
            InitializeForm();
        }

        private void InitializeForm()
        {
            // Form setup using AppTheme methods
            Text = "Create Sheets from Excel";
            Size = new Size(450, 160);
            Padding = AppTheme.FormPadding;

            // Apply form style using existing AppTheme method
            AppTheme.ApplyFormStyle(this);

            // Create controls using AppTheme factory methods
            CreateControls();
        }

        private void CreateControls()
        {
            // Instruction label using AppTheme
            instructionLabel = AppTheme.CreateNormalLabel("Select Excel file containing sheet data:");
            instructionLabel.Location = new Point(15, 15);
            instructionLabel.Size = new Size(400, 20);

            // File path textbox using AppTheme
            filePathTextBox = AppTheme.CreateModernTextBox();
            filePathTextBox.Location = new Point(15, 40);
            filePathTextBox.Size = new Size(300, 20);
            filePathTextBox.ReadOnly = true;

            // Browse button using AppTheme
            browseButton = AppTheme.CreateModernButton("Browse", 75, 25);
            browseButton.Location = new Point(325, 38);

            // OK button using AppTheme
            okButton = AppTheme.CreateAccentButton("OK", 75, 25);
            okButton.Location = new Point(240, 80);

            // Cancel button using AppTheme
            cancelButton = AppTheme.CreateModernButton("Cancel", 75, 25);
            cancelButton.Location = new Point(325, 80);

            // Add controls to form
            Controls.Add(instructionLabel);
            Controls.Add(filePathTextBox);
            Controls.Add(browseButton);
            Controls.Add(okButton);
            Controls.Add(cancelButton);

            // Wire events
            browseButton.Click += BrowseButton_Click;
            okButton.Click += OkButton_Click;
            cancelButton.Click += CancelButton_Click;
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files (*.xlsx; *.xls)|*.xlsx;*.xls|All Files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePathTextBox.Text = openFileDialog.FileName;
                }
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(filePathTextBox.Text) || !File.Exists(filePathTextBox.Text))
            {
                MessageBox.Show("Please select a valid Excel file.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                DialogResult = DialogResult.None;
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}