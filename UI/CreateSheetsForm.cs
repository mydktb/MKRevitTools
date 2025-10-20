using System;
using System.Windows.Forms;
using System.Drawing;

namespace MKRevitTools.UI
{
    public partial class CreateSheetsForm : Form
    {
        public int SheetCount { get; private set; }
        public int StartNumber { get; private set; }
        public string NumberPrefix { get; private set; }
        public string NumberSuffix { get; private set; }
        public bool CreateSheets { get; private set; }

        private NumericUpDown numSheetCount;
        private NumericUpDown numStartNumber;
        private TextBox txtPrefix;
        private TextBox txtSuffix;
        private Button btnCreate;
        private Button btnCancel;
        private Label label1;
        private Label label2;
        private Label label3;
        private Label label4;
        private Label lblPreview;
        private GroupBox groupBox1;
        //private Font MKFontArialBold;
        //private Font MKFontArial;
        //private int fontSize;

        public CreateSheetsForm()
        {
            InitializeComponent();
            AttachEventHandlers(); // Attach event handlers after initialization
            SheetCount = 10;
            StartNumber = 1;
            NumberPrefix = "A";
            NumberSuffix = "";
            CreateSheets = false;
            UpdatePreview();
        }

        private void InitializeComponent()
        {
            this.numSheetCount = new System.Windows.Forms.NumericUpDown();
            this.numStartNumber = new System.Windows.Forms.NumericUpDown();
            this.txtPrefix = new System.Windows.Forms.TextBox();
            this.txtSuffix = new System.Windows.Forms.TextBox();
            this.btnCreate = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.lblPreview = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            ((System.ComponentModel.ISupportInitialize)(this.numSheetCount)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numStartNumber)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // numSheetCount
            // 
            this.numSheetCount.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.numSheetCount.Location = new System.Drawing.Point(160, 38);
            this.numSheetCount.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.numSheetCount.Name = "numSheetCount";
            this.numSheetCount.Size = new System.Drawing.Size(200, 26);
            this.numSheetCount.TabIndex = 1;
            this.numSheetCount.Value = new decimal(new int[] {
            10,
            0,
            0,
            0});
            // 
            // numStartNumber
            // 
            this.numStartNumber.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.numStartNumber.Location = new System.Drawing.Point(160, 73);
            this.numStartNumber.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.numStartNumber.Name = "numStartNumber";
            this.numStartNumber.Size = new System.Drawing.Size(200, 26);
            this.numStartNumber.TabIndex = 3;
            this.numStartNumber.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // txtPrefix
            // 
            this.txtPrefix.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.txtPrefix.Location = new System.Drawing.Point(160, 108);
            this.txtPrefix.Name = "txtPrefix";
            this.txtPrefix.Size = new System.Drawing.Size(200, 26);
            this.txtPrefix.TabIndex = 5;
            this.txtPrefix.Text = "A";
            // 
            // txtSuffix
            // 
            this.txtSuffix.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.txtSuffix.Location = new System.Drawing.Point(160, 143);
            this.txtSuffix.Name = "txtSuffix";
            this.txtSuffix.Size = new System.Drawing.Size(200, 26);
            this.txtSuffix.TabIndex = 7;
            // 
            // btnCreate
            // 
            this.btnCreate.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.btnCreate.Location = new System.Drawing.Point(340, 290);
            this.btnCreate.Name = "btnCreate";
            this.btnCreate.Size = new System.Drawing.Size(90, 30);
            this.btnCreate.TabIndex = 3;
            this.btnCreate.Text = "Create Sheets";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.btnCancel.Location = new System.Drawing.Point(240, 290);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(90, 30);
            this.btnCancel.TabIndex = 2;
            this.btnCancel.Text = "Cancel";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.label1.Location = new System.Drawing.Point(30, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Number of Sheets:";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.label2.Location = new System.Drawing.Point(30, 75);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(120, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "Starting Number:";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.label3.Location = new System.Drawing.Point(30, 110);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(120, 20);
            this.label3.TabIndex = 4;
            this.label3.Text = "Number Prefix:";
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F);
            this.label4.Location = new System.Drawing.Point(30, 145);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(120, 20);
            this.label4.TabIndex = 6;
            this.label4.Text = "Number Suffix:";
            // 
            // lblPreview
            // 
            this.lblPreview.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Italic);
            this.lblPreview.ForeColor = System.Drawing.Color.DarkBlue;
            this.lblPreview.Location = new System.Drawing.Point(20, 240);
            this.lblPreview.Name = "lblPreview";
            this.lblPreview.Size = new System.Drawing.Size(410, 40);
            this.lblPreview.TabIndex = 1;
            this.lblPreview.Text = "Preview: A101, A102, A103...";
            this.lblPreview.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.numSheetCount);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.numStartNumber);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtPrefix);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.txtSuffix);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold);
            this.groupBox1.Location = new System.Drawing.Point(20, 20);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(410, 200);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Sheet Settings";
            // 
            // CreateSheetsForm
            // 
            this.AcceptButton = this.btnCreate;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(500, 350);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lblPreview);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnCreate);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CreateSheetsForm";
            this.Padding = new System.Windows.Forms.Padding(20);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Create Multiple Sheets - MKRevitTools";
            this.Load += new System.EventHandler(this.CreateSheetsForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.numSheetCount)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numStartNumber)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        // New method to attach event handlers at runtime only
        private void AttachEventHandlers()
        {
            this.numSheetCount.ValueChanged += NumSheetCount_ValueChanged;
            this.numStartNumber.ValueChanged += NumStartNumber_ValueChanged;
            this.txtPrefix.TextChanged += TxtPrefix_TextChanged;
            this.txtSuffix.TextChanged += TxtSuffix_TextChanged;
            this.btnCreate.Click += BtnCreate_Click;
            this.btnCancel.Click += BtnCancel_Click;
        }

        // Separate event handler methods
        private void NumSheetCount_ValueChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void NumStartNumber_ValueChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void TxtPrefix_TextChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void TxtSuffix_TextChanged(object sender, EventArgs e)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            string prefix = string.IsNullOrEmpty(txtPrefix.Text) ? "" : txtPrefix.Text;
            string suffix = string.IsNullOrEmpty(txtSuffix.Text) ? "" : txtSuffix.Text;
            int start = (int)numStartNumber.Value;
            int count = (int)numSheetCount.Value;

            string preview = $"Preview: {prefix}{start:00}{suffix}";

            if (count > 1)
            {
                preview += $", {prefix}{start + 1:00}{suffix}";
            }
            if (count > 2)
            {
                preview += $", {prefix}{start + 2:00}{suffix}";
            }
            if (count > 3)
            {
                preview += $"... {prefix}{start + count - 1:00}{suffix}";
            }

            lblPreview.Text = preview;
        }

        private void BtnCreate_Click(object sender, EventArgs e)
        {
            SheetCount = (int)numSheetCount.Value;
            StartNumber = (int)numStartNumber.Value;
            NumberPrefix = txtPrefix.Text;
            NumberSuffix = txtSuffix.Text;
            CreateSheets = true;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            CreateSheets = false;
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void CreateSheetsForm_Load(object sender, EventArgs e)
        {

        }
    }
}