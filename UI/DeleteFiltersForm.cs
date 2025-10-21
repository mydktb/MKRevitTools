using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;

namespace MKRevitTools.UI
{
    public partial class DeleteFiltersForm : System.Windows.Forms.Form
    {
        private Document _document;
        private List<ParameterFilterElement> _matchingFilters;

        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Label titleLabel;
        private System.Windows.Forms.Label instructionLabel;
        private System.Windows.Forms.TextBox searchTextBox;
        private System.Windows.Forms.Button searchButton;
        private System.Windows.Forms.ListBox filtersListBox;
        private System.Windows.Forms.Label resultsLabel;
        private System.Windows.Forms.Button deleteSelectedButton;
        private System.Windows.Forms.Button deleteAllButton;
        private System.Windows.Forms.Button closeButton;

        public DeleteFiltersForm(Document document)
        {
            _document = document;
            _matchingFilters = new List<ParameterFilterElement>();

            InitializeComponent();
            ApplyModernStyling();
        }

        private void InitializeComponent()
        {
            // Initialize components first
            this.mainPanel = new System.Windows.Forms.Panel();
            this.titleLabel = new System.Windows.Forms.Label();
            this.instructionLabel = new System.Windows.Forms.Label();
            this.searchTextBox = new System.Windows.Forms.TextBox();
            this.searchButton = new System.Windows.Forms.Button();
            this.filtersListBox = new System.Windows.Forms.ListBox();
            this.resultsLabel = new System.Windows.Forms.Label();
            this.deleteSelectedButton = new System.Windows.Forms.Button();
            this.deleteAllButton = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();

            // Main Panel
            this.mainPanel.SuspendLayout();
            this.SuspendLayout();

            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.closeButton);
            this.mainPanel.Controls.Add(this.deleteAllButton);
            this.mainPanel.Controls.Add(this.deleteSelectedButton);
            this.mainPanel.Controls.Add(this.resultsLabel);
            this.mainPanel.Controls.Add(this.filtersListBox);
            this.mainPanel.Controls.Add(this.searchButton);
            this.mainPanel.Controls.Add(this.searchTextBox);
            this.mainPanel.Controls.Add(this.instructionLabel);
            this.mainPanel.Controls.Add(this.titleLabel);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Padding = new System.Windows.Forms.Padding(20);
            this.mainPanel.Size = new System.Drawing.Size(500, 450);
            this.mainPanel.TabIndex = 0;

            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold);
            this.titleLabel.ForeColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.titleLabel.Location = new System.Drawing.Point(20, 20);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(124, 25);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "Delete Filters";

            // 
            // instructionLabel
            // 
            this.instructionLabel.AutoSize = true;
            this.instructionLabel.ForeColor = System.Drawing.Color.LightGray;
            this.instructionLabel.Location = new System.Drawing.Point(20, 55);
            this.instructionLabel.Name = "instructionLabel";
            this.instructionLabel.Size = new System.Drawing.Size(350, 15);
            this.instructionLabel.TabIndex = 1;
            this.instructionLabel.Text = "Enter text to search for in filter names (case-insensitive):";

            // 
            // searchTextBox
            // 
            this.searchTextBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.searchTextBox.ForeColor = System.Drawing.Color.White;
            this.searchTextBox.Location = new System.Drawing.Point(20, 80);
            this.searchTextBox.Name = "searchTextBox";
            this.searchTextBox.Size = new System.Drawing.Size(300, 23);
            this.searchTextBox.TabIndex = 2;
            this.searchTextBox.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.SearchTextBox_KeyPress);

            // 
            // searchButton
            // 
            this.searchButton.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.searchButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.searchButton.ForeColor = System.Drawing.Color.White;
            this.searchButton.Location = new System.Drawing.Point(330, 80);
            this.searchButton.Name = "searchButton";
            this.searchButton.Size = new System.Drawing.Size(100, 23);
            this.searchButton.TabIndex = 3;
            this.searchButton.Text = "Search";
            this.searchButton.UseVisualStyleBackColor = false;
            this.searchButton.Click += new System.EventHandler(this.SearchButton_Click);

            // 
            // filtersListBox
            // 
            this.filtersListBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.filtersListBox.ForeColor = System.Drawing.Color.White;
            this.filtersListBox.FormattingEnabled = true;
            this.filtersListBox.ItemHeight = 15;
            this.filtersListBox.Location = new System.Drawing.Point(20, 130);
            this.filtersListBox.Name = "filtersListBox";
            this.filtersListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.filtersListBox.Size = new System.Drawing.Size(410, 184);
            this.filtersListBox.TabIndex = 4;

            // 
            // resultsLabel
            // 
            this.resultsLabel.AutoSize = true;
            this.resultsLabel.ForeColor = System.Drawing.Color.LightGray;
            this.resultsLabel.Location = new System.Drawing.Point(20, 110);
            this.resultsLabel.Name = "resultsLabel";
            this.resultsLabel.Size = new System.Drawing.Size(118, 15);
            this.resultsLabel.TabIndex = 5;
            this.resultsLabel.Text = "Matching filters: 0";

            // 
            // deleteSelectedButton
            // 
            this.deleteSelectedButton.BackColor = System.Drawing.Color.FromArgb(192, 0, 0);
            this.deleteSelectedButton.Enabled = false;
            this.deleteSelectedButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.deleteSelectedButton.ForeColor = System.Drawing.Color.White;
            this.deleteSelectedButton.Location = new System.Drawing.Point(20, 330);
            this.deleteSelectedButton.Name = "deleteSelectedButton";
            this.deleteSelectedButton.Size = new System.Drawing.Size(130, 35);
            this.deleteSelectedButton.TabIndex = 6;
            this.deleteSelectedButton.Text = "Delete Selected";
            this.deleteSelectedButton.UseVisualStyleBackColor = false;
            this.deleteSelectedButton.Click += new System.EventHandler(this.DeleteSelectedButton_Click);

            // 
            // deleteAllButton
            // 
            this.deleteAllButton.BackColor = System.Drawing.Color.FromArgb(192, 0, 0);
            this.deleteAllButton.Enabled = false;
            this.deleteAllButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.deleteAllButton.ForeColor = System.Drawing.Color.White;
            this.deleteAllButton.Location = new System.Drawing.Point(160, 330);
            this.deleteAllButton.Name = "deleteAllButton";
            this.deleteAllButton.Size = new System.Drawing.Size(130, 35);
            this.deleteAllButton.TabIndex = 7;
            this.deleteAllButton.Text = "Delete All";
            this.deleteAllButton.UseVisualStyleBackColor = false;
            this.deleteAllButton.Click += new System.EventHandler(this.DeleteAllButton_Click);

            // 
            // closeButton
            // 
            this.closeButton.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.closeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeButton.ForeColor = System.Drawing.Color.White;
            this.closeButton.Location = new System.Drawing.Point(300, 330);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(130, 35);
            this.closeButton.TabIndex = 8;
            this.closeButton.Text = "Close";
            this.closeButton.UseVisualStyleBackColor = false;
            this.closeButton.Click += new System.EventHandler(this.CloseButton_Click);

            // 
            // DeleteFiltersForm
            // 
            this.ClientSize = new System.Drawing.Size(500, 450);
            this.Controls.Add(this.mainPanel);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DeleteFiltersForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Delete Multiple Filters";

            this.mainPanel.ResumeLayout(false);
            this.mainPanel.PerformLayout();
            this.ResumeLayout(false);
        }

        private void ApplyModernStyling()
        {
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.ForeColor = System.Drawing.Color.White;
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            SearchFilters();
        }

        private void SearchTextBox_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)System.Windows.Forms.Keys.Enter)
            {
                SearchFilters();
                e.Handled = true;
            }
        }

        private void SearchFilters()
        {
            string searchText = searchTextBox.Text.Trim();

            if (string.IsNullOrEmpty(searchText))
            {
                System.Windows.Forms.MessageBox.Show("Please enter text to search for.", "Info",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            try
            {
                // Get all parameter filters in the project
                FilteredElementCollector collector = new FilteredElementCollector(_document);
                ICollection<Element> allFilters = collector.OfClass(typeof(ParameterFilterElement)).ToElements();

                _matchingFilters.Clear();
                filtersListBox.Items.Clear();

                // Find filters that contain the search text (case-insensitive)
                foreach (Element element in allFilters)
                {
                    ParameterFilterElement filter = element as ParameterFilterElement;
                    if (filter != null && filter.Name.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        _matchingFilters.Add(filter);
                    }
                }

                // Update the UI
                foreach (ParameterFilterElement filter in _matchingFilters)
                {
                    filtersListBox.Items.Add(filter.Name);
                }

                resultsLabel.Text = $"Matching filters: {_matchingFilters.Count}";

                // Enable/disable delete buttons based on results
                deleteSelectedButton.Enabled = _matchingFilters.Count > 0;
                deleteAllButton.Enabled = _matchingFilters.Count > 0;

                if (_matchingFilters.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show($"No filters found containing '{searchText}'.", "No Results",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error searching filters: {ex.Message}", "Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void DeleteSelectedButton_Click(object sender, EventArgs e)
        {
            if (filtersListBox.SelectedItems.Count == 0)
            {
                System.Windows.Forms.MessageBox.Show("Please select one or more filters to delete.", "Info",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return;
            }

            try
            {
                using (Transaction trans = new Transaction(_document, "Delete Selected Filters"))
                {
                    trans.Start();

                    int deletedCount = 0;
                    foreach (string selectedName in filtersListBox.SelectedItems)
                    {
                        ParameterFilterElement filterToDelete = _matchingFilters
                            .FirstOrDefault(f => f.Name == selectedName);

                        if (filterToDelete != null)
                        {
                            _document.Delete(filterToDelete.Id);
                            deletedCount++;
                        }
                    }

                    trans.Commit();

                    System.Windows.Forms.MessageBox.Show($"Successfully deleted {deletedCount} filter(s).", "Success",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                    // Refresh the search
                    SearchFilters();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error deleting filters: {ex.Message}", "Error",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        private void DeleteAllButton_Click(object sender, EventArgs e)
        {
            if (_matchingFilters.Count == 0)
                return;

            var result = System.Windows.Forms.MessageBox.Show(
                $"Are you sure you want to delete ALL {_matchingFilters.Count} matching filters? This action cannot be undone.",
                "Confirm Delete All",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Warning);

            if (result == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    using (Transaction trans = new Transaction(_document, "Delete All Matching Filters"))
                    {
                        trans.Start();

                        foreach (ParameterFilterElement filter in _matchingFilters)
                        {
                            _document.Delete(filter.Id);
                        }

                        trans.Commit();

                        System.Windows.Forms.MessageBox.Show($"Successfully deleted all {_matchingFilters.Count} filter(s).", "Success",
                            System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);

                        // Refresh the search
                        SearchFilters();
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show($"Error deleting filters: {ex.Message}", "Error",
                        System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}