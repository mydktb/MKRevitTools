using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;

namespace MKRevitTools.UI
{
    public partial class AddFiltersToViewTemplateForm : System.Windows.Forms.Form
    {
        private Document _document;
        private List<ParameterFilterElement> _allFilters;
        private List<Autodesk.Revit.DB.View> _viewTemplates;

        // Declare all controls at class level for designer
        private System.Windows.Forms.Panel mainPanel;
        private System.Windows.Forms.Label titleLabel;
        private System.Windows.Forms.Label filtersLabel;
        private System.Windows.Forms.ListBox availableFiltersListBox;
        private System.Windows.Forms.TextBox filterSearchTextBox;
        private System.Windows.Forms.Label templatesLabel;
        private System.Windows.Forms.ListBox viewTemplatesListBox;
        private System.Windows.Forms.TextBox templateSearchTextBox;
        private System.Windows.Forms.Button addButton;
        private System.Windows.Forms.Button executeButton;
        private System.Windows.Forms.Button closeButton;
        private System.Windows.Forms.Label statusLabel;

        public AddFiltersToViewTemplateForm(Document document)
        {
            _document = document;
            InitializeComponent();
            LoadData();
        }

        private void InitializeComponent()
        {
            // Initialize all components first (designer requirement)
            this.mainPanel = new System.Windows.Forms.Panel();
            this.titleLabel = new System.Windows.Forms.Label();
            this.filtersLabel = new System.Windows.Forms.Label();
            this.filterSearchTextBox = new System.Windows.Forms.TextBox();
            this.availableFiltersListBox = new System.Windows.Forms.ListBox();
            this.templatesLabel = new System.Windows.Forms.Label();
            this.templateSearchTextBox = new System.Windows.Forms.TextBox();
            this.viewTemplatesListBox = new System.Windows.Forms.ListBox();
            this.addButton = new System.Windows.Forms.Button();
            this.executeButton = new System.Windows.Forms.Button();
            this.closeButton = new System.Windows.Forms.Button();
            this.statusLabel = new System.Windows.Forms.Label();

            this.mainPanel.SuspendLayout();
            this.SuspendLayout();

            // 
            // mainPanel
            // 
            this.mainPanel.Controls.Add(this.statusLabel);
            this.mainPanel.Controls.Add(this.closeButton);
            this.mainPanel.Controls.Add(this.executeButton);
            this.mainPanel.Controls.Add(this.addButton);
            this.mainPanel.Controls.Add(this.viewTemplatesListBox);
            this.mainPanel.Controls.Add(this.templateSearchTextBox);
            this.mainPanel.Controls.Add(this.templatesLabel);
            this.mainPanel.Controls.Add(this.availableFiltersListBox);
            this.mainPanel.Controls.Add(this.filterSearchTextBox);
            this.mainPanel.Controls.Add(this.filtersLabel);
            this.mainPanel.Controls.Add(this.titleLabel);
            this.mainPanel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.mainPanel.Location = new System.Drawing.Point(0, 0);
            this.mainPanel.Name = "mainPanel";
            this.mainPanel.Padding = new System.Windows.Forms.Padding(20);
            this.mainPanel.Size = new System.Drawing.Size(700, 550);
            this.mainPanel.TabIndex = 0;

            // 
            // titleLabel
            // 
            this.titleLabel.AutoSize = true;
            this.titleLabel.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold);
            this.titleLabel.ForeColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.titleLabel.Location = new System.Drawing.Point(20, 20);
            this.titleLabel.Name = "titleLabel";
            this.titleLabel.Size = new System.Drawing.Size(243, 25);
            this.titleLabel.TabIndex = 0;
            this.titleLabel.Text = "Add Filters to View Templates";

            // 
            // filtersLabel
            // 
            this.filtersLabel.AutoSize = true;
            this.filtersLabel.ForeColor = System.Drawing.Color.LightGray;
            this.filtersLabel.Location = new System.Drawing.Point(20, 60);
            this.filtersLabel.Name = "filtersLabel";
            this.filtersLabel.Size = new System.Drawing.Size(95, 15);
            this.filtersLabel.TabIndex = 1;
            this.filtersLabel.Text = "Available Filters:";

            // 
            // filterSearchTextBox
            // 
            this.filterSearchTextBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.filterSearchTextBox.ForeColor = System.Drawing.Color.White;
            this.filterSearchTextBox.Location = new System.Drawing.Point(20, 80);
            this.filterSearchTextBox.Name = "filterSearchTextBox";
            this.filterSearchTextBox.Size = new System.Drawing.Size(300, 23);
            this.filterSearchTextBox.TabIndex = 2;
            this.filterSearchTextBox.Text = "Search filters...";
            this.filterSearchTextBox.ForeColor = System.Drawing.Color.Gray;
            this.filterSearchTextBox.TextChanged += new System.EventHandler(this.FilterSearchTextBox_TextChanged);
            this.filterSearchTextBox.Enter += new System.EventHandler(this.FilterSearchTextBox_Enter);
            this.filterSearchTextBox.Leave += new System.EventHandler(this.FilterSearchTextBox_Leave);

            // 
            // availableFiltersListBox
            // 
            this.availableFiltersListBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.availableFiltersListBox.ForeColor = System.Drawing.Color.White;
            this.availableFiltersListBox.FormattingEnabled = true;
            this.availableFiltersListBox.ItemHeight = 15;
            this.availableFiltersListBox.Location = new System.Drawing.Point(20, 110);
            this.availableFiltersListBox.Name = "availableFiltersListBox";
            this.availableFiltersListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.availableFiltersListBox.Size = new System.Drawing.Size(300, 289);
            this.availableFiltersListBox.TabIndex = 3;

            // 
            // templatesLabel
            // 
            this.templatesLabel.AutoSize = true;
            this.templatesLabel.ForeColor = System.Drawing.Color.LightGray;
            this.templatesLabel.Location = new System.Drawing.Point(350, 60);
            this.templatesLabel.Name = "templatesLabel";
            this.templatesLabel.Size = new System.Drawing.Size(89, 15);
            this.templatesLabel.TabIndex = 4;
            this.templatesLabel.Text = "View Templates:";

            // 
            // templateSearchTextBox
            // 
            this.templateSearchTextBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.templateSearchTextBox.ForeColor = System.Drawing.Color.White;
            this.templateSearchTextBox.Location = new System.Drawing.Point(350, 80);
            this.templateSearchTextBox.Name = "templateSearchTextBox";
            this.templateSearchTextBox.Size = new System.Drawing.Size(300, 23);
            this.templateSearchTextBox.TabIndex = 5;
            this.templateSearchTextBox.Text = "Search templates...";
            this.templateSearchTextBox.ForeColor = System.Drawing.Color.Gray;
            this.templateSearchTextBox.TextChanged += new System.EventHandler(this.TemplateSearchTextBox_TextChanged);
            this.templateSearchTextBox.Enter += new System.EventHandler(this.TemplateSearchTextBox_Enter);
            this.templateSearchTextBox.Leave += new System.EventHandler(this.TemplateSearchTextBox_Leave);

            // 
            // viewTemplatesListBox
            // 
            this.viewTemplatesListBox.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.viewTemplatesListBox.ForeColor = System.Drawing.Color.White;
            this.viewTemplatesListBox.FormattingEnabled = true;
            this.viewTemplatesListBox.ItemHeight = 15;
            this.viewTemplatesListBox.Location = new System.Drawing.Point(350, 110);
            this.viewTemplatesListBox.Name = "viewTemplatesListBox";
            this.viewTemplatesListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.viewTemplatesListBox.Size = new System.Drawing.Size(300, 289);
            this.viewTemplatesListBox.TabIndex = 6;

            // 
            // addButton
            // 
            this.addButton.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.addButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.addButton.ForeColor = System.Drawing.Color.White;
            this.addButton.Location = new System.Drawing.Point(20, 420);
            this.addButton.Name = "addButton";
            this.addButton.Size = new System.Drawing.Size(300, 35);
            this.addButton.TabIndex = 7;
            this.addButton.Text = "Add Selected Filters to Selected Templates";
            this.addButton.UseVisualStyleBackColor = false;
            this.addButton.Click += new System.EventHandler(this.AddButton_Click);

            // 
            // executeButton
            // 
            this.executeButton.BackColor = System.Drawing.Color.FromArgb(0, 122, 204);
            this.executeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.executeButton.ForeColor = System.Drawing.Color.White;
            this.executeButton.Location = new System.Drawing.Point(20, 465);
            this.executeButton.Name = "executeButton";
            this.executeButton.Size = new System.Drawing.Size(300, 35);
            this.executeButton.TabIndex = 8;
            this.executeButton.Text = "Execute Addition";
            this.executeButton.UseVisualStyleBackColor = false;
            this.executeButton.Click += new System.EventHandler(this.ExecuteButton_Click);

            // 
            // closeButton
            // 
            this.closeButton.BackColor = System.Drawing.Color.FromArgb(62, 62, 66);
            this.closeButton.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.closeButton.ForeColor = System.Drawing.Color.White;
            this.closeButton.Location = new System.Drawing.Point(350, 465);
            this.closeButton.Name = "closeButton";
            this.closeButton.Size = new System.Drawing.Size(300, 35);
            this.closeButton.TabIndex = 9;
            this.closeButton.Text = "Close";
            this.closeButton.UseVisualStyleBackColor = false;
            this.closeButton.Click += new System.EventHandler(this.CloseButton_Click);

            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.ForeColor = System.Drawing.Color.LightGray;
            this.statusLabel.Location = new System.Drawing.Point(20, 510);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(42, 15);
            this.statusLabel.TabIndex = 10;
            this.statusLabel.Text = "Ready";

            // 
            // AddFiltersToViewTemplateForm
            // 
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.ClientSize = new System.Drawing.Size(700, 550);
            this.Controls.Add(this.mainPanel);
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AddFiltersToViewTemplateForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Add Filters to View Templates";

            this.mainPanel.ResumeLayout(false);
            this.mainPanel.PerformLayout();
            this.ResumeLayout(false);
        }

        // Placeholder text event handlers
        private void FilterSearchTextBox_Enter(object sender, EventArgs e)
        {
            if (filterSearchTextBox.Text == "Search filters...")
            {
                filterSearchTextBox.Text = "";
                filterSearchTextBox.ForeColor = System.Drawing.Color.White;
            }
        }

        private void FilterSearchTextBox_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(filterSearchTextBox.Text))
            {
                filterSearchTextBox.Text = "Search filters...";
                filterSearchTextBox.ForeColor = System.Drawing.Color.Gray;
            }
        }

        private void TemplateSearchTextBox_Enter(object sender, EventArgs e)
        {
            if (templateSearchTextBox.Text == "Search templates...")
            {
                templateSearchTextBox.Text = "";
                templateSearchTextBox.ForeColor = System.Drawing.Color.White;
            }
        }

        private void TemplateSearchTextBox_Leave(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(templateSearchTextBox.Text))
            {
                templateSearchTextBox.Text = "Search templates...";
                templateSearchTextBox.ForeColor = System.Drawing.Color.Gray;
            }
        }

        // Rest of your methods remain the same...
        private void LoadData()
        {
            try
            {
                FilteredElementCollector filterCollector = new FilteredElementCollector(_document);
                _allFilters = filterCollector.OfClass(typeof(ParameterFilterElement))
                    .Cast<ParameterFilterElement>()
                    .OrderBy(f => f.Name)
                    .ToList();

                FilteredElementCollector viewCollector = new FilteredElementCollector(_document);
                _viewTemplates = viewCollector.OfClass(typeof(Autodesk.Revit.DB.View))
                    .Cast<Autodesk.Revit.DB.View>()
                    .Where(v => v.IsTemplate)
                    .OrderBy(v => v.Name)
                    .ToList();

                availableFiltersListBox.Items.Clear();
                viewTemplatesListBox.Items.Clear();

                foreach (var filter in _allFilters)
                {
                    availableFiltersListBox.Items.Add(filter.Name);
                }

                foreach (var template in _viewTemplates)
                {
                    viewTemplatesListBox.Items.Add(template.Name);
                }

                statusLabel.Text = $"Loaded {_allFilters.Count} filters and {_viewTemplates.Count} view templates";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading data: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void FilterSearchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (filterSearchTextBox.Text != "Search filters..." && !string.IsNullOrWhiteSpace(filterSearchTextBox.Text))
            {
                FilterListBox(availableFiltersListBox, _allFilters, filterSearchTextBox.Text);
            }
        }

        private void TemplateSearchTextBox_TextChanged(object sender, EventArgs e)
        {
            if (templateSearchTextBox.Text != "Search templates..." && !string.IsNullOrWhiteSpace(templateSearchTextBox.Text))
            {
                FilterListBox(viewTemplatesListBox, _viewTemplates, templateSearchTextBox.Text);
            }
        }

        private void FilterListBox<T>(ListBox listBox, List<T> sourceList, string searchText) where T : Element
        {
            listBox.Items.Clear();
            var filteredItems = sourceList.Where(item =>
                item.Name.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0);

            foreach (var item in filteredItems)
            {
                listBox.Items.Add(item.Name);
            }
        }

        private void AddButton_Click(object sender, EventArgs e)
        {
            int selectedFiltersCount = availableFiltersListBox.SelectedItems.Count;
            int selectedTemplatesCount = viewTemplatesListBox.SelectedItems.Count;

            if (selectedFiltersCount == 0 || selectedTemplatesCount == 0)
            {
                MessageBox.Show("Please select at least one filter and one view template.", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            statusLabel.Text = $"Ready to add {selectedFiltersCount} filter(s) to {selectedTemplatesCount} template(s)";
        }

        private void ExecuteButton_Click(object sender, EventArgs e)
        {
            int selectedFiltersCount = availableFiltersListBox.SelectedItems.Count;
            int selectedTemplatesCount = viewTemplatesListBox.SelectedItems.Count;

            if (selectedFiltersCount == 0 || selectedTemplatesCount == 0)
            {
                MessageBox.Show("Please select at least one filter and one view template.", "Info",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                using (Transaction trans = new Transaction(_document, "Add Filters to View Templates"))
                {
                    trans.Start();

                    int totalAdditions = 0;
                    var selectedFilters = _allFilters.Where(f =>
                        availableFiltersListBox.SelectedItems.Contains(f.Name)).ToList();
                    var selectedTemplates = _viewTemplates.Where(t =>
                        viewTemplatesListBox.SelectedItems.Contains(t.Name)).ToList();

                    foreach (Autodesk.Revit.DB.View template in selectedTemplates)
                    {
                        foreach (ParameterFilterElement filter in selectedFilters)
                        {
                            try
                            {
                                if (!template.GetFilters().Contains(filter.Id))
                                {
                                    template.AddFilter(filter.Id);
                                    totalAdditions++;
                                }
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error adding filter {filter.Name} to template {template.Name}: {ex.Message}");
                            }
                        }
                    }

                    trans.Commit();

                    MessageBox.Show(
                        $"Successfully completed!\n" +
                        $"Added {totalAdditions} filter assignments across {selectedTemplatesCount} template(s).\n" +
                        $"Some filters may have been skipped if they were already present.",
                        "Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    statusLabel.Text = $"Completed: {totalAdditions} filter assignments made";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error adding filters to templates: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}