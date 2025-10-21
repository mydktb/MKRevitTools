using System.Drawing;
using System.Windows.Forms;

namespace MKRevitTools.UI
{
    public static class AppTheme
    {
        // Color Palette
        public static Color BackgroundColor = Color.FromArgb(45, 45, 48);
        public static Color PanelColor = Color.FromArgb(62, 62, 66);
        public static Color AccentColor = Color.FromArgb(0, 122, 204);
        public static Color DangerColor = Color.FromArgb(192, 0, 0);
        public static Color TextColor = Color.White;
        public static Color SecondaryTextColor = Color.LightGray;
        public static Color PlaceholderTextColor = Color.Gray;

        // Fonts
        public static Font TitleFont = new Font("Segoe UI", 14F, FontStyle.Bold);
        public static Font HeaderFont = new Font("Segoe UI", 16F, FontStyle.Bold);
        public static Font NormalFont = new Font("Segoe UI", 9F);
        public static Font SmallFont = new Font("Segoe UI", 8F);

        // Sizes and Spacing
        public static Padding FormPadding = new Padding(20);
        public static Size ButtonSize = new Size(250, 35);
        public static Size SmallButtonSize = new Size(100, 23);
        public static int StandardSpacing = 10;

        // Form Defaults
        public static FormBorderStyle DefaultFormBorderStyle = FormBorderStyle.FixedDialog;
        public static FormStartPosition DefaultStartPosition = FormStartPosition.CenterScreen;

        // Methods to apply styling
        public static void ApplyFormStyle(Form form)
        {
            form.BackColor = BackgroundColor;
            form.ForeColor = TextColor;
            form.Font = NormalFont;
            form.FormBorderStyle = DefaultFormBorderStyle;
            form.StartPosition = DefaultStartPosition;
            form.MaximizeBox = false;
            form.MinimizeBox = false;
        }

        public static Button CreateModernButton(string text, int width = 0, int height = 0)
        {
            var button = new Button
            {
                Text = text,
                BackColor = PanelColor,
                ForeColor = TextColor,
                FlatStyle = FlatStyle.Flat,
                Font = NormalFont,
                Cursor = Cursors.Hand
            };

            if (width > 0 && height > 0)
                button.Size = new Size(width, height);
            else
                button.Size = ButtonSize;

            return button;
        }

        public static Button CreateAccentButton(string text, int width = 0, int height = 0)
        {
            var button = CreateModernButton(text, width, height);
            button.BackColor = AccentColor;
            return button;
        }

        public static Button CreateDangerButton(string text, int width = 0, int height = 0)
        {
            var button = CreateModernButton(text, width, height);
            button.BackColor = DangerColor;
            return button;
        }

        public static TextBox CreateModernTextBox(string placeholderText = "")
        {
            var textBox = new TextBox
            {
                BackColor = PanelColor,
                ForeColor = TextColor,
                BorderStyle = BorderStyle.FixedSingle
            };

            if (!string.IsNullOrEmpty(placeholderText))
            {
                textBox.Text = placeholderText;
                textBox.ForeColor = PlaceholderTextColor;
            }

            return textBox;
        }

        public static ListBox CreateModernListBox()
        {
            return new ListBox
            {
                BackColor = PanelColor,
                ForeColor = TextColor,
                BorderStyle = BorderStyle.FixedSingle
            };
        }

        public static Label CreateTitleLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = TitleFont,
                ForeColor = AccentColor,
                AutoSize = true
            };
        }

        public static Label CreateHeaderLabel(string text)
        {
            return new Label
            {
                Text = text,
                Font = HeaderFont,
                ForeColor = AccentColor,
                AutoSize = true
            };
        }

        public static Label CreateNormalLabel(string text)
        {
            return new Label
            {
                Text = text,
                ForeColor = TextColor,
                AutoSize = true
            };
        }

        public static Label CreateSecondaryLabel(string text)
        {
            return new Label
            {
                Text = text,
                ForeColor = SecondaryTextColor,
                AutoSize = true
            };
        }

        public static Panel CreateMainPanel()
        {
            return new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = BackgroundColor,
                Padding = FormPadding
            };
        }

        // Helper method to setup placeholder text behavior
        public static void SetupPlaceholderText(TextBox textBox, string placeholderText)
        {
            if (textBox.Text == placeholderText)
            {
                textBox.ForeColor = PlaceholderTextColor;
            }

            textBox.Enter += (sender, e) =>
            {
                if (textBox.Text == placeholderText)
                {
                    textBox.Text = "";
                    textBox.ForeColor = TextColor;
                }
            };

            textBox.Leave += (sender, e) =>
            {
                if (string.IsNullOrWhiteSpace(textBox.Text))
                {
                    textBox.Text = placeholderText;
                    textBox.ForeColor = PlaceholderTextColor;
                }
            };
        }
    }
}