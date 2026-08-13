using System;
using System.Collections.Generic;
using System.Windows;
// --- FIX: Use full namespaces to avoid ambiguity ---
using System.Windows.Controls;
using System.Text.RegularExpressions; // --- NEW: For sanitizing name ---

namespace CADWorkFlowTool
{
    public partial class PointsEntryDialog : Window
    {
        public Dictionary<string, double> PointsMap { get; private set; }
        private List<(string PartName, System.Windows.Controls.TextBox TextBox)> _partInputs;

        // --- NEW: Helper to sanitize names for WPF ---
        private string SanitizeName(string name)
        {
            // Replace invalid characters (like spaces, hyphens) with underscores
            string sanitized = Regex.Replace(name, @"[^a-zA-Z0-9_]", "_");
            // Ensure it starts with a letter or underscore
            if (sanitized.Length == 0) return "_"; // Handle empty string case
            if (!char.IsLetter(sanitized[0]) && sanitized[0] != '_')
            {
                sanitized = "_" + sanitized;
            }
            return sanitized;
        }

        public PointsEntryDialog(List<string> partNames)
        {
            InitializeComponent();

            _partInputs = new List<(string PartName, System.Windows.Controls.TextBox TextBox)>();
            PointsMap = null;

            foreach (var partName in partNames)
            {
                // --- FIX: Sanitize the part name BEFORE using it for the TextBox Name ---
                // Ensure partName is not null or empty before sanitizing
                if (string.IsNullOrWhiteSpace(partName)) continue;
                string sanitizedName = SanitizeName(partName);

                var label = new System.Windows.Controls.Label
                {
                    Content = $"{partName}:", // Display the ORIGINAL name
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(5, 5, 5, 0)
                };

                var textBox = new System.Windows.Controls.TextBox
                {
                    // Use the SANITIZED name for the control's Name property
                    Name = $"txt_{sanitizedName}",
                    MinWidth = 50,
                    Padding = new Thickness(3),
                    Margin = new Thickness(5, 0, 5, 5),
                    Text = "0.0"
                };

                var grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                Grid.SetColumn(label, 0);
                Grid.SetColumn(textBox, 1);

                grid.Children.Add(label);
                grid.Children.Add(textBox);

                PartsStackPanel.Children.Add(grid);

                // Store the ORIGINAL part name with the textbox
                _partInputs.Add((partName, textBox));
            }
        }

        private void BtnSubmit_Click(object sender, RoutedEventArgs e)
        {
            PointsMap = new Dictionary<string, double>();

            foreach (var (partName, textBox) in _partInputs)
            {
                // Ensure partName is valid before proceeding
                if (string.IsNullOrWhiteSpace(partName)) continue;

                if (double.TryParse(textBox.Text, out double points))
                {
                    if (points < 0)
                    {
                        // --- FIX: Use full namespace ---
                        System.Windows.MessageBox.Show($"Points for '{partName}' cannot be negative.",
                           "Invalid Input", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                        PointsMap = null;
                        return;
                    }
                    PointsMap[partName] = points; // Use the ORIGINAL part name here
                }
                else
                {
                    // --- FIX: Use full namespace ---
                    System.Windows.MessageBox.Show($"Invalid number for '{partName}'. Please enter a valid number (e.g., 10 or 5.5).",
                        "Invalid Input", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    PointsMap = null;
                    return;
                }
            }

            this.DialogResult = true;
            this.Close();
        }
    }
}