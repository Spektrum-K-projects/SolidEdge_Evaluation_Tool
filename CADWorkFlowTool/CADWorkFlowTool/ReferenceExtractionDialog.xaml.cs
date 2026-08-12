using System.Windows;
using Forms = System.Windows.Forms;

namespace CADWorkFlowTool
{
    public partial class ReferenceExtractionDialog : Window
    {
        public string InputPath { get; private set; }
        public string OutputPath { get; private set; }

        public ReferenceExtractionDialog()
        {
            InitializeComponent();

            BtnBrowseRefInput.Click += BtnBrowseRefInput_Click;
            BtnBrowseRefOutput.Click += BtnBrowseRefOutput_Click;
            BtnRunRefExtract.Click += BtnRunRefExtract_Click;
            BtnCancelRefExtract.Click += BtnCancelRefExtract_Click;
        }

        private void BtnBrowseRefInput_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                var res = dlg.ShowDialog();
                if (res == Forms.DialogResult.OK)
                {
                    TxtRefInput.Text = dlg.SelectedPath;
                }
            }
        }

        private void BtnBrowseRefOutput_Click(object sender, RoutedEventArgs e)
        {
            using (var dlg = new Forms.FolderBrowserDialog())
            {
                var res = dlg.ShowDialog();
                if (res == Forms.DialogResult.OK)
                {
                    TxtRefOutput.Text = dlg.SelectedPath;
                }
            }
        }

        private void BtnRunRefExtract_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtRefInput.Text) || string.IsNullOrWhiteSpace(TxtRefOutput.Text))
            {
                System.Windows.MessageBox.Show(this, "Please select both Input and Output folders.", "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            InputPath = TxtRefInput.Text.Trim();
            OutputPath = TxtRefOutput.Text.Trim();
            this.DialogResult = true;
            this.Close();
        }

        private void BtnCancelRefExtract_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}