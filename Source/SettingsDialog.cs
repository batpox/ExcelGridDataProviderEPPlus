using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExcelGridDataProviderEPPlus
{
    public partial class SettingsDialog : Form
    {
        public SettingsDialog()
        {
            InitializeComponent();
        }

        internal void SetSettings(ExcelGridDataSettings settings)
        {
            excelGridDataSettingsBindingSource.DataSource = settings;
        }

        private void fileBrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Excel Files (*.xls,*.xlsx,*.xlsm)|*.xls;*.xlsx;*.xlsm";
            dlg.Title = "Select File";

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                fileNameTextBox.Text = dlg.FileName;
            }
        }

        // This is to get around an issue with having the radio buttons databound. Apparently the AutoCheck functionality
        //  was interfering in the data-binding (or vise-versa) and you would have to click on a radio button twice to 
        //  get it to "take". This gets around the problem, by manually setting the Checked state for each button on a click.
        private void wholeWorksheetRadioButton_Click(object sender, EventArgs e)
        {
            wholeWorksheetRadioButton.Checked = !wholeWorksheetRadioButton.Checked;
        }
        private void namedRangeRadioButton_Click(object sender, EventArgs e)
        {
            namedRangeRadioButton.Checked = !namedRangeRadioButton.Checked;
        }
        private void specificRangeRadioButton_Click(object sender, EventArgs e)
        {
            specificRangeRadioButton.Checked = !specificRangeRadioButton.Checked;
        }

        private void SettingsDialog_Load(object sender, EventArgs e)
        {

        }
    }
}
