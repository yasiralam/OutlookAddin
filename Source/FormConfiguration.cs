using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
namespace OutlookAddIn
{
    public partial class FormConfiguration : Form
    {
        public FormConfiguration()
        {
            InitializeComponent();
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            var path = pathTextBox.Text.Trim();
            if (string.IsNullOrEmpty(path) || string.IsNullOrWhiteSpace(path))
            {
                MessageBox.Show("Please enter path", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else 
            {
                var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                config.AppSettings.Settings["Path"].Value = path;
                config.Save(ConfigurationSaveMode.Modified);
                MessageBox.Show("Settings Saved");
                this.Close();
            }
        }

        private void FormConfiguration_Load(object sender, EventArgs e)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            pathTextBox.Text = config.AppSettings.Settings["Path"].Value;
        }
    }
}
