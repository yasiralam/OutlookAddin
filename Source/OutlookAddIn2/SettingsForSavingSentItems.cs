using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace OutlookAddIn2
{
    public partial class SettingsForSavingSentItems : Form
    {
        public SettingsForSavingSentItems()
        {
            InitializeComponent();

            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            tbaddress.Text = config.AppSettings.Settings["FileAddress"].Value;
        }

        private void CancelBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SaveBtn_Click(object sender, EventArgs e)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["FileAddress"].Value = tbaddress.Text;
            config.Save(ConfigurationSaveMode.Modified);
            MessageBox.Show("Settings Saved");
        }
    }
}
