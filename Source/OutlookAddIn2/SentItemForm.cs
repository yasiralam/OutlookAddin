using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Data;
using System.Windows.Forms;
using System.Reflection;

namespace OutlookAddIn2
{
    public partial class SentItemForm
    {
        private void SentItemForm_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            SettingsForSavingSentItems settings = new SettingsForSavingSentItems();
            settings.Show();
        }
    }
}
