using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.Configuration;
using System.IO;

namespace OutlookAddIn2
{
    public partial class ThisAddIn
    {
        public Outlook.Application OutlookApplication;
        public Outlook.Inspectors OutlookInspectors;
        public Outlook.Inspector OutlookInspector;
        public Outlook.MailItem OutlookMailItem;
        private object applicationObject;
        private object addInInstance;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            OutlookApplication = Application as Outlook.Application;
            OutlookInspectors = OutlookApplication.Inspectors;
            OutlookInspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(OutlookInspectors_NewInspector);
            OutlookApplication.ItemSend += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemSendEventHandler(OutlookApplication_ItemSend);
        }

        void OutlookInspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            OutlookInspector = (Outlook.Inspector)Inspector;
            if (Inspector.CurrentItem is Outlook.MailItem)
            {
                OutlookMailItem = (Outlook.MailItem)Inspector.CurrentItem;
            }

        }
        void OutlookApplication_ItemSend(object Item, ref bool Cancel)
        {

            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            string filepath = config.AppSettings.Settings["FileAddress"].Value + DateTime.Now.Millisecond.ToString() + ".rfc";
            FileStream fstream = new FileStream(filepath, FileMode.Create);
            StreamWriter writer = new StreamWriter(fstream, System.Text.ASCIIEncoding.ASCII);
            writer.Write("From: " + OutlookMailItem.SenderEmailAddress + "\r");
            writer.Write("To: " + OutlookMailItem.To + "\r");
            writer.Write("Subject: " + OutlookMailItem.Subject + "\r\n");
            writer.Write(OutlookMailItem.Body);

            writer.Flush();
            writer.Dispose();

            //MessageBox.Show(strchk + "\r\n" + strchkTo);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
