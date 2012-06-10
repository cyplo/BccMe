using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace BccMe
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="email"></param>
        /// <returns> true if any changes done </returns>
        private bool AddBccIfNeeded(Outlook.MailItem email)
        {
            var address = Application.Session.CurrentUser.Address;
            var name = Application.Session.CurrentUser.Name;

            string bcc = email.BCC;
            if (bcc == null) { bcc = ""; }

            if (bcc.Contains(address)) { return false; }
            if (bcc.Contains(name)) { return false; }

            bcc += "  " + address +";";
            email.BCC = bcc;
            email.Recipients.ResolveAll();
            return true;
        }

        private void OnNewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            var email = Inspector.CurrentItem as Outlook.MailItem;
            if (email == null) { return; }
            if (email.EntryID != null) { return; }//we're interested only in freshly created messages

            AddBccIfNeeded(email);
        }

        private void OnSend(object item, ref bool cancel)
        {
            var email = item as Outlook.MailItem;
            if (email == null) { return; }
            cancel = AddBccIfNeeded(email);
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.Inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(OnNewInspector);
            Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(OnSend);
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
