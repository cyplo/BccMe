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

        private void AddMeOnBcc(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            var email = Inspector.CurrentItem as Outlook.MailItem;
            if (email == null) { return; }
            if (email.EntryID != null) { return; }//we're interested only in freshly created messages
            
            email.BCC += " " + Application.Session.CurrentUser.Address;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.Inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(AddMeOnBcc);
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
