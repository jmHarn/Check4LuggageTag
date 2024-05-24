using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace CheckForLuggageTag
{
    public partial class ThisAddIn
    {
        private void InternalStartup()
        {
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            if (Item is Outlook.MailItem)
            {
                Outlook.MailItem olMsg = (Outlook.MailItem)Item;
                // Specify the text you want to search for in the subject line
                string[] searchTerms = { "[HDP-stl_legal.FID", "[HDP-troy_legal.FID", "[HDP-dc_legal.FID", "[HDP-firm_admin.FID" };

                // Check if any of the search terms exist in the subject
                bool containsSearchTerm = searchTerms.Any(term => olMsg.Subject.IndexOf(term, StringComparison.OrdinalIgnoreCase) != -1);

                string Prompt = "***WARNING***\n\nThis message appears to have a luggage tag and might be filed to the DM automatically if you send this email. Are you sure you wish to send this message?";
                if (containsSearchTerm)
                {
                    if (System.Windows.Forms.MessageBox.Show(Prompt, "Check before Sending", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                    {
                        Cancel = true;
                    }
                }
            }
        }
    }
}
