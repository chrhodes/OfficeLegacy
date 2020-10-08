using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Outlook=Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Core;
using Microsoft.Office.Tools.Outlook;

namespace SupportTools_Outlook.User_Interface.Task_Panes
{
    public partial class TaskPane_AppUtil : UserControl
    {
        public TaskPane_AppUtil()
        {
            InitializeComponent();
        }

        private void AddRules()
        {
            Common.WriteToDebugWindow("AddRules");
            Outlook.Folders sessionFolders = Globals.ThisAddIn.Application.Session.Folders;
            Outlook.Folder inbox =  (Outlook.Folder)Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            Outlook.Folders inboxFolders = inbox.Folders;
            Outlook.Folder junkFolder = (Outlook.Folder)Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJunk);
            //Outlook.Folders advertisementFolders = Globals.ThisAddIn.Application.Session.GetFolderFromID("Advertisements").Folders;
            Outlook.Folder advertisementsFolder = (Outlook.Folder)sessionFolders["crhodes"].Folders["Advertisements"];
            Outlook.Folder edvantageFolder = (Outlook.Folder)sessionFolders["crhodes"].Folders["EDVantage"];

            foreach (Outlook.Folder folder in sessionFolders)
            {
                Common.WriteToDebugWindow(folder.Name);
            }

            Outlook.Folder victoriaFolder;

            try
            {
                victoriaFolder = (Outlook.Folder)inboxFolders["Victoria Secret"];
            }
            catch
            {
                victoriaFolder = (Outlook.Folder)inboxFolders.Add("Victoria Secret");
            }


            Outlook.AddressEntry currentUser = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            Outlook.Rules currentRules = Globals.ThisAddIn.Application.Session.DefaultStore.GetRules();

            Outlook.Rule victoriaRule;
            Outlook.Rule advertisementsRule;

            //try
            //{
            //    victoriaRule = currentRules["Victoria Secret"];
            //}
            //catch (Exception ex)
            //{
                
            //}

            victoriaRule = currentRules.Create("Victoria Secret", Outlook.OlRuleType.olRuleReceive);
            string[] victoriaAddress = { "VictoriasSecret@e.victoriassecret.com" };
            victoriaRule.Conditions.SenderAddress.Address = victoriaAddress;
            victoriaRule.Conditions.SenderAddress.Enabled = true;

            victoriaRule.Actions.MoveToFolder.Folder = victoriaFolder;
            victoriaRule.Actions.MoveToFolder.Enabled = true;

            advertisementsRule = currentRules.Create("Advertisements", Outlook.OlRuleType.olRuleReceive);
            string[] advertisersAddresses = 
            { 
                "mail@e.groupon.com"
                , "eBags@response.ebags.com"
                , "news@airage.com"
                , "info@em.surveynetwork.com"
                , "princesscruises@email.princess.com"
                , "Polytechnic@pushpage.org"
                , "marketing@pcrush.com"
                , "sales@parts-express.com"
                , "pantone@web.pantone.com"
                , "Panera@panera.fbmta.com"
                , "enews@email.ononesoftware.com"
                , "newsletter@onbeing.org"
                , "nuance@reply.digitalriver.com"
                , "northern_tool@email-northerntool.com"
                , "promo@e.newegg.com"
                , "Mouser@e.mouser.com"
                , "do-not-reply@email.globalspec.com"
                , "microsoftstore@microsoftstoreemail.com"
                , "eNews@email.microcentermedia.com"
                , "kohls@email.kohls.com"
                , "AskKelar@kelarpacific.com"
                , "InformationWeek@techwebregionalevents.com"
            };

            advertisementsRule.Conditions.SenderAddress.Address = advertisersAddresses;
            advertisementsRule.Conditions.SenderAddress.Enabled = true;

            advertisementsRule.Actions.MoveToFolder.Folder = advertisementsFolder;
            advertisementsRule.Actions.MoveToFolder.Enabled = true;

            try
            {
                currentRules.Save();
            }
            catch (Exception ex)
            {
                Common.WriteToDebugWindow(ex.ToString());
            }
             
        }
        private void btnAddRules_Click(object sender, EventArgs e)
        {
            AddRules();
        }

        private void ListFolders()
        {
            Common.WriteToDebugWindow("ListFolders");
            Outlook.Folders sessionFolders = Globals.ThisAddIn.Application.Session.Folders;

            foreach (Outlook.Folder folder in sessionFolders)
            {
                Common.WriteToDebugWindow(folder.Name);

                foreach (Outlook.Folder subFolder in folder.Folders)
                {
                    Common.WriteToDebugWindow("   " + subFolder.Name);
                }
            }
        }

        private void btnListFolders_Click(object sender, EventArgs e)
        {
            ListFolders();
        }
    }
}
