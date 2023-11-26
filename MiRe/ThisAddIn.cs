using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;

namespace MiRe
{
    public partial class ThisAddIn
    {

        public Outlook.Application application = null;
        public Outlook.Explorer activeExplorer = null;
        public Outlook.Inspectors inspectors = null;
        public Outlook.Inspector activeInspector = null;
        public Outlook.Explorers explorers = null;
        public InfoCard myUserControl1;
        public Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Get the Application object
            application = this.Application;

            // Get the Inspector object
            inspectors = application.Inspectors;

            // Get the active Inspector object
            activeInspector = application.ActiveInspector();

            // Get the Explorer objects
            explorers = application.Explorers;

            // Get the active Explorer object
            activeExplorer = application.ActiveExplorer();


            activeExplorer.SelectionChange += new Outlook
                .ExplorerEvents_10_SelectionChangeEventHandler
                (CurrentExplorer_Event);

            application.AdvancedSearchComplete += Application_AdvancedSearchComplete;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            application.AdvancedSearchComplete -= Application_AdvancedSearchComplete;
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

        private void CurrentExplorer_Event()
        {
            Outlook.MAPIFolder selectedFolder = this.Application.ActiveExplorer().CurrentFolder;
            try
            {
                if (this.Application.ActiveExplorer().Selection.Count > 0)
                {
                    Object selObject = this.Application.ActiveExplorer().Selection[1];

                    if (selObject is Outlook.AppointmentItem)
                    {
                        Outlook.AppointmentItem apptItem = (selObject as Outlook.AppointmentItem);
                        RunAdvancedSearch(application, apptItem.Subject, apptItem.ConversationID);

                    }
                }
            }
            catch (System.Exception ex)
            {
            }
            //MessageBox.Show(expMessage);
        }

        public string AdvancedSearchTag;
        //SEARCH Function
        Search RunAdvancedSearch(Outlook.Application Application, string wordInSubject, string advancedSearchTag)
        {
            AdvancedSearchTag = advancedSearchTag;
            string scope = "Inbox";
            //string filter = string.Join("", "[Subject] = '", wordInSubject, "'");
            string filter = "urn:schemas:mailheader:subject LIKE \'%" + wordInSubject + "%\'";
            //string filter = "urn:schemas:mailheader:subject = \'" + wordInSubject + "\'";
            Outlook.Search advancedSearch = null;
            Outlook.MAPIFolder folderInbox = null;
            Outlook.MAPIFolder folderSentMail = null;
            Outlook.NameSpace ns = null;
            try
            {
                ns = Application.GetNamespace("MAPI");
                folderInbox = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                folderSentMail = ns.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                scope = "\'" + folderInbox.FolderPath + "\',\'" + folderSentMail.FolderPath + "\'";
                //scope = "\'" + folderInbox.FolderPath + "\'";
                advancedSearch = Application.AdvancedSearch(scope, filter, true, AdvancedSearchTag);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "An eexception is thrown");
            }
            finally
            {
                if (advancedSearch != null) Marshal.ReleaseComObject(advancedSearch);
                //if (folderSentMail != null) Marshal.ReleaseComObject(folderSentMail);
                if (folderInbox != null) Marshal.ReleaseComObject(folderInbox);
                if (ns != null) Marshal.ReleaseComObject(ns);
            }

            return advancedSearch;
        }
        //Handle AdvancedSearchComplete event
        void Application_AdvancedSearchComplete(Outlook.Search SearchObject)
        {
            Outlook.Results advancedSearchResults = null;
            Outlook.MailItem resultItem = null;
            System.Text.StringBuilder strBuilder = null;
            string title = "";
            try
            {
                if (SearchObject.Tag == AdvancedSearchTag)
                {
                    advancedSearchResults = SearchObject.Results;


                    System.Diagnostics.Debug.WriteLine("Count: " + advancedSearchResults.Count);
                    if (advancedSearchResults.Count > 0)
                    {
                        strBuilder = new System.Text.StringBuilder();

                        for (int i = 1; i <= advancedSearchResults.Count; i++)
                        {
                            var item = advancedSearchResults[i];
                            if (item != null)
                            {
                                try
                                {
                                    if ((item.MessageClass == "IPM.Schedule.Meeting.Resp.Neg" ||
                                   item.MessageClass == "IPM.Schedule.Meeting.Resp.Pos" ||
                                   item.MessageClass == "IPM.Schedule.Meeting.Resp.Tent") &&
                                   !string.IsNullOrWhiteSpace(item?.Body))
                                    {
                                        title = item.ConversationTopic;
                                        try { strBuilder.AppendLine("Resp: " + GetResponseText(item.MessageClass) + " "); } catch { }
                                        try { strBuilder.AppendLine("|Received: " + item.ReceivedTime + " "); } catch { }
                                        try { strBuilder.AppendLine("|Subject: " + item.Subject + " "); } catch { }
                                        try { strBuilder.AppendLine("|SenderName: " + item.SenderName + " "); } catch { }
                                        try { strBuilder.AppendLine("|Body: " + item.Body.Replace("\n", "").Replace("\r", " ") + " "); } catch { }
                                        //try { strBuilder.Append("MessageClass: " + item.MessageClass); } catch { }
                                        //try { strBuilder.Append("SenderEmailType: " + item.SenderEmailType); } catch { }
                                        //try { strBuilder.Append("|CreationTime: " + item.CreationTime + " "); } catch { }
                                        //try { strBuilder.Append("|LastModificationTime: " + item.LastModificationTime + " "); } catch { }                                        
                                        //try { strBuilder.Append("|SentOn: " + item.SentOn + " "); } catch { }
                                        //try { strBuilder.Append("|ReceivedTime: " + item.ReceivedTime + " "); } catch { }
                                        try { strBuilder.Append("\n\r----------------------------\n\r"); } catch { }
                                    }
                                }
                                catch { }

                                Marshal.ReleaseComObject(item);
                            }

                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("There are no items found.");
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "An exception is occured");
            }
            finally
            {
                if (resultItem != null) Marshal.ReleaseComObject(resultItem);
                if (advancedSearchResults != null)
                    Marshal.ReleaseComObject(advancedSearchResults);

                if (strBuilder?.Length > 0)
                {
                    //string fileName = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".txt";
                    //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@fileName))
                    //{
                    //    file.WriteLine(strBuilder.ToString());
                    //}
                    //Process.Start(fileName);

                    myUserControl1 = new InfoCard(strBuilder.ToString().Split(Environment.NewLine.ToCharArray()));
                    var width = myUserControl1.Width;
                    myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "Response Notes (" + title ??"" + ")");
                    myCustomTaskPane.Visible = true;
                    myCustomTaskPane.Width = width + 10;
                }
            }
        }

        string GetResponseText(string text)
        {
            if (text == "IPM.Schedule.Meeting.Resp.Neg") return "Neg";
            if (text == "IPM.Schedule.Meeting.Resp.Pos") return "Pos";
            if (text == "IPM.Schedule.Meeting.Resp.Tent") return "Tent";
            return "None";
        }


        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
