using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

// TODO:  suivez ces étapes pour activer l'élément (XML) Ruban :

// 1. Copiez le bloc de code suivant dans la classe ThisAddin, ThisWorkbook ou ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon1();
//  }

// 2. Créez des méthodes de rappel dans la région "Rappels du ruban" de cette classe pour gérer les actions des utilisateurs
//    comme les clics sur un bouton. Remarque : si vous avez exporté ce ruban à partir du Concepteur de ruban,
//    vous devrez déplacer votre code des gestionnaires d'événements vers les méthodes de rappel et modifiez le code pour qu'il fonctionne avec
//    le modèle de programmation d'extensibilité du ruban (RibbonX).

// 3. Assignez les attributs aux balises de contrôle dans le fichier XML du ruban pour identifier les méthodes de rappel appropriées dans votre code.  

// Pour plus d'informations, consultez la documentation XML du ruban dans l'aide de Visual Studio Tools pour Office.


namespace OutlookAddIn1
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
       

        public Ribbon1()
        {
        }

        public Bitmap LoadImages(string imageName)
        {
            switch (imageName)
            {
                case "relancer":
                    return new Bitmap(Properties.Resources.relancer);
                default:
                    break;

            }
            return null;
        }
        public void OnActionMyButton_Click(Office.IRibbonControl control)
        {
            
            // Get selected calendar date
            Outlook.Application application = new Outlook.Application();
            Outlook.Explorer explorer = application.ActiveExplorer();
            Outlook.Folder folder = explorer.CurrentFolder as Outlook.Folder;
            Outlook.View view = explorer.CurrentView as Outlook.View;
            String recip_mail = null;
            String title_mail = null;
            String content_mail = null;

            if (view.ViewType == Outlook.OlViewType.olCalendarView)
            {
                //Outlook.CalendarView calView = view as Outlook.CalendarView;
                //DateTime calDateStart = calView.SelectedStartTime;
                //DateTime calDateEnd = calView.SelectedEndTime;

                // Do stuff with dates. 
                //MessageBox.Show("relance test"+ calDateStart+"---"+ calDateEnd);

                const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";
                try
                {
                    if (application.ActiveExplorer().Selection.Count > 0)
                    {
                        Object selObject = application.ActiveExplorer().Selection[1];


                        if (selObject is Outlook.AppointmentItem)
                        {
                            Outlook.AppointmentItem apptItem =
                                (selObject as Outlook.AppointmentItem);
                            title_mail = apptItem.Subject;
                            content_mail = apptItem.Body;
                            Outlook.Recipients recips = apptItem.Recipients;
                            string str = null;

                            foreach (Outlook.Recipient recip in recips)
                            {
                                if (recip.MeetingResponseStatus != Outlook.OlResponseStatus.olResponseAccepted)
                                {
                                    Outlook.PropertyAccessor pa = recip.PropertyAccessor;
                                    string smtpAddress =
                                        pa.GetProperty(PR_SMTP_ADDRESS).ToString();
                                    str += smtpAddress + ";";
                                }

                            }
                            if (str != null)
                            {
                                recip_mail = str;
                                //MessageBox.Show("Not responded: " + recip_mail);
                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }
                //MessageBox.Show("infomation: " + recip_mail+";"+ title_mail+";"+content_mail);
                if (recip_mail !=null && title_mail != null && content_mail != null)
                {
                    sendMail(recip_mail, title_mail, content_mail);
                }

            }
            


            
        }

        private void sendMail(string recip_mail, string title_mail, string content_mail)
        {
            Outlook.Application olApp = new Outlook.Application();
            Outlook.MailItem mailItem = (Outlook.MailItem)olApp.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.To = recip_mail;
            mailItem.Subject = title_mail;
            mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatHTML;
            mailItem.HTMLBody = content_mail;
            //((Outlook._MailItem)mailItem).Send();
            mailItem.Display(true);
            mailItem = null;
            olApp = null;
           // MessageBox.Show("Mail has been sent successfully!");
        }


        #region Membres IRibbonExtensibility

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn1.Ribbon1.xml");
        }

        #endregion

        #region Rappels du ruban
        //Créez des méthodes de rappel ici. Pour plus d'informations sur l'ajout de méthodes de rappel, consultez https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Programmes d'assistance

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
