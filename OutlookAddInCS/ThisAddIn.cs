using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;


namespace OutlookAddInCS
{
    public partial class ThisAddIn
    {
        private void AddAppointment(DateTime start, DateTime end, string loc, string body, bool allday, string subject)
        {
            try
            {
                Outlook.AppointmentItem newAppointment =
                    (Outlook.AppointmentItem)
                this.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
                newAppointment.Start = start;
                newAppointment.End = end;
                newAppointment.Location = loc;
                newAppointment.Body = body;
                newAppointment.AllDayEvent = allday;
                newAppointment.Subject = subject;
                newAppointment.Save();
                newAppointment.Display(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("The following error occurred: " + ex.Message);
            }
        }

        private void ExpAppoint()
        {
            TextWriter tw = new StreamWriter("c:/Users/Team Boeing/Downloads/TeamBoeingSalesApp/appdata.txt");
            Outlook.MAPIFolder calender = Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar);
            Outlook.Items calenderItems = calender.Items;
            foreach (Outlook.AppointmentItem appt in calenderItems)
            {
                if (appt != null)
                {
                    tw.WriteLine(appt.Start.ToString());
                    tw.WriteLine(appt.End.ToString());
                    tw.WriteLine(appt.Location);
                    tw.WriteLine(appt.Body);
                    tw.WriteLine(appt.AllDayEvent.ToString());
                    tw.WriteLine(appt.Subject);
                    //MessageBox.Show(appt.Body);
                }
            }
            tw.Close();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            AddAppointment(DateTime.Now.AddHours(42), DateTime.Now.AddHours(43), "EGR Lobby", "We are meeting because we are meeting", false, "Test Appointment #2");
            ExpAppoint();
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
