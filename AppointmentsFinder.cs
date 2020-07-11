using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace EwsTest
{
    class AppointmentsFinder
    {
        public void Run(List<User> users, DateTime start, DateTime end)
        {
            foreach (User user in users) {
                CalendarView view = new CalendarView(start, end);
                FindItemsResults<Appointment> results = user.Service.FindAppointments(new FolderId(WellKnownFolderName.Calendar), view);
                foreach (Appointment appointment in results) {
                    appointment.Load();
                    Console.WriteLine("User = " + user.Id);
                    Console.WriteLine("Subject = " + appointment.Subject);
                    Console.WriteLine("StartTime = " + appointment.Start);
                    Console.WriteLine("EndTime = " + appointment.End);
                    //Console.WriteLine("Body Preview = " + appointment.Body.Text);
                    Console.WriteLine("----------------------------------------------------------------");
                }
            }
        }
    }
}