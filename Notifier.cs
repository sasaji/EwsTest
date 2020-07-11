using System;
using System.Collections.Generic;
using System.Threading;
using Microsoft.Exchange.WebServices.Data;

namespace EwsTest
{
    class Notifier
    {
        public void Run(List<User> users, string watermark, int interval)
        {
            foreach (User user in users) {
                TimerCallback timerDelegate = new TimerCallback(Notify);
                Console.WriteLine("Head Watermark: " + watermark);
                Timer timer = new Timer(timerDelegate, user, 0, interval);
            }

            while (true) {
                if (Console.ReadLine().ToLower() == "end")
                    break;
            }
        }

        private void Notify(object o)
        {
            User user = (User)o;
            Console.WriteLine("Notify for " + user.Id + "...");
            GetEventsResults results = user.Subscription.GetEvents();
            Console.WriteLine("Current Watermark: " + user.Subscription.Watermark);

            foreach (ItemEvent eventItem in results.ItemEvents) {
                string eventType = String.Empty;
                if (eventItem.EventType == EventType.NewMail)
                    eventType = "新規";
                else if (eventItem.EventType == EventType.Deleted)
                    eventType = "削除";
                else if (eventItem.EventType == EventType.Moved)
                    eventType = "移動";
                Console.WriteLine("ItemId: " + eventItem.ItemId);
                try {
                    EmailMessage message = EmailMessage.Bind(user.Service, eventItem.ItemId);
                    Console.WriteLine(eventType + " : " + message.Subject);
                    Console.WriteLine("Saved Watermark: " + user.Subscription.Watermark);
                } catch (Exception exception) {
                    Console.WriteLine(exception.Message);
                } finally {
                }
            }
        }
    }
}
