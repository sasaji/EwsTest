using System;
using Microsoft.Exchange.WebServices.Data;

namespace EwsTest
{
    class ExchangeServiceFactory
    {
        public static ExchangeService CreateByNetworkCredential(string id, string password)
        {
            // EWS の接続設定
            ExchangeVersion version = new ExchangeVersion();
            version = ExchangeVersion.Exchange2013_SP1;
            ExchangeService service = new ExchangeService(version);
            service.Credentials = new System.Net.NetworkCredential(id, password);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            return service;
        }

        public static ExchangeService CreateByWebCredential(string id, string password)
        {
            // EWS の接続設定
            ExchangeVersion version = new ExchangeVersion();
            version = ExchangeVersion.Exchange2013_SP1;
            ExchangeService service = new ExchangeService(version);
            //service.UseDefaultCredentials = true;
            //service.Credentials = new WebCredentials(username, password, domain);
            service.Credentials = new WebCredentials(id, password);
            service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            return service;
        }

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }
    }
}
