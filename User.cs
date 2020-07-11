using Microsoft.Exchange.WebServices.Data;

namespace EwsTest
{
    class User
    {
        public ExchangeService Service;
        public PullSubscription Subscription;
        public readonly string Id;

        public User(string id, string password, string watermark)
        {
            Id = id;
            Service = ExchangeServiceFactory.CreateByNetworkCredential(id, password);
            Subscription = PullSubscriptionFactory.Create(Service, watermark);
        }
    }
}

