using Microsoft.Exchange.WebServices.Data;

namespace EwsTest
{
    class PullSubscriptionFactory
    {
        public static PullSubscription Create(ExchangeService service, string watermark)
        {
            // 通知の Subscribe
            // 第 1 引数 : 対象のフォルダーID
            // 第 2 引数 : タイムアウト (分)
            // 第 3 引数 : Watermark (初回は null)
            // 第 4 引数以降 : 購読するイベント (複数可能)
            // Watermark を指定すると、後からでも過去のアイテムにさかのぼって取得可能。
            return service.SubscribeToPullNotifications(
                new FolderId[] { new FolderId(WellKnownFolderName.Inbox), new FolderId(WellKnownFolderName.DeletedItems) },
                1, watermark, EventType.NewMail, EventType.Moved, EventType.Deleted);
        }
    }
}
