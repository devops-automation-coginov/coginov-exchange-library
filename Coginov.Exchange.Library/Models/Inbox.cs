using Coginov.Exchange.Library.Enums;

namespace Coginov.Exchange.Library.Models
{
    public class Inbox
    {
        public string User { get; set; }
        public string Password { get; set; }
        public string ReplyFrom { get; set; }
        public string ServerUrl { get; set; }
        public AuthenticationMethod? AuthenticationMethod { get; set; }
        public string TenantId { get; set; }
        public string ClientId { get; set; }
        public string ClientSecret { get; set; }
    }
}