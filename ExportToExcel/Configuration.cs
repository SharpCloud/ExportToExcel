using System.Configuration;

namespace ExportToExcel
{
    public class Configuration
    {
        public string Username { get; } = ConfigurationManager.AppSettings["Username"];
        public string Password { get; } = ConfigurationManager.AppSettings["Password"];
        public string Url { get; } = ConfigurationManager.AppSettings["Url"];
        public string StoryId { get; } = ConfigurationManager.AppSettings["StoryId"];
   
    }
}
