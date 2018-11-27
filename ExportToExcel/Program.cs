using SC.API.ComInterop;

namespace ExportToExcel
{
    public class Program
    {
        static void Main(string[] args)
        {
            var config = new Configuration();

            var api = new SharpCloudApi(
                config.Username,
                config.Password,
                config.Url);

            var story = api.LoadStory(config.StoryId);
        }
    }
}
