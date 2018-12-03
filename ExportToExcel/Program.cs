using SC.API.ComInterop;

namespace ExportToExcel
{
    public class Program
    {
        static void Main(string[] args)
        {
            // Load story and get data

            var config = new Configuration();

            var api = new SharpCloudApi(
                config.Username,
                config.Password,
                config.Url);

            var story = api.LoadStory(config.StoryId);
            var itemsData = story.GetItemsData();
            var relationshipsData = story.GetRelationshipsData();

            // Write data to Excel

            var writer = new ExcelWriter();
            writer.AddSheet("Items", itemsData);
            writer.AddSheet("Relationships", relationshipsData);
        }
    }
}
