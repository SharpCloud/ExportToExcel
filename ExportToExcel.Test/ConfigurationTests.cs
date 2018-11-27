using NUnit.Framework;

namespace ExportToExcel.Test
{
    [TestFixture]
    public class ConfigurationTests
    {
        [Test]
        public void ReadsFromConfigurationFile()
        {
            var config = new Configuration();

            Assert.AreEqual("username", config.Username);
            Assert.AreEqual("password", config.Password);
            Assert.AreEqual("url", config.Url);
            Assert.AreEqual("story-id", config.StoryId);
        }
    }
}
