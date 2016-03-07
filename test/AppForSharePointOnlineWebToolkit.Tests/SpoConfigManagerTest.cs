using AppForSharePointOnlineWebToolkit.Configs;

using FluentAssertions;

using Xunit;

namespace AppForSharePointOnlineWebToolkit.Tests
{
    /// <summary>
    /// This represents the test entity for the <see cref="SpoConfigManager"/> class.
    /// </summary>
    public class SpoConfigManagerTest
    {
        /// <summary>
        /// Tests whether the config file is properly loaded or not.
        /// </summary>
        [Fact]
        public void Given_AppConfig_Constructor_ShouldThrow_NoException()
        {
            var manager = new SpoConfigManager();
            var settings = manager.AppSettings;

            settings.Get("ClientSigningCertificatePassword").Should().Be("Pa$$W0rd");
        }
    }
}
