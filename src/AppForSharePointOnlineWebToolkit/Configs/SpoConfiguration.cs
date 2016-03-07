using System.Collections.Generic;
using System.Collections.Specialized;

using AppForSharePointOnlineWebToolkit.Extensions;

using Newtonsoft.Json;

namespace AppForSharePointOnlineWebToolkit.Configs
{
    /// <summary>
    /// This represents the root node entity of either app.config or web.config.
    /// </summary>
    public class SpoConfiguration
    {
        [JsonIgnore]
        public NameValueCollection AppSettings => this.AppSettingsList.ToNameValueCollection();

        /// <summary>
        /// Gets or sets the app settings.
        /// </summary>
        [JsonProperty("appSettings")]
        public List<AppSettingItem> AppSettingsList { get; set; }
    }

    /// <summary>
    /// This represents the child node entity of the appSettings node.
    /// </summary>
    public class AppSettingItem
    {
        /// <summary>
        /// Gets or sets the node key.
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// Gets or sets the node value.
        /// </summary>
        public string Value { get; set; }
    }
}