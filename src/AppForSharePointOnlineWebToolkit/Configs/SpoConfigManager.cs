using System.Collections.Specialized;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Xml;
using System.Xml.Serialization;

using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;

using Formatting = Newtonsoft.Json.Formatting;

namespace AppForSharePointOnlineWebToolkit.Configs
{
    /// <summary>
    /// This represents the configuration manager entity for SharePoint Online.
    /// </summary>
    public class SpoConfigManager
    {
        private const string SpoConfig = "spoconfig.json";

        private readonly JsonSerializerSettings _settings;
        private readonly SpoConfiguration _root;

        /// <summary>
        /// Initializes a new instance of the <see cref="SpoConfigManager"/> class.
        /// </summary>
        public SpoConfigManager()
        {
            this._settings = Init();
            this._root = this.Load();
        }

        /// <summary>
        /// Gets the app settings section of either app.config or web.config.
        /// </summary>
        public NameValueCollection AppSettings => this._root.AppSettings;

        private static JsonSerializerSettings Init()
        {
            var settings = new JsonSerializerSettings()
                               {
                                   Formatting = Formatting.Indented,
                                   ContractResolver = new CamelCasePropertyNamesContractResolver(),
                                   NullValueHandling = NullValueHandling.Ignore,
                                   MissingMemberHandling = MissingMemberHandling.Ignore,
                               };
            return settings;
        }

        private SpoConfiguration Load()
        {
            if (!File.Exists(SpoConfig))
            {
                throw new FileNotFoundException("spoconfig.json not found");
            }

            using (var stream = new FileStream(SpoConfig, FileMode.Open, FileAccess.Read))
            using (var reader = new StreamReader(stream))
            {
                var root = JsonConvert.DeserializeObject<SpoConfiguration>(reader.ReadToEnd(), this._settings);
                return root;
            }
        }
    }
}