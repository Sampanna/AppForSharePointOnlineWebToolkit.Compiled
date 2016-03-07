using System;
using System.Collections.Generic;
using System.Collections.Specialized;

using AppForSharePointOnlineWebToolkit.Configs;

namespace AppForSharePointOnlineWebToolkit.Extensions
{
    /// <summary>
    /// This represents the extension entity for the <see cref="List{T}"/> class.
    /// </summary>
    public static class ListExtensions
    {
        /// <summary>
        /// Converts the <see cref="List{AppSettingItem}"/> to <see cref="NameValueCollection"/>.
        /// </summary>
        /// <param name="items"><see cref="List{AppSettingItem}"/> instance.</param>
        /// <returns>Returns the <see cref="NameValueCollection"/> instance converted.</returns>
        public static NameValueCollection ToNameValueCollection(this List<AppSettingItem> items)
        {
            if (items == null)
            {
                throw new ArgumentNullException(nameof(items));
            }

            var nvc = new NameValueCollection();
            foreach (var item in items)
            {
                nvc.Add(item.Key, item.Value);
            }

            return nvc;
        }
    }
}