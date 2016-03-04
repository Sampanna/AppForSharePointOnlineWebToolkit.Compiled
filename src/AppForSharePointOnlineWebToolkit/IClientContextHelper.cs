using System;

using Microsoft.SharePoint.Client;

namespace AppForSharePointOnlineWebToolkit
{
    /// <summary>
    /// This provides interfaces to the <see cref="ClientContextHelper"/> class.
    /// </summary>
    public interface IClientContextHelper : IDisposable
    {
        /// <summary>
        /// Creates a new instance of the the <see cref="ClientContext"/> class.
        /// </summary>
        /// <param name="targetUri">Target site URL value.</param>
        /// <returns>Returns the <see cref="ClientContext"/> instance created.</returns>
        ClientContext CreateAppOnlyClientContext(string targetUri);

        /// <summary>
        /// Creates a new instance of the the <see cref="ClientContext"/> class.
        /// </summary>
        /// <param name="targetUri">Target site URI value.</param>
        /// <returns>Returns the <see cref="ClientContext"/> instance created.</returns>
        ClientContext CreateAppOnlyClientContext(Uri targetUri);
    }
}