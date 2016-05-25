using System;

using Microsoft.SharePoint.Client;

namespace AppForSharePointOnlineWebToolkit
{
    /// <summary>
    /// This represents the helper class for the <see cref="ClientContext"/> class.
    /// </summary>
    public class ClientContextHelper : IClientContextHelper
    {
        private bool _disposed;

        /// <summary>
        /// Creates a new instance of the the <see cref="ClientContextWrapper"/> class.
        /// </summary>
        /// <param name="targetUri">Target site URL value.</param>
        /// <returns>Returns the <see cref="ClientContextWrapper"/> instance created.</returns>
        public IClientContextWrapper CreateAppOnlyClientContext(string targetUri)
        {
            if (string.IsNullOrWhiteSpace(targetUri))
            {
                throw new ArgumentNullException(nameof(targetUri));
            }

            return this.CreateAppOnlyClientContext(new Uri(targetUri));
        }

        /// <summary>
        /// Creates a new instance of the the <see cref="ClientContextWrapper"/> class.
        /// </summary>
        /// <param name="targetUri">Target site URI value.</param>
        /// <returns>Returns the <see cref="ClientContextWrapper"/> instance created.</returns>
        public IClientContextWrapper CreateAppOnlyClientContext(Uri targetUri)
        {
            if (targetUri == null)
            {
                throw new ArgumentNullException(nameof(targetUri));
            }

            var realm = TokenHelper.GetRealmFromTargetUrl(targetUri);
            var response = TokenHelper.GetAppOnlyAccessToken(TokenHelper.SharePointPrincipal, targetUri.Authority, realm);
            var context = TokenHelper.GetClientContextWithAccessToken(targetUri.ToString(), response.AccessToken);
            return new ClientContextWrapper(context);
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (this._disposed)
            {
                return;
            }

            this._disposed = true;
        }
    }
}