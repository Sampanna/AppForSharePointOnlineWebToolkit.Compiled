using System;

namespace AppForSharePointOnlineWebToolkit
{
    /// <summary>
    /// This provides interfaces to the <see cref="ClientContextHelper"/> class.
    /// </summary>
    public interface IClientContextHelper : IDisposable
    {
        /// <summary>
        /// Creates a new instance of the the <see cref="ClientContextWrapper"/> class.
        /// </summary>
        /// <param name="targetUri">Target site URL value.</param>
        /// <returns>Returns the <see cref="ClientContextWrapper"/> instance created.</returns>
        ClientContextWrapper CreateAppOnlyClientContext(string targetUri);

        /// <summary>
        /// Creates a new instance of the the <see cref="ClientContextWrapper"/> class.
        /// </summary>
        /// <param name="targetUri">Target site URI value.</param>
        /// <returns>Returns the <see cref="ClientContextWrapper"/> instance created.</returns>
        ClientContextWrapper CreateAppOnlyClientContext(Uri targetUri);
    }
}