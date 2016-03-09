using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;

namespace AppForSharePointOnlineWebToolkit
{
    /// <summary>
    /// This provides interfaces to the <see cref="ClientContextWrapper"/> class.
    /// </summary>
    public interface IClientContextWrapper : IDisposable
    {
        /// <summary>
        /// Gets or sets the <see cref="ClientContext"/> instance.
        /// </summary>
        ClientContext ContextInstance { get; set; }

        /// <summary>
        /// Gets the <see cref="Web"/> instance.
        /// </summary>
        Web Web { get; }

        /// <summary>
        /// Gets the <see cref="Site"/> instance.
        /// </summary>
        Site Site { get; }

        /// <summary>
        /// Gets the <see cref="RequestResources"/> instance.
        /// </summary>
        RequestResources RequestResources { get; }

        /// <summary>
        /// Gets or sets a value indicating whether form digest handling enabled or not.
        /// </summary>
        bool FormDigestHandlingEnabled { get; set; }

        /// <summary>
        /// Gets the <see cref="Version"/> instance.
        /// </summary>
        Version ServerVersion { get; }

        /// <summary>
        /// Gets the URL.
        /// </summary>
        string Url { get; }

        /// <summary>
        /// Gets or sets the application name.
        /// </summary>
        string ApplicationName { get; set; }

        /// <summary>
        /// Gets or sets the client tag.
        /// </summary>
        string ClientTag { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to disable return value cache or not.
        /// </summary>
        bool DisableReturnValueCache { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether to validate on client or not.
        /// </summary>
        bool ValidateOnClient { get; set; }

        /// <summary>
        /// Gets or sets the authentication mode.
        /// </summary>
        ClientAuthenticationMode AuthenticationMode { get; set; }

        /// <summary>
        /// Gets or sets the forms authentication login info.
        /// </summary>
        FormsAuthenticationLoginInfo FormsAuthenticationLoginInfo { get; set; }

        /// <summary>
        /// Gets or sets the credentials.
        /// </summary>
        ICredentials Credentials { get; set; }

        /// <summary>
        /// Gets or sets the web request executor factory.
        /// </summary>
        WebRequestExecutorFactory WebRequestExecutorFactory { get; set; }

        /// <summary>
        /// Gets the pending request.
        /// </summary>
        ClientRequest PendingRequest { get; }

        /// <summary>
        /// Gets a value indicating whether the context has pending request or not.
        /// </summary>
        bool HasPendingRequest { get; }

        /// <summary>
        /// Gets or sets the tag.
        /// </summary>
        object Tag { get; set; }

        /// <summary>
        /// Gets or sets the request timeout.
        /// </summary>
        int RequestTimeout { get; set; }

        /// <summary>
        /// Gets the static objects.
        /// </summary>
        Dictionary<string, object> StaticObjects { get; }

        /// <summary>
        /// Gets the server schema version.
        /// </summary>
        Version ServerSchemaVersion { get; }

        /// <summary>
        /// Gets the server library version.
        /// </summary>
        Version ServerLibraryVersion { get; }

        /// <summary>
        /// Gets or sets the request schema version.
        /// </summary>
        Version RequestSchemaVersion { get; set; }

        /// <summary>
        /// Gets or sets the trace correlation Id.
        /// </summary>
        string TraceCorrelationId { get; set; }

        /// <summary>
        /// Gets the <see cref="FormDigestInfo"/> instance.
        /// </summary>
        /// <returns>Returns the <see cref="FormDigestInfo"/> instance.</returns>
        FormDigestInfo GetFormDigestDirect();

        /// <summary>
        /// Executes loaded query.
        /// </summary>
        void ExecuteQuery();

        /// <summary>
        /// Executes loaded query asynchronously.
        /// </summary>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        Task ExecuteQueryAsync();

        /// <summary>
        /// Casts the <see cref="ClientObject"/> instance to the given type.
        /// </summary>
        /// <param name="obj"><see cref="ClientObject"/> instance.</param>
        /// <typeparam name="T">Type to convert.</typeparam>
        /// <returns>Returns the converted type.</returns>
        T CastTo<T>(ClientObject obj) where T : ClientObject;

        /// <summary>
        /// Casts the <see cref="ClientObject"/> instance to the given type asynchronously.
        /// </summary>
        /// <param name="obj"><see cref="ClientObject"/> instance.</param>
        /// <typeparam name="T">Type to convert.</typeparam>
        /// <returns>Returns the converted type.</returns>
        Task<T> CastToAsync<T>(ClientObject obj) where T : ClientObject;

        /// <summary>
        /// Adds query into the context.
        /// </summary>
        /// <param name="query"><see cref="ClientAction"/> instance for query.</param>
        void AddQuery(ClientAction query);

        /// <summary>
        /// Adds query into the context asynchronously.
        /// </summary>
        /// <param name="query"><see cref="ClientAction"/> instance for query.</param>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        Task AddQueryAsync(ClientAction query);

        /// <summary>
        /// Adds query Id and result object into the context.
        /// </summary>
        /// <param name="id">Query Id.</param>
        /// <param name="obj">Result object.</param>
        void AddQueryIdAndResultObject(long id, object obj);

        /// <summary>
        /// Adds query Id and result object into the context asynchronously.
        /// </summary>
        /// <param name="id">Query Id.</param>
        /// <param name="obj">Result object.</param>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        Task AddQueryIdAndResultObjectAsync(long id, object obj);

        /// <summary>
        /// Parses object from the JSON string.
        /// </summary>
        /// <param name="json">JSON string value.</param>
        /// <returns>Returns the object parsed.</returns>
        object ParseObjectFromJsonString(string json);

        /// <summary>
        /// Parses object from the JSON string asynchronously.
        /// </summary>
        /// <param name="json">JSON string value.</param>
        /// <returns>Returns the object parsed.</returns>
        Task<object> ParseObjectFromJsonStringAsync(string json);

        /// <summary>
        /// Loads the client object.
        /// </summary>
        /// <param name="clientObject">Client object to load.</param>
        /// <param name="retrievals">Expression for retrieval.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        void Load<T>(T clientObject, params Expression<Func<T, object>>[] retrievals) where T : ClientObject;

        /// <summary>
        /// Loads the client object asynchronously.
        /// </summary>
        /// <param name="clientObject">Client object to load.</param>
        /// <param name="retrievals">Expression for retrieval.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        Task LoadAsync<T>(T clientObject, params Expression<Func<T, object>>[] retrievals) where T : ClientObject;

        /// <summary>
        /// Loads the client objects.
        /// </summary>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        IEnumerable<T> LoadQuery<T>(ClientObjectCollection<T> clientObjects) where T : ClientObject;

        /// <summary>
        /// Loads the client objects asynchronously.
        /// </summary>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        Task<IEnumerable<T>> LoadQueryAsync<T>(ClientObjectCollection<T> clientObjects) where T : ClientObject;

        /// <summary>
        /// Loads the client objects.
        /// </summary>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        IEnumerable<T> LoadQuery<T>(IQueryable<T> clientObjects) where T : ClientObject;

        /// <summary>
        /// Loads the client objects asynchronously.
        /// </summary>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        Task<IEnumerable<T>> LoadQueryAsync<T>(IQueryable<T> clientObjects) where T : ClientObject;
    }
}