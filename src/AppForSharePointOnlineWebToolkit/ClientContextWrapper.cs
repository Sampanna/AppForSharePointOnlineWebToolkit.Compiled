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
    /// This represents the wrapper entity for the <see cref="ClientContext"/> class.
    /// </summary>
    public class ClientContextWrapper : IClientContextWrapper
    {
        private bool _disposed;

        /// <summary>
        /// Initializes a new instance of the <see cref="ClientContextWrapper"/> class.
        /// </summary>
        /// <param name="context"><see cref="ClientContext"/> instance.</param>
        public ClientContextWrapper(ClientContext context = null)
        {
            if (context == null)
            {
                return;
            }

            this.ContextInstance = context;
        }

        /// <summary>
        /// Gets or sets the <see cref="ClientContext"/> instance.
        /// </summary>
        public ClientContext ContextInstance { get; set; }

        /// <summary>
        /// Gets the <see cref="Web"/> instance.
        /// </summary>
        public Web Web
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.Web;
            }
        }

        /// <summary>
        /// Gets the <see cref="Site"/> instance.
        /// </summary>
        public Site Site
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.Site;
            }
        }

        /// <summary>
        /// Gets the <see cref="RequestResources"/> instance.
        /// </summary>
        public RequestResources RequestResources
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.RequestResources;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether form digest handling enabled or not.
        /// </summary>
        public bool FormDigestHandlingEnabled
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.FormDigestHandlingEnabled;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.FormDigestHandlingEnabled = value;
            }
        }

        /// <summary>
        /// Gets the <see cref="Version"/> instance.
        /// </summary>
        public Version ServerVersion
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.ServerVersion;
            }
        }

        /// <summary>
        /// Gets the URL.
        /// </summary>
        public string Url
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.Url;
            }
        }

        /// <summary>
        /// Gets or sets the application name.
        /// </summary>
        public string ApplicationName
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.ApplicationName;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.ApplicationName = value;
            }
        }

        /// <summary>
        /// Gets or sets the client tag.
        /// </summary>
        public string ClientTag
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.ClientTag;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.ClientTag = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to disable return value cache or not.
        /// </summary>
        public bool DisableReturnValueCache
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.DisableReturnValueCache;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.DisableReturnValueCache = value;
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether to validate on client or not.
        /// </summary>
        public bool ValidateOnClient
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.ValidateOnClient;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.ValidateOnClient = value;
            }
        }

        /// <summary>
        /// Gets or sets the authentication mode.
        /// </summary>
        public ClientAuthenticationMode AuthenticationMode
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.AuthenticationMode;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.AuthenticationMode = value;
            }
        }

        /// <summary>
        /// Gets or sets the forms authentication login info.
        /// </summary>
        public FormsAuthenticationLoginInfo FormsAuthenticationLoginInfo
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.FormsAuthenticationLoginInfo;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.FormsAuthenticationLoginInfo = value;
            }
        }

        /// <summary>
        /// Gets or sets the credentials.
        /// </summary>
        public ICredentials Credentials
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.Credentials;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.Credentials = value;
            }
        }

        /// <summary>
        /// Gets or sets the web request executor factory.
        /// </summary>
        public WebRequestExecutorFactory WebRequestExecutorFactory
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.WebRequestExecutorFactory;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.WebRequestExecutorFactory = value;
            }
        }

        /// <summary>
        /// Gets the pending request.
        /// </summary>
        public ClientRequest PendingRequest
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.PendingRequest;
            }
        }

        /// <summary>
        /// Gets a value indicating whether the context has pending request or not.
        /// </summary>
        public bool HasPendingRequest
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.HasPendingRequest;
            }
        }

        /// <summary>
        /// Gets or sets the tag.
        /// </summary>
        public object Tag
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.Tag;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.Tag = value;
            }
        }

        /// <summary>
        /// Gets or sets the request timeout.
        /// </summary>
        public int RequestTimeout
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.RequestTimeout;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.RequestTimeout = value;
            }
        }

        /// <summary>
        /// Gets the static objects.
        /// </summary>
        public Dictionary<string, object> StaticObjects
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.StaticObjects;
            }
        }

        /// <summary>
        /// Gets the server schema version.
        /// </summary>
        public Version ServerSchemaVersion
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.ServerSchemaVersion;
            }
        }

        /// <summary>
        /// Gets the server library version.
        /// </summary>
        public Version ServerLibraryVersion
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.ServerLibraryVersion;
            }
        }

        /// <summary>
        /// Gets or sets the request schema version.
        /// </summary>
        public Version RequestSchemaVersion
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.RequestSchemaVersion;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.RequestSchemaVersion = value;
            }
        }

        /// <summary>
        /// Gets or sets the trace correlation Id.
        /// </summary>
        public string TraceCorrelationId
        {
            get
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                return this.ContextInstance.TraceCorrelationId;
            }

            set
            {
                if (this.ContextInstance == null)
                {
                    throw new InvalidOperationException();
                }

                this.ContextInstance.TraceCorrelationId = value;
            }
        }

        /// <summary>
        /// Gets the <see cref="FormDigestInfo"/> instance.
        /// </summary>
        /// <returns>Returns the <see cref="FormDigestInfo"/> instance.</returns>
        public FormDigestInfo GetFormDigestDirect()
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            return this.ContextInstance.GetFormDigestDirect();
        }

        /// <summary>
        /// Executes loaded query.
        /// </summary>
        public void ExecuteQuery()
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            this.ContextInstance.ExecuteQuery();
        }

        /// <summary>
        /// Executes loaded query asynchronously.
        /// </summary>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        public Task ExecuteQueryAsync()
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            return Task.Factory.StartNew(this.ExecuteQuery);
        }

        /// <summary>
        /// Casts the <see cref="ClientObject"/> instance to the given type.
        /// </summary>
        /// <param name="obj"><see cref="ClientObject"/> instance.</param>
        /// <typeparam name="T">Type to convert.</typeparam>
        /// <returns>Returns the converted type.</returns>
        public T CastTo<T>(ClientObject obj) where T : ClientObject
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (obj == null)
            {
                throw new ArgumentNullException(nameof(obj));
            }

            return this.ContextInstance.CastTo<T>(obj);
        }

        /// <summary>
        /// Casts the <see cref="ClientObject"/> instance to the given type asynchronously.
        /// </summary>
        /// <param name="obj"><see cref="ClientObject"/> instance.</param>
        /// <typeparam name="T">Type to convert.</typeparam>
        /// <returns>Returns the converted type.</returns>
        public Task<T> CastToAsync<T>(ClientObject obj) where T : ClientObject
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (obj == null)
            {
                throw new ArgumentNullException(nameof(obj));
            }

            var result = default(T);
            Task.Factory.StartNew(() => { result = this.CastTo<T>(obj); });
            return Task.FromResult(result);
        }

        /// <summary>
        /// Adds query into the context.
        /// </summary>
        /// <param name="query"><see cref="ClientAction"/> instance for query.</param>
        public void AddQuery(ClientAction query)
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (query == null)
            {
                throw new ArgumentNullException(nameof(query));
            }

            this.ContextInstance.AddQuery(query);
        }

        /// <summary>
        /// Adds query into the context asynchronously.
        /// </summary>
        /// <param name="query"><see cref="ClientAction"/> instance for query.</param>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        public Task AddQueryAsync(ClientAction query)
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (query == null)
            {
                throw new ArgumentNullException(nameof(query));
            }

            return Task.Factory.StartNew(() => { this.AddQuery(query); });
        }

        /// <summary>
        /// Adds query Id and result object into the context.
        /// </summary>
        /// <param name="id">Query Id.</param>
        /// <param name="obj">Result object.</param>
        public void AddQueryIdAndResultObject(long id, object obj)
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (obj == null)
            {
                throw new ArgumentNullException(nameof(obj));
            }

            this.ContextInstance.AddQueryIdAndResultObject(id, obj);
        }

        /// <summary>
        /// Adds query Id and result object into the context asynchronously.
        /// </summary>
        /// <param name="id">Query Id.</param>
        /// <param name="obj">Result object.</param>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        public Task AddQueryIdAndResultObjectAsync(long id, object obj)
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (obj == null)
            {
                throw new ArgumentNullException(nameof(obj));
            }

            return Task.Factory.StartNew(() => { this.AddQueryIdAndResultObject(id, obj); });
        }

        /// <summary>
        /// Parses object from the JSON string.
        /// </summary>
        /// <param name="json">JSON string value.</param>
        /// <returns>Returns the object parsed.</returns>
        public object ParseObjectFromJsonString(string json)
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentNullException(nameof(json));
            }

            return this.ContextInstance.ParseObjectFromJsonString(json);
        }

        /// <summary>
        /// Parses object from the JSON string asynchronously.
        /// </summary>
        /// <param name="json">JSON string value.</param>
        /// <returns>Returns the object parsed.</returns>
        public Task<object> ParseObjectFromJsonStringAsync(string json)
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (string.IsNullOrWhiteSpace(json))
            {
                throw new ArgumentNullException(nameof(json));
            }

            object result = null;
            Task.Factory.StartNew(() => { result = this.ParseObjectFromJsonString(json); });
            return Task.FromResult(result);
        }

        /// <summary>
        /// Loads the client object.
        /// </summary>
        /// <param name="clientObject">Client object to load.</param>
        /// <param name="retrievals">Expression for retrieval.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        public void Load<T>(T clientObject, params Expression<Func<T, object>>[] retrievals) where T : ClientObject
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (clientObject == null)
            {
                throw new ArgumentNullException(nameof(clientObject));
            }

            this.ContextInstance.Load(clientObject, retrievals);
        }

        /// <summary>
        /// Loads the client object asynchronously.
        /// </summary>
        /// <param name="clientObject">Client object to load.</param>
        /// <param name="retrievals">Expression for retrieval.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        public Task LoadAsync<T>(T clientObject, params Expression<Func<T, object>>[] retrievals) where T : ClientObject
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (clientObject == null)
            {
                throw new ArgumentNullException(nameof(clientObject));
            }

            return Task.Factory.StartNew(() => { this.Load(clientObject, retrievals); });
        }

        /// <summary>
        /// Loads the client objects.
        /// </summary>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        public IEnumerable<T> LoadQuery<T>(ClientObjectCollection<T> clientObjects) where T : ClientObject
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (clientObjects == null)
            {
                throw new ArgumentNullException(nameof(clientObjects));
            }

            var results = this.ContextInstance.LoadQuery(clientObjects);
            return results;
        }

        /// <summary>
        /// Loads the client objects asynchronously.
        /// </summary>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        public Task<IEnumerable<T>> LoadQueryAsync<T>(ClientObjectCollection<T> clientObjects) where T : ClientObject
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (clientObjects == null)
            {
                throw new ArgumentNullException(nameof(clientObjects));
            }

            IEnumerable<T> results = null;
            Task.Factory.StartNew(() => { results = this.LoadQuery(clientObjects); });
            return Task.FromResult(results);
        }

        /// <summary>
        /// Loads the client objects.
        /// </summary>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        public IEnumerable<T> LoadQuery<T>(IQueryable<T> clientObjects) where T : ClientObject
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (clientObjects == null)
            {
                throw new ArgumentNullException(nameof(clientObjects));
            }

            var results = this.ContextInstance.LoadQuery(clientObjects);
            return results;
        }

        /// <summary>
        /// Loads the client objects asynchronously.
        /// </summary>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        public Task<IEnumerable<T>> LoadQueryAsync<T>(IQueryable<T> clientObjects) where T : ClientObject
        {
            if (this.ContextInstance == null)
            {
                throw new InvalidOperationException();
            }

            if (clientObjects == null)
            {
                throw new ArgumentNullException(nameof(clientObjects));
            }

            IEnumerable<T> results = null;
            Task.Factory.StartNew(() => { results = this.LoadQuery(clientObjects); });
            return Task.FromResult(results);
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
