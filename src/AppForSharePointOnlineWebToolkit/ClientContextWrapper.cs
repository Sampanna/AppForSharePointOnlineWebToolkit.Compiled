using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
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
