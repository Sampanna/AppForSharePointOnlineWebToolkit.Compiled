using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
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
        /// Gets the <see cref="Site"/> instance.
        /// </summary>
        Site Site { get; }

        /// <summary>
        /// Gets the <see cref="Web"/> instance.
        /// </summary>
        Web Web { get; }

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