using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Threading.Tasks;

using Microsoft.SharePoint.Client;

namespace AppForSharePointOnlineWebToolkit.Extensions
{
    /// <summary>
    /// This represents the extensions entity for the <see cref="ClientContext"/> class.
    /// </summary>
    public static class ClientContextExtensions
    {
        /// <summary>
        /// Executes loaded query asynchronously.
        /// </summary>
        /// <param name="context"><see cref="ClientContext"/> instance to extend.</param>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        public static Task ExecuteQueryAsync(this ClientContext context)
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            return Task.Factory.StartNew(context.ExecuteQuery);
        }

        /// <summary>
        /// Loads the client object asynchronously.
        /// </summary>
        /// <param name="context"><see cref="ClientContext"/> instance to extend.</param>
        /// <param name="clientObject">Client object to load.</param>
        /// <param name="retrievals">Expression for retrieval.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the <see cref="Task"/>.</returns>
        public static Task LoadAsync<T>(this ClientContext context, T clientObject, params Expression<Func<T, object>>[] retrievals)
            where T : ClientObject
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (clientObject == null)
            {
                throw new ArgumentNullException(nameof(clientObject));
            }

            return Task.Factory.StartNew(() => { context.Load(clientObject, retrievals); });
        }

        /// <summary>
        /// Loads the client objects asynchronously.
        /// </summary>
        /// <param name="context"><see cref="ClientContext"/> instance to extend.</param>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        public static Task<IEnumerable<T>> LoadQueryAsync<T>(this ClientContext context, ClientObjectCollection<T> clientObjects)
            where T : ClientObject
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (clientObjects == null)
            {
                throw new ArgumentNullException(nameof(clientObjects));
            }

            IEnumerable<T> results = null;
            Task.Factory.StartNew(() => { results = context.LoadQuery(clientObjects); });
            return Task.FromResult(results);
        }

        /// <summary>
        /// Loads the client objects asynchronously.
        /// </summary>
        /// <param name="context"><see cref="ClientContext"/> instance to extend.</param>
        /// <param name="clientObjects">Client objects to load.</param>
        /// <typeparam name="T">Object type inheriting <see cref="ClientObject"/>.</typeparam>
        /// <returns>Returns the list of client objects.</returns>
        public static Task<IEnumerable<T>> LoadQueryAsync<T>(this ClientContext context, IQueryable<T> clientObjects)
            where T : ClientObject
        {
            if (context == null)
            {
                throw new ArgumentNullException(nameof(context));
            }

            if (clientObjects == null)
            {
                throw new ArgumentNullException(nameof(clientObjects));
            }

            IEnumerable<T> results = null;
            Task.Factory.StartNew(() => { results = context.LoadQuery(clientObjects); });
            return Task.FromResult(results);
        }
    }
}