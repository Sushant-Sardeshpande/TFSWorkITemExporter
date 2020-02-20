using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Common;
using Sedco.Products.TFSHelpers.WorkItemsDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using WindowsCredential = Microsoft.VisualStudio.Services.Common.WindowsCredential;

namespace Sedco.Products.TFSHelpers.WorkItemsDefaultExtractor
{
    public class WorkItemsExtractor : IWorkItemsExtractor
    {
        public NetworkCredential Credentials { get; set; } = null;

        private WorkItemStore store;
        private string url = String.Empty;

        public WorkItemsExtractor()
        {
            
        }

        public string TfsURL
        {
            get
            {
                return url;
            }

            set
            {
                if (string.IsNullOrEmpty(value))
                {
                    throw new ArgumentNullException("TfsURL");
                }

                url = value;
                TfsTeamProjectCollection collection;

                // Connect to TFS
                if (Credentials != null && !string.IsNullOrEmpty(Credentials.Domain) && !string.IsNullOrEmpty(Credentials.UserName))
                {
                    collection = new TfsTeamProjectCollection(new Uri(url), Credentials);
                    collection.EnsureAuthenticated();
                }
                else
                {
                    var tfsClientCredentials = TfsClientCredentials.LoadCachedCredentials(new Uri(url), false, true);
                    collection = new TfsTeamProjectCollection(new Uri(url), tfsClientCredentials);
                    collection.EnsureAuthenticated();

                    //collection = new TfsTeamProjectCollection(new Uri(url));
                    //var authTask = Task.Run(() => collection.Authenticate());
                    //authTask.Wait();
                }

                // Get the work item store service
                store = new WorkItemStore(collection);
            }
        }

        public IWorkItemSummary GetWorkItemById(int workItemId)
        {
            try
            {
                WorkItem workItem = store.GetWorkItem(workItemId);

                DefaultWorkItemSummary item = new DefaultWorkItemSummary(workItem);
                return item;
            }
            catch (Exception ex)
            {
                Exception exception = new Exception("Make sure you set TfsURL before calling this method, and that workItemId is valid. Check inner exception for more details.", ex);
                throw exception;
            }
        }

        public IEnumerable<IWorkItemSummary> GetWorkItemsByQueryText(string query)
        {
            try
            {
                WorkItemCollection workItemCollection = store.Query(query);
                List<IWorkItemSummary> result = new List<IWorkItemSummary>();

                foreach (WorkItem workItem in workItemCollection)
                {
                    result.Add(new DefaultWorkItemSummary(workItem));
                    break;
                }

                return result;
            }
            catch (Exception ex)
            {
                Exception exception = new Exception("Make sure you set TfsURL before calling this method, and that the query is valid. Check inner exception for more details.", ex);
                throw exception;
            }
        }

        public IEnumerable<IWorkItemSummary> GetWorkItemsBySavedQuery(string projectName, string queryFolderName, string queryName)
        {
            try
            {
                List<IWorkItemSummary> result = new List<IWorkItemSummary>();

                QueryHierarchy queryRoot = store.Projects[projectName].QueryHierarchy;
                QueryFolder folder = (QueryFolder)queryRoot[queryFolderName];
                QueryDefinition query1 = (QueryDefinition)folder[queryName];
                WorkItemCollection queryResultsqueryResults = store.Query(query1.QueryText);

                foreach (WorkItem workItem in queryResultsqueryResults)
                {
                    result.Add(new DefaultWorkItemSummary(workItem));
                }

                return result;
            }
            catch (Exception ex)
            {
                Exception exception = new Exception("Make sure you set TfsURL before calling this method, and that the query is valid. Check inner exception for more details.", ex);
                throw exception;
            }
        }

        public List<string> GetAvailableProjects()
        {
            List<string> result = new List<string>();
            try
            {
                foreach (Project project in store.Projects)
                {
                    result.Add(project.Name);
                }

                return result;
            }
            catch (Exception ex)
            {
                Exception exception = new Exception("Make sure you set TfsURL before calling this method, and that the query is valid. Check inner exception for more details.", ex);
                throw exception;
            }
        }

        public List<string> GetAvailableQueryFolders(string projectName)
        {
            List<string> result = new List<string>();
            try
            {
                QueryHierarchy queryRoot = store.Projects[projectName].QueryHierarchy;

                foreach (QueryFolder folder in queryRoot)
                {
                    result.Add(folder.Name);
                }

                return result;
            }
            catch (Exception ex)
            {
                Exception exception = new Exception("Make sure you set TfsURL before calling this method, and that the query is valid. Check inner exception for more details.", ex);
                throw exception;
            }
        }

        public List<string> GetAvailableQueries(string projectName, string queryFolderName)
        {
            List<string> result = new List<string>();
            try
            {
                QueryHierarchy queryRoot = store.Projects[projectName].QueryHierarchy;
                var entity = queryRoot[queryFolderName];
                QueryFolder root = (QueryFolder)queryRoot[queryFolderName];
                var queries = GetQueries(root);
                return queries.Select(o => o.Name).ToList();
            }
            catch (Exception ex)
            {
                Exception exception = new Exception("Make sure you set TfsURL before calling this method, and that the query is valid. Check inner exception for more details.", ex);
                throw exception;
            }
        }


        private List<QueryDefinition> GetQueries(QueryFolder root)
        {
            var queries = new List<QueryDefinition>();
            foreach (var query in root)
            {
                if (typeof(QueryFolder) == query.GetType())
                {
                    queries.AddRange(GetQueries((QueryFolder)query));
                }
                else
                {
                    queries.Add((QueryDefinition)query);
                }
            }
            return queries;
        }
    }
}
