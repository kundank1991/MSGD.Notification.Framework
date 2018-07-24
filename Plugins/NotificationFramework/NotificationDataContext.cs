// <copyright file="NotificationdataContext.cs" company="Microsoft">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author></author>
// <date>8/27/2017 12:08:08 PM</date>
// <summary>Implementation of NotificationdataContext Class.</summary>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;

    /// <summary>
    /// Class to create Notification Data Context
    /// </summary>
    public static class NotificationDataContext
    {
        /// <summary>
        /// Execute FetchXMl query to retrieve related data
        /// </summary>
        /// <param name="relatedQuery">Related query</param>
        /// <param name="primaryEntity">primary entity</param>
        /// <param name="parameters">list of parameters attribute</param>
        /// <param name="filter">list of filter which need to apply</param>
        /// <param name="service">organization service</param>
        /// <returns>collection of related entities</returns>
        internal static EntityCollection RetrieveRelatedEntity(Entity relatedQuery, Entity primaryEntity, Dictionary<string, dynamic> parameters, List<QueryFilter> filter, IOrganizationService service)
        {
            if (relatedQuery.Contains("didd_query") && relatedQuery.Contains("didd_sequencenumber"))
            {
                // Query string in from of fetchXMl
                string queryString = Convert.ToString(relatedQuery.Attributes["didd_query"], CultureInfo.InvariantCulture);

                // Replace Placeholder in query string
                string query = NotificationTokenReplacer.ReplacePlaceHolders(null, primaryEntity, parameters, string.Empty, queryString);

                // get filter to apply for query string
                var queryFilter = filter.Where(x => x.SequenceNumber == (int)relatedQuery.Attributes["didd_sequencenumber"]).OrderBy(o => o.SubsequenceNumber).ToList<QueryFilter>();

                // add filter in query string
                List<string> param = new List<string>();
                foreach (var lst in queryFilter)
                {
                    param.Add(lst.Value);
                }

                // Execute query string
                FetchExpression fetchQuery = new FetchExpression(string.Format(CultureInfo.InvariantCulture, query, param.ToArray<string>()));
                return service.RetrieveMultiple(fetchQuery);
            }
            else
            {
                throw new InvalidPluginExecutionException(string.Format(CultureInfo.InvariantCulture, "Notification Template Data Query is not configured properly. Please contact administrator."));
            }
        }

        /// <summary>
        /// Return Notification template Master Data
        /// </summary>
        /// <param name="templateId">Notification Template Guid</param>
        /// <param name="service">Organization service</param>
        /// <returns>Notification Template Entity</returns>
        internal static EntityCollection RetrieveNotificationTemplate(Guid templateId, IOrganizationService service)
        {
            // Set the condition for the retrieval  based on Notification template Guid
            ConditionExpression condition = new ConditionExpression("didd_notificationtemplateid", ConditionOperator.Equal, templateId);

            FilterExpression filter = new FilterExpression();
            filter.Conditions.Add(condition);
            QueryExpression query = new QueryExpression("didd_notificationtemplate");

            // notification party Link Entity
            query.LinkEntities.Add(new LinkEntity("didd_notificationtemplate", "didd_notificationtemplateparty", "didd_notificationtemplateid", "didd_notificationtemplate", JoinOperator.Inner));
            query.LinkEntities[0].Columns.AddColumns("didd_partylistquery", "didd_participationtype", "didd_notificationtemplatepartyid", "didd_user", "didd_queue");
            query.LinkEntities[0].EntityAlias = "NotificationParty";
            query.LinkEntities[0].JoinOperator = JoinOperator.LeftOuter;

            // add Fiter Criteria
            query.Criteria = filter;

            // Query Column set
            string[] columnValue = new string[] 
                                            { 
                                                "didd_notificationtemplateid", "didd_name", "didd_taskdescription", 
                                                "didd_task", "didd_subject", "statuscode", "statecode", "didd_sendemail", 
                                                "didd_regardingentity", "didd_primaryentity", "didd_emaildescription", "didd_email", 
                                                "didd_duedatemodifierattribure", "didd_duedatemodifier", "didd_duration", "didd_duedate",
                                                "didd_emailtemplate", "didd_emailtemplateoption"
                                            };
            ColumnSet cols = new ColumnSet(columnValue);
            query.ColumnSet = cols;
            query.NoLock = true;
            return service.RetrieveMultiple(query);
        }

        /// <summary>
        /// Get Data of Notification template query entity
        /// </summary>
        /// <param name="templateId">Notification Template Id</param>
        /// <param name="service">organization service</param>
        /// <returns>Return Notification Template Query Entity Data</returns>
        internal static EntityCollection RetrieveRelatedQueries(Guid templateId, IOrganizationService service)
        {
            // Get Deta of Notification template query entity
            QueryExpression query = new QueryExpression("didd_notificationtemplatequery");
            ColumnSet columns = new ColumnSet("didd_query", "didd_notificationtemplate", "didd_name", "didd_sequencenumber");
            query.ColumnSet = columns;

            FilterExpression filter = new FilterExpression();
            ConditionExpression condition = new ConditionExpression("didd_notificationtemplate", ConditionOperator.Equal, templateId);
            filter.Conditions.Add(condition);

            query.Criteria = filter;
            query.NoLock = true;
            return service.RetrieveMultiple(query);
        }

        /// <summary>
        /// Get Party List through custom query
        /// </summary>
        /// <param name="customQuery">custom query</param>
        /// <param name="primaryEntity">primary record</param>
        /// <param name="parameters">list of parameters attribute</param>
        /// <param name="service">organization service</param>
        /// <returns>collection of custom party</returns>
        internal static EntityCollection GetCustomPartyList(string customQuery, Entity primaryEntity, Dictionary<string, dynamic> parameters, IOrganizationService service)
        {
            // execute fetch query dynamically for custom party list
            string fetchPartyListQuery = NotificationTokenReplacer.ReplacePlaceHolders(null, primaryEntity, parameters, null, customQuery);
            FetchExpression fetchQuery = new FetchExpression(fetchPartyListQuery);
            return service.RetrieveMultiple(fetchQuery);
        }

        /// <summary>
        /// Get Primary Entity details
        /// </summary>
        /// <param name="entityLogicalName">primary entity logical name</param>
        /// <param name="primaryEntityId">primary entity id</param>
        /// <param name="service">organization service</param>
        /// <returns>primary entity</returns>
        internal static Entity GetPrimaryEntity(string entityLogicalName, Guid primaryEntityId, IOrganizationService service)
        {
            // Get primary entity details
            Entity entity = service.Retrieve(entityLogicalName, primaryEntityId, new ColumnSet(true));
            return entity;
        }

        /// <summary>
        /// Get global data from notification template entity
        /// </summary>
        /// <param name="service">organization service</param>
        /// <returns>dictionary of global data in key value pair</returns>
        internal static Dictionary<string, dynamic> GetNotificationGlobalData(IOrganizationService service)
        {
            // extract all global data
            QueryExpression query = new QueryExpression("didd_notificationglobalconfiguration");
            query.ColumnSet = new ColumnSet("didd_notificationglobalconfigurationid", "didd_value", "didd_name", "didd_user", "didd_description");
            query.NoLock = true;
            EntityCollection entCollection = service.RetrieveMultiple(query);

            // arrange Notification global data in form of string
            Dictionary<string, dynamic> globaldata = new Dictionary<string, dynamic>();
            foreach (Entity entity in entCollection.Entities)
            {
                string name = string.Empty;
                if (entity.Contains("didd_name"))
                {
                    name = (string)entity.Attributes["didd_name"];
                }

                // add value
                if (entity.Contains("didd_value"))
                {
                    globaldata.Add(name, (string)entity.Attributes["didd_value"]);
                }

                // add user
                if (entity.Contains("didd_user"))
                {
                    globaldata.Add(name, (EntityReference)entity.Attributes["didd_user"]);
                }
            }

            return globaldata;
        }

        /// <summary>
        /// Get Notification Template Id based on Notification  template Name
        /// </summary>
        /// <param name="templateName">Template name</param>
        /// <param name="service">organization service</param>
        /// <returns>Guid of notification template</returns>
        internal static Guid GetNotificationTemplateId(string templateName, IOrganizationService service)
        {
            Guid templateId = Guid.Empty;
            QueryExpression query = new QueryExpression("didd_notificationtemplate");
            query.ColumnSet = new ColumnSet("didd_notificationtemplateid");
            query.Criteria.AddCondition("didd_name", ConditionOperator.Equal, templateName);
            query.NoLock = true;

            EntityCollection entCollection = service.RetrieveMultiple(query);

            if (entCollection != null && entCollection.Entities.Count == 1)
            {
                templateId = entCollection.Entities[0].Id;
                return templateId;
            }
            else
            {
                throw new InvalidPluginExecutionException(string.Format(CultureInfo.InvariantCulture, "Error in getting Notification Template Record. Please COntact Administrator."));
            }
        }

        /// <summary>
        /// Return Filter Criteria list value
        /// </summary>
        /// <param name="criteria">Criteria string</param>
        /// <returns>Return Filter Criteria list</returns>
        internal static List<QueryFilter> GetFilterCriteria(string criteria)
        {
            // Filter Criteria format
            // didd_ispopeneddate=06-01-2017&didd_enddate=12-31-2017|didd_ispopeneddate=08-01-2017&didd_enddate=12-31-2017
            List<QueryFilter> lst = new List<QueryFilter>();

            if (string.IsNullOrEmpty(criteria))
            {
                return lst;
            }

            // split filter string for different query with '|'
            var sequence = criteria.Split('|');
            for (int i = 0; i < sequence.Length; i++)
            {
                // split filter string for different Parameter with '&'
                var str = sequence[i].Split('&');
                for (int j = 0; j < str.Length; j++)
                {
                    // split filter string for different key value pair with '='
                    var substr = str[j].Split('=');
                    QueryFilter filter = new QueryFilter();
                    filter.SequenceNumber = i;
                    filter.SubsequenceNumber = j;
                    filter.Key = substr[0];
                    filter.Value = substr[1];
                    lst.Add(filter);
                }
            }

            return lst;
        }
    }
}
