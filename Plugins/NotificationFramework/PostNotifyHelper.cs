// <copyright file="PostNotifyHelper.cs" company="Microsoft Corporation">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Microsoft Corporation</author>
// <date>6/14/2017 7:34:28 AM</date>
// <summary>Helper class of PostNotify Plugin.</summary>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using System.Text;
    using System.Text.RegularExpressions;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;

    /// <summary>
    /// Notification Framework Helper class
    /// </summary>
    public class PostNotifyHelper
    {
        /// <summary>
        /// Tracing service
        /// </summary>
        private static ITracingService tracingService;

        /// <summary>
        /// Notification Global Configuration data
        /// </summary>
        private Dictionary<string, dynamic> globalConfiguration;

        /// <summary>
        /// Send Notifications to stake holders based on Configuration data
        /// </summary>
        /// <param name="pluginContext">execution context</param>
        /// <param name="service">organization service</param>
        /// <param name="tracing">tracing service</param>
        public void SendNotifications(IExecutionContext pluginContext, IOrganizationService service, ITracingService tracing)
        {
            if (tracing == null || pluginContext == null || service == null)
            {
                return;
            }

            tracing.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: Entering NotifyStakeHolders.", "PostNotifyHelper"));
            tracingService = tracing;

            // get Notification global configuration data
            this.globalConfiguration = NotificationDataContext.GetNotificationGlobalData(service);

            // Extracting Input parameter from Action
            Guid recordId = new Guid((string)pluginContext.InputParameters["EntityRecordId"]);
            string filterCondition = (string)pluginContext.InputParameters["FilterCondition"];
            string templateValue = (string)pluginContext.InputParameters["TemplateId"];
            string entityLogicalName = (string)pluginContext.InputParameters["EntityLogicalName"];
            string parameterCollections = (string)pluginContext.InputParameters["ParameterCollections"];
            string partyListParameters = (string)pluginContext.InputParameters["PartyListParameters"];

            if (string.IsNullOrEmpty(parameterCollections))
            {
                parameterCollections = "[]";
            }

            if (string.IsNullOrEmpty(partyListParameters))
            {
                partyListParameters = "[]";
            }

            tracingService.Trace("Party List Parameters :- " + partyListParameters);
            Dictionary<string, dynamic> parameter = (Dictionary<string, dynamic>)JsonHelper.GetObject(parameterCollections, new Dictionary<string, dynamic>().GetType());
            Guid templateId = NotificationDataContext.GetNotificationTemplateId(templateValue, service);

            List<QueryFilter> filter = NotificationDataContext.GetFilterCriteria(filterCondition);

            Entity primaryEntity = NotificationDataContext.GetPrimaryEntity(entityLogicalName, recordId, service);

            // Get Notification Template Entity data
            EntityCollection notificationTemplate = NotificationDataContext.RetrieveNotificationTemplate(templateId, service);

            // Get Notification Template Query Entity data
            EntityCollection relatedQueries = NotificationDataContext.RetrieveRelatedQueries(templateId, service);
            int countQueries = relatedQueries.Entities.Count;
            List<Entity> relatedEntities = new List<Entity>();

            // if no related entity is found, work on primary entity
            if (countQueries == 0)
            {
                relatedQueries = new EntityCollection();
                relatedQueries.EntityName = entityLogicalName;
                relatedQueries.Entities.Add(primaryEntity);
                relatedEntities.AddRange(relatedQueries.Entities);
            }

            // Execute Related Queries
            for (int i = 0; i < countQueries; i++)
            {
                var relatedEntityColl = NotificationDataContext.RetrieveRelatedEntity(relatedQueries.Entities[i], primaryEntity, parameter, filter, service);

                if (relatedEntityColl.Entities.Count > 0)
                {
                    relatedEntities.AddRange(relatedEntityColl.Entities);
                }
            }

            // remove duplicates from related Query
            relatedEntities = relatedEntities.GroupBy(x => x.Id).Select(r => r.First()).ToList<Entity>();
            Guid? emailTemplateId = null;
            int? emailTemplateOption = null;

            if (notificationTemplate[0].Contains("didd_emailtemplate") && notificationTemplate[0].Contains("didd_emailtemplateoption"))
            {
                emailTemplateId = RenderEmailTemplate.GetTemplateId((string)notificationTemplate[0].Attributes["didd_emailtemplate"], service);
                emailTemplateOption = ((OptionSetValue)notificationTemplate[0].Attributes["didd_emailtemplateoption"]).Value;
            }

            if (relatedEntities == null || relatedEntities.Count == 0)
            {
                // no relatedentities records found
                return;
            }

            // Get setup data from notification template
            bool isTask = (bool)notificationTemplate.Entities[0].Attributes["didd_task"];
            bool isEmail = (bool)notificationTemplate.Entities[0].Attributes["didd_email"];
            bool isSendEmail = (bool)notificationTemplate.Entities[0].Attributes["didd_sendemail"];

            // Replace token in notification template 
            NotificationTokenReplacer ndr = new NotificationTokenReplacer();
            ndr.ReplaceTemplateData(primaryEntity, relatedEntities, parameter, notificationTemplate.Entities[0], this.globalConfiguration["RecordBaseURL"]);

            // Get all party List defined in notification template party
            List<PartyList> partylist = NotificationPartyHandler.GetTemplateParties(notificationTemplate, primaryEntity, parameter, partyListParameters, this.globalConfiguration, service);

            if (isTask && isEmail)
            {
                // Create Email and task both and embed link of task into email
                EntityCollection entColl = NotificationPartyHandler.GenerateEmailAndTask(relatedEntities, partylist, this.globalConfiguration, emailTemplateId, emailTemplateOption, service);
                RenderEmailTemplate.ProcessEmail(entColl, isSendEmail, emailTemplateId, emailTemplateOption, service);
            }
            else if (isTask && !isEmail)
            {
                // Create task
                NotificationPartyHandler.GenerateTask(relatedEntities, partylist, service);
            }
            else if (!isTask && isEmail)
            {
                // Create Email
                EntityCollection entColl = NotificationPartyHandler.GenerateEmail(relatedEntities, partylist, this.globalConfiguration, emailTemplateId, emailTemplateOption, service);
                RenderEmailTemplate.ProcessEmail(entColl, isSendEmail, emailTemplateId, emailTemplateOption, service);
            }

            tracingService.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: Exiting NotifyStakeHolders.", "PostNotifyHelper"));
        }
    }
}