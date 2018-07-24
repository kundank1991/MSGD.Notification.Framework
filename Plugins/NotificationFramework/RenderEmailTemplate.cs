// <copyright file="RenderEmailTemplate.cs" company="Microsoft">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author></author>
// <date>8/27/2017 12:08:08 PM</date>
// <summary>Implementation of RenderEmailTemplate Class.</summary>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Xml.Serialization;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;

    /// <summary>
    /// Render Email Template class
    /// </summary>
    public static class RenderEmailTemplate
    {
        /// <summary>
        /// Render email template data
        /// </summary>
        /// <param name="templateId">template id</param>
        /// <param name="primaryEntityId">primary entity id</param>
        /// <param name="primaryEntityName">primary entity name</param>
        /// <param name="service">organization service</param>
        /// <returns>rendered email object</returns>
        internal static Entity RenderEmailTemplateData(Guid templateId, Guid primaryEntityId, string primaryEntityName, IOrganizationService service)
        {
            // Instantiate Email template and return Email object
            InstantiateTemplateRequest instTemplateReq = new InstantiateTemplateRequest
            {
                TemplateId = templateId,
                ObjectId = primaryEntityId,
                ObjectType = primaryEntityName
            };
            InstantiateTemplateResponse instTemplateResp = (InstantiateTemplateResponse)service.Execute(instTemplateReq);
            return instTemplateResp.EntityCollection[0];
        }

        /// <summary>
        /// Get Email Template unique identifier 
        /// </summary>
        /// <param name="templateTitle">Email template name</param>
        /// <param name="service">organization service</param>
        /// <returns>return Email Template unique identifier</returns>
        internal static Guid GetTemplateId(string templateTitle, IOrganizationService service)
        {
            // Get Email Template unique identifier 
            QueryExpression query = new QueryExpression("template");
            query.ColumnSet = new ColumnSet("templateid", "title");
            FilterExpression filter = new FilterExpression(LogicalOperator.And);
            filter.Conditions.Add(new ConditionExpression("title", ConditionOperator.Equal, templateTitle));
            filter.Conditions.Add(new ConditionExpression("templatetypecode", ConditionOperator.Equal, "systemuser"));
            query.Criteria = filter;
            EntityCollection entCollection = service.RetrieveMultiple(query);
            if (entCollection != null && entCollection.Entities.Count != 1)
            {
                throw new InvalidPluginExecutionException("Invalid Email Template.");
            }

            return (Guid)entCollection[0].Attributes["templateid"];
        }

        /// <summary>
        /// Create and send email
        /// </summary>
        /// <param name="entityCollection">email object entity collection</param>
        /// <param name="isSendEmail">is email need to be send or not</param>
        /// <param name="templateId">email template unique identifier</param>
        /// <param name="emailTemplateOption">email template option</param>
        /// <param name="service">organization service</param>
        internal static void ProcessEmail(EntityCollection entityCollection, bool isSendEmail, Guid? templateId, int? emailTemplateOption, IOrganizationService service)
        {
            // if email is not need to be sent using email template
            if ((templateId == null) || (templateId != null && emailTemplateOption != null && emailTemplateOption.Value == 273310000))
            {
                foreach (Entity email in entityCollection.Entities)
                {
                    var emailid = service.Create(email);
                    email.Id = emailid;
                }

                // Call Send email activity action to Send email 
                if (isSendEmail)
                {
                    OrganizationRequest req = new OrganizationRequest("didd_SendEmailActivity");
                    req["EmailActivities"] = entityCollection;
                    service.Execute(req);
                }

                return;
            }

            // send email using OOB email template
            foreach (Entity email in entityCollection.Entities)
            {
                EntityReference regardingObject = (EntityReference)email.Attributes["regardingobjectid"];
                SendEmailUsingTemplate(email, templateId.Value, regardingObject.Id, regardingObject.LogicalName, service);
            }
        }

        /// <summary>
        /// send email using email template
        /// </summary>
        /// <param name="email">email object</param>
        /// <param name="templateId">email template unique identifier</param>
        /// <param name="primaryEntityId">primary entity id</param>
        /// <param name="primaryEntityName">primary entity logical name</param>
        /// <param name="service">organization service</param>
        internal static void SendEmailUsingTemplate(Entity email, Guid templateId, Guid primaryEntityId, string primaryEntityName, IOrganizationService service)
        {
            SendEmailFromTemplateRequest emailUsingTemplateReq = new SendEmailFromTemplateRequest
            {
                Target = email,

                // Use a built-in Email Template of type "contact".
                TemplateId = templateId,

                // The regarding Id is required, and must be of the same type as the Email Template.
                RegardingId = primaryEntityId,
                RegardingType = primaryEntityName
            };

            service.Execute(emailUsingTemplateReq);
        }
    }
}
