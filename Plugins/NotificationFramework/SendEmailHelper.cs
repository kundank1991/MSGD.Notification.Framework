// <copyright file="SendEmailHelper.cs" company="Microsoft">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author></author>
// <date>8/27/2017 12:08:08 PM</date>
// <summary>Implementation of Send Email Activity plugin</summary>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.Globalization;
    using System.ServiceModel;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Query;

    /// <summary>
    /// Send Email helper class for notification Framework
    /// </summary>
    public static class SendEmailHelper
    {
        /// <summary>
        /// Send Email in Bulk
        /// </summary>
        /// <param name="context">execution context</param>
        /// <param name="service">organization service</param>
        /// <param name="tracing">tracing service</param>
        public static void SendEmail(IExecutionContext context, IOrganizationService service, ITracingService tracing)
        {
            if (tracing == null || context == null || service == null)
            {
                return;
            }

            tracing.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: Entering SendEmail {1}", "SendEmailHelper", DateTime.Now));

            // Get Email Entity Collection from Action Input parameter
            EntityCollection emailCollection = (EntityCollection)context.InputParameters["EmailActivities"];

            /*ExecuteMultipleRequest bulkEmailSendRequest = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses. 
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = true,
                    ReturnResponses = true
                },

                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };*/

            // add send Email Request
            foreach (Entity ent in emailCollection.Entities)
            {
                SendEmailRequest sendEmail = new SendEmailRequest()
                {
                    EmailId = ent.Id,
                    TrackingToken = string.Empty,
                    IssueSend = true
                };

                // bulkEmailSendRequest.Requests.Add(sendEmail);
                service.Execute(sendEmail);
            }

            // send bulk email
            // service.Execute(bulkEmailSendRequest);
            tracing.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: Exiting SendEmail {1}", "SendEmailHelper", DateTime.Now));
        }
    }
}
