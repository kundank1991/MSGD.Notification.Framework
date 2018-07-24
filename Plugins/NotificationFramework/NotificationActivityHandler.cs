// <copyright file="NotificationActivityHandler.cs" company="Microsoft">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author></author>
// <date>8/27/2017 12:08:08 PM</date>
// <summary>Implementation of NotificationActivityHandler Class.</summary>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;

    /// <summary>
    /// Notification Activity Handler
    /// </summary>
    public static class NotificationActivityHandler
    {
        /// <summary>
        /// Create Email
        /// </summary>
        /// <param name="relatedEntity">Related entity</param>
        /// <param name="toParty">Email to</param>
        /// <param name="carbonCopyParty">Email carbon copy</param>
        /// <param name="bccParty">Email blank carbon copy</param>
        /// <param name="task">created task entity</param>
        /// <param name="globalConfiguration">global config details</param>
        /// <param name="emailObject">email object</param>
        /// <returns>created email record id</returns>
        internal static Entity CreateEmailInstance(Entity relatedEntity, List<PartyList> toParty, List<PartyList> carbonCopyParty, List<PartyList> bccParty, Entity task, Dictionary<string, dynamic> globalConfiguration, Entity emailObject)
        {
            // initialize email entity components
            var fromActivityParty = new EntityCollection();
            var toActivityParty = new EntityCollection();
            var carbonCopyActivityParty = new EntityCollection();
            var bccActivityParty = new EntityCollection();

            // get email instance
            Entity email = null;
            if (emailObject == null)
            {
                email = new Entity("email");
            }
            else
            {
                email = emailObject;
            }

            // Add email subject
            if (relatedEntity.Contains("emailsubject"))
            {
                email.Attributes["subject"] = (string)relatedEntity.Attributes["emailsubject"];
            }

            // Add From 
            Entity activityParty = new Entity("activityparty");
            activityParty.Attributes["partyid"] = (EntityReference)globalConfiguration["SendEmailFrom"];
            fromActivityParty.Entities.Add(activityParty);

            // Add to in activity party list
            foreach (PartyList lst in toParty)
            {
                Entity activityparty = new Entity("activityparty");
                activityparty.Attributes["partyid"] = lst.PartyId;
                toActivityParty.Entities.Add(activityparty);
            }

            // Add cc in activity party list
            foreach (PartyList lst in carbonCopyParty)
            {
                Entity activityparty = new Entity("activityparty");
                activityparty.Attributes["partyid"] = lst.PartyId;
                carbonCopyActivityParty.Entities.Add(activityparty);
            }

            // add bcc in activity party list
            foreach (PartyList lst in bccParty)
            {
                Entity activityparty = new Entity("activityparty");
                activityparty.Attributes["partyid"] = lst.PartyId;
                bccActivityParty.Entities.Add(activityparty);
            }

            if (toActivityParty.Entities.Count > 0)
            {
                email.Attributes["to"] = toActivityParty;
            }

            if (carbonCopyActivityParty.Entities.Count > 0)
            {
                email.Attributes["cc"] = carbonCopyActivityParty;
            }

            if (bccActivityParty.Entities.Count > 0)
            {
                email.Attributes["bcc"] = bccActivityParty;
            }

            if (fromActivityParty.Entities.Count > 0)
            {
                email.Attributes["from"] = fromActivityParty;
            }

            // add email description
            if (task == null && relatedEntity.Contains("emaildescription"))
            {
                email.Attributes["description"] = (string)relatedEntity.Attributes["emaildescription"];
            }
            else if (task != null && relatedEntity.Contains("emaildescription"))
            {
                // if task is associated with email, then add task URL
                string taskURL = NotificationTokenReplacer.ReplaceURL(task, (string)globalConfiguration["RecordBaseURL"]);
                string emailDescription = (string)relatedEntity.Attributes["emaildescription"];
                emailDescription = emailDescription.Replace("{!%URL}", taskURL);
                email.Attributes["description"] = emailDescription;
            }

            // add regarding
            if (relatedEntity.Contains("regarding"))
            {
                email.Attributes["regardingobjectid"] = (EntityReference)relatedEntity.Attributes["regarding"];
            }

            // Guid emailId = service.Create(email);
            return email;
        }

        /// <summary>
        /// Create Task entity
        /// </summary>
        /// <param name="relatedEntity">related Entity</param>
        /// <param name="party">Task Party list</param>
        /// <param name="service">organization service</param>
        /// <returns>created task entity</returns>
        internal static Entity CreateTask(Entity relatedEntity, PartyList party, IOrganizationService service)
        {
            Entity task = new Entity("task");

            // Add task subject
            if (relatedEntity.Contains("tasksubject"))
            {
                task.Attributes["subject"] = (string)relatedEntity.Attributes["tasksubject"];
            }

            // add task due date
            if (relatedEntity.Contains("taskduedate"))
            {
                task.Attributes["scheduledend"] = (DateTime)relatedEntity.Attributes["taskduedate"];
            }

            // add tasl description
            if (relatedEntity.Contains("taskdescription"))
            {
                task.Attributes["description"] = (string)relatedEntity.Attributes["taskdescription"];
            }

            // add task regarding
            if (relatedEntity.Contains("regarding"))
            {
                task.Attributes["regardingobjectid"] = (EntityReference)relatedEntity.Attributes["regarding"];
            }

            // specify ownership of task
            if (party.PartyId != null)
            {
                task.Attributes["ownerid"] = party.PartyId;
            }

            Guid taskid = service.Create(task);
            task.Id = taskid;
            return task;
        }
    }
}
