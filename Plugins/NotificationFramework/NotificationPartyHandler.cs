// <copyright file="NotificationPartyHandler.cs" company="Microsoft">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author></author>
// <date>8/27/2017 12:08:08 PM</date>
// <summary>Implementation of NotificationPartyHandler Class.</summary>
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
    /// Notification template party list handler class
    /// </summary>
    public static class NotificationPartyHandler
    {
        /// <summary>
        /// Get all template Parties
        /// </summary>
        /// <param name="notificationTemplate">notification template entity</param>
        /// <param name="primaryEntity">primary record</param>
        /// <param name="parameters">list of parameters attribute</param>
        /// <param name="partyListParam">party list parameters in form of JSON string</param>
        /// <param name="service">organization service</param>
        /// <returns>list of party list</returns>
        internal static List<PartyList> GetTemplateParties(EntityCollection notificationTemplate, Entity primaryEntity, Dictionary<string, dynamic> parameters, string partyListParam, Dictionary<string, dynamic> globalConfig, IOrganizationService service)
        {
            List<PartyList> partylist = new List<PartyList>();

            // iterate over Notification template party query entity
            foreach (Entity notification in notificationTemplate.Entities)
            {
                int participationType = 0;

                // get participation type from notification template entity
                if (notification.Attributes.Contains("NotificationParty.didd_participationtype"))
                {
                    participationType = ((OptionSetValue)((AliasedValue)notification.Attributes["NotificationParty.didd_participationtype"]).Value).Value;
                }

                List<Guid> userPartyList = new List<Guid>();
                List<Guid> teamPartyList = new List<Guid>();
                if (notification.Contains("NotificationParty.didd_notificationtemplatepartyid"))
                {
                    // add users to partylist
                    userPartyList = GetUserPartyList((Guid)((AliasedValue)notification.Attributes["NotificationParty.didd_notificationtemplatepartyid"]).Value, service);
                    AddUserToPartyList(userPartyList, partylist, participationType);

                    // add team, to party list
                    teamPartyList = GetTeamPartyList((Guid)((AliasedValue)notification.Attributes["NotificationParty.didd_notificationtemplatepartyid"]).Value, service);
                    AddTeamToPartyList(teamPartyList, partylist, participationType, service);
                }

                // linked entity attrib -if participationtype is To then create task for user
                if (notification.Attributes.Contains("NotificationParty.didd_partylistquery"))
                {
                    string query = (string)((AliasedValue)notification.Attributes["NotificationParty.didd_partylistquery"]).Value;
                    EntityCollection entColl = NotificationDataContext.GetCustomPartyList(query, primaryEntity, parameters, service);
                    AddPartyListFromCustomQuery(entColl, partylist, participationType, service);
                }

                // get From party Type 
                if (participationType == 273310003)
                {
                    GetFromPartyType(notification, globalConfig);
                }
            }

            // Add partylist from parameter
            AddPartyListFromParameters(partyListParam, partylist);

            return partylist;
        }

        /// <summary>
        /// overwrite from party type in Notification global Configuration record
        /// </summary>
        /// <param name="templateParty">Notification template party entity</param>
        /// <param name="globalConfig">global configuration</param>
        internal static void GetFromPartyType(Entity templateParty, Dictionary<string, dynamic> globalConfig)
        {
            // get user or queue from notification template party object
            EntityReference eref = null;
            if (templateParty.Contains("NotificationParty.didd_user"))
            {
                eref = (EntityReference)((AliasedValue)templateParty.Attributes["NotificationParty.didd_user"]).Value;
            }
            else if (templateParty.Contains("NotificationParty.didd_queue"))
            {
                eref = (EntityReference)((AliasedValue)templateParty.Attributes["NotificationParty.didd_queue"]).Value;
            }

            // if notification global config already contains SendEmailFrom, overwrite it with notification template party object
            if (globalConfig.ContainsKey("SendEmailFrom") && eref != null)
            {
                globalConfig.Remove("SendEmailFrom");
                globalConfig.Add("SendEmailFrom", eref);
            }
            else if (eref != null)
            {
                globalConfig.Add("SendEmailFrom", eref);
            }

        }

        /// <summary>
        /// Add party list passed from parameters
        /// </summary>
        /// <param name="partyListParamValue">party list parameters value from JSON string</param>
        /// <param name="partylist">party list object</param>
        internal static void AddPartyListFromParameters(string partyListParamValue, List<PartyList> partylist)
        {
            List<PartyList> partyListParams = (List<PartyList>)JsonHelper.GetObject(partyListParamValue, partylist.GetType());
            partylist.AddRange(partyListParams);
        }

        /// <summary>
        /// Generate PartyList for Task
        /// </summary>
        /// <param name="relatedEntities">Related entities details</param>
        /// <param name="partylist">party list</param>
        /// <param name="service">Organization service</param>
        internal static void GenerateTask(List<Entity> relatedEntities, List<PartyList> partylist, IOrganizationService service)
        {
            // Create task for each related query
            var toUserPartyList = GetToUserPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();
            foreach (Entity ent in relatedEntities)
            {
                var toUserRelatedPartyList = GetToUserRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();
                toUserRelatedPartyList.AddRange(toUserPartyList);
                toUserRelatedPartyList = toUserRelatedPartyList.GroupBy(x => x.PartyId).Select(p => p.First()).ToList<PartyList>();

                foreach (PartyList lst in toUserRelatedPartyList)
                {
                    // create task for each party
                    NotificationActivityHandler.CreateTask(ent, lst, service);
                }
            }
        }

        /// <summary>
        /// Generate PartyList when it is only intended to be create Email
        /// </summary>
        /// <param name="relatedEntities">Related entities collection</param>
        /// <param name="partylist">party list collection</param>
        /// <param name="globalConfiguration">Dictionary of notification global configuration</param>
        /// <param name="templateId">Email Template Id unique identifier</param>
        /// <param name="emailTemplateOption">Email Template option</param>
        /// <param name="service">organization service</param>
        /// <returns>Collection of email objects</returns>
        internal static EntityCollection GenerateEmail(List<Entity> relatedEntities, List<PartyList> partylist, Dictionary<string, dynamic> globalConfiguration, Guid? templateId, int? emailTemplateOption, IOrganizationService service)
        {
            // To partyList only for user entity
            var toUserPartyList = GetToUserPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();
            var carbonCopyPartyList = GetCCPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();
            var bccPartyList = GetBCCPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();

            // Create Entity Collection to send email in bulk through action calling
            EntityCollection sendEmailList = new EntityCollection();
            sendEmailList.EntityName = "email";

            var toOtherPartyList = GetToOtherPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();
            foreach (Entity ent in relatedEntities)
            {
                // Get Email object using Email Template
                Entity emailObject = GetEmailUsingTemplate(ent, templateId, emailTemplateOption, service);

                var toUserRelatedPartyList = GetToUserRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();
                var toOtherRelatedPartyList = GetToOtherRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();

                toOtherRelatedPartyList.AddRange(toUserPartyList);
                toOtherRelatedPartyList.AddRange(toUserRelatedPartyList);
                toOtherRelatedPartyList.AddRange(toOtherPartyList);
                toOtherRelatedPartyList = toOtherRelatedPartyList.GroupBy(x => x.PartyId).Select(p => p.First()).ToList<PartyList>();

                var carbonCopyRelatedPartyList = GetCCRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();
                carbonCopyRelatedPartyList.AddRange(carbonCopyPartyList);
                carbonCopyRelatedPartyList = carbonCopyRelatedPartyList.GroupBy(x => x.PartyId).Select(p => p.First()).ToList<PartyList>();

                var bccRelatedPartyList = GetBCCRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();
                bccRelatedPartyList.AddRange(bccPartyList);
                bccRelatedPartyList = bccRelatedPartyList.GroupBy(x => x.PartyId).Select(p => p.First()).ToList<PartyList>();

                Entity emailEnt = NotificationActivityHandler.CreateEmailInstance(ent, toOtherRelatedPartyList, carbonCopyRelatedPartyList, bccRelatedPartyList, null, globalConfiguration, emailObject);
                sendEmailList.Entities.Add(emailEnt);
                //// sendEmailList.Entities.Add(new Entity("email") { Id = emailId });
            }

            return sendEmailList;
            //// Call Glocal action SendEmailActivity to send bulk email
            /*if (isSendEmail)
            {
                OrganizationRequest req = new OrganizationRequest("didd_SendEmailActivity");
                req["EmailActivities"] = sendEmailList;
                service.Execute(req);
            }*/
        }

        /// <summary>
        /// Generate Email and Task both 
        /// </summary>
        /// <param name="relatedEntities">list of related entities</param>
        /// <param name="partylist">party list</param>
        /// <param name="globalConfiguration">notification global configuration object</param>
        /// <param name="templateId">Email template unique identifier</param>
        /// <param name="emailTemplateOption">email template option</param>
        /// <param name="service">organization service</param>
        /// <returns>Collection of email objects</returns>
        internal static EntityCollection GenerateEmailAndTask(List<Entity> relatedEntities, List<PartyList> partylist, Dictionary<string, dynamic> globalConfiguration, Guid? templateId, int? emailTemplateOption, IOrganizationService service)
        {
            var toUserPartyList = GetToUserPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();
            var carbonCopyPartyList = GetCCPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();
            var bccPartyList = GetBCCPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();

            var toOtherPartyList = GetToOtherPartyList(partylist, relatedEntities[0].LogicalName).ToList<PartyList>();

            // Create Entity Collection to send email in bulk through action calling
            EntityCollection sendEmailList = new EntityCollection();
            sendEmailList.EntityName = "email";

            foreach (Entity ent in relatedEntities)
            {
                // Get Email object using Email Template
                Entity emailObject = GetEmailUsingTemplate(ent, templateId, emailTemplateOption, service);

                var toUserRelatedPartyList = GetToUserRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();
                toUserRelatedPartyList.AddRange(toUserPartyList);
                toUserRelatedPartyList = toUserRelatedPartyList.GroupBy(x => x.PartyId).Select(p => p.First()).ToList<PartyList>();

                var carbonCopyRelatedPartyList = GetCCRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();
                carbonCopyRelatedPartyList.AddRange(carbonCopyPartyList);
                carbonCopyRelatedPartyList = carbonCopyRelatedPartyList.GroupBy(x => x.PartyId).Select(p => p.First()).ToList<PartyList>();

                var bccRelatedPartyList = GetBCCRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();
                bccRelatedPartyList.AddRange(bccPartyList);
                bccRelatedPartyList = bccRelatedPartyList.GroupBy(x => x.PartyId).Select(p => p.First()).ToList<PartyList>();

                var toOtherRelatedPartyList = GetToOtherRelatedPartyList(partylist, ent.LogicalName, ent.Id).ToList<PartyList>();
                toOtherRelatedPartyList.AddRange(toOtherPartyList);
                toOtherRelatedPartyList = toOtherRelatedPartyList.GroupBy(x => x.PartyId).Select(p => p.First()).ToList<PartyList>();

                foreach (PartyList lst in toUserRelatedPartyList)
                {
                    Entity task = NotificationActivityHandler.CreateTask(ent, lst, service);
                    List<PartyList> plist = new List<PartyList>();
                    plist.Add(lst);
                    Entity emailEnt = NotificationActivityHandler.CreateEmailInstance(ent, plist, carbonCopyRelatedPartyList, bccRelatedPartyList, task, globalConfiguration, emailObject);
                    sendEmailList.Entities.Add(emailEnt);
                }

                if (toOtherPartyList.Count > 0)
                {
                    Entity emailEnt = NotificationActivityHandler.CreateEmailInstance(ent, toOtherRelatedPartyList, carbonCopyRelatedPartyList, bccRelatedPartyList, null, globalConfiguration, emailObject);
                    sendEmailList.Entities.Add(emailEnt);
                }
            }

            // Call Glocal action SendEmailActivity to send bulk email
            /*if (isSendEmail)
            {
                OrganizationRequest req = new OrganizationRequest("didd_SendEmailActivity");
                req["EmailActivities"] = sendEmailList;
                service.Execute(req);
            }*/

            return sendEmailList;
        }

        /// <summary>
        /// get email object by using email instantiate object
        /// </summary>
        /// <param name="relatedEntity">related entity</param>
        /// <param name="templateId">email template unique identifier</param>
        /// <param name="emailTemplateOption">email template option</param>
        /// <param name="service">organization service</param>
        /// <returns>return email object by using email instantiate object</returns>
        private static Entity GetEmailUsingTemplate(Entity relatedEntity, Guid? templateId, int? emailTemplateOption, IOrganizationService service)
        {
            Entity email = null;
            if (templateId == null || templateId == Guid.Empty || !relatedEntity.Contains("regarding") || (emailTemplateOption != null && emailTemplateOption.Value == 273310001))
            {
                return email;
            }

            email = RenderEmailTemplate.RenderEmailTemplateData(templateId.Value, relatedEntity.Id, relatedEntity.LogicalName, service);
            return email;
        }

        /// <summary>
        /// Get To user Party List
        /// </summary>
        /// <param name="partylist">collection of all party list</param>
        /// <param name="relatedEntityName">related Entity Name</param>
        /// <returns>return user party list which is not associated to related entity</returns>
        private static IEnumerable<PartyList> GetToUserPartyList(List<PartyList> partylist, string relatedEntityName)
        {
            var toUserPartyList = partylist.Where(p => (p.ParticipationType == 273310000 && p.PartyId.LogicalName == "systemuser" && p.RelatedEntityLogicalName != relatedEntityName));
            return toUserPartyList;
        }

        /// <summary>
        /// Get To user Related Party List
        /// </summary>
        /// <param name="partylist">collection of all party list</param>
        /// <param name="relatedEntityName">related Entity Name</param>
        /// <param name="relatedEntityId">related entity id</param>
        /// <returns>return to user party list which is associated to related entity</returns>
        private static IEnumerable<PartyList> GetToUserRelatedPartyList(List<PartyList> partylist, string relatedEntityName, Guid relatedEntityId)
        {
            var toUserRelatedPartyList = partylist.Where(p => (p.ParticipationType == 273310000 && p.PartyId.LogicalName == "systemuser" && p.RelatedEntityLogicalName == relatedEntityName && p.RelatedEntityId == relatedEntityId));
            return toUserRelatedPartyList;
        }

        /// <summary>
        /// Get To party list other than user entity
        /// </summary>
        /// <param name="partylist">collection of all party list</param>
        /// <param name="relatedEntityName">related Entity Name</param>
        /// <returns>return to party list other than user entity which is not associated to related entity</returns>
        private static IEnumerable<PartyList> GetToOtherPartyList(List<PartyList> partylist, string relatedEntityName)
        {
            var toOtherPartyList = partylist.Where(p => p.ParticipationType == 273310000 && p.PartyId.LogicalName != "systemuser" && p.RelatedEntityLogicalName != relatedEntityName);
            return toOtherPartyList;
        }

        /// <summary>
        /// Get to party list other than user entity and associated to related entity
        /// </summary>
        /// <param name="partylist">collection of all party list</param>
        /// <param name="relatedEntityName">related Entity Name</param>
        /// <param name="relatedEntityId">related entity id</param>
        /// <returns>return party list other than user entity which is associated to related entity</returns>
        private static IEnumerable<PartyList> GetToOtherRelatedPartyList(List<PartyList> partylist, string relatedEntityName, Guid relatedEntityId)
        {
            var toOtherRelatedPartyList = partylist.Where(p => p.ParticipationType == 273310000 && p.PartyId.LogicalName != "systemuser" && p.RelatedEntityLogicalName == relatedEntityName && p.RelatedEntityId == relatedEntityId);
            return toOtherRelatedPartyList;
        }

        /// <summary>
        /// Get carbon copy party list and not associated to related entity
        /// </summary>
        /// <param name="partylist">collection of all party list</param>
        /// <param name="relatedEntityName">related entity name</param>
        /// <returns>return carbon copy party list which is not associated to related entity</returns>
        private static IEnumerable<PartyList> GetCCPartyList(List<PartyList> partylist, string relatedEntityName)
        {
            var partyList = partylist.Where(p => p.ParticipationType == 273310001 && p.RelatedEntityLogicalName != relatedEntityName);
            return partyList;
        }

        /// <summary>
        /// Get carbon copy party list and associated to related entity
        /// </summary>
        /// <param name="partylist">collection of all party list</param>
        /// <param name="relatedEntityName">related entity name</param>
        /// <param name="relatedEntityId">related entity id</param>
        /// <returns>return carbon copy party list which is associated to related entity</returns>
        private static IEnumerable<PartyList> GetCCRelatedPartyList(List<PartyList> partylist, string relatedEntityName, Guid relatedEntityId)
        {
            var partyList = partylist.Where(p => p.ParticipationType == 273310001 && p.RelatedEntityLogicalName == relatedEntityName && p.RelatedEntityId == relatedEntityId);
            return partyList;
        }

        /// <summary>
        /// Get Blank carbon copy party list and not associated to related entity
        /// </summary>
        /// <param name="partylist">collection of all party list</param>
        /// <param name="relatedEntityName">related entity name</param>
        /// <returns>return blank carbon copy party list which is not associated to related entity</returns>
        private static IEnumerable<PartyList> GetBCCPartyList(List<PartyList> partylist, string relatedEntityName)
        {
            var partyList = partylist.Where(p => p.ParticipationType == 273310002 && p.RelatedEntityLogicalName != relatedEntityName);
            return partyList;
        }

        /// <summary>
        /// Get Blank carbon copy party list and associated to related entity
        /// </summary>
        /// <param name="partylist">collection of all party list</param>
        /// <param name="relatedEntityName">related entity name</param>
        /// <param name="relatedEntityId">related entity id</param>
        /// <returns>return blank carbon copy party list which is associated to related entity</returns>
        private static IEnumerable<PartyList> GetBCCRelatedPartyList(List<PartyList> partylist, string relatedEntityName, Guid relatedEntityId)
        {
            var partyList = partylist.Where(p => p.ParticipationType == 273310002 && p.RelatedEntityLogicalName == relatedEntityName && p.RelatedEntityId == relatedEntityId);
            return partyList;
        }

        /// <summary>
        /// Get all user party list from Many to many relationship by executing fetch query
        /// </summary>
        /// <param name="templatePartyId">Notification template party id</param>
        /// <param name="service">organization service</param>
        /// <returns>List of all user party list</returns>
        private static List<Guid> GetUserPartyList(Guid templatePartyId, IOrganizationService service)
        {
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                              <entity name='didd_notificationtemplateparty'>
                                <attribute name='didd_notificationtemplatepartyid' />
                                <filter type='and'>
                                  <condition attribute='didd_notificationtemplatepartyid' operator='eq' value='{0}' />
                                </filter>
                                <link-entity name='didd_didd_notificationtemplateparty_systemus' from='didd_notificationtemplatepartyid' to='didd_notificationtemplatepartyid' visible='false' intersect='true' alias='sys'>
                                  <link-entity name='systemuser' from='systemuserid' to='systemuserid' alias='ab' />
	                            <attribute name='systemuserid' />
                                </link-entity>
                              </entity>
                            </fetch>";
            EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(string.Format(CultureInfo.InvariantCulture, fetchXML, templatePartyId)));
            return entColl.Entities.Select(s => (Guid)((AliasedValue)s.Attributes["sys.systemuserid"]).Value).ToList<Guid>();
        }

        /// <summary>
        /// Get all team party list by executing Many to many relationship
        /// </summary>
        /// <param name="templatePartyId">Notification template party id</param>
        /// <param name="service">organization service</param>
        /// <returns>List of all team party list</returns>
        private static List<Guid> GetTeamPartyList(Guid templatePartyId, IOrganizationService service)
        {
            string fetchXML = @"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                              <entity name='didd_notificationtemplateparty'>
                                <attribute name='didd_notificationtemplatepartyid' />
                                <filter type='and'>
                                  <condition attribute='didd_notificationtemplatepartyid' operator='eq' value='{0}' />
                                </filter>
                                <link-entity name='didd_didd_notificationtemplateparty_team' from='didd_notificationtemplatepartyid' to='didd_notificationtemplatepartyid' visible='false' intersect='true' alias='sys'>
                                  <link-entity name='team' from='teamid' to='teamid' alias='aa'>
                                    <attribute name='teamid' />
                                  </link-entity>
                                </link-entity>
                              </entity>
                            </fetch>";
            EntityCollection entColl = service.RetrieveMultiple(new FetchExpression(string.Format(CultureInfo.InvariantCulture, fetchXML, templatePartyId)));
            return entColl.Entities.Select(s => (Guid)((AliasedValue)s.Attributes["aa.teamid"]).Value).ToList<Guid>();
        }

        /// <summary>
        /// Add party list from custom Party query
        /// </summary>
        /// <param name="entColl">collection of entities retrieved from custom party query</param>
        /// <param name="partylist">collection of party list</param>
        /// <param name="participationType">participation type</param>
        /// <param name="service">organization service</param>
        private static void AddPartyListFromCustomQuery(EntityCollection entColl, List<PartyList> partylist, int participationType, IOrganizationService service)
        {
            foreach (Entity ent in entColl.Entities)
            {
                if (ent.LogicalName == "systemuser")
                {
                    // If entity collection is of type system user
                    EntityReference user = new EntityReference(ent.LogicalName, ent.Id);
                    partylist.Add(new PartyList() { ParticipationType = participationType, PartyId = user, RelatedEntityId = user.Id, RelatedEntityLogicalName = user.LogicalName });
                }
                else if (ent.LogicalName == "team")
                {
                    // if entity collection type is of Team
                    EntityReference team = new EntityReference(ent.LogicalName, ent.Id);
                    AddTeamMembertoPartyList(team, null, partylist, participationType, service);
                }
                else
                {
                    if (ent.Attributes.Count == 1 && ent.Attributes.Where(a => a.Value.GetType().ToString() == "System.Guid").Count() == 1)
                    {
                        // if Custom query has only primary guid attribute
                        EntityReference eref = new EntityReference(ent.LogicalName, ent.Id);
                        partylist.Add(new PartyList() { ParticipationType = participationType, PartyId = eref, RelatedEntityId = eref.Id, RelatedEntityLogicalName = eref.LogicalName });
                    }
                    else
                    {
                        // seperate out Entityreference for normal and Aliased type of attributes
                        var attribute = ent.Attributes.Where(a => (a.Value.GetType().ToString() == "Microsoft.Xrm.Sdk.EntityReference"
                                || (a.Value.GetType().ToString() == "Microsoft.Xrm.Sdk.AliasedValue"
                                && ((AliasedValue)a.Value).Value.GetType().ToString() == "Microsoft.Xrm.Sdk.EntityReference")));

                        foreach (KeyValuePair<string, object> val in attribute)
                        {
                            // get only entity reference type attribute
                            EntityReference eref = null;
                            if (val.Value.GetType().ToString() == "Microsoft.Xrm.Sdk.EntityReference")
                            {
                                eref = (EntityReference)val.Value;
                            }
                            else
                            {
                                eref = (EntityReference)((AliasedValue)val.Value).Value;
                            }

                            if (eref.LogicalName == "team")
                            {
                                // add team member to party list
                                AddTeamMembertoPartyList(eref, ent, partylist, participationType, service);
                            }
                            else
                            {
                                // add user and other entity attribute to partylist
                                partylist.Add(new PartyList { ParticipationType = participationType, PartyId = eref, RelatedEntityId = ent.Id, RelatedEntityLogicalName = ent.LogicalName });
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Add user to party list
        /// </summary>
        /// <param name="partyListUser">user party list</param>
        /// <param name="partylist">party list</param>
        /// <param name="participationType">participation type</param>
        private static void AddUserToPartyList(List<Guid> partyListUser, List<PartyList> partylist, int participationType)
        {
            foreach (Guid id in partyListUser)
            {
                EntityReference partyid = new EntityReference("systemuser", id);
                partylist.Add(new PartyList() { ParticipationType = participationType, PartyId = partyid, RelatedEntityId = id, RelatedEntityLogicalName = partyid.LogicalName });
            }
        }

        /// <summary>
        /// Add team to party list
        /// </summary>
        /// <param name="partyListTeam">party list for team</param>
        /// <param name="partylist">party list</param>
        /// <param name="participationType">participation type</param>
        /// <param name="service">organization service</param>
        private static void AddTeamToPartyList(List<Guid> partyListTeam, List<PartyList> partylist, int participationType, IOrganizationService service)
        {
            foreach (Guid id in partyListTeam)
            {
                EntityReference team = new EntityReference("team", id);
                AddTeamMembertoPartyList(team, null, partylist, participationType, service);
            }
        }

        /// <summary>
        /// Add member of team to party list
        /// </summary>
        /// <param name="team">team object</param>
        /// <param name="relatedEntity">related entity list</param>
        /// <param name="partylist">party list</param>
        /// <param name="participationType">participation type</param>
        /// <param name="service">organization service</param>
        private static void AddTeamMembertoPartyList(EntityReference team, Entity relatedEntity, List<PartyList> partylist, int participationType, IOrganizationService service)
        {
            // get all team member
            EntityCollection teamMember = GetTeamMember(team, service);

            // iterate over team member and add it to party list
            foreach (Entity member in teamMember.Entities)
            {
                EntityReference partyid = new EntityReference("systemuser", (Guid)member.Attributes["systemuserid"]);

                if (relatedEntity == null)
                {
                    partylist.Add(new PartyList() { ParticipationType = participationType, PartyId = partyid, RelatedEntityId = team.Id, RelatedEntityLogicalName = team.LogicalName });
                }
                else
                {
                    partylist.Add(new PartyList() { ParticipationType = participationType, PartyId = partyid, RelatedEntityId = relatedEntity.Id, RelatedEntityLogicalName = relatedEntity.LogicalName });
                }
            }
        }

        /// <summary>
        /// Fetch query to get all team member
        /// </summary>
        /// <param name="team">team object</param>
        /// <param name="service">organization service</param>
        /// <returns>return all team members collections</returns>
        private static EntityCollection GetTeamMember(EntityReference team, IOrganizationService service)
        {
            string teamMembershipQuery = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>" +
                                          "<entity name='systemuser'>" +
                                            "<attribute name='systemuserid' />" +
                                            "<attribute name='fullname' />" +
                                            "<link-entity name='teammembership' from='systemuserid' to='systemuserid' visible='false' intersect='true'>" +
                                              "<link-entity name='team' from='teamid' to='teamid' alias='ab'>" +
                                                "<filter type='and'>" +
                                                  "<condition attribute='teamid' operator='eq' value='" + team.Id + "' />" +
                                                "</filter>" +
                                              "</link-entity>" +
                                            "</link-entity>" +
                                          "</entity>" +
                                        "</fetch>";

            FetchExpression fetch = new FetchExpression(teamMembershipQuery);
            EntityCollection entColl = service.RetrieveMultiple(fetch);
            return entColl;
        }
    }
}
