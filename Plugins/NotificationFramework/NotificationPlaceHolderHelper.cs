using Microsoft.Xrm.Sdk;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MSGD.Notification.Framework.Plugins
{
    public class NotificationPlaceHolderHelper
    {
        /// <summary>
        /// Gets or sets Task Description
        /// </summary>
        private string TaskDescription { get; set; }

        /// <summary>
        /// Gets or sets Email Description
        /// </summary>
        private string EmailDescripiton { get; set; }

        /// <summary>
        /// Gets or sets Subject
        /// </summary>
        private string Subject { get; set; }

        /// <summary>
        /// Gets or sets Due Date
        /// </summary>
        private DateTime? DueDate { get; set; }

        /// <summary>
        /// Gets or sets Duration
        /// </summary>
        private long? Duration { get; set; }

        /// <summary>
        /// Gets or sets Due Date Modifier
        /// </summary>
        private int? DueDateModifier { get; set; }

        /// <summary>
        /// Gets or sets Due Date Modifier Attribute
        /// </summary>
        private string DueDateModifierAttribute { get; set; }

        /// <summary>
        /// Gets or sets RegardingEntity Attribute
        /// </summary>
        private string RegardingEntity { get; set; }

        private ITracingService tracingService { get; set; }

        public void ReplaceTokenData(IPluginExecutionContext context, IOrganizationService service, ITracingService tracing)
        {
            tracingService = tracing;
            EntityCollection relatedEntityCollection = (EntityCollection)context.InputParameters["RelatedEntityCollection"];
            Entity notificationTemplate = (Entity)context.InputParameters["NotificationTemplate"];
            Entity primaryEntity = (Entity)context.InputParameters["PrimaryEntity"];
            string parameters = (string)context.InputParameters["Parameters"];
            string recordBaseURL = (string)context.InputParameters["RecordBaseURL"];

            tracing.Trace("parameters are ->  " + parameters);

            Dictionary<string, dynamic> parameterDictionary = (Dictionary<string, dynamic>)JsonHelper.GetObject(parameters, new Dictionary<string, dynamic>().GetType());
            ReplaceTemplateData(primaryEntity, relatedEntityCollection, parameterDictionary, notificationTemplate, recordBaseURL);

            context.OutputParameters["RelatedEntityCollectionOut"] = relatedEntityCollection;
        }

        /// <summary>
        /// Replace all string place holders in Template data
        /// </summary>
        /// <param name="relatedEntity">related Entity object</param>
        /// <param name="primaryEntity">primary entity object</param>
        /// <param name="parameters">list of parameter attribute</param>
        /// <param name="recordBaseUrl">record base URL</param>
        /// <param name="dataToReplace">actual template data which need to be replaced</param>
        /// <returns>return replaced data</returns>
        internal static string ReplacePlaceHolders(Entity relatedEntity, Entity primaryEntity, Dictionary<string, dynamic> parameters, string recordBaseUrl, string dataToReplace)
        {
            // Regex Match format. specified format is {entityLogicalName!AttributeName}
            string regexPattern = "{[\\w*!.]+}";
            MatchCollection matchCollections = Regex.Matches(dataToReplace, regexPattern);
            string entityName = string.Empty;
            string attributeName = string.Empty;

            // dictionary object to store all match
            Dictionary<string, string> matchValueCollection = new Dictionary<string, string>();

            foreach (Match match in matchCollections)
            {
                // split match with entity logical name and attribute
                var splitter = match.Value.TrimStart('{').TrimEnd('}').Split('!');
                if (splitter.Length > 1)
                {
                    entityName = splitter[0];
                    attributeName = splitter[1];

                    if (relatedEntity != null && relatedEntity.LogicalName == entityName && relatedEntity.Contains(attributeName))
                    {
                        // match with related entity
                        matchValueCollection.Add(match.Value, GetAttributeValue(attributeName, relatedEntity));
                    }
                    else if (primaryEntity != null && primaryEntity.LogicalName == entityName && primaryEntity.Contains(attributeName))
                    {
                        // match with primary entity
                        matchValueCollection.Add(match.Value, GetAttributeValue(attributeName, primaryEntity));
                    }
                    else if (parameters != null && entityName == "parameter" && parameters.ContainsKey(match.Value.TrimStart('{').TrimEnd('}')))
                    {
                        matchValueCollection.Add(match.Value, (string)parameters[match.Value.TrimStart('{').TrimEnd('}')]);
                    }
                    else if (relatedEntity != null && relatedEntity.LogicalName == entityName && attributeName.ToUpper(CultureInfo.InvariantCulture) == "URL")
                    {
                        // URL match for related entity
                        matchValueCollection.Add(match.Value, ReplaceURL(relatedEntity, recordBaseUrl));
                    }
                    else if (primaryEntity != null && primaryEntity.LogicalName == entityName && attributeName.ToUpper(CultureInfo.InvariantCulture) == "URL")
                    {
                        // URL match for primary entity
                        matchValueCollection.Add(match.Value, ReplaceURL(primaryEntity, recordBaseUrl));
                    }
                    else
                    {
                        // no match
                        matchValueCollection.Add(match.Value, string.Empty);
                    }
                }
                else
                {
                    throw new InvalidPluginExecutionException("Specified token is not in correct format. Please contact Administrator.");
                }
            }

            return ReplaceTokens(dataToReplace, matchValueCollection, regexPattern);
        }

        /// <summary>
        /// Replace place holders for regarding attributes
        /// </summary>
        /// <param name="primaryEntity">primary entity</param>
        /// <param name="relatedEntity">related entity</param>
        /// <param name="regardingAttribute">regarding attribute</param>
        /// <returns>Entity Reference for regarding attribute</returns>
        internal static EntityReference ReplaceRegardingAttribute(Entity primaryEntity, Entity relatedEntity, string regardingAttribute)
        {
            string entityName = string.Empty;
            string attributeName = string.Empty;

            if (!string.IsNullOrEmpty(regardingAttribute))
            {
                // Regex Match format. specified format is {entityLogicalName!AttributeName}
                MatchCollection found = Regex.Matches(regardingAttribute, "[\\w*!.]+");

                if (found.Count != 1)
                {
                    // place holder should contain maximum one attribute value
                    throw new InvalidPluginExecutionException(string.Format(CultureInfo.InvariantCulture, "Replacement string format for 'Regarding Entity' is not in correct format. Please contact Administrator."));
                }

                // split placeholders in form of attributeName and entityName
                var replacementString = found[0].Value.Split('!');

                if (replacementString.Length != 2)
                {
                    throw new InvalidPluginExecutionException(string.Format(CultureInfo.InvariantCulture, "Replacement string format for 'Regarding Entity' is not in correct format. Please contact Administrator."));
                }

                entityName = replacementString[0];
                attributeName = replacementString[1];
            }

            EntityReference regardingId = null;

            if (relatedEntity.LogicalName == entityName && relatedEntity.Contains(attributeName))
            {
                regardingId = GetRegardingAttributeValue(attributeName, relatedEntity);
            }
            else if (primaryEntity.LogicalName == entityName && primaryEntity.Contains(attributeName))
            {
                regardingId = GetRegardingAttributeValue(attributeName, primaryEntity);
            }

            return regardingId;
        }

        /// <summary>
        /// replace URL in Template data for email
        /// </summary>
        /// <param name="entity">entity object</param>
        /// <param name="baseURL">record base URL</param>
        /// <returns>return replaced URL</returns>
        internal static string ReplaceURL(Entity entity, string baseURL)
        {
            string url = string.Format(CultureInfo.InvariantCulture, baseURL, entity.LogicalName, entity.Id);
            string hyperlink = "<a href='" + url + "'>" + "Click Here" + "</a>";
            return hyperlink;
        }

        /// <summary>
        /// replace place holders of Template date
        /// </summary>
        /// <param name="primaryEntity">primary entity</param>
        /// <param name="relatedEntities">related entities</param>
        /// <param name="parameter">list of parameter attribute</param>
        /// <param name="notificationTemplate">notification template entity</param>
        /// <param name="recordBaseUrl">record base URL</param>
        internal void ReplaceTemplateData(Entity primaryEntity, EntityCollection relatedEntities, Dictionary<string, dynamic> parameter, Entity notificationTemplate, string recordBaseUrl)
        {
            if (relatedEntities == null || relatedEntities.Entities.Count == 0)
            {
                return;
            }

            // Extract data from Notification template entity
            this.ExtractContext(notificationTemplate);

            // Iterate over related entities to replace subject, decsription and due date placeholders
            foreach (Entity relatedEntity in relatedEntities.Entities)
            {
                // Replace placeholders for subject field of activity entity
                string activitySubject = ReplacePlaceHolders(relatedEntity, primaryEntity, parameter, recordBaseUrl, this.Subject);

                // get regarding value for task or email
                EntityReference regardingId = ReplaceRegardingAttribute(primaryEntity, relatedEntity, this.RegardingEntity);

                // replace placeholder for activity regarding
                if (regardingId != null)
                {
                    relatedEntity.Attributes["regarding"] = regardingId;
                }

                // if notification template is selected for task entity
                if (notificationTemplate.Contains("didd_task") && (bool)notificationTemplate.Attributes["didd_task"] == true)
                {
                    // Replace placeholders for task Descripiton
                    relatedEntity.Attributes["taskdescription"] = ReplacePlaceHolders(relatedEntity, primaryEntity, parameter, recordBaseUrl, this.TaskDescription);

                    // replace placeholders for task subject
                    relatedEntity.Attributes["tasksubject"] = activitySubject;

                    // replace placeholder for task due date
                    relatedEntity.Attributes["taskduedate"] = RelpcaeDueDatePlaceHolder(primaryEntity, relatedEntity, parameter, this.DueDate, this.Duration, this.DueDateModifier, this.DueDateModifierAttribute);
                }

                // if notification template is selected for email entity
                if (notificationTemplate.Contains("didd_email") && (bool)notificationTemplate.Attributes["didd_email"] == true)
                {
                    // Replace placeholders for email Descripiton
                    relatedEntity.Attributes["emaildescription"] = ReplacePlaceHolders(relatedEntity, primaryEntity, parameter, recordBaseUrl, this.EmailDescripiton);

                    // Replace placeholders for email subject
                    relatedEntity.Attributes["emailsubject"] = activitySubject;
                }
            }
        }

        /// <summary>
        /// Replace placeholders tokens
        /// </summary>
        /// <param name="template">template value</param>
        /// <param name="replacements">replacement tokens dictionary</param>
        /// <param name="regexPattern">regex pattern</param>
        /// <returns>return replaced value of template</returns>
        private static string ReplaceTokens(string template, Dictionary<string, string> replacements, string regexPattern)
        {
            // regex pattern
            var rex = new Regex(regexPattern);
            return rex.Replace(
                template,
                delegate(Match m)
                {
                    string key = m.Value;
                    string rep = replacements.ContainsKey(key) ? replacements[key] : string.Empty;
                    return rep;
                });
        }

        /// <summary>
        /// Replace due date place holders
        /// </summary>
        /// <param name="primaryEntity">primary entity</param>
        /// <param name="relatedEntity">related entity</param>
        /// <param name="parameters">list of parameter attribute</param>
        /// <param name="dueDate">Due Date</param>
        /// <param name="duration">Duration Attribute</param>
        /// <param name="dueDateModifier">Due Date Modifier</param>
        /// <param name="dueDateModifierAttribute">Due Date Modifier Attribute</param>
        /// <returns>return Replaced Due Date</returns>
        private DateTime? RelpcaeDueDatePlaceHolder(Entity primaryEntity, Entity relatedEntity, Dictionary<string, dynamic> parameters, DateTime? dueDate, long? duration, int? dueDateModifier, string dueDateModifierAttribute)
        {
            string entityName = string.Empty;
            string attributeName = string.Empty;
            string replacementStringValue = string.Empty;

            if (!string.IsNullOrEmpty(dueDateModifierAttribute))
            {
                // Regex Match format. specified format is {entityLogicalName!AttributeName}
                MatchCollection found = Regex.Matches(dueDateModifierAttribute, "[\\w*!.]+");

                if (found.Count != 1)
                {
                    // place holder should contain maximum one attribute value
                    throw new InvalidPluginExecutionException(string.Format(CultureInfo.InvariantCulture, "Replacement string format for due date is not in correct format. Please contact Administrator."));
                }

                // split placeholders in form of attributeName and entityName
                replacementStringValue = found[0].Value;
                var replacementString = replacementStringValue.Split('!');

                if (replacementString.Length != 2)
                {
                    throw new InvalidPluginExecutionException(string.Format(CultureInfo.InvariantCulture, "Replacement string format for due date is not in correct format. Please contact Administrator."));
                }

                entityName = replacementString[0];
                attributeName = replacementString[1];
            }

            DateTime? dueDateValue = null;

            if (relatedEntity.LogicalName == entityName && relatedEntity.Contains(attributeName))
            {
                // due date with related entity placeholders
                dueDateValue = FormDueDate(dueDate, duration, dueDateModifier, Convert.ToDateTime(GetAttributeValue(attributeName, relatedEntity), CultureInfo.InvariantCulture));
            }
            else if (primaryEntity.LogicalName == entityName && primaryEntity.Contains(attributeName))
            {
                // due date with primary entity placeholders
                dueDateValue = FormDueDate(dueDate, duration, dueDateModifier, Convert.ToDateTime(GetAttributeValue(attributeName, primaryEntity), CultureInfo.InvariantCulture));
            }
            else if (parameters != null && entityName == "parameter" && parameters.ContainsKey(replacementStringValue))
            {
                dueDateValue = Convert.ToDateTime((string)parameters[replacementStringValue], CultureInfo.InvariantCulture);
                tracingService.Trace("Due Date -> " + dueDateValue.Value);
            }
            else
            {
                // default data
                dueDateValue = FormDueDate(dueDate, duration, dueDateModifier, null);
            }

            if (dueDateValue != null)
            {
                return dueDateValue.Value;
            }

            return null;
        }

        /// <summary>
        /// Form due date with help of duration, due date, due date modifier and due date Modifier attribute
        /// </summary>
        /// <param name="dueDate">due date attribute</param>
        /// <param name="duration">duration attribute</param>
        /// <param name="dueDateModifier">due date modifier</param>
        /// <param name="modifierAttributeValue">due date modifier attribute value</param>
        /// <returns>Form due date with given data</returns>
        private static DateTime? FormDueDate(DateTime? dueDate, long? duration, int? dueDateModifier, DateTime? modifierAttributeValue)
        {
            DateTime? dueDateValue = null;
            if (dueDate != null)
            {
                // if system only contains exact due date
                dueDateValue = dueDate;
            }
            else if (duration != null && dueDateModifier != null && modifierAttributeValue != null)
            {
                // if template contains duretion, DueDateModifier and DueDateModifierAttribute value
                TimeSpan span = new TimeSpan(duration.Value * 60 * 1000 * 10000);

                if (dueDateModifier == 273310000)
                {
                    // Before Modifier attribute
                    dueDateValue = modifierAttributeValue.Value.Subtract(span);
                }
                else
                {
                    // after modifier attribute
                    dueDateValue = modifierAttributeValue.Value.Add(span);
                }
            }
            else if (modifierAttributeValue != null && dueDateModifier == null && duration == null)
            {
                // if ReplaceTemplateData contains only modifierAttributeValue 
                dueDateValue = modifierAttributeValue;
            }
            else if (duration != null && modifierAttributeValue == null && dueDateModifier == null)
            {
                // some fixed duration
                TimeSpan span = new TimeSpan(duration.Value * 60 * 1000 * 10000);
                dueDateValue = DateTime.Today.Add(span);
            }

            return dueDateValue;
        }

        /// <summary>
        /// Get Attribute value from entity
        /// </summary>
        /// <param name="attribute">attribute name</param>
        /// <param name="entity">Entity object</param>
        /// <returns>return attribute value from entity</returns>
        private static string GetAttributeValue(string attribute, Entity entity)
        {
            string attributeType = Convert.ToString(entity.Attributes[attribute].GetType(), CultureInfo.InvariantCulture);
            string attributeValue = string.Empty;

            switch (attributeType)
            {
                // EntityReference
                case "Microsoft.Xrm.Sdk.EntityReference":
                    attributeValue = ((EntityReference)entity.Attributes[attribute]).Name;
                    break;

                // Guid Value
                case "System.Guid":
                    attributeValue = ((Guid)entity.Attributes[attribute]).ToString();
                    break;

                // Option Set
                case "Microsoft.Xrm.Sdk.OptionSetValue":
                    attributeValue = entity.FormattedValues[attribute].ToString();
                    break;

                // Money
                case "Microsoft.Xrm.Sdk.Money":
                    attributeValue = ((Money)entity.Attributes[attribute]).Value.ToString(CultureInfo.InvariantCulture);
                    break;

                // Aliased Value
                case "Microsoft.Xrm.Sdk.AliasedValue":
                    attributeValue = GetAliasedValue(((AliasedValue)entity.Attributes[attribute]).Value, entity, attribute);
                    break;

                // DateTime and others attribute
                default:
                    if (attributeType == "System.DateTime")
                    {
                        attributeValue = ((DateTime)entity.Attributes[attribute]).Date.ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        attributeValue = entity.Attributes[attribute].ToString();
                    }

                    break;
            }

            return attributeValue;
        }

        /// <summary>
        /// Get attribute value for regarding attribute
        /// </summary>
        /// <param name="attribute">attribute name</param>
        /// <param name="entity">entity object</param>
        /// <returns>Entity reference for regarding attribute</returns>
        private static EntityReference GetRegardingAttributeValue(string attribute, Entity entity)
        {
            string attributeType = Convert.ToString(entity.Attributes[attribute].GetType(), CultureInfo.InvariantCulture);
            EntityReference attributeValue = null;

            switch (attributeType)
            {
                // EntityReference
                case "Microsoft.Xrm.Sdk.EntityReference":
                    attributeValue = (EntityReference)entity.Attributes[attribute];
                    break;

                // Guid Value
                case "System.Guid":
                    attributeValue = new EntityReference(entity.LogicalName, (Guid)entity.Attributes[attribute]);
                    break;

                // Aliased Value
                case "Microsoft.Xrm.Sdk.AliasedValue":
                    string aliasedAttributeType = Convert.ToString(((AliasedValue)entity.Attributes[attribute]).Value.GetType(), CultureInfo.InvariantCulture);
                    if (aliasedAttributeType == "Microsoft.Xrm.Sdk.EntityReference")
                    {
                        attributeValue = (EntityReference)((AliasedValue)entity.Attributes[attribute]).Value;
                    }
                    else if (aliasedAttributeType == "System.Guid")
                    {
                        AliasedValue attributeValueAlias = (AliasedValue)entity.Attributes[attribute];
                        attributeValue = new EntityReference(attributeValueAlias.EntityLogicalName, (Guid)attributeValueAlias.Value);
                    }

                    break;
            }

            return attributeValue;
        }

        /// <summary>
        /// Get Value of Aliased attribute from entity
        /// </summary>
        /// <param name="attribure">attribute type</param>
        /// <param name="entity">entity object</param>
        /// <param name="attributeName">attribute logical name</param>
        /// <returns>return value of aliased attribute from entity</returns>
        private static string GetAliasedValue(object attribure, Entity entity, string attributeName)
        {
            // get type of attribute
            string attributeType = Convert.ToString(attribure.GetType(), CultureInfo.InvariantCulture);
            string attributeValue = string.Empty;

            switch (attributeType)
            {
                // Entity Reference
                case "Microsoft.Xrm.Sdk.EntityReference":
                    attributeValue = ((EntityReference)attribure).Name;
                    break;

                case "System.Guid":
                    attributeValue = ((Guid)attribure).ToString();
                    break;

                // Option Set
                case "Microsoft.Xrm.Sdk.OptionSetValue":
                    attributeValue = entity.FormattedValues[attributeName].ToString();
                    break;

                // Money Attribute
                case "Microsoft.Xrm.Sdk.Money":
                    attributeValue = ((Money)attribure).Value.ToString(CultureInfo.InvariantCulture);
                    break;

                // DateTime and others attributes
                default:
                    if (attributeType == "System.DateTime")
                    {
                        attributeValue = ((DateTime)attribure).ToString(CultureInfo.InvariantCulture);
                    }
                    else
                    {
                        attributeValue = attribure.ToString();
                    }

                    break;
            }

            return attributeValue;
        }

        /// <summary>
        /// Extract data from notification template Entity, which need to be processed
        /// </summary>
        /// <param name="notificationTemplate">Notification Template Entity</param>
        private void ExtractContext(Entity notificationTemplate)
        {
            if (notificationTemplate.Contains("didd_taskdescription"))
            {
                // Task Description
                this.TaskDescription = (string)notificationTemplate.Attributes["didd_taskdescription"];
            }

            if (notificationTemplate.Contains("didd_emaildescription"))
            {
                // Email Description
                this.EmailDescripiton = (string)notificationTemplate.Attributes["didd_emaildescription"];
            }

            if (notificationTemplate.Contains("didd_subject"))
            {
                // subject
                this.Subject = (string)notificationTemplate.Attributes["didd_subject"];
            }

            if (notificationTemplate.Contains("didd_duedate"))
            {
                // Due Date
                this.DueDate = (DateTime)notificationTemplate.Attributes["didd_duedate"];
            }

            if (notificationTemplate.Contains("didd_duration"))
            {
                // Duration 
                this.Duration = (long)((int)notificationTemplate.Attributes["didd_duration"]);
            }

            if (notificationTemplate.Contains("didd_duedatemodifier"))
            {
                // Due Date Modifier
                this.DueDateModifier = ((OptionSetValue)notificationTemplate.Attributes["didd_duedatemodifier"]).Value;
            }

            if (notificationTemplate.Contains("didd_duedatemodifierattribure"))
            {
                // Due Date Modifier attribute
                this.DueDateModifierAttribute = (string)notificationTemplate.Attributes["didd_duedatemodifierattribure"];
            }

            if (notificationTemplate.Contains("didd_regardingentity"))
            {
                // regarding Entity Details
                this.RegardingEntity = (string)notificationTemplate.Attributes["didd_regardingentity"];
            }
        }
    }
}
