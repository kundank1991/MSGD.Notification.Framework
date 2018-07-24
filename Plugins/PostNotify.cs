// <copyright file="PostNotify.cs" company="">
// Copyright (c) 2018 All Rights Reserved
// </copyright>
// <author></author>
// <date>7/24/2018 1:23:38 PM</date>
// <summary>Implements the PostNotify Plugin.</summary>
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.1
// </auto-generated>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Messages;
    using System.Collections;
    using System.Collections.Generic;
    using Microsoft.Crm.Sdk.Messages;
    using System.Xml;
    using System.Text.RegularExpressions;
    using System.Globalization;

    /// <summary>
    /// PostNotify Plugin.
    /// </summary>    
    public class PostNotify: PluginBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PostNotify"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PostNotify(string unsecure, string secure)
            : base(typeof(PostNotify))
        {
            
           // TODO: Implement your custom configuration handling.
        }


        /// <summary>
        /// Executes the plug-in.
        /// </summary>
        /// <param name="localContext">The <see cref="LocalPluginContext"/> which contains the
        /// <see cref="IPluginExecutionContext"/>,
        /// <see cref="IOrganizationService"/>
        /// and <see cref="ITracingService"/>
        /// </param>
        /// <remarks>
        /// For improved performance, Microsoft Dynamics CRM caches plug-in instances.
        /// The plug-in's Execute method should be written to be stateless as the constructor
        /// is not called for every invocation of the plug-in. Also, multiple system threads
        /// could execute the plug-in at the same time. All per invocation state information
        /// is stored in the context. This means that you should not use global variables in plug-ins.
        /// </remarks>
        protected void ExecutePostNotify(LocalPluginContext localContext)
        {
            ITracingService tracingservice = null;
            try
            {
                if (localContext == null)
                {
                    throw new InvalidPluginExecutionException("Local context");
                }

                localContext.TracingService.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: getting the plugin execution context. {1}", "PostNotify", DateTime.Now));

                ////Get the plugin context.
                IPluginExecutionContext context = localContext.PluginExecutionContext;

                ////Get the Tracing service.
                tracingservice = localContext.TracingService;

                ////Get the IOrganization Service
                IOrganizationService service = localContext.OrganizationService;

                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: Entering ExecutePostNotify {1}", "PostNotify", DateTime.Now));

                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: Exiting ExecutePostNotify {1}", "PostNotify", DateTime.Now));

                PostNotifyHelper postNotify = new PostNotifyHelper();
                postNotify.SendNotifications(context, service, tracingservice);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: encountred an isssue and exiting. {1}", "PostNotify", DateTime.Now));
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: exited with the following error {1}. {2}", "PostNotify", ex, DateTime.Now));
                throw new InvalidPluginExecutionException("Ended up with Organization Service Fault. Following is the error.", ex);
            }
            catch (TimeoutException ex)
            {
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: . {1}", "PostNotify", DateTime.Now));
                throw new InvalidPluginExecutionException("Plugin not able to complete its execution with in the specified time limits.", ex);
            }
            catch (Exception ex)
            {
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: encountred an isssue and exiting. {1}", "PostNotify", DateTime.Now));
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: exited with the following error {1}. {2}", "PostNotify", ex, DateTime.Now));
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }
    }
}
