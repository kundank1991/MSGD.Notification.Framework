// <copyright file="PreNotificationPlaceHolder.cs" company="">
// Copyright (c) 2018 All Rights Reserved
// </copyright>
// <author></author>
// <date>7/24/2018 1:50:30 PM</date>
// <summary>Implements the PreNotificationPlaceHolder Plugin.</summary>
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.1
// </auto-generated>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using System.Globalization;

    /// <summary>
    /// PreNotificationPlaceHolder Plugin.
    /// </summary>    
    public class PreNotificationPlaceHolder: PluginBase
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PreNotificationPlaceHolder"/> class.
        /// </summary>
        /// <param name="unsecure">Contains public (unsecured) configuration information.</param>
        /// <param name="secure">Contains non-public (secured) configuration information. 
        /// When using Microsoft Dynamics 365 for Outlook with Offline Access, 
        /// the secure string is not passed to a plug-in that executes while the client is offline.</param>
        public PreNotificationPlaceHolder(string unsecure, string secure)
            : base(typeof(PreNotificationPlaceHolder))
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
        protected void ExecutePreNotificationPlaceHolder(LocalPluginContext localContext)
        {
            if (localContext == null)
            {
                throw new ArgumentNullException("localContext");
            }

            localContext.TracingService.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: getting the plugin execution context. {1}", "PreNotificationPlaceHolder", DateTime.Now));

            ////Get the plugin context.
            IPluginExecutionContext context = localContext.PluginExecutionContext;

            ////Get the Tracing service.
            ITracingService tracingservice = localContext.TracingService;

            ////Get the IOrganization Service
            IOrganizationService service = localContext.OrganizationService;

            tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: Entering ExecutePreNotificationPlaceHolder {1}", "PreNotificationPlaceHolder", DateTime.Now));

            try
            {
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: Exiting ExecutePreNotificationPlaceHolder {1}", "PreNotificationPlaceHolder", DateTime.Now));

                NotificationPlaceHolderHelper notify = new NotificationPlaceHolderHelper();
                notify.ReplaceTokenData(context, service, tracingservice);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: encountred an isssue and exiting. {1}", "PreNotificationPlaceHolder", DateTime.Now));
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: exited with the following error {1}. {2}", "PreNotificationPlaceHolder", ex, DateTime.Now));
                throw new InvalidPluginExecutionException("Ended up with Organization Service Fault. Following is the error.", ex);
            }
            catch (TimeoutException ex)
            {
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: . {1}", "PreNotificationPlaceHolder", DateTime.Now));
                throw new InvalidPluginExecutionException("Plugin not able to complete its execution with in the specified time limits.", ex);
            }
            catch (Exception ex)
            {
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: encountred an isssue and exiting. {1}", "PreNotificationPlaceHolder", DateTime.Now));
                tracingservice.Trace(string.Format(CultureInfo.InvariantCulture, "{0}: exited with the following error {1}. {2}", "PostNotify", ex, DateTime.Now));
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
        }
    }
}
