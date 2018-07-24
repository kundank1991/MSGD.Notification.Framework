// <copyright file="PartyList.cs" company="Microsoft">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author></author>
// <date>7/9/2017 12:08:08 PM</date>
// <summary>Party List</summary>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Microsoft.Xrm.Sdk;

    /// <summary>
    /// Party List class for Notification Framework
    /// </summary>
    public class PartyList
    {
        /// <summary>
        /// Gets or sets Participation Type
        /// </summary>
        public int ParticipationType { get; set; }

        /// <summary>
        /// Gets or sets related Entity Logical Name
        /// </summary>
        public string RelatedEntityLogicalName { get; set; }

        /// <summary>
        /// Gets or sets Party Id
        /// </summary>
        public EntityReference PartyId { get; set; }

        /// <summary>
        /// Gets or sets Related Entity Id
        /// </summary>
        public Guid RelatedEntityId { get; set; }
    }
}
