// <copyright file="QueryFilter.cs" company="Microsoft">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author></author>
// <date>7/9/2017 12:08:08 PM</date>
// <summary>Query Filter</summary>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;

    /// <summary>
    /// Filter Class for Fetch Expression
    /// </summary>
    public class QueryFilter
    {
        /// <summary>
        /// Gets or sets Sequence Number
        /// </summary>
        public int SequenceNumber { get; set; }

        /// <summary>
        /// Gets or sets Sub Sequence Number
        /// </summary>
        public int SubsequenceNumber { get; set; }

        /// <summary>
        /// Gets or sets Key
        /// </summary>
        public string Key { get; set; }

        /// <summary>
        /// Gets or sets Value
        /// </summary>
        public string Value { get; set; }
    }
}
