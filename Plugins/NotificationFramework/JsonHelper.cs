// <copyright file="JsonHelper.cs" company="Microsoft">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author>Microsoft</author>
// <date>11/1/2017</date>
// <summary>Helper class for JSON string.</summary>
namespace MSGD.Notification.Framework.Plugins
{
    using System;
    using System.IO;
    using System.Runtime.Serialization.Json;
    using System.Text;

    /// <summary>
    /// Class to convert JSON text to Object and vice versa
    /// </summary>
    public static class JsonHelper
    {
        /// <summary>
        /// Converts JSON String to Object
        /// </summary>
        /// <param name="jsonValue">JSON Text to convert into desired Object</param>
        /// <param name="type">Type of Object</param>
        /// <returns>Returns Object</returns>
        public static object GetObject(string jsonValue, Type type)
        {
            using (MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonValue)))
            {
                DataContractJsonSerializer jsonSerializer = new DataContractJsonSerializer(type);
                object objResponse = jsonSerializer.ReadObject(stream);
                return objResponse;
            }
        }

        /// <summary>
        /// Converts Object to JSON string
        /// </summary>
        /// <param name="value">Type of Object to convert into JSON text</param>
        /// <param name="type">Type of Object</param>
        /// <returns>Returns JSON Text</returns>
        public static string GetJsonStringFor(object value, Type type)
        {
            string result = string.Empty;
            using (MemoryStream stream = new MemoryStream())
            {
                DataContractJsonSerializer ser = new DataContractJsonSerializer(type);
                ser.WriteObject(stream, value);
                stream.Position = 0;
                StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                result = reader.ReadToEnd();
            }

            return result;
        }
    }
}
