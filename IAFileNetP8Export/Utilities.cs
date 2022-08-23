using System;
using System.Collections.Generic;
using System.Text;
using Oracle.ManagedDataAccess.Client;

namespace CNO.BPA.IAFileNetP8Export
{
    /// <summary>
    /// Provides some utility functions for working with the database.
    /// </summary>
    internal class DBUtilities
    {
        /// <summary>
        /// Creates a new parameter for a command
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="type"></param>
        /// <param name="direction"></param>
        /// <param name="size"></param>
        /// <param name="command"></param>
        public static void CreateAndAddParameter(string name, object value, OracleDbType type, System.Data.ParameterDirection direction, int size, OracleCommand command)
        {
            OracleParameter parameter = command.CreateParameter();
            parameter.ParameterName = name;
            parameter.Value = value;
            parameter.OracleDbType = type;
            parameter.Direction = direction;
            parameter.Size = size;
            command.Parameters.Add(parameter);
        }

        /// <summary>
        /// Creates a new parameter for a command
        /// </summary>
        /// <param name="name"></param>
        /// <param name="value"></param>
        /// <param name="type"></param>
        /// <param name="direction"></param>
        /// <param name="command"></param>
        public static void CreateAndAddParameter(string name, object value, OracleDbType type, System.Data.ParameterDirection direction, OracleCommand command)
        {
            OracleParameter parameter = command.CreateParameter();
            parameter.ParameterName = name;
            parameter.Value = value;
            parameter.OracleDbType = type;
            //parameter.OracleType = type;
            parameter.Direction = direction;
            command.Parameters.Add(parameter);
        }

        /// <summary>
        /// Creates a new parameter for a command
        /// </summary>
        /// <param name="name"></param>
        /// <param name="type"></param>
        /// <param name="direction"></param>
        /// <param name="command"></param>
        public static void CreateAndAddParameter(string name, OracleDbType type, System.Data.ParameterDirection direction, OracleCommand command)
        {
            CreateAndAddParameter(name, null, type, direction, command);
        }

        /// <summary>
        /// Creates a new parameter for a command
        /// </summary>
        /// <param name="name"></param>
        /// <param name="type"></param>
        /// <param name="direction"></param>
        /// <param name="size"></param>
        /// <param name="command"></param>
        public static void CreateAndAddParameter(string name, OracleDbType type, System.Data.ParameterDirection direction, int size, OracleCommand command)
        {
            CreateAndAddParameter(name, null, type, direction, size, command);
        }
    }
}
