using System;
using System.Globalization;
using System.Text;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using System.Collections.Generic;


namespace CNO.BPA.IAFileNetP8Export
{
    public class DataAccess
   {

      #region Procedure Names
      private const string SQL_SELECT_DD_ITEM = "BPA_APPS.PKG_IAFNP8EXP.SELECT_DD_ITEM";
      private const string SQL_SELECT_DD_ITEM_INDEXES = "BPA_APPS.PKG_IAFNP8EXP.SELECT_DD_ITEM_INDEXES";
      private const string SQL_UPDATE_DOCID = "BPA_APPS.PKG_IAFNP8EXP.UPDATE_FNP8_DOCID";
      private const string SQL_UPDATE_STATUS = "BPA_APPS.PKG_IAFNP8EXP.UPDATE_STATUS";

      #endregion
      #region Variables
      private CNO.BPA.Framework.Cryptography crypto = new CNO.BPA.Framework.Cryptography();
      private OracleConnection _connection = null;
      private string _connectionString = null;
      private OracleTransaction _transaction = null;


      private string _DSN = "";
      private string _DBUser = "";
      private string _DBPass = "";
      #endregion

      public DataAccess(string dsn, string user, string pass)
      {

         //next grab a copy of each of the db values
         _DSN = dsn;
         _DBUser = crypto.Decrypt(user);
         _DBPass = crypto.Decrypt(pass);

         //check to see that we have values for the db info
         if (_DSN.Length != 0 & _DBUser.Length != 0 &
             _DBPass.Length != 0)
         {
            //build the connection string
            _connectionString = "Data Source=" + _DSN + ";Persist Security Info=True;User ID="
               + _DBUser + ";Password=" + _DBPass + "";
         }
         else
         {
            throw new ArgumentNullException("-266363925; Database information could "
               + "not be found in the setup for this instance step.");
         }
      }
      public void UpdateDocId(DataRow dr, string docId)
      {
         try
         {
            CreateConnection();
            using (OracleCommand cmd = GenerateCommand(SQL_UPDATE_DOCID, CommandType.StoredProcedure, false))
            {                
               DBUtilities.CreateAndAddParameter("p_in_app_name", "IABPAFileNetP8Export",
                  OracleDbType.Varchar2, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_machine_name", System.Environment.MachineName.ToString().ToUpper(),
                  OracleDbType.Varchar2, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_user_id", System.Environment.UserName.ToString().ToUpper(),
                  OracleDbType.Varchar2, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_fnp8_docid", docId,
                  OracleDbType.Varchar2, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_dd_item_seq", dr["DD_ITEM_SEQ"],
                  OracleDbType.Int32, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_source_id", dr["SOURCE_ID"],
                  OracleDbType.Int32, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_docustream_request_id", dr["DOCUSTREAM_REQUEST_ID"],
                  OracleDbType.Int32, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_front_office_request_id", dr["FRONT_OFFICE_REQUEST_ID"],
                  OracleDbType.Int32, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_third_party_request_id", dr["THIRD_PARTY_REQUEST_ID"],
                  OracleDbType.Int32, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_invoice_request_id", dr["INVOICE_REQUEST_ID"],
                  OracleDbType.Int32, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_out_result", OracleDbType.Varchar2,
                  ParameterDirection.Output, 255, cmd);
               DBUtilities.CreateAndAddParameter("p_out_error_message", OracleDbType.Varchar2,
                  ParameterDirection.Output, 4000, cmd);
                    cmd.BindByName = true;
                    OpenConnection();
               cmd.ExecuteNonQuery();
               CloseConnection();
               if (cmd.Parameters["p_out_result"].Value.ToString().ToUpper() != "SUCCESSFUL")
               {
                  throw new Exception("-266363929; Procedure Error: " +
                     cmd.Parameters["p_out_result"].Value.ToString() + "; Oracle Error: " +
                     cmd.Parameters["p_out_error_message"].Value.ToString());
               }
            }
         }
         catch (Exception ex)
         {
            throw new Exception(ex.Message.ToString());
         }

      }
      public void UpdateStatus(string batchNo, string nodeId)
      {
         try
         {
            CreateConnection();
            using (OracleCommand cmd = GenerateCommand(SQL_UPDATE_STATUS, CommandType.StoredProcedure, false))
            {                
               DBUtilities.CreateAndAddParameter("p_in_app_name", "IABPAFileNetP8Export",
                  OracleDbType.Varchar2, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_machine_name", System.Environment.MachineName.ToString().ToUpper(),
                  OracleDbType.Varchar2, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_user_id", System.Environment.UserName.ToString().ToUpper(),
                  OracleDbType.Varchar2, ParameterDirection.Input, cmd);

               DBUtilities.CreateAndAddParameter("p_in_batch_no",
                    batchNo, OracleDbType.Varchar2, ParameterDirection.Input, cmd);
               DBUtilities.CreateAndAddParameter("p_in_doc_xref",
                    nodeId, OracleDbType.Varchar2, ParameterDirection.Input, cmd);

               DBUtilities.CreateAndAddParameter("p_out_result", OracleDbType.Varchar2,
                  ParameterDirection.Output, 255, cmd);
               DBUtilities.CreateAndAddParameter("p_out_error_message", OracleDbType.Varchar2,
                  ParameterDirection.Output, 4000, cmd);
                    cmd.BindByName = true;
                    OpenConnection();
               cmd.ExecuteNonQuery();
               CloseConnection();
               if (cmd.Parameters["p_out_result"].Value.ToString().ToUpper() != "SUCCESSFUL")
               {
                  throw new Exception("-266363929; Procedure Error: " +
                     cmd.Parameters["p_out_result"].Value.ToString() + "; Oracle Error: " +
                     cmd.Parameters["p_out_error_message"].Value.ToString());
               }
            }
         }
         catch (Exception ex)
         {
            throw new Exception(ex.Message.ToString());
         }

      }
      private OracleDataAdapter GenerateAdapter(string commandText, System.Data.CommandType commandType)
      {
         OracleDataAdapter cmd = new OracleDataAdapter(commandText, _connectionString);
         cmd.SelectCommand.CommandType = commandType;
         return cmd;
      }

      internal void CreateConnection()
      {
         _connection = new OracleConnection();
         _connection.ConnectionString = _connectionString;
      }
      internal void OpenConnection()
      {
         try
         {
            _connection.Open();
         }
         catch (Exception ex)
         {
            throw new Exception("An error occurred while connecting to the database.", ex);
         }
      }
      internal void CloseConnection()
      {
         try
         {
            _connection.Close();
            _connection.Dispose();
         }
         catch
         {
         }
      }
      internal void BeginTran()
      {
         _connection = new OracleConnection();
         _connection.ConnectionString = _connectionString;
         try
         {
            _connection.Open();
            _transaction = _connection.BeginTransaction();

         }
         catch (Exception ex)
         {
            throw new Exception("An error occurred while connecting to the database.", ex);
         }
      }
      /// <summary>
      /// Commits the current transaction and disconnects from the database.
      /// </summary>
      internal void EndTran()
      {
         try
         {
            if (null != _connection)
            {

               _transaction.Commit();
               _transaction.Dispose();
               _transaction = null;
               _connection.Close();
               _connection.Dispose();
               _connection = null;

            }
         }
         catch { } // ignore an error here
      }
      /// <summary>
      /// Commits all of the data changes to the database.
      /// </summary>
      internal void Commit()
      {
         _transaction.Commit();
      }
      /// <summary>
      /// Cancels the transaction and voids any changes to the database.
      /// </summary>
      internal void Cancel()
      {
         _transaction.Rollback();
         _connection.Close();
         _connection.Dispose();
         _transaction.Dispose();
         _connection = null;
         _transaction = null;
      }
      /// <summary>
      /// Generates the command object and associates it with the current transaction object
      /// </summary>
      /// <param name="commandText"></param>
      /// <param name="commandType"></param>
      /// <returns></returns>
      private OracleCommand GenerateCommand(string commandText, System.Data.CommandType commandType, Boolean useTransaction)
      {
         OracleCommand cmd = new OracleCommand(commandText, _connection);
         if (true == useTransaction)
         {
            cmd.Transaction = _transaction;
         }
         cmd.CommandType = commandType;
         cmd.BindByName = true;
         return cmd;
      }

      internal DataSet getDD_ITEMDataSet(string BatchNo, string NodeID)
      {
         DataSet dsBatchData = new DataSet();

         OracleDataAdapter odaBD = GenerateAdapter(SQL_SELECT_DD_ITEM, CommandType.StoredProcedure);
         odaBD.SelectCommand.CommandType = CommandType.StoredProcedure;
         try
         {

            DBUtilities.CreateAndAddParameter("p_in_batch_no",
                 BatchNo, OracleDbType.Varchar2, ParameterDirection.Input, odaBD.SelectCommand);
            DBUtilities.CreateAndAddParameter("p_in_doc_xref",
                 NodeID, OracleDbType.Varchar2, ParameterDirection.Input, odaBD.SelectCommand);
            DBUtilities.CreateAndAddParameter("p_out_ref_cursor", OracleDbType.RefCursor,
                 ParameterDirection.Output, odaBD.SelectCommand);
            DBUtilities.CreateAndAddParameter("p_out_result", OracleDbType.Varchar2,
                 ParameterDirection.Output, 255, odaBD.SelectCommand);
            DBUtilities.CreateAndAddParameter("p_out_error_message", OracleDbType.Varchar2,
                 ParameterDirection.Output, 4000, odaBD.SelectCommand);

            odaBD.Fill(dsBatchData);

            if (odaBD.SelectCommand.Parameters["p_out_result"].Value.ToString().ToUpper() != "SUCCESSFUL")
            {
               throw new Exception("-266363929; Procedure Error: " +
                  odaBD.SelectCommand.Parameters["p_out_result"].Value.ToString() + "; Oracle Error: " +
                  odaBD.SelectCommand.Parameters["p_out_error_message"].Value.ToString());
            }

            return dsBatchData;
         }
         catch (Exception ex)
         {
            throw new Exception(ex.Message.ToString());
         }


      }
      internal DataSet getDD_ITEM_INDEXESDataSet(string DDItemSeq)
      {
         DataSet dsBatchData = new DataSet();

         OracleDataAdapter odaBD = GenerateAdapter(SQL_SELECT_DD_ITEM_INDEXES, CommandType.StoredProcedure);
         odaBD.SelectCommand.CommandType = CommandType.StoredProcedure;
         try
         {

            DBUtilities.CreateAndAddParameter("p_in_dd_item_seq",
                 DDItemSeq, OracleDbType.Varchar2, ParameterDirection.Input, odaBD.SelectCommand);

            DBUtilities.CreateAndAddParameter("p_out_ref_cursor", OracleDbType.RefCursor,
                 ParameterDirection.Output, odaBD.SelectCommand);

            DBUtilities.CreateAndAddParameter("p_out_result", OracleDbType.Varchar2,
                 ParameterDirection.Output, 255, odaBD.SelectCommand);
            DBUtilities.CreateAndAddParameter("p_out_error_message", OracleDbType.Varchar2,
                 ParameterDirection.Output, 4000, odaBD.SelectCommand);

            odaBD.Fill(dsBatchData);

            if (odaBD.SelectCommand.Parameters["p_out_result"].Value.ToString().ToUpper() != "SUCCESSFUL")
            {
               throw new Exception("-266363929; Procedure Error: " +
                  odaBD.SelectCommand.Parameters["p_out_result"].Value.ToString() + "; Oracle Error: " +
                  odaBD.SelectCommand.Parameters["p_out_error_message"].Value.ToString());
            }

            return dsBatchData;
         }
         catch (Exception ex)
         {
            throw new Exception(ex.Message.ToString());
         }


      }
   }

}
