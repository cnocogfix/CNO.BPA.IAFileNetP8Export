// Copyright © 2003–2009 EMC Corporation. All rights reserved.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
//----------------------------------------------------------------
// InputAccel namespaces - QuickModule, Processing Helper,
// Script Helper, Scripting interface, and Workflow Client.
//----------------------------------------------------------------
using Emc.InputAccel.QuickModule;
using Emc.InputAccel.QuickModule.Helpers.Processing;
using Emc.InputAccel.QuickModule.Plugins.Processing;
using Emc.InputAccel.Workflow.Client;
//----------------------------------------------------------------
// Custom script event namespace defined in the ExportCS.Scripting project.
//----------------------------------------------------------------
//dt using Sample.ExportCS.Scripting;

namespace CNO.BPA.IAFileNetP8Export
{
   //----------------------------------------------------------------
   // The IModule base interface is the main interface for QuickModule modules.
   // The IModuleInformation base interface is used to provide information when
   // installing the module as a service.
   //----------------------------------------------------------------
   public sealed class ExportModule : IModule, IModuleInformation
   {
      //---------------------------------------------------------------- 
      // TODO: Specify the following constants per your custom module.
      // 
      // If you will sell your module, contact the EMC Captiva support department
      // to register your ModuleName (MODULE_MDF_NAME) and get a unique, non-zero
      // ModuleId (MODULE_ID).
      // 
      // ModuleName (MODULE_MDF_NAME) must match the Module definition name in the
      // MDF such as "EXPORT" in "Module EXPORT(0)" as shown in the included 
      // Export.mdf file.
      //----------------------------------------------------------------
      private const int MODULE_ID = 4620;
      private const string MODULE_MDF_NAME = "FNP8EXP";                 // 8 character maximum
      private const string MODULE_SHORT_TITLE = "IA BPA FileNet P8 Export";          // 31 character maximum
      private const string MODULE_LONG_TITLE = "IA BPA FileNet P8 Export Custom Module";

      private const string INPUTFILE_KEY = "InputFile";

      //----------------------------------------------------------------
      // The following private members reference various objects during the
      // lifetime of the module.
      //----------------------------------------------------------------
      private IHelper quickModulehelper = null;
      private IProcessingHelper processingHelper = null;

      //Error Cosntants
      public const int IA_SUCCESS = 0;
      public const int IA_ERR_UNKNOWN = -4523;
      public const int IA_ERR_NOFUNC = -4518;
      public const int IA_ERR_CANCEL = -4526;
      public const int IA_ERR_NORETRY = -6112;
      public const int IA_ERR_RETRYSOME = -6113;
      public const int IA_ERR_RETRY = -6114;
      public const int IA_ERR_ACCESS = -4505;
      
      //----------------------------------------------------------------
      // The following private members keep track of the processing state of 
      // the current task.
      //----------------------------------------------------------------

      private ITask currentTask = null;                       // The task currently being processed.
      private string _DSN = String.Empty;
      private string _DBUser = String.Empty;
      private string _DBPass = String.Empty;
      private string _FNP8Domain = String.Empty;
      private string _FNP8User = String.Empty;
      private string _FNP8Pass = String.Empty;
      private string _FNP8URI = String.Empty;

      private DataAccess da = null;

      private FileNetP8 fnP8 = null;

      private List<Image> replicateImages = new List<Image>();

      IValueProvider _stepValueProvider = null;
      IValueProvider _taskValueProvider = null;

      //----------------------------------------------------------------
      // Miscellaneous private members. 
      //----------------------------------------------------------------
      private Random random = null;           // Used to generate a psuedo session ID for our fake third-party system.

      public ExportModule()
      {
         this.random = new Random(System.DateTime.Now.Millisecond);
      }

      #region IModule Members

      public void Initialize(IHelper helper, IModuleOptions options)
      {
         this.quickModulehelper = helper;

         //---------------------------------------------------------------- 
         // Module Options
         // 
         // See the comments at the beginning of this class explaining the constants used below.
         //----------------------------------------------------------------
         options.ModuleId = this.ModuleId;
         options.ModuleName = this.ModuleName;
         options.ShortTitle = this.ShortTitle;
         options.LongTitle = this.LongTitle;

         //----------------------------------------------------------------
         // Set the help file name to the general InputAccel help file name.
         //
         // TODO: If you create your own CHM file, replace the file name below with your own.
         //----------------------------------------------------------------
         options.HelpNamespace = "ia_en-us.chm";

         //----------------------------------------------------------------
         // Assign an event handler to BeginSetup in order to add one or more module-specific
         // setup panels to setup mode.
         //----------------------------------------------------------------
         options.BeginSetup += new EventHandler<BeginSetupEventArgs>(options_BeginSetup);

         options.OnServerReconnect += new EventHandler<ServerConnectionStateEventArgs>(options_OnServerReconnect);
         options.OnServerDisconnect += new EventHandler<ServerConnectionStateEventArgs>(options_OnServerDisconnect);

         //----------------------------------------------------------------
         // Processing Helper
         // 
         // Create the Processing Helper, add it to the external helpers collection, 
         // and initialize it. 
         // 
         // The Processing Helper aids in receiving tasks, processing nodes, and handling errors.
         // The Processing Helper also provides the "Errors" tab in setup mode in order to
         // configure error handling options.
         //----------------------------------------------------------------
         this.processingHelper = new ProcessingHelper(options);
         options.ExternalHelpers.Add(this.processingHelper);
         InitializeProcessingHelper();

         //----------------------------------------------------------------
         // Script Helper
         // 
         // Create the Script Helper, add it to the external helpers collection, 
         // and initialize it.
         // 
         // The Script Helper provides scripting capabilities which in turn allows your
         // customers to write scripts to extend the functionality of your module. You may
         // define custom script events and/or enable the default QuickModule script events.
         // 
         // The Script Helper adds the "Scripting" tab in setup mode which allows customers
         // to map script events to their own scripts.
         // 
         // TODO: Decide if you want scripting in your module. In most cases you should provide
         // scripting. If you do not want scripting in your module, remove the following
         // three lines of code, the scriptHelper variable, and the InitializeScriptHelper function.
         //----------------------------------------------------------------


      }

      #endregion

      #region Initialization Functions

      private void InitializeProcessingHelper()
      {
         /*
          * The following initialization is needed for production mode only.
          */
         if (!this.quickModulehelper.SetupMode)
         {
            /* 
             * The ProcessNodeLevel property maps to the BeginNode and FinishNode events. You can choose which
             * levels (0 to 7) receive these two events. This sample receives these events for level 0 only 
             * because it only deals with InputFile defined at level 0 in the MDF. This module, however, can 
             * be defined as a Step at any level in an IPP.
             */
            this.processingHelper.ProcessNodeLevel[0] = false;
            this.processingHelper.ProcessNodeLevel[1] = true;
            this.processingHelper.ProcessNodeLevel[2] = false;
            this.processingHelper.ProcessNodeLevel[3] = true;
            this.processingHelper.ProcessNodeLevel[4] = false;
            this.processingHelper.ProcessNodeLevel[5] = false;
            this.processingHelper.ProcessNodeLevel[6] = false;
            this.processingHelper.ProcessNodeLevel[7] = false;

            /*
             * Select which operations are shown on the left-hand side of the main production window. By default,
             * no operations are provided. See Emc.InputAccel.QuickModule.Helpers.Processing.OperationType for
             * details on each operation type.
             */
            this.processingHelper.Operations[OperationType.RunSingleBatch] = true;
            this.processingHelper.Operations[OperationType.RunAllBatches] = true;
            this.processingHelper.Operations[OperationType.Exit] = true;
            this.processingHelper.Operations[OperationType.Stop] = true;

            /*
              * Processing Helper Event Handlers
              * 
              * Define and register event handlers for Processing Helper events. 
              * 
              * Depending on the needs of your module, you may not handle all events. For example, this module
              * only needs BeginTask and BeginNode.
              */
            this.processingHelper.BeginTask += new EventHandler<BeginTaskEventsArgs>(processingHelper_BeginTask);
            this.processingHelper.BeginNode += new EventHandler<BeginNodeEventsArgs>(processingHelper_BeginNode);
            this.processingHelper.FinishTask += new EventHandler<FinishTaskEventsArgs>(processingHelper_FinishTask);
            this.processingHelper.BeginProduction += new EventHandler<EventArgs>(processingHelper_BeginProduction);
            this.processingHelper.FinishProduction += new EventHandler<EventArgs>(processingHelper_FinishProduction);
         }
      }



      #endregion

      #region IModuleOptions (Setup) Event Handlers

      //----------------------------------------------------------------
      // BeginSetup is fired when the module is started in setup mode to setup a Step in a batch
      // or process.
      // 
      // Handle this event to add your own setup panels as shown below. Setup panels are 
      // automatically shown in setup mode.
      // 
      // To create a new setup panel:
      // 
      // 1) Add a new "User Control" item to the project.
      // 2) Add "using Emc.InputAccel.QuickModule;" and "using Emc.InputAccel.Workflow.Client;" 
      //    to the new user control cs file.
      // 3) Add IPanel as a base interface to the user control class as in
      //    "public partial class MySetupPanel : UserControl, IPanel".
      // 4) Modify the default constructor to take two parameters - IHelper and IStep - as in
      //    "public MySetupPanel(IHelper helper, IStep step)".
      // 5) Implement and add code where needed to the IPanel members.
      // 6) Create and add an instance of the panel to the SetupPanels List in BeginSetup as shown below.
      //----------------------------------------------------------------
      void options_BeginSetup(object sender, BeginSetupEventArgs e)
      {

         IPanel panel = new SetupPanel(this.quickModulehelper, e.Step);
         e.SetupPanels.Add(panel);
      }

      void options_OnServerReconnect(object sender, ServerConnectionStateEventArgs e)
      {

      }

      void options_OnServerDisconnect(object sender, ServerConnectionStateEventArgs e)
      {

      }

      #endregion

      #region Processing Helper Event Handlers

      //----------------------------------------------------------------
      // BeginNode is fired for a node before its children (if applicable) are processed with
      // BeginNode and FinishNode.
      //----------------------------------------------------------------
      void processingHelper_BeginNode(object sender, BeginNodeEventsArgs e)
      {
         try
         {
            if (e.Node.Level.Number == 3)
            {


               List<string> lstNodes = new List<string>();
               List<Stream> docImages = new List<Stream>();
               List<Stream> replicateImages = new List<Stream>();
               TiffUtility tu = new TiffUtility();

               INode envNode = e.Node;
               string BatchNo = envNode.Tree.Batch.Name;

               ILevel level1 = this.currentTask.Batch.Level(1);
               ILevel level0 = this.currentTask.Batch.Level(0);
               foreach (INode docNode in envNode.Children(level1))
               {
                  string NodeID = docNode.Id.ToString();
                  INodeValueProvider nodeValueProvider = docNode.Value(this.currentTask.Step);
                  string replicate = nodeValueProvider.Get("$instance=Standard_MDF/D_REPLICATE_FLAG", "not_found");
                  if (replicate == "1")
                  {
                     foreach (Stream img in replicateImages)
                     {
                        img.Dispose();
                     }
                     replicateImages.Clear();

                     //save image for future docs in envelope
                     foreach (INode pageNode in docNode.Children(level0))
                     {
                        INodeValueProvider valueProvider = pageNode.Value(this.currentTask.Step);
                        replicateImages.Add(valueProvider.Get<Stream>("CurrentImgBW", null));
                     }
                  }
                  else
                  {
                     foreach (Stream img in docImages)
                     {
                        img.Dispose();
                     }
                     docImages.Clear();
                     //add pages in doc
                     foreach (INode pageNode in docNode.Children(level0))
                     {
                        INodeValueProvider valueProvider = pageNode.Value(this.currentTask.Step);
                        docImages.Add(valueProvider.Get<Stream>("CurrentImgBW", null));

                     }

                     //commit to P8
                     DataSet dsDDI = null;
                     DataSet dsDDII = null;

                     dsDDI = da.getDD_ITEMDataSet(BatchNo, NodeID);
                     //keep track of NodeIDs process for this task to update status at finish_task
                     lstNodes.Add(NodeID);
                     string DocClass;
                     string DDItemSeq;
                     string ObjectStoreName = string.Empty;
                     string FolderName = string.Empty;
                     foreach (DataRow dr in dsDDI.Tables[0].Rows)
                     {
                        if (dr["FNP8_DOCCLASSNAME"] != null && dr["FNP8_DOCCLASSNAME"].ToString() != "")
                        {
                           DocClass = dr["FNP8_DOCCLASSNAME"].ToString();
                           DDItemSeq = dr["DD_ITEM_SEQ"].ToString();
                           this.quickModulehelper.LogMessage("Processing Node: " + NodeID + " DDItemSeq: " + DDItemSeq , LogType.Information);

                           ObjectStoreName = dr["FNP8_OBJECTSTORE"].ToString();
                           if (dr["FNP8_FOLDER"] != null)
                           {
                              FolderName = dr["FNP8_FOLDER"].ToString();
                           }
                           //commit to filenet
                           //logon to FileNet P8
                           fnP8.logon(FolderName, ObjectStoreName);
                           dsDDII = da.getDD_ITEM_INDEXESDataSet(DDItemSeq);

                           Dictionary<string, string> DictIndexValues = getIndexes(dsDDII);
                           string DocID;
                           using (Stream str = tu.JoinImages(ref docImages,ref replicateImages ))
                           {
                              str.Position = 0;
                              DocID = fnP8.createDocument(str, DictIndexValues, DocClass, DDItemSeq);
                              this.quickModulehelper.LogMessage("Committed DocID: " + DocID, LogType.Information);

                           }
                           //now that we have the docID lets update the DB
                           try
                           {
                              da.updateDocID(dr, DocID);
                           }
                           catch (Exception ex1)
                           {
                              handleError(ex1, "-266363921");
                              return;
                           }

                        }
                        else
                        {
                           throw new ArgumentNullException("-266088525; FileNet Document Class not specified. Batch: " + BatchNo + " Node: " + NodeID);
                        }
                     }
                     if (dsDDI != null)
                     {
                        dsDDI.Dispose();
                     }
                     if (dsDDII != null)
                     {
                        dsDDII.Dispose();
                     }


                  }



               }
               foreach (Stream img in docImages)
               {
                  img.Dispose();
               }
               foreach (Stream img in replicateImages)
               {
                  img.Dispose();
               }
               //update status
               foreach (INode docNode in envNode.Children(level1))
               {
                  //database call to update status
                  try
                  {
                     da.updateStatus(BatchNo, docNode.Id.ToString());
                  }
                  catch (Exception ex1)
                  {
                     handleError(ex1, "-266363921");
                     return;
                  }

               }

            }
         }
         catch (Exception ex1)
         {
            handleError(ex1, "-266363900");
            return;
         }

      }
      private Dictionary<string, string> getIndexes(DataSet DDII)
      {
         Dictionary<string, string> DictIndexValues = new Dictionary<string, string>();
         string indexName = "";
         string indexValue = "";

         foreach (DataRow dr in DDII.Tables[0].Rows)
         {
            indexName = dr["FNP8_INDEX_NAME"].ToString();
            indexValue = dr["INDEX_VALUE"].ToString();
            if (string.IsNullOrEmpty(indexName) == false && string.IsNullOrEmpty(indexValue) == false)
            {
               //if there are values in both fields, add the index
               DictIndexValues.Add(indexName, indexValue);
            }

         }


         return DictIndexValues;

      }


      //----------------------------------------------------------------
      // BeginTask is the first event fired when a task is received. It is fired before BeginNode.
      //----------------------------------------------------------------
      void processingHelper_BeginTask(object sender, BeginTaskEventsArgs e)
      {
         this.currentTask = e.Task;
         _stepValueProvider = this.currentTask.Step.Value();
         _taskValueProvider = this.currentTask.Value();

         try
         {


            //----------------------------------------------------------------
            // Load third party logon settings from the server.
            //----------------------------------------------------------------
            //----------------------------------------------------------------
            // Load the setup settings from the server.
            //----------------------------------------------------------------
            this._DSN = _stepValueProvider.Get(SetupPanel.DSN, String.Empty);
            this._DBUser = _stepValueProvider.Get(SetupPanel.DBUSER, String.Empty);
            this._DBPass = _stepValueProvider.Get(SetupPanel.DBPASS, String.Empty);
            this._FNP8URI = _stepValueProvider.Get(SetupPanel.FNP8URI, String.Empty);
            this._FNP8Domain = _stepValueProvider.Get(SetupPanel.FNP8DOMAIN, String.Empty);
            this._FNP8User = _stepValueProvider.Get(SetupPanel.FNP8USER, String.Empty);
            this._FNP8Pass = _stepValueProvider.Get(SetupPanel.FNP8PASS, String.Empty);


            try
            {
               da = new DataAccess(_DSN, _DBUser, _DBPass);
               fnP8 = new FileNetP8(_FNP8URI,_FNP8Domain,_FNP8User,_FNP8Pass);
            }
            catch (Exception ex1)
            {
               handleError(ex1, "-266363922" );
               return;
            }


         }
         catch (Exception ex2)
         {
            handleError(ex2, "-266363900");
            return;
         }

         _taskValueProvider.Set("QMResult", IA_SUCCESS.ToString());

      }
      private void handleError(Exception ex, string errNo)
      {
         _taskValueProvider.Set("QMResult", IA_ERR_UNKNOWN.ToString());
         _taskValueProvider.Set("QMErrorDesc", ex.Message);
         _taskValueProvider.Set("QMErrorNo", errNo);
         this.quickModulehelper.LogMessage(errNo + " " + ex.Message, LogType.Error);

      }

      //----------------------------------------------------------------
      // FinishTask is fired after the task node and its children (if applicable) have been processed
      // with BeginNode and FinishNode.
      //----------------------------------------------------------------
      void processingHelper_FinishTask(object sender, FinishTaskEventsArgs e)
      {
      }

      //----------------------------------------------------------------
      // BeginProduction is fired when the module is started in production mode (versus setup mode).
      //----------------------------------------------------------------
      void processingHelper_BeginProduction(object sender, EventArgs e)
      {

      }

      //----------------------------------------------------------------
      // FinishProduction is fired when the module is about to close.
      //----------------------------------------------------------------
      void processingHelper_FinishProduction(object sender, EventArgs e)
      {

      }

      #endregion



      //----------------------------------------------------------------
      // The IModuleInformation interface provides information about the module during installation
      // as a service. It is not required to implement this interface in order for your module to be
      // installed as a service.
      //----------------------------------------------------------------
      #region IModuleInformation Members

      //----------------------------------------------------------------
      // Description is used as the service description in the Services control panel. If not
      // specified the service display name is used as the description (see LongTitle).
      //----------------------------------------------------------------
      public string Description
      {
         get { return String.Empty; }
      }

      //----------------------------------------------------------------
      // LongTitle is used as the service display name in the Services control panel. If the service
      // name is specified with "serviceName" in the "-install[:serviceName]" command line parameter,
      // the service display name will appear as "LongTitle - serviceName".
      //----------------------------------------------------------------
      public string LongTitle
      {
         get { return MODULE_LONG_TITLE; }
      }

      public int ModuleId
      {
         get { return MODULE_ID; }
      }

      //----------------------------------------------------------------
      // ModuleName is used as the service name if the "-install[:serviceName]" command line parameter
      // does not include the optional "serviceName" parameter.
      //----------------------------------------------------------------
      public string ModuleName
      {
         get { return MODULE_MDF_NAME; }
      }

      public string ShortTitle
      {
         get { return MODULE_SHORT_TITLE; }
      }

      #endregion

      #region IDisposable Members

      public void Dispose()
      {

      }

      #endregion
   }
}
