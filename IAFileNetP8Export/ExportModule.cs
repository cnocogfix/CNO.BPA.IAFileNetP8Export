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
// BPA Custom
using CNO.BPA.Framework;


namespace CNO.BPA.IAFileNetP8Export
{
    //----------------------------------------------------------------
    // The IModule base interface is the main interface for QuickModule modules.
    // The IModuleInformation base interface is used to provide information when
    // installing the module as a service.
    //----------------------------------------------------------------
    public sealed class ExportModule : IModule, IModuleInformation
    {
        #region Variables
        //----------------------------------------------------------------       
        // ModuleName (MODULE_MDF_NAME) maximum of 8 characters
        // ModuleId (MODULE_ID) a unique, non-zero value obtained from EMC (zero can be used and equates to unlicensed). 
        // ModuleName (MODULE_MDF_NAME) must match the Module definition name in the MDF
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
        private IHelper _quickModulehelper = null;
        private IProcessingHelper _processingHelper = null;
        private ITask _currentTask = null; // The task currently being processed.      
        private IValueProvider _stepValueProvider = null;
        private IValueProvider _taskValueProvider = null;

        //----------------------------------------------------------------
        // Miscellaneous private members. 
        //----------------------------------------------------------------
        private Random _random = null; // Used to generate a psuedo session ID for our fake third-party system.

        //BPA custom variables
        private Export _p8Export = null;
        private DataAccess _da = null;
        private CommonParameters _cp = null;

        #endregion

        public ExportModule()
        {
            this._random = new Random(System.DateTime.Now.Millisecond);
        }

        #region BPA Custom Functions

        private void handleError(Exception ex, string errNo)
        {
            _taskValueProvider.Set("QMResult", CommonParameters.IA_ERR_UNKNOWN.ToString());
            _taskValueProvider.Set("QMErrorDesc", ex.Message);
            _taskValueProvider.Set("QMErrorNo", errNo);
            this._quickModulehelper.LogMessage(errNo + " " + ex.Message, LogType.Error);
        }

        #endregion

        #region IModule Members

        public void Initialize(IHelper helper, IModuleOptions options)
        {
            this._quickModulehelper = helper;

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
            this._processingHelper = new ProcessingHelper(options);
            options.ExternalHelpers.Add(this._processingHelper);
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
            if (!this._quickModulehelper.SetupMode)
            {
                /* 
                 * The ProcessNodeLevel property maps to the BeginNode and FinishNode events. You can choose which
                 * levels (0 to 7) receive these two events. This sample receives these events for level 0 only 
                 * because it only deals with InputFile defined at level 0 in the MDF. This module, however, can 
                 * be defined as a Step at any level in an IPP.
                 */
                this._processingHelper.ProcessNodeLevel[0] = false;
                this._processingHelper.ProcessNodeLevel[1] = true;
                this._processingHelper.ProcessNodeLevel[2] = false;
                this._processingHelper.ProcessNodeLevel[3] = true;
                this._processingHelper.ProcessNodeLevel[4] = false;
                this._processingHelper.ProcessNodeLevel[5] = false;
                this._processingHelper.ProcessNodeLevel[6] = false;
                this._processingHelper.ProcessNodeLevel[7] = false;

                /*
                 * Select which operations are shown on the left-hand side of the main production window. By default,
                 * no operations are provided. See Emc.InputAccel.QuickModule.Helpers.Processing.OperationType for
                 * details on each operation type.
                 */
                this._processingHelper.Operations[OperationType.RunSingleBatch] = true;
                this._processingHelper.Operations[OperationType.RunAllBatches] = true;
                this._processingHelper.Operations[OperationType.Exit] = true;
                this._processingHelper.Operations[OperationType.Stop] = true;

                /*
                  * Processing Helper Event Handlers
                  * 
                  * Define and register event handlers for Processing Helper events. 
                  * 
                  * Depending on the needs of your module, you may not handle all events. For example, this module
                  * only needs BeginTask and BeginNode.
                  */
                this._processingHelper.BeginTask += new EventHandler<BeginTaskEventsArgs>(processingHelper_BeginTask);
                this._processingHelper.BeginNode += new EventHandler<BeginNodeEventsArgs>(processingHelper_BeginNode);
                this._processingHelper.FinishTask += new EventHandler<FinishTaskEventsArgs>(processingHelper_FinishTask);
                this._processingHelper.BeginProduction += new EventHandler<EventArgs>(processingHelper_BeginProduction);
                this._processingHelper.FinishProduction += new EventHandler<EventArgs>(processingHelper_FinishProduction);
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

            IPanel panel = new SetupPanel(this._quickModulehelper, e.Step);
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
                    string result = _p8Export.ProcessEnvelope(e);
                    if (result == "SUCCESS")
                    {
                        _taskValueProvider.Set("QMResult", CommonParameters.IA_SUCCESS.ToString());

                    }
                    else
                    {
                        _taskValueProvider.Set("QMResult", CommonParameters.IA_ERR_UNKNOWN.ToString());
                        _taskValueProvider.Set("QMErrorDesc", result);
                        _taskValueProvider.Set("QMErrorNo", "-266367810");
                        _quickModulehelper.LogMessage("-266367810;" + result, LogType.Error);
                    }
                }
            }
            catch (Exception ex1)
            {
                handleError(ex1, "-266367805");
                return;
            }

        }

        //----------------------------------------------------------------
        // BeginTask is the first event fired when a task is received. It is fired before BeginNode.
        //----------------------------------------------------------------
        void processingHelper_BeginTask(object sender, BeginTaskEventsArgs e)
        {
            try
            {
                //makes sense to grab a local handle for the current task
                this._currentTask = e.Task;
                //once we have the task, pull back the value provider so we can pull back the stored values.
                _stepValueProvider = this._currentTask.Step.Value();
                _taskValueProvider = this._currentTask.Value();
                //we need to create a new object to hold our common parameters that we need to pass around
                _cp = new CommonParameters();
                //next we'll set all the values
                _cp.DSN = _stepValueProvider.Get(SetupPanel.DSN, String.Empty);
                _cp.DBUser = _stepValueProvider.Get(SetupPanel.DBUSER, String.Empty);
                _cp.DBPass = _stepValueProvider.Get(SetupPanel.DBPASS, String.Empty);
                _cp.FNP8URI = _stepValueProvider.Get(SetupPanel.FNP8URI, String.Empty);
                _cp.FNP8Domain = _stepValueProvider.Get(SetupPanel.FNP8DOMAIN, String.Empty);
                _cp.FNP8User = _stepValueProvider.Get(SetupPanel.FNP8USER, String.Empty);
                _cp.FNP8Pass = _stepValueProvider.Get(SetupPanel.FNP8PASS, String.Empty);
                ////the user credentials are stored encrypted, so thos now need to be decrypted
                //Cryptography crypto = new Cryptography();
                //_cp.DBUser = crypto.Decrypt(_cp.DBUser);
                //_cp.DBPass = crypto.Decrypt(_cp.DBPass);
                //_cp.FNP8User = crypto.Decrypt(_cp.FNP8User);
                //_cp.FNP8Pass = crypto.Decrypt(_cp.FNP8Pass);
            }
            catch (Exception ex1)
            {
                handleError(ex1, "-266367801");
                throw;
            }
            try
                {
                    //we only want to connect to P8 one time per execution
                    CNO.BPA.FNP8.IUserConnection uc = new CNO.BPA.FNP8.UserConnection();
                    uc.logon(_cp.FNP8URI, _cp.FNP8Domain, _cp.FNP8User, _cp.FNP8Pass);
                    //pass the connection out to our common paramenters object
                    _cp.UserConnection = uc;
                    //now connect to our db
                    _da = new DataAccess(_cp.DSN, _cp.DBUser, _cp.DBPass);
                    //and pass that out to our common parameters object
                    _cp.DbConnection = _da;
                    //now we can create a new export object
                    _p8Export = new Export(this._quickModulehelper, this._currentTask, this._cp);
                }

            catch (Exception ex2)
            {
                handleError(ex2, "-266367803");
                return;
            }

            _taskValueProvider.Set("QMResult", CommonParameters.IA_SUCCESS.ToString());

        }

        //----------------------------------------------------------------
        // FinishTask is fired after the task node and its children (if applicable) have been processed
        // with BeginNode and FinishNode.
        //----------------------------------------------------------------

        void processingHelper_FinishTask(object sender, FinishTaskEventsArgs e)
        {
        }

        #region unused
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
