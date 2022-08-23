// Copyright © 2003–2009 EMC Corporation. All rights reserved.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
//----------------------------------------------------------------
// InputAccel namespaces - QuickModule and Client.
//----------------------------------------------------------------
using Emc.InputAccel.QuickModule;
using Emc.InputAccel.Workflow.Client;

namespace CNO.BPA.IAFileNetP8Export
{
    //----------------------------------------------------------------
    // A QuickModule setup panel must implement the IPanel interface as shown below.
    // IPanelTitle provides a panel title of the same text style as other QuickModule panels.
    //----------------------------------------------------------------
    public partial class SetupPanel : UserControl, IPanel, IPanelTitle
    {
        //---------------------------------------------------------------- 
        // IAValue names used to store file name and path properties.
        //----------------------------------------------------------------
        public static readonly string DBUSER = "_DBUser";
        public static readonly string DBPASS= "_DBPass";
        public static readonly string DSN = "_DSN";

        public static readonly string FNP8USER = "_p8username";
        public static readonly string FNP8PASS = "_p8password";
        public static readonly string FNP8URI = "_p8uri";
        public static readonly string FNP8DOMAIN = "_p8domainName";



        //----------------------------------------------------------------
        // The following private members reference various objects during the
        // lifetime of this setup panel.
        //----------------------------------------------------------------
        private IStep step = null;
        private IHelper helper = null;
        private IPanelsSheet panelsSheet = null;
        private IValueProvider valueProvider = null;

        private CNO.BPA.Framework.Cryptography crypto = new CNO.BPA.Framework.Cryptography();
        //----------------------------------------------------------------
        // The default constructor is modified to take two parameters, IHelper helper
        // and IStep step, that are referenced during the lifetime of the panel.
        //----------------------------------------------------------------
        public SetupPanel(IHelper helper, IStep step)
        {
            InitializeComponent();

            this.helper = helper;
            this.step = step;
            this.valueProvider = step.Value();

            this.txtPassword.TextChanged += new EventHandler(ValueChanged);
            this.txtUserID.TextChanged += new EventHandler(ValueChanged);
            this.txtDSN.TextChanged += new EventHandler(ValueChanged);
            this.txtP8DomainName.TextChanged += new EventHandler(ValueChanged);
            this.txtP8URI.TextChanged += new EventHandler(ValueChanged);
            this.txtP8User.TextChanged += new EventHandler(ValueChanged);
            this.txtP8Password.TextChanged += new EventHandler(ValueChanged);




        }

        #region IPanel Members

        //----------------------------------------------------------------
        // ApplyChanges() is called after the user clicks the OK or Apply button and ValidateContent()
        // has returned TRUE (validated) for all panels. Therefore, data is assumed to be valid and 
        // must be saved in ApplyChanges().
        //----------------------------------------------------------------
        public void ApplyChanges()
        {
            //----------------------------------------------------------------
            // Copy the desired image naming properties over to the server.
            //----------------------------------------------------------------
            this.valueProvider.Set(DSN , this.txtDSN.Text);
            this.valueProvider.Set(DBUSER , crypto.Encrypt(this.txtUserID.Text));
            this.valueProvider.Set(DBPASS , crypto.Encrypt(this.txtPassword.Text));

            this.valueProvider.Set(FNP8USER, crypto.Encrypt(this.txtP8User.Text));
            this.valueProvider.Set(FNP8PASS, crypto.Encrypt(this.txtP8Password.Text));
            this.valueProvider.Set(FNP8URI, this.txtP8URI.Text);
            this.valueProvider.Set(FNP8DOMAIN, this.txtP8DomainName.Text);



        }

        //----------------------------------------------------------------
        // OnKillActive() is called when the panel is active and
        // 
        // 1) The panel is about to become inactive because the setup window is switching to another
        //    setup panel, or
        // 2) The Cancel button is clicked.
        // 
        // Returns TRUE to leave the panel, otherwise FALSE to remain on the panel.
        //----------------------------------------------------------------
        public bool OnKillActive()
        {
            return true;
        }

        //----------------------------------------------------------------
        // OnLoad() is called one time during intialization. Setup values should be fetched and the
        // user interface populated.
        //----------------------------------------------------------------
        public void OnLoad(IPanelsSheet Sheet)
        {
            this.txtDSN.Text = this.valueProvider.Get(DSN, String.Empty);
            this.txtP8DomainName.Text = this.valueProvider.Get(FNP8DOMAIN, String.Empty);
            this.txtP8URI.Text = this.valueProvider.Get(FNP8URI, String.Empty);




            if (this.valueProvider.Get(DBUSER, String.Empty).Length > 0)
            {
               this.txtUserID.Text = crypto.Decrypt(this.valueProvider.Get(DBUSER, String.Empty));
            }
            if (this.valueProvider.Get(DBPASS, String.Empty).Length > 0)
            {
               this.txtPassword.Text = crypto.Decrypt(this.valueProvider.Get(DBPASS, String.Empty));
            }
            if (this.valueProvider.Get(FNP8USER, String.Empty).Length > 0)
            {
               this.txtP8User.Text = crypto.Decrypt(this.valueProvider.Get(FNP8USER, String.Empty));
            }
            if (this.valueProvider.Get(FNP8PASS, String.Empty).Length > 0)
            {
               this.txtP8Password.Text = crypto.Decrypt(this.valueProvider.Get(FNP8PASS, String.Empty));
            }

            this.panelsSheet = Sheet;
        }

        //----------------------------------------------------------------
        // OnSetActive() is called when the panel has become the active panel.
        //----------------------------------------------------------------
        public void OnSetActive()
        {
        }

        //----------------------------------------------------------------
        // OnUnload() is called after ApplyChanges() is called (triggered by the user clicking the OK
        // button), or after OnKillActive() is called (triggered by the user clicking the Cancel button).
        //----------------------------------------------------------------
        public void OnUnload()
        {
        }

        //----------------------------------------------------------------
        // The PanelIcon (optional, can be null) is displayed on the navigation option on the setup window.
        //----------------------------------------------------------------
        public Image PanelIcon
        {
            get { return null; }
        }

        //----------------------------------------------------------------
        // The PanelName is displayed on the navigation option on the setup window.
        // 
        // TODO: Set the panel name to a word or phrase that describes the feature(s) the user will set up.
        // Add an ampersand sign '&' before a letter in the panel name to create a keyboard shortcut. Make
        // sure it does not conflict with other shortcut keys on your panel or other QuickModule panels you
        // are using.
        //----------------------------------------------------------------
        public string PanelName
        {
            get { return Resource.SetupPanelName; }
        }

        //----------------------------------------------------------------
        // The Subpanel (optional, can be null) is displayed on the left-hand side of the setup window under
        // the navigation options.
        //----------------------------------------------------------------
        public Control Subpanel
        {
            get { return null; }
        }

        //----------------------------------------------------------------
        // ValidateContent() is called to verify if the data on the panel is valid. Returns TRUE if the data is
        // valid, or FALSE if the data is invalid.
        // 
        // ValidateContent() should not save or modify data, however it is appropriate to use IHelper.PopupError
        // to notify the user of invalid or missing data.
        //----------------------------------------------------------------
        public bool ValidateContent()
        {
            bool valid = true;

            //----------------------------------------------------------------
            // The only required settings are the name of the folder and the base name of the images.
            // Display an error popup next to the UI element if the values are not present.
            //----------------------------------------------------------------
            //if (String.IsNullOrEmpty(this.fileName.Text))
            //{
            //    valid = false;
            //    this.helper.PopupError(Resource.ValidationFailed, Resource.ValidationEmpty, this.fileName);
            //}

            //if (String.IsNullOrEmpty(this.filePath.Text))
            //{
            //    valid = false;
            //    this.helper.PopupError(Resource.ValidationFailed, Resource.ValidationEmpty, this.filePath);
            //}

            return valid;
        }

        #endregion

        #region IPanelTitle Members

        //----------------------------------------------------------------
        // The PanelTitle property is displayed above the panel. The text is formatted in the same style
        // as other QuickModule panels such as Information, Error, and Scripting.
        // 
        // TODO: Set the panel title to a word or phrase that describes the feature(s) of the panel. In most
        // cases this will be the same as IPanel.PanelName (without the ampersand sign).
        //----------------------------------------------------------------
        public string PanelTitle
        {
            get { return Resource.SetupPanelTitle; }
        }

        #endregion


        //----------------------------------------------------------------
        // Helper function to display a dialog to allow the user to pick an IA Value.
        // Returns a value name in the format STEP.VALUE or an empty string if the user 
        // canceled the dialog.
        //----------------------------------------------------------------
        private string GetIAValue()
        {
            IVarListDialog valueDialog = this.helper.QuickModuleDialog.Get(typeof(IVarListDialog)) as IVarListDialog;

            if (valueDialog != null)
            {
                valueDialog.BatchProcessId = this.step.BatchProcess;
                valueDialog.ShowOnlyValues(VarListDialogValueTypes.Value);
                if (valueDialog.ShowVarListDialog(step) == DialogResult.OK)
                {
                    return valueDialog.SelectedValue;
                }
                else
                {
                    return String.Empty;
                }
            }
            else
            {
                throw new NullReferenceException();
            }
        }



        void ValueChanged(object sender, EventArgs e)
        {
            SetModified();
        }



        //----------------------------------------------------------------
        // Helper function to set the modified flag on the IPanelsSheet object. When the modified flag
        // is set the Apply button on the user interface becomes enabled.
        //----------------------------------------------------------------
        void SetModified()
        {
            if (this.panelsSheet != null)
            {
                this.panelsSheet.SetModified();
            }
        }

        private void SetupPanel_Load(object sender, EventArgs e)
        {

        }


    }
}
