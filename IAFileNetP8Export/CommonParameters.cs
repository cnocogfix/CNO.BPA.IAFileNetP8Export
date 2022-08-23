using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using CNO.BPA.FNP8;
using FileNet.Api.Security;

namespace CNO.BPA.IAFileNetP8Export
{
    public class CommonParameters: IDisposable
    {
        #region Variables
        //Error Cosntants
        public const int IA_SUCCESS = 0;
        public const int IA_ERR_UNKNOWN = -4523;
        public const int IA_ERR_NOFUNC = -4518;
        public const int IA_ERR_CANCEL = -4526;
        public const int IA_ERR_NORETRY = -6112;
        public const int IA_ERR_RETRYSOME = -6113;
        public const int IA_ERR_RETRY = -6114;
        public const int IA_ERR_ACCESS = -4505;
        //strings
        private string _dsn = String.Empty;
        private string _dbUser = String.Empty;
        private string _dbPass = String.Empty;
        private string _fnp8Domain = String.Empty;
        private string _fnp8User = String.Empty;
        private string _fnp8Pass = String.Empty;
        private string _fnp8Uri = String.Empty;
        private string _batchNo = String.Empty;
        private string _nodeId = String.Empty;
        private string _tiffType = String.Empty;
        private string _fileExtension = String.Empty;
        private string _folderPath = String.Empty;
        //objects
        private IUserConnection _userConnection = null;
        private DataAccess _da = null;
        

        #endregion

        #region Properties
        public string DSN
        {
            get { return _dsn; }
            set { _dsn = value; }
        }
        public string DBUser
        {
            get { return _dbUser; }
            set { _dbUser = value; }
        }
        public string DBPass
        {
            get { return _dbPass; }
            set { _dbPass = value; }
        }
        public string FNP8Domain
        {
            get { return _fnp8Domain; }
            set { _fnp8Domain = value; }
        }
        public string FNP8User
        {
            get { return _fnp8User; }
            set { _fnp8User = value; }
        }
        public string FNP8Pass
        {
            get { return _fnp8Pass; }
            set { _fnp8Pass = value; }
        }
        public string FNP8URI
        {
            get { return _fnp8Uri; }
            set { _fnp8Uri = value; }
        }
        public string BatchNo
        {
            get { return _batchNo; }
            set { _batchNo = value; }
        }
        public string NodeId
        {
            get { return _nodeId; }
            set { _nodeId = value; }         
        }
        public string TiffType
        {
            get { return _tiffType; }
            set { _tiffType = value; }
        }
        public string FileExtension
        {
            get { return _fileExtension; }
            set { _fileExtension = value; }
        }
        public string FolderPath
        {
            get { return _folderPath; }
            set { _folderPath = value; }
        }
        public IUserConnection UserConnection
        {
            get { return _userConnection; }
            set { _userConnection = value; }
        }
        public DataAccess DbConnection
        {
            get { return _da; }
            set { _da = value; }
        }
        #endregion

        public void Dispose()
        {
            
         _dsn = String.Empty;
         _dbUser = String.Empty;
         _dbPass = String.Empty;
        _fnp8Domain = String.Empty;
        _fnp8User = String.Empty;
        _fnp8Pass = String.Empty;
        _fnp8Uri = String.Empty;
        _batchNo = String.Empty;
        _nodeId = String.Empty;
            _da = null;
            _userConnection = null;
        }
    }
}
