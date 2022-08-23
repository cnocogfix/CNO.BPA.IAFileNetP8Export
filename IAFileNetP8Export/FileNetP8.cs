using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using FileNet.Api.Collection;
using FileNet.Api.Constants;
using FileNet.Api.Core;
using FileNet.Api.Property;
using FileNet.Api.Admin;
using FileNet.Api.Meta;
using FileNet.Api.Util;
using FileNet.Api.Authentication;



namespace CNO.BPA.IAFileNetP8Export
{
    class FileNetP8
    {
        IConnection _conn = null;
        IDomain _domain = null;
        IObjectStore _objectStore = null;
        IFolder _folder = null;

        string _userName = string.Empty;
        string _password = string.Empty;
        string _uri = string.Empty;
        string _domainName = string.Empty;
        private Framework.Cryptography crypto = new Framework.Cryptography();


        public FileNetP8(string uri, string domain, string user, string pass)
        {
            _userName = crypto.Decrypt(user);
            _password = crypto.Decrypt(pass);
            _uri = uri;
            _domainName = domain;

        }

        public void logon(string FolderName, string ObjectStoreName)
        {
            try
            {
                if (null == _conn)
                {
                    _conn = getConnectionEDU(_uri, _userName, _password);
                    _domain = getDomainEDU(_conn, _domainName);
                }
                if (_objectStore == null || ObjectStoreName != _objectStore.Name)
                {
                    _objectStore = getObjectStoreEDU(_domain, ObjectStoreName);
                    _folder = null;
                }
                if (FolderName.Length == 0)
                {
                    _folder = null;
                }
                else if (_folder == null || FolderName != _folder.PathName)
                {
                    _folder = getFolderEDU(_objectStore, FolderName);
                }

            }
            catch (Exception ex)
            {
                throw new Exception("FileNetP8.logon: " + ex.Message);
            }


        }
        public IConnection getConnectionEDU(string uri, string user, string pword)
        {
            try
            {
                UsernameCredentials creds = new UsernameCredentials(user, pword);
                ClientContext.SetProcessCredentials(creds);

                IConnection conn = Factory.Connection.GetConnection(uri);
                return conn;
            }
            catch (Exception ex)
            {
                throw new Exception("FileNetP8.getConnectionEDU: " + ex.Message);
            }
        }
        public IDomain getDomainEDU(IConnection conn, string domainName)
        {
            try
            {
                IDomain domain = null;
                domain = Factory.Domain.FetchInstance(conn, domainName, null);
                return domain;
            }
            catch (Exception ex)
            {
                throw new Exception("FileNetP8.getDomainEDU: " + ex.Message);
            }
        }


        public IObjectStore getObjectStoreEDU(IDomain domain, string objectStoreName)
        {
            try
            {

                IObjectStore store = null;
                store = Factory.ObjectStore.FetchInstance(domain, objectStoreName, null);
                return store;
            }
            catch (Exception ex)
            {
                throw new Exception("FileNetP8.getObjectStoreEDU: " + ex.Message);
            }

        }

        public IFolder getFolderEDU(IObjectStore store, string folderName)
        {
            try
            {
                IFolder folder = null;
                folder = Factory.Folder.FetchInstance(store, folderName, null);
                folderName = folder.FolderName;
                return folder;
            }
            catch (Exception ex)
            {
                throw new Exception("FileNetP8.getFolderEDU: " + ex.Message);
            }

        }

        public string createDocument(Stream docStream, Dictionary<string, string> indexValues, string DocClass, string DDItemSeq)
        {
            IDocument myDoc = null;
            try
            {

                myDoc = Factory.Document.CreateInstance(_objectStore, DocClass);
                IContentElementList contentList = null;
                contentList = Factory.ContentElement.CreateList();
                IContentTransfer content = null;
                content = Factory.ContentTransfer.CreateInstance();
                string fileName = DDItemSeq + ".tif";

                content.SetCaptureSource(docStream);
                content.RetrievalName = fileName; //needed for downloading with name from workplace
                contentList.Add(content);

                myDoc.ContentElements = contentList;
                myDoc.Checkin(AutoClassify.DO_NOT_AUTO_CLASSIFY, CheckinType.MAJOR_VERSION);


                IClassDescription myClassDesc = null;
                myClassDesc = Factory.ClassDescription.FetchInstance(_objectStore, DocClass, null);
                //add indexes to properties
                IProperties properties = myDoc.Properties;
                properties["DocumentTitle"] = fileName;
                foreach (string indexName in indexValues.Keys)
                {
                    object indexValue = indexValues[indexName];
                    Cardinality propCardinality = new FileNet.Api.Constants.Cardinality();
                    bool validIndex = validProperty(indexName, ref indexValue, myClassDesc.PropertyDescriptions, ref propCardinality);
                    if (validIndex == true)
                    {

                        if (propCardinality == Cardinality.LIST)
                        {
                            IStringList listValue = Factory.StringList.CreateList();
                            //assume indexValue is a pipe delimited string
                            string[] values = indexValue.ToString().Split(new Char[]{'|'});
                            foreach (string value in values)
                            {
                                listValue.Add(value);
                            }
                            properties[indexName] = listValue;
                        }
                        else
                        {
                            //Index value passed validation
                            properties[indexName] = indexValue;
                        }

                    }
                }


                myDoc.MimeType = "image/tiff";
                myDoc.Save(RefreshMode.REFRESH);



                //file to folder if needed
                if (null != _folder)
                {
                    IReferentialContainmentRelationship rel = null;
                    rel = _folder.File(myDoc, AutoUniqueName.AUTO_UNIQUE, fileName, DefineSecurityParentage.DEFINE_SECURITY_PARENTAGE);
                    rel.Save(RefreshMode.NO_REFRESH);
                }
                return myDoc.Id.ToString();

            }
            catch (Exception ex)
            {
                try
                {
                    if (null != myDoc)
                    {
                        myDoc.Delete();
                    }
                }
                catch { }
                throw new Exception("FileNetP8.createDocument: " + ex.Message);

            }
        }
        private bool validProperty(string indexName, ref object indexValue, IPropertyDescriptionList propertyDescs, ref Cardinality propCardinality)
        {
            try
            {
                bool foundProperty = false;
                foreach (IPropertyDescription propDesc in propertyDescs)
                {
                    if (propDesc.SymbolicName == indexName)
                    {
                        foundProperty = true;
                        propCardinality = propDesc.Cardinality;
                        switch (propDesc.DataType)
                        {
                            #region DataTypes
                            case TypeID.DATE:
                                {
                                    #region DataType = Date
                                    DateTime parsedDate;
                                    if (DateTime.TryParse(indexValue.ToString(), out parsedDate))
                                    {
                                        parsedDate = parsedDate.ToUniversalTime();
                                        indexValue = parsedDate;
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                    #endregion
                                }

                            case TypeID.BOOLEAN:
                                {
                                    #region DataType = boolean
                                    bool parsedBool;
                                    if (Boolean.TryParse(indexValue.ToString(), out parsedBool))
                                    {
                                        indexValue = parsedBool;
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                    #endregion
                                }
                            case TypeID.DOUBLE:
                                {
                                    #region DataType = double
                                    double parsedDouble;
                                    if (Double.TryParse(indexValue.ToString(), out parsedDouble))
                                    {
                                        indexValue = parsedDouble;
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                    #endregion
                                }
                            case TypeID.LONG:
                                {
                                    #region DataType = long
                                    long parsedLong;
                                    if (long.TryParse(indexValue.ToString(), out parsedLong))
                                    {
                                        indexValue = parsedLong;
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                    #endregion
                                }
                            case TypeID.STRING:
                                {
                                    #region DataType = string
                                    //trim value to size if needed
                                    IPropertyDescriptionString propDescString = (IPropertyDescriptionString)propDesc;
                                    int maxLength = 0;
                                    int.TryParse(propDescString.MaximumLengthString.ToString(), out maxLength);
                                    if (maxLength > 0 && indexValue.ToString().Length > maxLength)
                                    {
                                        indexValue = indexValue.ToString().Substring(0, maxLength);
                                    }
                                    //check if it is a menu...
                                    if (propDesc.ChoiceList != null)
                                    {
                                        //if a choice list exists we should ensure the value is contained within the list
                                        for (int i = 0; i < propDesc.ChoiceList.ChoiceValues.Count; i++)
                                        {
                                            IChoice choice = (IChoice)propDesc.ChoiceList.ChoiceValues[i];
                                            if (propCardinality == Cardinality.LIST)
                                            //assume the value will be pipe delimited
                                            {
                                                string[] values = indexValue.ToString().Split(new Char[] { '|' });
                                                foreach (string value in values)
                                                {
                                                    if (choice.ChoiceStringValue == value)
                                                    {
                                                        return true;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (choice.ChoiceStringValue == indexValue.ToString())
                                                {
                                                    return true;
                                                }
                                            }
                                        }
                                        //if we make it here we did not find a match
                                        return false;
                                    }
                                    else
                                    {
                                        return true;
                                    }
                                    #endregion
                                }
                            #endregion
                        }
                        break;
                    }
                }

                if (foundProperty == false)
                {
                    return false;
                }

                return true;

            }
            catch (Exception ex)
            {
                throw new Exception("FileNetP8.validProperty: " + ex.Message);

            }

        }

    }
}
