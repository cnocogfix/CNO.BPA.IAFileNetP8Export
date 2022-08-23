using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;

//----------------------------------------------------------------
// InputAccel namespaces - QuickModule, Processing Helper,
// Script Helper, Scripting interface, and Workflow Client.
//----------------------------------------------------------------
using Emc.InputAccel.QuickModule;
using Emc.InputAccel.QuickModule.Helpers.Processing;
using Emc.InputAccel.QuickModule.Plugins.Processing;
using Emc.InputAccel.Workflow.Client;

//BPA Custom
using CNO.BPA.FNP8;

namespace CNO.BPA.IAFileNetP8Export
{

    internal class Export
    {
        //IA Objects        
        private IHelper _helper = null;
        private ITask _task = null;
        private IStep _step = null;
        private IValueProvider _taskValueProvider = null;
        //bpa custom objects
        private IUserConnection _userConnection = null;
        
        private CommonParameters _cp = null;

        public Export(IHelper helper, ITask task, CommonParameters commonParameters)
        {
            _helper = helper;
            _task = task;
            _step = _task.Step;
            _cp = commonParameters;
        }

        public string ProcessEnvelope(BeginNodeEventsArgs e)
        {
            try
            {               
                MemoryStream[] docImages = new MemoryStream[1];
                MemoryStream[] replicateImages = new MemoryStream[1];
                Stream inputFile = null;

                //pull a copy of the node into a local object
                INode envNode = e.Node;
                //Now grab a copy of the level 1 node and the level 0 node for use later
                ILevel level1 = _task.Batch.Level(1);
                ILevel level0 = _task.Batch.Level(0);
                //pull back the batch number
                _cp.BatchNo = envNode.Tree.Batch.Name;
                //loop through each document(1) node within the envelope(3) node
                foreach (INode docNode in envNode.Children(level1))
                {
                    //grab the current document node
                    _cp.NodeId = docNode.Id.ToString(CultureInfo.InvariantCulture);
                    INodeValueProvider nodeValueProvider = docNode.Value(_task.Step);
                    string replicate = nodeValueProvider.Get("$instance=Standard_MDF/D_REPLICATE_FLAG", "not_found");
                    try
                    {
                        inputFile = nodeValueProvider.Get<Stream>("Level1_InputFile", null);
                    }
                    catch (Exception ex2)
                    {
                      //if an error occurred trying to pull in the stream, for now we will ignore it and continue
                    }

                    if (replicate == "1")
                    {
                        foreach (MemoryStream img in replicateImages)
                        {
                            if (null != img)
                            {
                                img.Dispose();
                            }
                        }                                   
                        //now we need to handle whether or not the export should use the node or a passed in file
                        if (null != inputFile)
                        {
                            MemoryStream memStream1 = new MemoryStream();  
                            replicateImages = new MemoryStream[1];
                            memStream1.SetLength(inputFile.Length);
                            inputFile.Read(memStream1.GetBuffer(), 0, (int)inputFile.Length);
                            replicateImages[0] = memStream1;
                            //memStream1.Dispose();
                        }
                        else
                        {
                            replicateImages = new MemoryStream[docNode.ChildCount(level0)];
                            int countR = 0;
                            //save image for future docs in envelope
                            foreach (INode pageNode in docNode.Children(level0))
                            {
                                MemoryStream memStream2 = new MemoryStream();                              
                                INodeValueProvider valueProvider = pageNode.Value(_task.Step);
                                Stream rStream = valueProvider.Get<Stream>("CurrentImgBW", null);
                                memStream2.SetLength(rStream.Length);
                                rStream.Read(memStream2.GetBuffer(), 0, (int)rStream.Length);
                                replicateImages[countR] = memStream2;
                                countR++;
                                //memStream2.Dispose();
                            }
                        }
                    }
                    else
                    {
                        foreach (MemoryStream img in docImages)
                        {
                            if (null != img)
                            {
                                img.Dispose();
                            }
                        }
                        //now we need to handle whether or not the export should use the node or a passed in file
                        if (null != inputFile)
                        {
                            MemoryStream memStream3 = new MemoryStream();   
                            docImages = new MemoryStream[1];
                            memStream3.SetLength(inputFile.Length);
                            inputFile.Read(memStream3.GetBuffer(), 0, (int)inputFile.Length);
                            docImages[0] = memStream3;
                            //memStream3.Dispose();
                        }
                        else
                        {
                            docImages = new MemoryStream[docNode.ChildCount(level0)];
                            int countD = 0;
                            //add pages in doc
                            foreach (INode pageNode in docNode.Children(level0))
                            {
                                MemoryStream memStream4 = new MemoryStream();
                                INodeValueProvider valueProvider = pageNode.Value(_task.Step);
                                Stream dStream = valueProvider.Get<Stream>("CurrentImgBW", null);
                                memStream4.SetLength(dStream.Length);
                                dStream.Read(memStream4.GetBuffer(), 0, (int)dStream.Length);
                                docImages[countD] = memStream4;
                                countD++;
                                //memStream4.Dispose();
                            }
                        }                     
                        //now we are ready to commit the document                       
                        string returnValue = commitDocument(ref docImages, replicateImages);

                        foreach (MemoryStream img in docImages)
                        {
                            if (null != img)
                            {
                                img.Dispose();
                            }
                        }

                        if (returnValue != "SUCCESS")
                        {
                            return returnValue;
                        }

                        GC.Collect();
                    }
                }
                
                foreach (MemoryStream img in replicateImages)
                {
                    if (null != img)
                    {
                        img.Dispose();
                    }
                }
               
                //update status
                foreach (INode docNode in envNode.Children(level1))
                {
                    //database call to update status
                    try
                    {
                        _cp.DbConnection.UpdateStatus(_cp.BatchNo, docNode.Id.ToString());
                    }
                    catch (Exception ex1)
                    {
                        handleError(ex1, "-266363921");
                        return ex1.Message;
                    }
                }
                
                return "SUCCESS";
            }
            catch (Exception ex)
            {
                handleError(ex, "-266367832");
                return ex.Message;               
            }
        }
        private string commitDocument(ref MemoryStream[] DocImages, MemoryStream[] ReplicateImages)        
        {
            List<string> lstNodes = new List<string>();
            //List<Stream> documentStreams = new List<Stream>();
            MemoryStream[] docStreams = null;
            TiffUtility tu = new TiffUtility();

            try
            {
                //commit to P8
                DataSet dsDDI = null;
                DataSet dsDDII = null;

                dsDDI = _cp.DbConnection.getDD_ITEMDataSet(_cp.BatchNo, _cp.NodeId);

                if (dsDDI.Tables[0].Rows.Count == 0)
                {
                    return "No records returned for batch number (" + _cp.BatchNo + ") and Node Id (" + _cp.NodeId + ").";
                }
                else if (dsDDI.Tables[0].Rows[0]["FNP8_DOCID"].ToString().Length > 0)
                {
                    return "SUCCESS";
                }

                //keep track of NodeIDs process for this task to update status at finish_task
                lstNodes.Add(_cp.NodeId);
               // loop through each dd item record and commit the document with the specified indexes.
                foreach (DataRow dr in dsDDI.Tables[0].Rows)
                {
                    if (dr["FNP8_DOCCLASSNAME"] != null && dr["FNP8_DOCCLASSNAME"].ToString() != "")
                    {
                        string docClass = dr["FNP8_DOCCLASSNAME"].ToString();
                        string ddItemSeq = dr["DD_ITEM_SEQ"].ToString();
                        _helper.LogMessage("Processing Node: " + _cp.NodeId + " DDItemSeq: " + ddItemSeq, LogType.Information);

                        string objectStoreName = dr["FNP8_OBJECTSTORE"].ToString();
                        if (dr["FNP8_FOLDER"] != null)
                        {
                            _cp.FolderPath = dr["FNP8_FOLDER"].ToString();
                        }
                        
                        dsDDII = _cp.DbConnection.getDD_ITEM_INDEXESDataSet(ddItemSeq);

                        Dictionary<string, string> dictIndexValues = getIndexes(dsDDII);
                        string DocID;
                        if (null != ReplicateImages[0])                        
                        {
                            docStreams = new MemoryStream[ReplicateImages.Count() + DocImages.Count()];
                            int count = 0;
                            foreach (MemoryStream dStream in DocImages)
                            {
                                docStreams[count] = dStream;
                                count++;
                            }
                            foreach (MemoryStream rStream in ReplicateImages)
                            {
                                docStreams[count] = rStream;
                                count++;
                            }
                        }
                        else
                        {
                            docStreams = new MemoryStream[DocImages.Count()];
                            int count = 0;
                            foreach (MemoryStream dStream in DocImages)
                            {
                                docStreams[count] = dStream;
                                count++;
                            }
                        }
                        FNP8.IDocInfo myDocInfo = new DocInfo();
                        FNP8.IDocCreate myDocCreate = new DocCreate();
                        myDocInfo.DocumentClassName = docClass;
                        myDocInfo.ObjectStore = objectStoreName;
                        switch (_cp.TiffType)
                        {
                            case "SINGLE":
                                myDocInfo.IsMulti = false;
                                break;
                            case "MULTI":
                                myDocInfo.IsMulti = true;
                                break;
                            default:
                                myDocInfo.IsMulti = false;
                                break;
                        }
                        myDocInfo.RetrievalName = _cp.BatchNo + "_" + _cp.NodeId;
                        myDocInfo.FolderPath = _cp.FolderPath;
                        myDocInfo.Extension = _cp.FileExtension.Length > 0 ? _cp.FileExtension : "tif";
                        myDocInfo.Properties = dictIndexValues;
                        myDocInfo.Properties.Add("DocumentTitle", myDocInfo.RetrievalName);
                        myDocCreate.createDocument(docStreams, _cp.UserConnection, myDocInfo);

                        DocID = myDocInfo.VersionSeriesID;

                            _helper.LogMessage("Committed DocID: " + DocID, LogType.Information);
                        //now that we have the docID lets update the DB
                        try
                        {
                            _cp.DbConnection.UpdateDocId(dr, DocID);
                        }
                        catch (Exception ex1)
                        {
                            handleError(ex1, "-266363921");
                            return ex1.Message;
                        }

                    }
                    else
                    {
                        throw new ArgumentNullException("-266088525; FileNet Document Class not specified. Batch: " + _cp.BatchNo + " Node: " + _cp.NodeId);
                    }
                }
                //foreach (MemoryStream stream in docStreams)
                //{
                  //  stream.Dispose();
                //}                
                
                dsDDI.Dispose();

                if (dsDDII != null)
                {
                    dsDDII.Dispose();
                }

                return "SUCCESS";
            }
            catch (Exception ex)
            {
                return ex.Message;
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
        private void handleError(Exception ex, string errNo)
        {
            _taskValueProvider.Set("QMResult", "-4523");
            _taskValueProvider.Set("QMErrorDesc", ex.Message);
            _taskValueProvider.Set("QMErrorNo", errNo);
            _helper.LogMessage(errNo + " " + ex.Message, LogType.Error);
        }
    }
}