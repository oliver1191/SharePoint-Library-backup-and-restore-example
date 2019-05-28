using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using System.Xml;
using System.IO;

namespace _2018081301MyLibraryDocument
{
    class MyLibraryDocument
    {
        /// <summary>
        /// Write the directory tree of document library to a file of XML
        /// </summary>
        /// <param name="siteUrl">the url of site</param>
        /// <param name="libraryName">the name of document library</param>
        public void WriteXml(string siteUrl, string libraryName)
        {
            using (SPSite mysite = new SPSite(siteUrl))
            {
                using (SPWeb myWeb = mysite.RootWeb)
                {
                    //SPList myList = myWeb.Lists["Documents"];
                    SPFolder myFolder = myWeb.GetFolder(libraryName);
                    //myFolder = myList.RootFolder;
                    if (!myFolder.Exists)
                    {
                        SPListTemplate myListTemplate = myWeb.ListTemplates["Document Library"];
                        //SPDocTemplate myDocTemplate = (from SPDocTemplate dt in myWeb.DocTemplates where dt.Type == 122 select dt).FirstOrDefault();
                        //Guid myGuid = myWeb.Lists.Add(libraryName, String.Format("create document library named {0} successfully", libraryName), myListTemplate,myDocTemplate);
                        Guid myGuid = myWeb.Lists.Add(libraryName, String.Format("create document library named {0} successfully", libraryName), myListTemplate);
                        SPDocumentLibrary myLibrary = myWeb.Lists[myGuid] as SPDocumentLibrary;
                        myLibrary.OnQuickLaunch = true;
                        myLibrary.Update();
                        myFolder = myWeb.GetFolder(libraryName);
                        Console.WriteLine("success");
                    }
                    string changeStr = ChangeString(libraryName);
                    XmlDocument myXmlDoc = new XmlDocument();
                    XmlDeclaration myXmlDeclaration = myXmlDoc.CreateXmlDeclaration("1.0", "utf-8", null);
                    myXmlDoc.AppendChild(myXmlDeclaration);
                    XmlElement addWeb = myXmlDoc.CreateElement("Web");
                    XmlElement addLibrary = myXmlDoc.CreateElement("Library");
                    XmlElement myXmlNode = myXmlDoc.CreateElement(changeStr);
                    SPList list = myFolder.DocumentLibrary;
                    foreach (SPField item in list.Fields)
                    {
                        if (!item.ReadOnlyField)
                        {
                            myXmlNode.SetAttribute(ChangeString(item.Title), item.Type.ToString());
                        }
                    }
                    addLibrary.AppendChild(GetSubNodeAllDirectory(myXmlNode, myFolder, myXmlDoc));
                    addWeb.AppendChild(addLibrary);
                    myXmlDoc.AppendChild(addWeb);
                    string sFilePath = "C:\\Users\\administrator.SPCARTOON\\Desktop\\XML";
                    string sFileName = "MySharePointXML2";
                    sFileName = sFilePath + @"\\" + sFileName; //文件的绝对路径
                    using (XmlTextWriter xmlTextWriter = new XmlTextWriter(@sFileName, Encoding.UTF8)
                    {
                        Formatting = Formatting.Indented,
                        IndentChar = '\t',
                        Indentation = 1
                    })
                    {
                        myXmlDoc.Save(xmlTextWriter);
                        Console.WriteLine("sss");
                    }

                }
            }
        }
        /// <summary>
        /// get the sub node of all directories and files from document library
        /// </summary>
        /// <param name="myNode">the xmlNode to save the attributes and innerText</param>
        /// <param name="myFolder">the SPFolder that you want to get its all sub directories and files</param>
        /// <param name="myXmlDoc">the XmlDocument to save your information</param>
        /// <returns></returns>
        private XmlNode GetSubNodeAllDirectory(XmlNode myNode, SPFolder myFolder,  XmlDocument myXmlDoc)
        {
            SPFolderCollection myFolderCollection = myFolder.SubFolders;
            SPFileCollection myFileCollection = myFolder.Files;
            int myFolderLength = myFolderCollection.Count;
            int myFileLength = myFileCollection.Count;
            int totalLength = myFolderLength + myFileLength;
            if (myFolderLength > 0 || myFileLength > 0)
            {
                XmlNode[] myXmlNode = new XmlNode[totalLength];
                if (myFolderLength > 0)
                {
                    for (int i = 0; i < myFolderLength; i++)
                    {
                        if (myFolderCollection[i].Name != "Forms")
                        {
                            myXmlNode[i] = myXmlDoc.CreateElement(ChangeString(myFolderCollection[i].Name));
                            SPList myList = myFolder.DocumentLibrary;

                            SPListItem myListItem = null;
                            //myListItem = myFolder.Item;
                            for (int k = 0; k < myList.Folders.Count; k++)
                            {
                                if (myList.Folders[k].Name.ToString() == myFolderCollection[i].Name.ToString())
                                {
                                    myListItem = myList.Folders[k];
                                }
                            }
                            AddProperitys( myXmlNode[i],  myXmlDoc, myListItem);
                            myFolder = myFolderCollection[i];
                            GetSubNodeAllDirectory(myXmlNode[i], myFolder, myXmlDoc);
                            myFolder = myFolder.ParentFolder;
                        }
                    }
                }
                for (int j = 0; j < myFileLength; j++)
                {
                    SPList myList = myFileCollection[j].DocumentLibrary;
                    SPListItem myListItem = null;
                    // myListItem = myFileCollection[j].Item;
                    for (int k = 0; k < myList.Items.Count; k++)
                    {
                        if (myList.Items[k].Name.ToString() == myFileCollection[j].Name.ToString())
                        {
                            myListItem = myList.Items[k];
                        }

                    }
                    myXmlNode[j + myFolderLength] = myXmlDoc.CreateElement(ChangeString(myFileCollection[j].Name));
                    AddProperitys( myXmlNode[j + myFolderLength],  myXmlDoc, myListItem);

                }
                for (int i = 0; i < totalLength; i++)
                {
                    if (myXmlNode[i] != null)
                    {
                        myNode.AppendChild(myXmlNode[i]);
                    }
                }
            }
            return myNode;
        }
        private string ChangeString(string str)
        {
            string[] strSplit = str.Split(' ');
            string newStr = null;
            for (int i = 0; i < strSplit.Length; i++)
            {
                if (i == 0)
                {
                    newStr += strSplit[i];
                }
                else
                {
                    newStr += "_" + strSplit[i];
                }
            }
            return newStr;
        }
        private string ChangeBackString(string str)
        {
            string[] strSplit = str.Split('_');
            string newStr = null;
            for (int i = 0; i < strSplit.Length; i++)
            {
                if (i == 0)
                {
                    newStr += strSplit[i];
                }
                else
                {
                    newStr += " " + strSplit[i];
                }
            }
            return newStr;
        }
        private void AddProperity(ref XmlNode myXmlNode, ref XmlDocument myXmlDoc, SPListItem myListItem)
        {
            string[] str = { "Name", "Modified", "Modified By", "Content Type", "ID", "Version" };
            string newStr = null;
            foreach (string item in str)
            {
                newStr = ChangeString(item);
                XmlNode xmlNode = myXmlDoc.CreateElement(newStr);
                xmlNode.InnerText = myListItem[item].ToString();
                myXmlNode.AppendChild(xmlNode);
            }

        }
        private void AddProperitys( XmlNode myXmlNode,  XmlDocument myXmlDoc, SPListItem myListItem)
        {
            string[] str = { "Name", "Modified", "Modified By", "Content Type", "ID", "Title", "Version", "Source Path", "MyDate" };
            string newStr = null;
            foreach (string item in str)
            {
                newStr = ChangeString(item);
                if (myListItem[item] != null)
                {
                    if (newStr == "Source_Path")
                    {
                        // XmlNode xmlNode = myXmlDoc.CreateElement(newStr);
                        myXmlNode.InnerText = myListItem[item].ToString();
                        // myXmlNode.AppendChild(xmlNode);
                    }
                    else
                    {
                        XmlAttribute xmlAttribute = myXmlDoc.CreateAttribute(newStr);
                        xmlAttribute.Value = myListItem[item].ToString();
                        myXmlNode.Attributes.Append(xmlAttribute);



                        //XmlElement ele = myXmlNode as XmlElement;
                        //ele.SetAttribute("", "");
                        //myXmlDoc.AppendChild(ele as XmlNode);
                    }
                }
                //XmlNode xmlNode = myXmlDoc.CreateElement(newStr);
                //xmlNode.InnerText = myListItem[item].ToString();
                // myXmlNode.AppendChild(xmlNode);                
            }

        }
        public void ReadXml(string siteUrl, string libraryName, string xmlPath)
        {
            using (SPSite mySite = new SPSite(siteUrl))
            {
                using (SPWeb myWeb = mySite.RootWeb)
                {
                    SPFolder myFolder = myWeb.GetFolder(libraryName);
                    if (!myFolder.Exists)
                    {
                        SPListTemplate myListTemplate = myWeb.ListTemplates["Document Library"];
                        //SPDocTemplate myDocTemplate = (from SPDocTemplate dt in myWeb.DocTemplates where dt.Type == 122 select dt).FirstOrDefault();
                        //Guid myGuid = myWeb.Lists.Add(libraryName, String.Format("create document library named {0} successfully", libraryName), myListTemplate,myDocTemplate);
                        Guid myGuid = myWeb.Lists.Add(libraryName, String.Format("create document library named {0} successfully", libraryName), myListTemplate);
                        SPDocumentLibrary myLibrary = myWeb.Lists[myGuid] as SPDocumentLibrary;
                        myLibrary.OnQuickLaunch = true;
                        myLibrary.Update();
                        myFolder = myWeb.GetFolder(libraryName);
                        Console.WriteLine("success");
                    }
                    SPList myList = myFolder.DocumentLibrary;
                    try
                    {
                        if (!IsExist(myList.Fields, "Source Path"))
                        {
                            string str1 = myList.Fields.Add("Source Path", SPFieldType.Text, false);
                            //myList.Fields[str1].Description = "Oliver Content Type" + DateTime.Now.ToString();
                            myList.Update();
                            //myList = subFile.DocumentLibrary;
                            Console.WriteLine("Source Path not exist,but create");
                        }
                    }
                    catch (Exception e)
                    {

                        WriteLog(e.Message);
                    }
                    XmlDocument myXmlDoc = new XmlDocument();
                    myXmlDoc.Load(xmlPath);
                    string changeStr = "Web/Library/" + ChangeString(libraryName);
                    XmlNode myXmlNode = myXmlDoc.SelectSingleNode(changeStr);
                    AddFileToClient(myFolder, myXmlNode);

                }
            }
        }
        public void AddFile(string sourceFile, SPFolder myFolder, string fileName)
        {
            using (FileStream fs = new FileStream(@sourceFile, FileMode.Open))
            {

                byte[] content = new byte[fs.Length];
                fs.Read(content, 0, (int)fs.Length);
                try
                {
                    myFolder.Files.Add(fileName, content);
                    Console.WriteLine("{0} success add", fileName);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }
                fs.Close();
            }
        }
        public void AddFileToClient(SPFolder myFolder, XmlNode myXmlNode)
        {
            //SPFolder subFolder = null;
            SPFolder tempFolder = myFolder;
            XmlNodeList myXmlNodeList = myXmlNode.ChildNodes;
            for (int i = 0; i < myXmlNodeList.Count; i++)
            {
                XmlNode item = myXmlNodeList[i];
                string contentType = item.Attributes["Content_Type"].Value.ToString();
                //string filePath = item.InnerText.ToString();
                if (contentType == "Folder")
                {
                    string str = item.Name.ToString();
                    string newStr = ChangeBackString(str);
                    // SPFolderCollection subFolderCollection = tempFolder.SubFolders;
                    SPFolder subFolder = null;
                    bool isSubFolderExist = false;
                    foreach (SPFolder test in tempFolder.SubFolders)
                    {
                        if (test.Name.ToString() == newStr)
                        {
                            isSubFolderExist = true;
                            subFolder = tempFolder.SubFolders[newStr];
                            break;
                        }
                    }
                    if (!isSubFolderExist)
                    {
                        tempFolder.SubFolders.Add(newStr);

                        subFolder = tempFolder.SubFolders[newStr];
                    }
                    SPList myList = tempFolder.DocumentLibrary;
                    //SPListItem myListItem = null;
                    int length = myList.Folders.Count;
                    int temp = 0;
                    bool isExistFolder = false;
                    for (int k = 0; k < length; k++)
                    {
                        if (myList.Folders[k].Name.ToString() == newStr)
                        {
                            temp = k;
                            isExistFolder = true;
                            break;
                        }
                    }
                    if (isExistFolder)
                    {
                        // myListItem = myList.Folders[temp];
                        // XmlAttributeCollection myXmlAttributeCollection = item.Attributes;
                        //foreach (XmlAttribute xmlAttribute in myXmlAttributeCollection)
                        // {
                        //     string columnName = ChangeBackString(xmlAttribute.Name.ToString());
                        //     //Console.WriteLine(columnName);
                        //     if (columnName == "Modified By")
                        //    {
                        //         //myListItem[columnName] = xmlAttribute.Value;
                        //         myListItem[columnName] =" oliver.lin" ;
                        //         Console.WriteLine(myListItem[columnName]);                                                            
                        //    }
                        //}
                        // //myListItem["Source Path"] = item.InnerText.ToString();
                        //// Console.Write("Source Path:");
                        //// Console.Write(myListItem["Source Path"]);
                        //// Console.WriteLine();
                        // myListItem.Update();


                        AddFileToClient(subFolder, item);
                        subFolder = subFolder.ParentFolder;
                        item = item.ParentNode;
                    }
                    else
                    {
                        Console.WriteLine("{0} folder not exists", newStr);
                    }

                }
                else if (contentType == "Document")
                {
                    string filePath = item.InnerText.ToString();
                    //Console.WriteLine(filePath.ToString());
                    string str = item.Name.ToString();
                    string newStr = ChangeBackString(str);
                    // SPFolderCollection subFolderCollection = tempFolder.SubFolders;
                    SPFile subFile = null;
                    bool isFile = false;
                    foreach (SPFile iFile in tempFolder.Files)
                    {
                        if (iFile.Name.ToString() == newStr)
                        {
                            subFile = tempFolder.Files[newStr];
                            isFile = true;
                            break;
                        }
                    }
                    using (FileStream fs = new FileStream(@filePath, FileMode.Open))
                    {
                        const int bufferSize = 4 * 1024 * 1024;
                        byte[] fileContentBuffer = new byte[bufferSize];
                        //tempFolder.Files.Add(newStr, fileContentBuffer);
                        int data;
                        if (!isFile)
                        {
                            byte[] nullByte = new byte[0];
                            tempFolder.Files.Add(newStr, nullByte);
                            subFile = tempFolder.Files[newStr];
                        }
                        do
                        {
                            data = fs.Read(fileContentBuffer, 0, fileContentBuffer.Length);

                            if (data < bufferSize)
                            {
                                byte[] tempByte = new byte[data];
                                for (int b = 0; b < data; b++)
                                {
                                    tempByte[b] = fileContentBuffer[b];
                                }
                                //Console.WriteLine(data);
                                //Console.WriteLine(fileContentBuffer.Length);
                                //Console.WriteLine(tempByte.Length);
                                tempFolder.Files[newStr].SaveBinary(tempByte);
                            }
                            else
                            {
                                tempFolder.Files[newStr].SaveBinary(fileContentBuffer);
                            }
                        } while (data > 0);
                    }

                    //AddFile(filePath, tempFolder, newStr);
                    SPList myList = subFile.DocumentLibrary;
                    SPListItem myListItem = null;
                    int fileLength = myList.Items.Count;
                    int tempFileKey = 0;
                    bool isExist = false;
                    for (int k = 0; k < fileLength; k++)
                    {
                        if (myList.Items[k].Name.ToString() == newStr)
                        {
                            tempFileKey = k;
                            isExist = true;
                            break;
                        }
                    }
                    if (isExist)
                    {
                        myListItem = myList.Items[tempFileKey];
                        XmlAttributeCollection myXmlAttributeCollection = item.Attributes;
                        foreach (XmlAttribute xmlAttribute in myXmlAttributeCollection)
                        {

                            string columnName = ChangeBackString(xmlAttribute.Name.ToString());
                            //Console.WriteLine(columnName);
                            if (columnName == "Name" || columnName == "Title")
                            {
                                myListItem[columnName] = xmlAttribute.Value;
                                //Console.WriteLine(myListItem[columnName]);

                            }
                            if (columnName == "MyDate")
                            {

                                //                if (columnName == "MyDate")
                                //                {
                                //                    if (myList.Fields[columnName] != null)){
                                ////Console.WriteLine(myListItem[columnName].GetType());
                                //myListItem[columnName] = Convert.ToDateTime(xmlAttribute.Value);
                                //Console.WriteLine(myListItem[columnName]);
                                
                                if(!IsExist(myList.Fields,columnName))
                                {
                                    myList.Fields.Add(columnName, SPFieldType.DateTime, false);
                                    Console.WriteLine("not exist");
                                }          
                                //Console.WriteLine(myListItem[columnName].GetType());
                                myListItem[columnName] = Convert.ToDateTime(xmlAttribute.Value);
                                Console.WriteLine(myListItem[columnName]);
                            }
                        }
                        myListItem["Source Path"] = item.InnerText.ToString();
                        //Console.Write("Source Path:");
                        //Console.Write(myListItem["Source Path"]);
                        //Console.WriteLine();
                        myListItem.Update();
                    }
                    else
                    {
                        Console.WriteLine("{0} file not exists", newStr);
                    }

                }

            }
        }
        public void WriteLog(string strLog)
        {
            string sFilePath = "C:\\Users\\administrator.SPCARTOON\\Desktop\\XML";
            string sFileName = DateTime.Now.ToString("yyyy-MM-dd") + ".log";
            sFileName = sFilePath + @"\\" + sFileName; //文件的绝对路径
            if (!Directory.Exists(sFilePath))//验证路径是否存在
            {
                Directory.CreateDirectory(sFilePath);
                //set if directory not exist
            }
            FileStream fs;
            StreamWriter sw;
            if (File.Exists(sFileName))
            //append or set
            {
                fs = new FileStream(sFileName, FileMode.Append, FileAccess.Write);
            }
            else
            {
                fs = new FileStream(sFileName, FileMode.Create, FileAccess.Write);
            }
            sw = new StreamWriter(fs);
            sw.WriteLine(strLog);
            sw.Close();
            fs.Close();
        }
        /// <summary>
        /// Copy  files of document library to your place
        /// </summary>
        /// <param name="siteUrl">the url of site </param>
        /// <param name="destinationPath">the path to save files of document library</param>
        public void CreateDirectoryTree(string siteUrl, string destinationPath)
        {
            using (SPSite mySite = new SPSite(siteUrl))
            {
                using (SPWeb myWeb = mySite.RootWeb)
                {
                    string tempDirectory = destinationPath;
                    SPList myList = myWeb.Lists["Documents"];
                    SPFolder myFolder = myList.RootFolder;
                    DirectoryInfo myDirectory = new DirectoryInfo(tempDirectory);
                    if (myDirectory.Exists)
                    {
                        tempDirectory += "\\" + "Web";
                        if (!Directory.Exists(tempDirectory))
                        {
                            myDirectory.CreateSubdirectory("Web");
                        }
                        tempDirectory += "\\" + "Library";
                        if (!Directory.Exists(tempDirectory))
                        {
                            Directory.CreateDirectory(tempDirectory);
                        }
                        tempDirectory += "\\" + myFolder.Name.ToString();
                        if (!Directory.Exists(tempDirectory))
                        {
                            Directory.CreateDirectory(tempDirectory);
                        }
                        CreateSubDirectory(myFolder, tempDirectory);
                    }
                    // Console.WriteLine(myFolder.Name);
                }
            }
        }
        /// <summary>
        /// Copy subfolders and files of folder to a path 
        /// </summary>
        /// <param name="myFolder">the folder that you want to get its subfolder and files</param>
        /// <param name="destinationPath">the path that you want to copy files of folder to</param>
        private void CreateSubDirectory(SPFolder myFolder, string destinationPath)
        {
            //get the document library of myFolder
            SPList myList = myFolder.DocumentLibrary;
            int myListLength = myList.Items.Count;
            SPListItem myListItem = myList.Items[0];
            string tempPath = destinationPath;
            SPFolderCollection myFolderCollection = myFolder.SubFolders;
            SPFileCollection myFileCollection = myFolder.Files;
            int subFolderLength = myFolderCollection.Count;
            int subFileLength = myFileCollection.Count;
            if (subFolderLength > 0 || subFileLength > 0)
            {
                if (subFolderLength > 0)
                {
                    for (int i = 0; i < subFolderLength; i++)
                    {
                        if (myFolderCollection[i].Name != "Forms")
                        {
                            tempPath += "\\" + myFolderCollection[i].Name;
                            if (!Directory.Exists(tempPath))
                            {
                                Directory.CreateDirectory(tempPath);
                            }
                            CreateSubDirectory(myFolderCollection[i], tempPath);
                            tempPath = destinationPath;
                        }
                    }
                }
                foreach (SPFile item in myFileCollection)
                {
                    DirectoryInfo myDirectoryInfo = new DirectoryInfo(destinationPath);
                    string tempFilePath = destinationPath;
                    tempFilePath += "\\" + item.Name;
                    if (File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                    using (FileStream myFileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write))
                    {
                        const int bufferSize = 4 * 1024 * 1024;
                        //byte[] fileContentBuffer = new byte[bufferSize];
                        byte[] fileContent = item.OpenBinary();
                        int num = fileContent.Length / bufferSize;
                        int data = fileContent.Length % bufferSize;
                        int start = 0;
                        while (start < num)
                        {
                            myFileStream.Write(fileContent, start * bufferSize, bufferSize);
                            start++;
                        }
                        myFileStream.Write(fileContent, start * bufferSize, data);
                    }
                    for (int a = 0; a < myListLength; a++)
                    {
                        if (myList.Items[a].Name == item.Name)
                        {
                            myListItem = myList.Items[a];
                            break;
                        }
                    }
                    myListItem["Source Path"] = tempFilePath;
                    myListItem.Update();
                    tempFilePath = destinationPath;
                }
            }
        }
        private bool IsExist(SPFieldCollection myFieldCollection,string str)
        {
            bool isExist = false;
            try
            {
                Console.WriteLine(myFieldCollection[str].InternalName);
                isExist = true;
            }
            catch (Exception e)
            {                
                isExist = false;
                return isExist;
            }
            
            return isExist;
            
            
            //foreach (SPField item in myFieldCollection)
            //{
            //    if (item.Title == str)
            //    {
            //        isExist = true;
            //        break;
            //    }
            //}
           
        }
        public void showListColumn(string siteUrl, string listName)
        {
            using (SPSite mySite=new SPSite(siteUrl))
            {
                using (SPWeb myWeb=mySite.RootWeb)
                {
                    SPList myList = myWeb.Lists[listName];
                    SPFieldCollection myFields = myList.Fields;
                    foreach (SPField item in myFields)
                    {
                        if (!item.Hidden)
                        {
                            Console.WriteLine(item.Title);
                        }
                            
                       
                        
                    }
                }
            }
        }
    }
}
