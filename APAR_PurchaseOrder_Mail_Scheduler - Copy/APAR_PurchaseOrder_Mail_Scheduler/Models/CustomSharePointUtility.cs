using System;
using System.Collections.Generic;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using UserInformation;
using MSC = Microsoft.SharePoint.Client;
namespace APAR_PurchaseOrder_Mail_Scheduler.Models
{
    public static class CustomSharePointUtility
    {
        static UserOperation _UserOperation = new UserOperation();
        public static StreamWriter logFile;
        static byte[] bytes = ASCIIEncoding.ASCII.GetBytes("ZeroCool");


        public static string Decrypt(string cryptedString)
        {
            if (String.IsNullOrEmpty(cryptedString))
            {
                throw new ArgumentNullException("The string which needs to be decrypted can not be null.");
            }

            DESCryptoServiceProvider cryptoProvider = new DESCryptoServiceProvider();
            MemoryStream memoryStream = new MemoryStream(Convert.FromBase64String(cryptedString));
            CryptoStream cryptoStream = new CryptoStream(memoryStream, cryptoProvider.CreateDecryptor(bytes, bytes), CryptoStreamMode.Read);
            StreamReader reader = new StreamReader(cryptoStream);

            return reader.ReadToEnd();
        }
        public static MSC.ClientContext GetContext(string siteUrl)
        {
            try
            {
                AppConfiguration _AppConfiguration = GetSharepointCredentials(siteUrl);
                var securePassword = new SecureString();
                foreach (char c in _AppConfiguration.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                var onlineCredentials = new MSC.SharePointOnlineCredentials(_AppConfiguration.ServiceUserName, securePassword);
                var context = new MSC.ClientContext(_AppConfiguration.ServiceSiteUrl);
                context.Credentials = onlineCredentials;
                return context;
            }
            catch (Exception ex)
            {
                WriteLog("Error in  CustomSharePointUtility.GetContext: " + ex.ToString());
                return null;
            }
        }
        public static MSC.ClientContext GetEmpContext(string RootsiteUrl)
        {
            try
            {
                AppConfiguration _NewAppConfiguration = GetSharepointRootCredentials(RootsiteUrl);
                var securePassword = new SecureString();
                foreach (char c in _NewAppConfiguration.ServicePassword)
                {
                    securePassword.AppendChar(c);
                }
                var onlineCredentials = new MSC.SharePointOnlineCredentials(_NewAppConfiguration.ServiceUserName, securePassword);
                var context = new MSC.ClientContext(_NewAppConfiguration.ServiceSiteUrl);
                context.Credentials = onlineCredentials;
                return context;
            }
            catch (Exception ex)
            {
                WriteLog("Error in  CustomSharePointUtility.GetEmpContext: " + ex.ToString());
                return null;
            }
        }

        public static void WriteLog(string logmsg)
        {
            // StreamWriter logFile;

            try
            {

                string LogString = DateTime.Now.ToString("dd/MM/yyyy HH:MM") + " " + logmsg.ToString();

                //  logFile.WriteLine(DateTime.Now);
                //  logFile.WriteLine(logmsg.ToString());
                logFile.WriteLine(LogString);

                //logFile.Close();
            }
            catch (Exception ex)
            {
                WriteLog(ex.ToString());

            }

        }

        public static AppConfiguration GetSharepointCredentials(string siteUrl)
        {
            AppConfiguration _AppConfiguration = new AppConfiguration();

            _AppConfiguration.ServiceSiteUrl = siteUrl;// _UserOperation.ReadValue("SP_Address");
            _AppConfiguration.ServiceUserName = _UserOperation.ReadValue("SP_USER_ID_Live");
            _AppConfiguration.ServicePassword = Decrypt(_UserOperation.ReadValue("SP_Password_Live"));
            return _AppConfiguration;
        }

        public static AppConfiguration GetSharepointRootCredentials(string RootsiteUrl)
        {
            AppConfiguration _NewAppConfiguration = new AppConfiguration();

            _NewAppConfiguration.ServiceSiteUrl = RootsiteUrl;// _UserOperation.ReadValue("SP_Address");
            _NewAppConfiguration.ServiceUserName = _UserOperation.ReadValue("SP_USER_ID_Live");
            _NewAppConfiguration.ServicePassword = Decrypt(_UserOperation.ReadValue("SP_Password_Live"));
            return _NewAppConfiguration;
        }

        #region old code
        //public static List<PurchaseOrder> GetAll_PurchaseOrderFromSharePoint(string siteUrl, string listName, string DaysDifference)
        //{
        //    List<PurchaseOrder> _retList = new List<PurchaseOrder>();
        //    try
        //    {
        //        using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
        //        {
        //            if (context != null)
        //            {
        //                MSC.List list = context.Web.Lists.GetByTitle(listName);
        //                MSC.ListItemCollectionPosition itemPosition = null;
        //                while (true)
        //                {
        //                    var dataDateValue = DateTime.Now.AddDays(-Convert.ToInt32(DaysDifference));
        //                    MSC.CamlQuery camlQuery = new MSC.CamlQuery();
        //                    camlQuery.ListItemCollectionPosition = itemPosition;
        //                    camlQuery.ViewXml = @"<View>
        //                         <Query>
        //                            <Where>
        //                                <And>
        //                                    <Or> 
        //                                        <Or>  
        //                                            <Eq>
        //                                                <FieldRef Name='ApprovalStatus'/>
        //                                                <Value Type='text'>Submitted</Value>
        //                                            </Eq>
        //                                            <Eq>
        //                                                <FieldRef Name='ApprovalStatus'/>
        //                                                <Value Type='text'>Approved By Functional Head</Value>
        //                                            </Eq>
        //                                        </Or>
        //                                        <Or>  
        //                                            <Eq>
        //                                                <FieldRef Name='ApprovalStatus'/>
        //                                                <Value Type='text'>Approved By Purchase Head</Value>
        //                                            </Eq>
        //                                            <Eq>
        //                                                <FieldRef Name='ApprovalStatus'/>
        //                                                <Value Type='text'>Approved By Plant Head</Value>
        //                                            </Eq>
        //                                        </Or>
        //                                    </Or> 
        //                                    <Leq><FieldRef Name='Modified'/><Value Type='DateTime'>" + dataDateValue.ToString("o") + "</Value></Leq>";


        //                    camlQuery.ViewXml += @"</And></Where></Query>
        //                        <RowLimit>5000</RowLimit>
        //                        <ViewFields>
        //                        <FieldRef Name='ID'/>
        //                        <FieldRef Name='POReferenceNumber'/>
        //                        <FieldRef Name='Author'/>
        //                        <FieldRef Name='Created'/>
        //                        <FieldRef Name='DepartmentName'/>
        //                        <FieldRef Name='DepartmentID'/>
        //                        <FieldRef Name='LocationName'/>
        //                        <FieldRef Name='LocationID'/>
        //                        <FieldRef Name='LocationType'/>
        //                        <FieldRef Name='DivisionName'/>
        //                        <FieldRef Name='DivisionID'/>
        //                        <FieldRef Name='PONumber'/>
        //                        <FieldRef Name='POCost'/>
        //                        <FieldRef Name='MaterialDetails'/>
        //                        <FieldRef Name='ApprovalStatus'/>
        //                        <FieldRef Name='FHCode'/>
        //                        <FieldRef Name='Modified'/>
        //                        </ViewFields></View>";
        //                    MSC.ListItemCollection Items = list.GetItems(camlQuery);

        //                    context.Load(Items);
        //                    context.ExecuteQuery();
        //                    itemPosition = Items.ListItemCollectionPosition;
        //                    foreach (MSC.ListItem item in Items)
        //                    {
        //                        _retList.Add(new PurchaseOrder
        //                        {
        //                            Id = Convert.ToInt32(item["ID"]),
        //                            POReferenceNumber = Convert.ToString(item["POReferenceNumber"]).Trim(),
        //                            Author = Convert.ToString((item["Author"] as Microsoft.SharePoint.Client.FieldUserValue).LookupValue),
        //                            Created = Convert.ToString(item["Created"]).Trim(),
        //                            DepartmentName = Convert.ToString(item["DepartmentName"]).Trim(),
        //                            DepartmentID = Convert.ToString(item["DepartmentID"]).Trim(),
        //                            LocationName = Convert.ToString(item["LocationName"]).Trim(),
        //                            LocationID = Convert.ToString(item["LocationID"]).Trim(),
        //                            LocationType = Convert.ToString(item["LocationType"]).Trim(),
        //                            ////Author = Convert.ToString((item["Author"] as Microsoft.SharePoint.Client.FieldUserValue).LookupId),
        //                            //Author = Convert.ToString((item["Author"] as Microsoft.SharePoint.Client.FieldUserValue).LookupValue),
        //                            //FunctionalHead = Convert.ToString((item["FunctionalHead"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
        //                            //HRHead = Convert.ToString((item["HRHead"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
        //                            //HRHeadOnly = Convert.ToString((item["HRHeadOnly"] as Microsoft.SharePoint.Client.FieldUserValue).LookupId),
        //                            //MDorJMD = item["MDorJMD"] == null ? "" : Convert.ToString((item["MDorJMD"] as Microsoft.SharePoint.Client.FieldUserValue).LookupId),
        //                            //Recruiter = Convert.ToString((item["Recruiter"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
        //                            DivisionName = Convert.ToString(item["DivisionName"]).Trim(),
        //                            DivisionID = Convert.ToString(item["DivisionID"]).Trim(),
        //                            PONumber = Convert.ToString(item["PONumber"]).Trim(),
        //                            POCost = Convert.ToString(item["POCost"]).Trim(),
        //                            MaterialDetails = Convert.ToString(item["MaterialDetails"]).Trim(),
        //                            ApprovalStatus = Convert.ToString(item["ApprovalStatus"]).Trim(),
        //                            FHCode = Convert.ToString(item["FHCode"]).Trim(),
        //                            Modified = Convert.ToString(item["Modified"]).Trim(),
        //                        });
        //                    }
        //                    if (itemPosition == null)
        //                    {
        //                        break; // TODO: might not be correct. Was : Exit While
        //                    }

        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        CustomSharePointUtility.WriteLog("Error in  GetAll_ManpowerRequisitionFromSharePoint()" + " Error:" + ex.Message);
        //    }
        //    return _retList;
        //}

        #endregion
        /// New Code to Get Employee
        /// 

        public static List<Employee> GetEmployees(string siteUrl, string listName)
        {
            List<Employee> _returnEmployee = new List<Employee>();
            try
            {
                using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
                {
                    if (context != null)
                    {
                        MSC.List HODList = context.Web.Lists.GetByTitle(listName);
                        MSC.ListItemCollectionPosition itemPosition = null;

                        while (true)
                        {
                            MSC.CamlQuery camlQuery = new MSC.CamlQuery();
                            camlQuery.ListItemCollectionPosition = itemPosition;
                            camlQuery.ViewXml = @"<View><Query><Where>
                                                    <Eq><FieldRef Name='Reminder' /><Value Type='Choice'>Yes</Value></Eq>
                                                    </Where></Query>
                                                    <RowLimit>5000</RowLimit>
                                                     <FieldRef Name='ID'/>
                                                    <FieldRef Name='EmployeeName'/>
                                                    <FieldRef Name='Author'/>
                                                    <FieldRef Name='Created'/>
                                                </View>";
                            MSC.ListItemCollection Items = HODList.GetItems(camlQuery);

                            context.Load(Items);
                            context.ExecuteQuery();
                            itemPosition = Items.ListItemCollectionPosition;

                            foreach (MSC.ListItem item in Items)
                            {
                                _returnEmployee.Add(new Employee
                                {

                                    EmployeeName = Convert.ToString((item["EmployeeName"] as Microsoft.SharePoint.Client.FieldUserValue).LookupValue),
                                    EmployeeID = Convert.ToString((item["EmployeeName"] as Microsoft.SharePoint.Client.FieldUserValue).LookupId)
                                });



                            }
                            if (itemPosition == null)
                            {
                                break; // TODO: might not be correct. Was : Exit While
                            }
                        }


                    }

                }
            }
            catch (Exception ex)
            {
            }

            return _returnEmployee;
        }

        public static List<Employee> GetEnterEstimate(string month, string year, string listname, string siteUrl)
        {
            List<Employee> _returnEmployee = new List<Employee>();
            try
            {
                using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
                {
                    if (context != null)
                    {
                        MSC.List SalesEstimateEmployee = context.Web.Lists.GetByTitle(listname);
                        MSC.ListItemCollectionPosition itemPosition = null;

                        string fromday = year + "-" + month + "-" + "1";
                        string Today = year + "-" + month + "31";

                        while (true)
                        {
                            MSC.CamlQuery camlQuery = new MSC.CamlQuery();
                            camlQuery.ListItemCollectionPosition = itemPosition;
                            camlQuery.ViewXml = @"<View><Query><Where>
                                                    <Eq><FieldRef Name='Date' />
                                                    <Value IncludeTimeValue='FLASE' Type='DateTime'>" + fromday + "</Value></Eq>";
                            camlQuery.ViewXml += @" </Where></Query>
                                                    <RowLimit>5000</RowLimit>
                                                     <FieldRef Name='ID'/>
                                                    <FieldRef Name='EmployeeName'/>
                                                    <FieldRef Name='Author'/>
                                                    <FieldRef Name='Created'/>
                                                </View>";
                            MSC.ListItemCollection Items = SalesEstimateEmployee.GetItems(camlQuery);
                            context.Load(Items);
                            context.ExecuteQuery();
                            itemPosition = Items.ListItemCollectionPosition;

                            foreach (MSC.ListItem item in Items)
                            {
                                _returnEmployee.Add(new Employee
                                {

                                    EmployeeName = Convert.ToString((item["EmployeeName"] as Microsoft.SharePoint.Client.FieldUserValue).LookupValue),
                                    EmployeeID = Convert.ToString((item["EmployeeName"] as Microsoft.SharePoint.Client.FieldUserValue).LookupId)
                                });



                            }
                            if (itemPosition == null)
                            {
                                break; // TODO: might not be correct. Was : Exit While
                            }



                        }

                    }


                }



            }
            catch (Exception ex)
            {

            }
            return _returnEmployee;
        }

        public static List<Employee> GetFilter(List<Employee> Employeelist, List<Employee> SalesEmployeelist)
        {
            List<Employee> _returnEMployee = new List<Employee>();


            for (int i = 0; i < Employeelist.Count; i++)
            {
                var find = "No";
                for (int p = 0; p < SalesEmployeelist.Count; p++)
                {
                    if (Employeelist[i].EmployeeID == SalesEmployeelist[p].EmployeeID)
                    {
                        find = "Yes";
                    }
                }

                if (find == "No")
                {
                    _returnEMployee.Add(new Employee
                    {

                        EmployeeName = Employeelist[i].EmployeeName,
                        EmployeeID = Employeelist[i].EmployeeID
                    });
                }
            }
            return _returnEMployee;
        }


        public static bool EmailData(List<Employee>PendingEmployee, string siteUrl, string listName)
        {
            bool retValue = false;
            try
            {

                using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
                {
                    //List<Mailing> varx = new List<Mailing>();

                    MSC.List list = context.Web.Lists.GetByTitle(listName);
                   
                    MSC.ListItem listItem = null;

                    MSC.ListItemCreationInformation itemCreateInfo = new MSC.ListItemCreationInformation();
                    listItem = list.AddItem(itemCreateInfo);


                    // MSC.UserCollection users = new MSC.UserCollection();
                    //MSC.FieldUserValue fieldUserValue = new MSC.FieldUserValue();

                   
                    MSC.FieldUserValue[] userValueCollection = new MSC.FieldUserValue[PendingEmployee.Count];
                    for (var i = 0; i < PendingEmployee.Count; i++)
                    {
                        MSC.FieldUserValue fieldUserVal = new MSC.FieldUserValue();
                        fieldUserVal.LookupId = Convert.ToInt32(PendingEmployee[i].EmployeeID);
                        userValueCollection.SetValue(fieldUserVal,i);
                    }


                        //var _From = "";
                        var _To = "";
                    //var _Cc = "";
                    var _Body = "";
                    var _Subject = "";

                     _To = "1247";

                    _Subject = "Sales Estimate is not submitted";
                    _Body += "Dear User, <br><br>This is to inform you that still your Sales Estimate is not submitted kindly fill and submit on priority.<br><br>Kindly click on below link for submit estimate";
                    _Body += "<br><a href="+siteUrl+"/SitePages/NewRequest.aspx>Submit Estimate</ a>";
                                      



                    listItem["ToUser"] = userValueCollection;
                    listItem["SubjectDesc"] = _Subject;
                    listItem["BodyDesc"] = _Body;
                    listItem.Update();

                    
                    try
                    {
                       context.ExecuteQuery();
                        retValue = true;

                    }
                    catch (Exception ex)
                    {
                        CustomSharePointUtility.WriteLog(string.Format("Error in  InsertUpdate_EmployeeMaster ( context.ExecuteQuery();): Error ({0}) ", ex.Message));
                        return false;
                        //continue;
                    }
                }

                //        }
                //    }
                //}

                //}
            }

            catch (Exception ex)
            {
                CustomSharePointUtility.WriteLog(string.Format("Error in  InsertUpdate_EmployeeMaster: Error ({0}) ", ex.Message));
            }
            return retValue;

        }



        //    public static List<EmployeeMaster> EmployeeMasterData(PurchaseOrder purchaseDataFinal, string RootsiteUrl, string siteUrl, string listName, string EmaillistName)
        //    {
        //        List<EmployeeMaster> _returnList = new List<EmployeeMaster>();
        //        try
        //        {
        //            using (MSC.ClientContext context = CustomSharePointUtility.GetEmpContext(RootsiteUrl))
        //            {
        //                if (context != null)
        //                {
        //                    MSC.List list = context.Web.Lists.GetByTitle(listName);
        //                    MSC.ListItemCollectionPosition itemPosition = null;
        //                    while (true)
        //                    {
        //                        //var dataDateValue = DateTime.Now.AddDays(-Convert.ToInt32(DaysDifference));
        //                        MSC.CamlQuery camlQuery = new MSC.CamlQuery();
        //                        camlQuery.ListItemCollectionPosition = itemPosition;
        //                        camlQuery.ViewXml = @"<View>
        //                             <Query>
        //                                <Where>
        //                                    <Eq>
        //                                        <FieldRef Name='Employee_x0020_Code'/>
        //                                        <Value Type='text'>"+ purchaseDataFinal.FHCode + "</Value></Eq>";

        //                        camlQuery.ViewXml += @"</Where>
        //                             </Query>
        //                            <RowLimit>5000</RowLimit>
        //                            <ViewFields>
        //                            <FieldRef Name='ID'/>
        //                            <FieldRef Name='Employee_x0020_Code'/>
        //                            <FieldRef Name='Employee_x0020_Email'/>
        //                            </ViewFields></View>";
        //                        MSC.ListItemCollection Items = list.GetItems(camlQuery);

        //                        context.Load(Items);
        //                        context.ExecuteQuery();
        //                        itemPosition = Items.ListItemCollectionPosition;
        //                        foreach (MSC.ListItem item in Items)
        //                        {
        //                            _returnList.Add(new EmployeeMaster
        //                            {
        //                                Employee_x0020_Code = Convert.ToString(item["Employee_x0020_Code"]).Trim(),
        //                                Employee_x0020_Email = Convert.ToString(item["Employee_x0020_Email"]).Trim(),

        //                            });
        //                        }
        //                        if (itemPosition == null)
        //                        {
        //                            break; // TODO: might not be correct. Was : Exit While
        //                        }

        //                    }
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            CustomSharePointUtility.WriteLog("Error in  EmployeeMasterData()" + " Error:" + ex.Message);
        //        }
        //        var success = CustomSharePointUtility.EmailData(purchaseDataFinal, _returnList[0].Employee_x0020_Email, "", "", "", siteUrl, EmaillistName);
        //        if (success) { }
        //            return _returnList;
        //    }
        //    public static List<PurchaseApprovers> PurchaseApproversData(PurchaseOrder purchaseDataFinal, string RootsiteUrl, string siteUrl, string listName, string EmaillistName)
        //    {
        //        List<PurchaseApprovers> _returnList = new List<PurchaseApprovers>();
        //        try
        //        {
        //            using (MSC.ClientContext context = CustomSharePointUtility.GetEmpContext(RootsiteUrl))
        //            {
        //                if (context != null)
        //                {
        //                    MSC.List list = context.Web.Lists.GetByTitle(listName);
        //                    MSC.ListItemCollectionPosition itemPosition = null;
        //                    while (true)
        //                    {
        //                        //var dataDateValue = DateTime.Now.AddDays(-Convert.ToInt32(DaysDifference));
        //                        MSC.CamlQuery camlQuery = new MSC.CamlQuery();
        //                        camlQuery.ListItemCollectionPosition = itemPosition;
        //                        camlQuery.ViewXml = @"<View>
        //                             <Query>
        //                                <Where>
        //                                   <And>
        //                                    <Eq>
        //                                        <FieldRef Name='Division' LookupId='True'/>
        //                                        <Value Type='Lookup'>" + purchaseDataFinal.DivisionID + "</Value></Eq>";
        //                        camlQuery.ViewXml += @"<Eq>
        //                                        <FieldRef Name ='Location' LookupId='True'/>
        //                                        <Value Type='Lookup'>" + purchaseDataFinal.LocationID + "</Value></Eq>";

        //                        camlQuery.ViewXml += @"</And></Where>
        //                             </Query>
        //                            <RowLimit>5000</RowLimit>
        //                            <ViewFields>
        //                            <FieldRef Name='ID'/>
        //                            <FieldRef Name='ApproverName'/>
        //                            <FieldRef Name='ApproverType'/>
        //                            </ViewFields></View>";
        //                        MSC.ListItemCollection Items = list.GetItems(camlQuery);

        //                        context.Load(Items);
        //                        context.ExecuteQuery();
        //                        itemPosition = Items.ListItemCollectionPosition;
        //                        foreach (MSC.ListItem item in Items)
        //                        {
        //                            _returnList.Add(new PurchaseApprovers
        //                            {
        //                                ApproverName = item["ApproverName"] == null ? "" : Convert.ToString((item["ApproverName"] as Microsoft.SharePoint.Client.FieldUserValue[])[0].LookupId),
        //                                ApproverType = Convert.ToString(item["ApproverType"]).Trim(),
        //                            });
        //                        }
        //                        if (itemPosition == null)
        //                        {
        //                            break; // TODO: might not be correct. Was : Exit While
        //                        }

        //                    }
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            CustomSharePointUtility.WriteLog("Error in  PurchaseApproversData()" + " Error:" + ex.Message);
        //        }
        //        for (var i = 0; i < _returnList.Count; i++) {
        //            if (_returnList[i].ApproverType == "MD" || _returnList[i].ApproverType == "JMD") {
        //                String MDorJMD = _returnList[i].ApproverName;
        //                var success = CustomSharePointUtility.EmailData(purchaseDataFinal, "", "", "", MDorJMD, siteUrl, EmaillistName);
        //                if (success) { }
        //            }
        //            else if (_returnList[i].ApproverType == "PURCHASE HEAD") {
        //                String PurchaseHead = _returnList[i].ApproverName;
        //                var success = CustomSharePointUtility.EmailData(purchaseDataFinal, "", PurchaseHead, "", "", siteUrl, EmaillistName);
        //                if (success) { }
        //            }
        //            else if (_returnList[i].ApproverType == "PLANT HEAD")
        //            {
        //                String PlantHead = _returnList[i].ApproverName;
        //                var success = CustomSharePointUtility.EmailData(purchaseDataFinal, "", "", PlantHead, "", siteUrl, EmaillistName);
        //                if (success) { }
        //            }
        //        }
        //        return _returnList;
        //    }
        //    public static bool EmailData(PurchaseOrder updationList, string FunctionalHeadEmail, string PurchaseHead, string PlantHead, string MD, string siteUrl, string listName)
        //    {
        //        bool retValue = false;
        //        try
        //        {

        //            using (MSC.ClientContext context = CustomSharePointUtility.GetContext(siteUrl))
        //            {
        //                //List<Mailing> varx = new List<Mailing>();

        //                MSC.List list = context.Web.Lists.GetByTitle(listName);
        //                //for (var i = 0; i < updationList.Count; i++)
        //                //{
        //                //var updateList = updationList.Skip(i).Take(1).ToList();
        //                //if (updateList != null && updateList.Count > 0)
        //                //{
        //                //    foreach (var updateItem in updateList)
        //                //    {
        //                MSC.ListItem listItem = null;

        //                MSC.ListItemCreationInformation itemCreateInfo = new MSC.ListItemCreationInformation();
        //                listItem = list.AddItem(itemCreateInfo);

        //                var obj = new Object();
        //                //Mailing data = new Mailing();

        //                //var _From = "";
        //                var _To = "";
        //                //var _Cc = "";
        //                var _Body = "";
        //                var _Subject = "";
        //                if (updationList.ApprovalStatus == "Submitted" && updationList.LocationType == "Office")
        //                {
        //                    _To = FunctionalHeadEmail;
        //                }
        //                else if (updationList.ApprovalStatus == "Approved By Functional Head" && updationList.LocationType=="Office")
        //                {
        //                    _To = MD;
        //                }
        //                else if (updationList.ApprovalStatus == "Submitted" && updationList.LocationType == "Plant")
        //                {
        //                    _To = PurchaseHead;
        //                }
        //                else if (updationList.ApprovalStatus == "Approved By Purchase Head" && updationList.LocationType == "Plant")
        //                {
        //                    _To = PlantHead;
        //                }
        //                else if (updationList.ApprovalStatus == "Approved By Plant Head" && Convert.ToInt32(updationList.POCost) > 50000 && updationList.LocationType == "Plant")
        //                {
        //                    _To = MD;
        //                }

        //                    _Subject = "Gentle Reminder";
        //                    _Body += "Dear User, <br><br>This is to inform you that below request is pending for your Approval.";
        //                    _Body += "<br><b>Workflow Name :</b> Purchase Order ";
        //                    _Body += "<br><b>Voucher No :</b>  " + updationList.POReferenceNumber;
        //                    _Body += "<br><b>Date of Creation :</b>  " + updationList.Created;
        //                    _Body += "<br><b>Employee : </b> " + updationList.Author;

        //                    _Body += "<br><b>Department :</b> " + updationList.DepartmentName;
        //                    _Body += "<br><b>Location :</b> " + updationList.LocationName;
        //                    _Body += "<br><b>PO No : </b> " + updationList.PONumber;
        //                    _Body += "<br><b> PO Amount : </b> " + updationList.POCost;
        //                    _Body += "<br><b> Material Details : </b> " + updationList.MaterialDetails;

        //                    if (updationList.ApprovalStatus == "Submitted" && updationList.LocationType == "Office")
        //                    {
        //                        _Body += "<br><b>Status :</b> Pending With Functional Head";
        //                    }
        //                    else if (updationList.ApprovalStatus == "Approved By Functional Head" && updationList.LocationType == "Office")
        //                    {
        //                        _Body += "<br><b>Status :</b> Pending With MD";
        //                    }
        //                    else if (updationList.ApprovalStatus == "Submitted" && updationList.LocationType == "Plant")
        //                    {
        //                        _Body += "<br><b>Status :</b> Pending With Purchase Head";
        //                    }
        //                    else if (updationList.ApprovalStatus == "Approved By Purchase Head" && updationList.LocationType == "Plant")
        //                    {
        //                        _Body += "<br><b>Status :</b> Pending With Plant Head";
        //                    }
        //                    else if (updationList.ApprovalStatus == "Approved By Plant Head" && Convert.ToInt32(updationList.POCost) > 50000 && updationList.LocationType == "Plant")
        //                    {
        //                        _Body += "<br><b>Status :</b> Pending With MD";
        //                    }

        //                    _Body += "<br><h3>Kindly provide your approval</h3>";
        //                    _Body += "<br><h3>For Approval Please Click in the below link</h3>";
        //                    if (updationList.ApprovalStatus == "Submitted" && updationList.LocationType == "Office")
        //                    {
        //                        _Body += "<br><a href=\"https://aparindltd.sharepoint.com/PurchaseOrder/SitePages/PendingWithFunctionalHead.aspx\">View Link</a>";
        //                    }
        //                    else if (updationList.ApprovalStatus == "Approved By Functional Head" && updationList.LocationType == "Office")
        //                    {
        //                        _Body += "<br><a href=\"https://aparindltd.sharepoint.com/PurchaseOrder/SitePages/PendingWithMD.aspx\">View Link</a>";
        //                    }
        //                    else if (updationList.ApprovalStatus == "Submitted" && updationList.LocationType == "Plant")
        //                    {
        //                        _Body += "<br><a href=\"https://aparindltd.sharepoint.com/PurchaseOrder/SitePages/PendingWithPurchaseHead.aspx\">View Link</a>";
        //                    }
        //                    else if (updationList.ApprovalStatus == "Approved By Purchase Head" && updationList.LocationType == "Plant")
        //                    {
        //                        _Body += "<br><a href=\"https://aparindltd.sharepoint.com/PurchaseOrder/SitePages/PendingWithPlantHead.aspx\">View Link</a>";
        //                    }
        //                    else if (updationList.ApprovalStatus == "Approved By Plant Head" && updationList.LocationType == "Plant")
        //                    {
        //                        _Body += "<br><a href=\"https://aparindltd.sharepoint.com/PurchaseOrder/SitePages/PendingWithMD.aspx\">View Link</a>";
        //                    }

        //                //data.MailTo = _From;
        //                //data.MailTo = _To;
        //                //data.MailCC = _Cc;
        //                //data.MailSubject = _Subject;
        //                //data.MailBody = _Body;
        //                //varx.Add(data);
        //                listItem["ToUser"] = _To;
        //                listItem["SubjectDesc"] = _Subject;
        //                listItem["BodyDesc"] = _Body;
        //            if (_To != "")
        //            {
        //                listItem.Update();
        //            }
        //                try
        //                {
        //                    context.ExecuteQuery();
        //                    retValue = true;

        //                }
        //                catch (Exception ex)
        //                {
        //                    CustomSharePointUtility.WriteLog(string.Format("Error in  InsertUpdate_EmployeeMaster ( context.ExecuteQuery();): Error ({0}) ", ex.Message));
        //                    return false;
        //                    //continue;
        //                }
        //            }

        //            //        }
        //            //    }
        //            //}

        //        //}
        //        }

        //        catch (Exception ex)
        //        {
        //            CustomSharePointUtility.WriteLog(string.Format("Error in  InsertUpdate_EmployeeMaster: Error ({0}) ", ex.Message));
        //        }
        //        return retValue;

        //    }
        //}

    }
    public class AppConfiguration
    {
        public string ServiceSiteUrl;
        public string ServiceUserName;
        public string ServicePassword;
    }
}
