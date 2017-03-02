/*
 * Developed By: Logixcess
 * Dated: 30-11-2015
 * All rights reserved by Logixcess.
 */


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OutLook = Microsoft.Office.Interop.Outlook;
using System.Configuration;
using System.Collections.Specialized;
using System.IO;
using System.Threading;
using System.Globalization;
using System.Diagnostics;

namespace AccesstoOutlook
{
    public partial class Form1 : Form
    {

        static DataTable dt = new DataTable();
   //     static bool allDone = false;
        long startIndex = 0, count  = 0;
        int totalRecords = 0;
        OleDbConnection connection;
        OleDbDataAdapter adapter;
        List<OutLook.ContactItem> contactList;
        // new instance of OUTLOOK
        OutLook.MAPIFolder folder;
        Microsoft.Office.Interop.Outlook.Application OutLookApp;
        Microsoft.Office.Interop.Outlook.MAPIFolder Folder_Contacts;
        //string format = "M/d/yyyy hh:mm:ss tt";
        CultureInfo format;
        string path = AppDomain.CurrentDomain.BaseDirectory + "log.ini";
        string dirPath = AppDomain.CurrentDomain.BaseDirectory ;
        public Form1()
        {
            InitializeComponent();
            tbTableName.Text = "Contacts";
            tbContactsFolder.Text = "";
            
            tbContactsFolder.Text = "Test";
            format = new CultureInfo("en-US");
            comboBox1.SelectedIndex = 1;
            comboBox1.Text = comboBox1.Items[1].ToString();
          
        }

        private void btnDbPath_Click(object sender, EventArgs e)
        {
            FileDialog fd = new OpenFileDialog();
            fd.Filter = "Access Database file |*.accdb";
            if (fd.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
            {
                tbAccessDBPath.Text = fd.FileName;
                Passvalues.connectionString = string.Format("Provider=Microsoft.ACE.Oledb.12.0;Data Source=" + tbAccessDBPath.Text.Trim() + ";");
                connection = DB_Connection.GetDBConnection();
            }
        }

        private void btnContactsFolder_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("The System will automatically detect the default folder" + Environment.NewLine + "Click Yes if you want to change it.", "Confirm!", MessageBoxButtons.YesNo) == System.Windows.Forms.DialogResult.Yes)
            {
                FileDialog fd = new OpenFileDialog();

                if (fd.ShowDialog(this) == System.Windows.Forms.DialogResult.OK)
                {
                    tbContactsFolder.Text = fd.FileName;
                }
            }
        }

        private void transfer() {
            LBWait.Visible = true;
            
            checkForAccurateData();
            getContacts();
            
            addContact();
            DisposeExcelInstance();
            LBWait.Visible = false;
        }

        

        private void btnTransfer_Click(object sender, EventArgs e)
        {
            
            if (tbAccessDBPath.Text.Trim() != String.Empty && Passvalues.connectionString != string.Empty)
            {
                //isInserted = false;
                if(OutLookApp == null)
                     OutLookApp = new Microsoft.Office.Interop.Outlook.Application();
                if(Folder_Contacts == null)
                    Folder_Contacts = (OutLook.MAPIFolder)OutLookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
                if ( (folder = searchFolder(tbContactsFolder.Text)) == null )
                    CreateCustomFolder();
                    progressBar1.Value = 0;
               
                    ProgressBarHandler_Timer.Enabled = true;
                
                    ProgressBarHandler_Timer.Start();
                    transfer();
                    ProgressBarHandler_Timer.Stop();
                    ProgressBarHandler_Timer.Enabled = false;
                    
               
            }
            
            else
                MessageBox.Show("Please enter DB Path.");
        }


        

        private void updateRecords() {
            try
            {
                if (connection == null || connection.ConnectionString == null || connection.ConnectionString.Trim() == "")
                    connection = DB_Connection.GetDBConnection();
                if (connection.State == ConnectionState.Closed)
                    connection.Open();
                setLastEditDate();
                
                
                string query = "SELECT * FROM [" + tbTableName.Text.Trim() + "]  WHERE  [Edit Date] > #" + Passvalues.lastUpdateDate + "#";
                adapter = new OleDbDataAdapter(query, connection);
                
                dt.Dispose();
                dt = new DataTable();
                //adapter.Fill(Convert.ToInt32(startIndex), 100, dt);
                adapter.Fill(dt);
                progressBar1.Maximum = dt.Rows.Count;
                adapter.Dispose();
                contactList = new List<OutLook.ContactItem>();
                MethodInvoker m = new MethodInvoker(() => progressBar1.Maximum = Convert.ToInt16(dt.Rows.Count));
                progressBar1.Invoke(m);
                

            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.Message);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (connection.State != ConnectionState.Closed)
                {

                    adapter.Dispose();
                    connection.Close();
                    connection.Dispose();
                }
            }
        }


        private bool updateSingleRecord(string ID)
        {
            try
            {
                if (connection == null || connection.ConnectionString == null || connection.ConnectionString.Trim() == "")
                    connection = DB_Connection.GetDBConnection();
                if (connection.State == ConnectionState.Closed)
                    connection.Open();
                setLastEditDate();

                string query = "SELECT * FROM [" + tbTableName.Text.Trim() + "]  WHERE  [Edit Date] > #" + Passvalues.lastUpdateDate + "# AND [ID] ="+ ID +"";
                adapter = new OleDbDataAdapter(query, connection);

                
                DataTable datarow = new DataTable();
                //adapter.Fill(Convert.ToInt32(startIndex), 100, dt);
                adapter.Fill(datarow);
                if (datarow.Rows.Count > 0)
                    return true;
                else
                    return false;
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.Message);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (connection.State != ConnectionState.Closed)
                {
                    if(adapter != null)
                        adapter.Dispose();
                    connection.Close();
                    connection.Dispose();
                }
            }
            return false;
        }




        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (OutLookApp == null)
                OutLookApp = new Microsoft.Office.Interop.Outlook.Application();
            if (Folder_Contacts == null)
                Folder_Contacts = (OutLook.MAPIFolder)OutLookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
            if (tbAccessDBPath.Text.Trim() != String.Empty && tbContactsFolder.Text.Trim() != string.Empty )
            {
                LBWait.Visible = true;
                totalRecords = 0;
                int newtotalRecords = 0;
                string logText;
                OutLook.ContactItem contact;
                progressBar1.Value = 0;
                
                updateRecords();
                
                count = 0;
                totalRecords = 0;
                startIndex = 0;
                
                progressBar1.Maximum = Convert.ToInt16(count);
                foreach (DataRow dr in dt.Rows) {
                    contact = FindContactEmailByID(dr["ID"].ToString());
                    
                    if (contact != null)
                    {
                        contact = saveContact(contact, dr);
                        contact.Save();
                        totalRecords++;
                        if (contact != null)
                            contact = null;
                        if (progressBar1.Value < progressBar1.Maximum)
                        {
                            MethodInvoker m = new MethodInvoker(() => progressBar1.Value++);
                            progressBar1.Invoke(m);
                        }
                        if (comboBox1.Text == "Level 2")
                        {
                            logText = "Contact updated (" + dr["FirstName"] + " " + dr["LastName"] + ")";
                            logIt(logText);
                        }
                    }
                    else
                    {
                        contact = saveContact(contact, dr);
                        contact.Save();
                        newtotalRecords++;
                        
                        if (contact != null)
                            contact = null;
                        if (progressBar1.Value < progressBar1.Maximum)
                        {
                            MethodInvoker m = new MethodInvoker(() => progressBar1.Value++);
                            progressBar1.Invoke(m);
                        }
                        
                        
                        if (comboBox1.Text == "Level 2")
                        {
                            logText = "New Contact Added (" + dr["FirstName"] + " " + dr["LastName"] + ")";
                            logIt(logText);
                        }
                    }
                }
                LBWait.Visible = false;
                progressBar1.Value = progressBar1.Maximum;
                Passvalues.totalRecords = "" + totalRecords;
                Passvalues.message = "Records Successfully Updated!";
                Notification notification = new Notification();
                notification.Show();


                if (totalRecords > 0)
                {
                    if (comboBox1.Text == "Level 1")
                    {
                        logText = totalRecords + " Contacts Changed," + newtotalRecords + " Contacts Added, 0 Deleted";
                        logIt(logText);
                    }
                    logIt(DateTime.Now.ToString(format));
                }
                totalRecords = 0;
                DisposeExcelInstance();
            }
            else
                MessageBox.Show("Please enter DB Path.");
        }

        private OutLook.ContactItem FindContactEmailByID(String ID)
        {
            OutLook.NameSpace outlookNameSpace = OutLookApp.GetNamespace("MAPI");
            if(folder == null)
                folder = OutLookApp.Session.GetDefaultFolder(
                OutLook.OlDefaultFolders.olFolderContacts).Folders[
                 tbContactsFolder.Text] as OutLook.Folder;

            OutLook.Items contactItems = folder.Items;

            try
            {
                OutLook.ContactItem contact =
                    (OutLook.ContactItem)contactItems.
                    Find(String.Format("[Organizational ID]='{0}'",ID));
                if (contact != null)
                {
                    return contact;
                }
                else
                {
                    
                }
            }
            catch (Exception ex)
            {
                            }
            return null; 
        }

        private void setCount() {
            try
            {
                if (connection == null || connection.ConnectionString == null || connection.ConnectionString.Trim() == "")
                    connection = DB_Connection.GetDBConnection();
                if (connection.State == ConnectionState.Closed)
                    connection.Open();
                string query = "SELECT count(*) FROM [" + tbTableName.Text.Trim() + "]";
                adapter = new OleDbDataAdapter(query, connection);
                dt.Dispose();
                dt = new DataTable();
                adapter.Fill(dt);
                count = Convert.ToInt64(dt.Rows[0][0].ToString());
                
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.Message);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (connection.State != ConnectionState.Closed)
                {

                    connection.Close();
                    connection.Dispose();
                }

            }

        
        }

        private void getContacts()
        {
            try
            {
                if (count == 0)
                    setCount();
                if (connection == null || connection.ConnectionString == null || connection.ConnectionString.Trim() == "")
                    connection = DB_Connection.GetDBConnection();
                if (connection.State == ConnectionState.Closed)
                    connection.Open();

                MethodInvoker m = new MethodInvoker(() => progressBar1.Maximum =Convert.ToInt16(count));
                progressBar1.Invoke(m); 
                string query = "SELECT * FROM [" + tbTableName.Text.Trim() + "] ORDER BY ID";    
                adapter = new OleDbDataAdapter(query, connection);
                dt.Dispose();
                dt = new DataTable();
                adapter.Fill(Convert.ToInt32(startIndex), 100, dt);
                adapter.Dispose();
                contactList = new List<OutLook.ContactItem>();
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.Message);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (connection.State != ConnectionState.Closed)
                {
                    adapter.Dispose();
                    connection.Close();
                    connection.Dispose();
                }
            }
        }




        // outlook work
        
        private OutLook.MAPIFolder searchFolder(string folderName) {
           
                OutLook.NameSpace nameSpace = OutLookApp.GetNamespace("MAPI");
                OutLook.Folders inboxSubfolders = Folder_Contacts.Folders;
                for (int i = 1; inboxSubfolders.Count >= i; i++)
                {
                    OutLook.MAPIFolder subfolderInbox = inboxSubfolders[i];
                    if (subfolderInbox.Name == folderName)
                    {
                        try
                        {
                            return subfolderInbox;
                        }
                        catch (COMException exception)
                        {
                            System.Windows.Forms.MessageBox.Show(exception.Message);
                        }
                    }
                    if (subfolderInbox != null) Marshal.ReleaseComObject(subfolderInbox);
                }
                if (inboxSubfolders != null) Marshal.ReleaseComObject(inboxSubfolders);
                if (nameSpace != null) Marshal.ReleaseComObject(nameSpace);
            
            return null;
        }

        

        private void CreateCustomFolder()
        {
            
            if ((folder = searchFolder( tbContactsFolder.Text)) == null)
            {
                try
                {
                    object missing = System.Reflection.Missing.Value;
                    folder = Folder_Contacts.Folders.Add( tbContactsFolder.Text, OutLook.OlDefaultFolders.olFolderContacts);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            folder = OutLookApp.Session.GetDefaultFolder(
                OutLook.OlDefaultFolders.olFolderContacts).Folders[
                 tbContactsFolder.Text] as OutLook.Folder;
        }




        private void addContact()
        {
            progressBar1.Value = 0;
                Microsoft.Office.Interop.Outlook.ContactItem newContact;
            again:
                string logText = "";
                bool checker = true;

                foreach (DataRow dr in dt.Rows)
                {

                    try
                    {
                        
                        
                        // if contact is new.
                        if ((newContact = FindContactEmailByID(dr["ID"].ToString())) == null)
                        {
                            newContact =
                            folder.Items.Add(
                            "IPM.Contact." + tbContactsFolder.Text) as OutLook.ContactItem;
                            newContact = saveContact(newContact, dr);
                            newContact.Save();
                            totalRecords++;
                            if (progressBar1.Value < progressBar1.Maximum)
                            {
                                MethodInvoker m = new MethodInvoker(() => progressBar1.Value++);
                                progressBar1.Invoke(m);
                            }
                            if (comboBox1.Text == "Level 2")
                            {
                                logText = "New Contact added (Name: " + dr["FirstName"] + " " + dr["LastName"] + ")";
                                logIt(logText);
                            }
                        }
                        else if (newContact != null && updateSingleRecord(dr["ID"].ToString()))
                        {
                            // update contact
                            newContact = saveContact(newContact, dr);
                            newContact.Save();
                            totalRecords++;
                            if (newContact != null)
                                newContact = null;
                            if (progressBar1.Value < progressBar1.Maximum)
                            {
                                MethodInvoker m = new MethodInvoker(() => progressBar1.Value++);
                                progressBar1.Invoke(m);

                            }

                            if (comboBox1.Text == "Level 2")
                            {
                                logText = "Contact Updated (" + dr["FirstName"] + " " + dr["LastName"] + ")";
                                logIt(logText);
                            }
                        }
                        else
                        {
                            checker = false;
                            if(newContact != null)
                                newContact = null;
                        }
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally {
                        newContact = null;
                        dr.Delete();
                    }
                }
                if (totalRecords < count)
                {
                    if(startIndex < totalRecords || !checker)
                        startIndex += 100;
                    if (startIndex >= count)
                        goto end;
                    contactList = null;
                    getContacts();
                    goto again;
                }
            end:
                if (totalRecords > 0)
                    logIt(DateTime.Now.ToString(format));
                progressBar1.Value = progressBar1.Maximum;
                Passvalues.totalRecords = "" + totalRecords;
                Passvalues.message = "Records Successfully Added!";
                Notification notification = new Notification();
                notification.Show();
                startIndex = 0;
                count = 0;
                logIt("");
                totalRecords = 0;
           
        }




        private void checkForAccurateData()
        {
            if ((folder = searchFolder(tbContactsFolder.Text)) == null)
                CreateCustomFolder();
            checkForUpdatedItemsInDB();
            progressBar1.Value = progressBar1.Maximum;
            Passvalues.message = "Data is Accurate!";
            Notification notification = new Notification();
            notification.Show();
            startIndex = 0;
            count = 0;
            logIt("");
            totalRecords = 0;
        }


        // save new Contact OR update 
        private OutLook.ContactItem saveContact(OutLook.ContactItem newContact, DataRow dr)
        {
            newContact.FirstName = dr["Firstname"].ToString();
            newContact.LastName = dr["Lastname"].ToString();
            newContact.Email1Address = dr["Email"].ToString();
            newContact.Account = dr["Account"].ToString();
            newContact.Title = dr["Title"].ToString();
            newContact.WebPage = dr["Website"].ToString();

            newContact.Email2AddressType = "Private";
            newContact.Email2DisplayName = "Private Email";
            newContact.Email2Address = dr["Email Private"].ToString();
            
            newContact.HomeAddress = dr["Address 1 Home"].ToString();
            newContact.HomeAddressCity = dr["City Home"].ToString();
            newContact.HomeAddressCountry = dr["Country Home"].ToString();
            newContact.HomeAddressPostalCode = dr["Postal Code Home"].ToString();

            newContact.UserProperties.Add("Organizational ID", OutLook.OlUserPropertyType.olText, true, OutLook.OlUserPropertyType.olText);
            newContact.UserProperties.Add("Note Description", OutLook.OlUserPropertyType.olText, true, OutLook.OlUserPropertyType.olText);
            newContact.UserProperties.Add("Private Mobile", OutLook.OlUserPropertyType.olText, true, OutLook.OlUserPropertyType.olText);
            //newContact.UserProperties.Add("Edit Date", OutLook.OlUserPropertyType.olText, true, OutLook.OlUserPropertyType.olText);
            //newContact.UserProperties.Add("Create Date", OutLook.OlUserPropertyType.olText, true, OutLook.OlUserPropertyType.olText);
            

            newContact.UserProperties["Private Mobile"].Value = dr["Mobile Phone Private"].ToString();
            newContact.UserProperties["Organizational ID"].Value = dr["ID"].ToString();
            newContact.UserProperties["Note Description"].Value = dr["Note"].ToString();
            //newContact.UserProperties["Edit Date"].Value = dr["Edit Date"].ToString();
            //newContact.UserProperties["Create Date"].Value = dr["CreateDate"].ToString();

            newContact.BusinessAddressPostalCode = dr["Postal Code"].ToString();
            newContact.BusinessFaxNumber = dr["Fax"].ToString();
            newContact.MobileTelephoneNumber = dr["Mobile Phone"].ToString();
            newContact.PrimaryTelephoneNumber = dr["Phone"].ToString();

            newContact.OtherTelephoneNumber = dr["Other Phone"].ToString();
            newContact.BusinessAddress = dr["Address 1"].ToString();

            return newContact;
        }


        private OutLook.ContactItem checkForUpdatedItemsInDB()
        {
            Microsoft.Office.Interop.Outlook.Items OutlookItems;
            //OutLook.MAPIFolder Folder_Contacts;
           // Folder_Contacts = (OutLook.MAPIFolder)OutLookApp.Session.GetDefaultFolder(OutLook.OlDefaultFolders.olFolderContacts);
            OutlookItems = folder.Items;
            progressBar1.Maximum = folder.Items.Count;
            for (int i = 1; i < OutlookItems.Count; i++)
            {
                object missing = System.Reflection.Missing.Value;
                Microsoft.Office.Interop.Outlook.ContactItem contact = (Microsoft.Office.Interop.Outlook.ContactItem)OutlookItems[i];
                if (progressBar1.Value < progressBar1.Maximum)
                {
                    MethodInvoker m = new MethodInvoker(() => progressBar1.Value++);
                    progressBar1.Invoke(m);

                }
                if (contact != null)
                {
                    OutLook.UserProperty userProperty2 = contact.UserProperties.Find("Organizational ID", missing);
                    if (userProperty2 != null && userProperty2.Value != string.Empty)
                    {
                        if (!ifExistsInDB(userProperty2.Value))
                            contact.Delete();
                    }
                }       
               }
                MethodInvoker m2 = new MethodInvoker(() => progressBar1.Value = 0);
                progressBar1.Invoke(m2);

            return null;
        }

        private bool ifExistsInDB(string ID)
        {
            try
            {
                if (connection == null || connection.ConnectionString == null || connection.ConnectionString.Trim() == "")
                    connection = DB_Connection.GetDBConnection();
                if (connection.State == ConnectionState.Closed)
                    connection.Open();
                setLastEditDate();

                string query = "SELECT * FROM [" + tbTableName.Text.Trim() + "]  WHERE  [ID] =" + ID + "";
                adapter = new OleDbDataAdapter(query, connection);


                DataTable datarow = new DataTable();
                //adapter.Fill(Convert.ToInt32(startIndex), 100, dt);
                adapter.Fill(datarow);
                if (datarow.Rows.Count > 0)
                    return true;
                else
                    return false;
            }
            catch (OleDbException e)
            {
                MessageBox.Show(e.Message);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (connection.State != ConnectionState.Closed)
                {
                    if (adapter != null)
                        adapter.Dispose();
                    connection.Close();
                    connection.Dispose();
                }
            }
            return false;
        }
        private void ProgressBarHandler_Timer_Tick(object sender, EventArgs e)
        {
            if (dt.Rows.Count > 0)
            {
                progressBar1.Maximum = dt.Rows.Count;
                if (progressBar1.Value < progressBar1.Maximum)
                    progressBar1.Value++;
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            // delete All Contacts.
            if (tbAccessDBPath.Text.Trim() != string.Empty)
            {
                if (MessageBox.Show("Are you sure to Delete All Contacts From " + tbContactsFolder.Text + " Folder?", "Confirm!"
                    , MessageBoxButtons.YesNo)
                    == System.Windows.Forms.DialogResult.Yes)
                {
                    if (OutLookApp == null)
                        OutLookApp = new Microsoft.Office.Interop.Outlook.Application();
                    if (Folder_Contacts == null)
                        Folder_Contacts = (OutLook.MAPIFolder)OutLookApp.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);

                    folder = OutLookApp.Session.GetDefaultFolder(
                                    OutLook.OlDefaultFolders.olFolderContacts).Folders[
                                    tbContactsFolder.Text] as OutLook.Folder;
                    count = 0;
                    totalRecords = 0;
                    startIndex = 0;
                    getContacts();
                    bool checker = true ;
                    progressBar1.Value = 0;
                    progressBar1.Maximum = Convert.ToInt16(count);
                    Microsoft.Office.Interop.Outlook.ContactItem newContact;
                again:
                    string logText = "";
                    

                    foreach (DataRow dr in dt.Rows)
                    {

                        try
                        {
                            // if contact is new.
                            //newContact = OutLookApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olContactItem)
                            //as Microsoft.Office.Interop.Outlook.ContactItem;
                            if (folder.Items.Count < 1)
                            {
                                MessageBox.Show("No Records in Outlook");
                                checker = false;
                                break;
                            }
                            if ((newContact = FindContactEmailByID(dr["ID"].ToString())) != null)
                            {
                                newContact.Delete();
                                totalRecords++;
                                //newContact.Close(OutLook.OlInspectorClose.olSave);

                                if (progressBar1.Value < progressBar1.Maximum)
                                {
                                    MethodInvoker m = new MethodInvoker(() => progressBar1.Value += 1);
                                    progressBar1.Invoke(m);
                                }
                                if (comboBox1.Text == "Level 2")
                                {

                                    logText = "Contacts Deleted (Name: " + dr["FirstName"] + " " + dr["LastName"] + ")";
                                    logIt(logText);

                                }
                                

                            }
                        }
                        catch (System.Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                        finally
                        {
                            //  newContact.Close(OutLook.OlInspectorClose.olSave);
                            newContact = null;
                            dr.Delete();

                        }
                    }

                    

                    if (totalRecords < count && checker)
                    {
                        //if (startIndex < totalRecords)
                            startIndex += 100;

                        //foreach (OutLook.ContactItem item in contactList) {
                        //    item.Save();
                        //}
                        contactList = null;
                        getContacts();
                        goto again;
                    }

                    progressBar1.Value = 0;

                   
                    
                    if (checker)
                    {
                        MessageBox.Show("Contacts Deleted");
                        logIt("All Contacts have been deleted!");
                        dt.Dispose();
                        count = 0;
                        adapter.Dispose();
                        startIndex = 0;
                        totalRecords = 0;
                        if (comboBox1.Text == "Level 1")
                        {

                            logText = "Contacts Deleted " + totalRecords;
                            logIt(logText);
                          

                        }
                        logIt(DateTime.Now.ToString(format));

                    }
                    //               folder.Delete();
                    if (Folder_Contacts != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(Folder_Contacts);
                    if (folder != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(folder);
                   
                    
                    DisposeExcelInstance();

                }
            }
            else { MessageBox.Show("Please Select Database"); }
        }

        // maintain logging
        private void logIt(string logText)
        {
            using (StreamWriter outputFile = new StreamWriter(path, true))
            {
                if (logText == "")
                {
                    outputFile.WriteLine(DateTime.Now.ToString(format));
                    Passvalues.lastUpdateDate = DateTime.Now.ToString(format);
                }
                else
                {
                    if (comboBox1.Text == "Level 1" || comboBox1.Text == "Level 2" || comboBox1.Text == "Level 0")
                        if (logText != "" & logText != "/n")
                            outputFile.WriteLine(logText);
                }
                
            }
        }

        private void setLastEditDate(){
            try
            {

                if (File.Exists(path))
                {
                    using (StreamReader sr = new StreamReader(path))
                    {
                        string contents = sr.ReadToEnd();
                        if (contents.Length < 10)
                        {
                            Passvalues.lastUpdateDate = DateTime.Now.ToString(format);
                        }
                    }


                    string[] array = File.ReadAllLines(path).Reverse().Take(5).ToArray();
                    foreach (string val in array)
                    {
                        if (val != "")
                        {
                            Passvalues.lastUpdateDate = val;
                            break;
                        }
                    }
                }
                else
                {
                    logIt("");
                }
            }
            catch (IOException ex)
            {}
        }

        // dispose
        public void DisposeExcelInstance()
        {
           if(connection.State == ConnectionState.Open)
            connection.Dispose();

           if (OutLookApp != null)
           {
               try
               {
                   System.Runtime.InteropServices.Marshal.ReleaseComObject(OutLookApp);
                   OutLookApp.Quit();
               }
               catch (Exception ex)
               { }
               finally
               {
                   OutLookApp = null;
               }
           }
           if (Folder_Contacts != null)
           {
               try{
               System.Runtime.InteropServices.Marshal.ReleaseComObject(Folder_Contacts);
               }
               catch (Exception ex)
               { }
               finally
               {
                   Folder_Contacts = null;
               }
           }
            GC.Collect(); 
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.LoggingLevel = comboBox1.SelectedIndex.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dirPath != "")
                    Process.Start(dirPath);
            }
            catch(Exception ex){}
        }
    }
}
