using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using Spire.Doc;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Net.NetworkInformation;
using System.Drawing;

namespace ResumeEngineV2
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public List<string> namesOrdered = new List<string>();
        public List<string> linksOrdered = new List<string>();
        Label lblMinusTextBox;
        TextBox txtBoxSecondKeyword;

        public Form1()
        {
            InitializeComponent();
            //Overlays progress bar ontop of gridview where results are displayed
            progressBar1.BringToFront();

            //Set combo box to default value, 100%
            cmbWeight.SelectedIndex = 9;

            //Add tool tip to + label
            ToolTip tt = new ToolTip();
            tt.SetToolTip(lblAddTextBox, "Click to add another keyword field");

            tt.SetToolTip(lblUsername, "Example: jbraham@aecon.com");

            //Checks to see if creds.xml exists, if not creates file
            if (System.IO.File.Exists("creds.xml") == false)
            {
                using (FileStream fs = System.IO.File.Create("creds.xml"))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>" + Environment.NewLine + "<credentials>" + Environment.NewLine + "<username>***</username>" + Environment.NewLine + "<password>***</password>" + Environment.NewLine + "</credentials>");
                    fs.Write(info, 0, info.Length);
                }
                Encrypt();
            }
            XmlDocument doc;
            //Loads in xml data in creds.xml
            try
            {
                Decrypt();
                doc = new XmlDocument();
                doc.Load("creds.xml");
                Encrypt();
            }
            //Problem with file, delete it create new one with *** as username and pass which will force user to login in again
            catch
            {
                MessageBox.Show("Failed to open data file. You will need to login in again!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                System.IO.File.Delete("creds.xml");
                using (FileStream fs = System.IO.File.Create("creds.xml"))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>" + Environment.NewLine + "<credentials>" + Environment.NewLine + "<username>***</username>" + Environment.NewLine + "<password>***</password>" + Environment.NewLine + "</credentials>");
                    fs.Write(info, 0, info.Length);
                }
                doc = new XmlDocument();
                doc.Load("creds.xml");
                Encrypt();
            }

            //Checks to see if data is '***' rather than actual data which is the case when the user logouts and does not log back in, or if the file was just created
            if (doc.DocumentElement.SelectSingleNode("/credentials/password").InnerText == "***")
            {
                //Only display login fields
                lblEnterKeyword.Visible = false;
                txtBoxKeyword.Visible = false;
                btnKeywordSubmit.Visible = false;
                lblResults.Visible = false;
                resultsView.Visible = false;
                progressBar1.Visible = false;
                btnLogout.Visible = false;
                lblAddTextBox.Visible = false;
                cmbWeight.Visible = false;
                txtBoxExperience.Visible = false;
                lblExperience.Visible = false;
                this.AcceptButton = btnLoginSubmit;
            }
            else
            {
                //Do not display login fields
                lblLogin.Visible = false;
                textBoxUsername.Visible = false;
                textBoxPassword.Visible = false;
                lblUsername.Visible = false;
                lblPassword.Visible = false;
                btnLoginSubmit.Visible = false;
                this.Text = "Resume Search Engine - Logged in as " + doc.DocumentElement.SelectSingleNode("/credentials/username").InnerText;
                this.AcceptButton = btnKeywordSubmit;
            }
        }

        private void btnLoginSubmit_Click(object sender, EventArgs e)
        {
            //Verify login credentials
            string targetSiteURL = @"https://aecon1.sharepoint.com/sites/bd/resume/";

            var login = textBoxUsername.Text;
            var password = textBoxPassword.Text;

            var securePassword = new SecureString();

            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }
            //User tries to login
            try
            {
                SharePointOnlineCredentials onlineCredentials = new SharePointOnlineCredentials(login, securePassword);
                ClientContext ctx = new ClientContext(targetSiteURL);
                ctx.Credentials = onlineCredentials;
                var web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();

                //Load creds.xml and add user login credentials
                try
                {
                    Decrypt();
                }
                catch
                {
                    MessageBox.Show("Failed to decrypt credentials file. Logging out!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    if (System.IO.File.Exists("creds.xml") == true)
                    {
                        System.IO.File.Delete("creds.xml");
                    }

                    using (FileStream fs = System.IO.File.Create("creds.xml"))
                    {
                        Byte[] info = new UTF8Encoding(true).GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>" + Environment.NewLine + "<credentials>" + Environment.NewLine + "<username>***</username>" + Environment.NewLine + "<password>***</password>" + Environment.NewLine + "</credentials>");
                        fs.Write(info, 0, info.Length);
                    }
                    Encrypt();

                    btnLogout_Click(sender, e);
                    return;
                }

                XmlDocument doc = new XmlDocument();
                doc.Load("creds.xml");
                doc.DocumentElement.SelectSingleNode("/credentials/username").InnerText = textBoxUsername.Text;
                doc.DocumentElement.SelectSingleNode("/credentials/password").InnerText = textBoxPassword.Text;
                doc.Save("creds.xml");
                Encrypt();

                //Hide login fields
                lblLogin.Visible = false;
                textBoxUsername.Visible = false;
                textBoxPassword.Visible = false;
                lblUsername.Visible = false;
                lblPassword.Visible = false;
                btnLoginSubmit.Visible = false;

                lblEnterKeyword.Visible = true;
                txtBoxKeyword.Visible = true;
                btnKeywordSubmit.Visible = true;
                lblResults.Visible = true;
                resultsView.Visible = true;
                progressBar1.BringToFront();
                progressBar1.Visible = true;
                btnLogout.Visible = true;
                lblAddTextBox.Visible = true;
                cmbWeight.Visible = true;
                txtBoxExperience.Visible = true;
                lblExperience.Visible = true;
                this.AcceptButton = btnKeywordSubmit;
                this.Text = "Resume Search Engine - Logged in as " + textBoxUsername.Text;

                txtBoxKeyword.Focus();
            }
            //Bad credentials, get user to try and login again
            catch (Exception ex)
            {
                MessageBox.Show("Failed to authenticate username or password! Please try again.\n\nDetails:\n" + ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBoxUsername.Text = "Example: jbraham@aecon.com";
                textBoxUsername.ForeColor = SystemColors.ButtonShadow;
                textBoxPassword.Text = "";

                btnLoginSubmit.Focus();
            }
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            //Load xml and set credentials to ***
            try
            {
                Decrypt();
            }
            catch
            {
                MessageBox.Show("Failed to decrypt credentials file. Logging out!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                if (System.IO.File.Exists("creds.xml") == true)
                {
                    System.IO.File.Delete("creds.xml");
                }

                using (FileStream fs = System.IO.File.Create("creds.xml"))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>" + Environment.NewLine + "<credentials>" + Environment.NewLine + "<username>***</username>" + Environment.NewLine + "<password>***</password>" + Environment.NewLine + "</credentials>");
                    fs.Write(info, 0, info.Length);
                }
                Encrypt();

                btnLogout_Click(sender, e);
                return;
            }
            XmlDocument doc = new XmlDocument();
            doc.Load("creds.xml");
            doc.DocumentElement.SelectSingleNode("/credentials/username").InnerText = "***";
            doc.DocumentElement.SelectSingleNode("/credentials/password").InnerText = "***";
            doc.Save("creds.xml");
            Encrypt();

            //Whipe data stored in fields
            this.Text = "Resume Search Engine";
            this.AcceptButton = btnLoginSubmit;
            txtBoxKeyword.Text = "";
            textBoxUsername.Text = "Example: jbraham@aecon.com";
            textBoxUsername.ForeColor = SystemColors.ButtonShadow;
            textBoxPassword.Text = "";
            lblResults.Text = "Results:";
            cmbWeight.ResetText();
            if (cmbWeight.Enabled == true)
            {
                cmbWeight.Items.Insert(9, "100%");
            }
            cmbWeight.SelectedIndex = 9;
            txtBoxExperience.Text = "0";

            //Only show login fields
            lblLogin.Visible = true;
            textBoxUsername.Visible = true;
            textBoxPassword.Visible = true;
            lblUsername.Visible = true;
            lblPassword.Visible = true;
            btnLoginSubmit.Visible = true;

            lblEnterKeyword.Visible = false;
            txtBoxKeyword.Visible = false;
            btnKeywordSubmit.Visible = false;
            lblResults.Visible = false;
            resultsView.Visible = false;
            progressBar1.Visible = false;
            btnLogout.Visible = false;
            if (lblAddTextBox.Visible == false)
            {
                lblMinusTextBox.Visible = false;
                txtBoxSecondKeyword.Text = "";
                txtBoxSecondKeyword.Visible = false;
            }
            else
            {
                lblAddTextBox.Visible = false;
            }
            cmbWeight.Visible = false;
            cmbWeight.Enabled = false;
            txtBoxExperience.Visible = false;
            lblExperience.Visible = false;
        }

        private void btnKeywordSubmit_Click(object sender, EventArgs e)
        {
            //Library of keywords
            string[] energyLib = { "Energy", "Bruce", "Cogeneration", "Fabrication", "Gas", "Module", "Modules", "Nuclear", "Oil", "OPG", "Ontario Power Generation", "Pipeline", "Pipelines", "Utilities" };
            string[] infrastructureLib = { "Infrastructure", "Airport", "Airports", "Asphalt", "Bridge", "Bridges", "Hydroelectric", "Rail", "Road", "Roads", "Transit", "Tunnel", "Tunnels", "Water Treatment" };
            string[] miningLib = { "Mining", "Fabrication", "Mechanical Works", "Mine Site Development", "Module", "Modules", "Overburden Removal", "Processing Facilities", "Reclamation" };
            string[] concessionsLib = { "Concessions", "Accounting", "Bank", "Banks", "Equity Investments", "Maintenance", "Operations", "Project Financing", "Project Development", "Public Private Partnership", "P3" };
            string[] otherLib = { "Advisor", "Boilermaker", "Buyer", "CAD", "Carpenter", "Concrete", "Contract", "Controller", "Controls", "Coordinator", "Counsel", "Craft Recruiter", "Customer Service Representative", "Designer", "Dockmaster", "Document Control", "Draftsperson", "E&I", "Electrical and Instrumentation", "EHS", "Environmental health and safety", "Electrician", "Engineer", "Environment", "Equipment", "Estimator", "Field Support", "Network Support", "Fitter", "Welder", "Foreperson", "Foreman", "Inspector", "Ironwork", "Labourer", "Lead", "Locator", "Material", "Operator", "Pavement", "PEng", "Professional Engineer", "Planner", "Plumber", "Project Design", "Purchaser", "Requisitioner", "Risk", "Scheduler", "Specialist", "Splicer", "Superintendent", "Supervisor", "Support", "Surveyor", "Technical Services", "Technician", "Turnover", "Vendor" };

            //Whipe global Lists
            namesOrdered.Clear();
            linksOrdered.Clear();

            //User must enter something for search to work, if extra field is up user must fill in something in both fields for submit to work
            if (txtBoxKeyword.Text == "" || txtBoxKeyword.Text.Contains("\"") || txtBoxKeyword.Text.Contains("\\"))
            {
                MessageBox.Show("Please enter a valid keyword in first text field!\n\nKeyword can not be empty\nKeyword can not contain the following characters:\n\" (double quotation mark) or \\ (Backslash)", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (lblAddTextBox.Visible == false && (txtBoxSecondKeyword.Text == "" || txtBoxSecondKeyword.Text.Contains("\"") || txtBoxSecondKeyword.Text.Contains("\\")))
            {
                MessageBox.Show("Please enter a valid keyword in second text field!\n\nKeyword can not be empty\nKeyword can not contain the following characters:\n\" (double quotation mark) or \\ (Backslash)", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (!int.TryParse(txtBoxExperience.Text, out int tempOut) || tempOut < 0)
            {
                MessageBox.Show("Please enter a valid number for years of experience greater then or equal to zero!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            //If user enters first keyword in our lib, second keyword must also be in the lib
            else if (lblAddTextBox.Visible == false && (energyLib.Contains(txtBoxKeyword.Text, StringComparer.OrdinalIgnoreCase) || infrastructureLib.Contains(txtBoxKeyword.Text, StringComparer.OrdinalIgnoreCase) || miningLib.Contains(txtBoxKeyword.Text, StringComparer.OrdinalIgnoreCase) || concessionsLib.Contains(txtBoxKeyword.Text, StringComparer.OrdinalIgnoreCase) || otherLib.Contains(txtBoxKeyword.Text, StringComparer.OrdinalIgnoreCase)) && (!energyLib.Contains(txtBoxSecondKeyword.Text, StringComparer.OrdinalIgnoreCase) && !infrastructureLib.Contains(txtBoxSecondKeyword.Text, StringComparer.OrdinalIgnoreCase) && !miningLib.Contains(txtBoxSecondKeyword.Text, StringComparer.OrdinalIgnoreCase) && !concessionsLib.Contains(txtBoxSecondKeyword.Text, StringComparer.OrdinalIgnoreCase) && !otherLib.Contains(txtBoxSecondKeyword.Text, StringComparer.OrdinalIgnoreCase)))
            {
                MessageBox.Show("Your first keyword is in our library and so the second keyword must also be in the library!\n\nList of Keywords:\n\nEnergy, Bruce, Cogeneration, Fabrication, Gas, Modules, Nuclear, Oil, OPG, Ontario Power Generation, Pipelines, Utilities\n\nInfrastructure, Airports, Asphalt, Bridges, Hydroelectric, Rail, Road, Transit, Tunnels, Water Treatment\n\nMining, Fabrication, Mechanical Works, Mine Site Development, Modules, Overburden Removal, Processing Facilities, Reclamation\n\nConcessions, Accounting, Bank,Equity Investments, Maintenance, Operations, Project Financing, Project Development, Public Private Partnership, P3s\n\nAdvisor, Boilermaker, Buyer, CAD, Carpenter, Concrete, Contract, Controller, Controls, Coordinator,Counsel, Craft Recruiter, Customer Service Representative, Designer, Dockmaster, Document Control, Draftsperson, E & I, Electrical and Instrumentation, EHS, Environmental health and safety, Electrician, Engineer, Environment, Equipment, Estimator, Field Support, Network Support, Fitter, Welder, Foreperson, Foreman, Inspector, Ironwork, Labourer, Lead, Locator, Material, Operator, Pavement, PEng, Professional Engineer, Planner, Plumber, Project Design, Purchaser, Requisitioner, Risk, Scheduler,Specialist, Splicer, Superintendent, Supervisor, Support, Surveyor, Technical Services, Technician, Turnover, Vendor", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else
            {
                progressBar1.Visible = true;

                //Fixes weird bug where label is cut off
                lblResults.Text = "";
                lblResults.Text = "Results:";

                //Whipes results GridView of old data
                resultsView.Rows.Clear();
                resultsView.Refresh();

                //Disables fields to stop user from using them while results are being fetched
                btnKeywordSubmit.Enabled = false;
                txtBoxKeyword.Enabled = false;
                btnLogout.Enabled = false;
                cmbWeight.Enabled = false;
                if (lblAddTextBox.Visible == true)
                {
                    lblAddTextBox.Enabled = false;
                }
                else
                {
                    lblMinusTextBox.Enabled = false;
                    txtBoxSecondKeyword.Enabled = false;
                }
                txtBoxExperience.Enabled = false;

                string targetSiteURL = @"https://aecon1.sharepoint.com/sites/bd/resume/";

                //Read credentials from creds.xml
                //Decrypt failed, delete creds file, get user to login again
                try
                {
                    Decrypt();
                }
                catch
                {
                    MessageBox.Show("Failed to decrypt credentials file. Logging out!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    if (System.IO.File.Exists("creds.xml") == true)
                    {
                        System.IO.File.Delete("creds.xml");
                    }
                        
                    using (FileStream fs = System.IO.File.Create("creds.xml"))
                    {
                        Byte[] info = new UTF8Encoding(true).GetBytes("<?xml version=\"1.0\" encoding=\"utf-8\"?>" + Environment.NewLine + "<credentials>" + Environment.NewLine + "<username>***</username>" + Environment.NewLine + "<password>***</password>" + Environment.NewLine + "</credentials>");
                        fs.Write(info, 0, info.Length);
                    }
                    Encrypt();

                    btnKeywordSubmit.Enabled = true;
                    txtBoxKeyword.Enabled = true;
                    btnLogout.Enabled = true;
                    cmbWeight.Enabled = true;
                    lblAddTextBox.Enabled = true;
                    txtBoxExperience.Enabled = true;
                    btnLogout_Click(sender, e);
                    return;
                }
                XmlDocument doc = new XmlDocument();
                doc.Load("creds.xml");
                Encrypt();

                var login = doc.DocumentElement.SelectSingleNode("/credentials/username").InnerText;
                var password = doc.DocumentElement.SelectSingleNode("/credentials/password").InnerText;
                string term = txtBoxKeyword.Text;
                string term2 = "";
                if (lblAddTextBox.Visible == false)
                {
                    term2 = txtBoxSecondKeyword.Text;
                }

                var securePassword = new SecureString();

                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                }

                ClientContext ctx;
                Web web;
                //Try and connect SharePoint
                try
                {
                    SharePointOnlineCredentials onlineCredentials = new SharePointOnlineCredentials(login, securePassword);

                    ctx = new ClientContext(targetSiteURL);
                    ctx.Credentials = onlineCredentials;
                    web = ctx.Web;
                    ctx.Load(web);
                    ctx.ExecuteQuery();
                }
                //Could not connect probably because of invalid credentials which could occur if user logged in to ResumeEngine but credentials were revoked in SharePoint later on
                catch
                {
                    MessageBox.Show("Could not authenticate credentials. Logging out!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    btnKeywordSubmit.Enabled = true;
                    txtBoxKeyword.Enabled = true;
                    btnLogout.Enabled = true;
                    cmbWeight.Enabled = true;
                    lblAddTextBox.Enabled = true;
                    txtBoxExperience.Enabled = true;
                    btnLogout_Click(sender, e);
                    return;
                }

                //Gets all folders under Documents
                var list = web.Lists.GetByTitle("Documents");
                ctx.Load(list);
                ctx.Load(list.RootFolder);
                ctx.Load(list.RootFolder.Folders);
                ctx.ExecuteQuery();
                FolderCollection fcol = list.RootFolder.Folders;
                List<string> lstFile = new List<string>();

                List<string> names = new List<string>();

                //Loops through each folder
                foreach (Folder f in fcol)
                {
                    //If folder is named Text
                    if (f.Name == "Original")
                    {
                        //Get all files under text folder
                        ctx.Load(f.Files);
                        ctx.ExecuteQuery();
                        var listItems = f.Files;

                        List<object> arguments = new List<object>();
                        arguments.Add(listItems);
                        arguments.Add(web);
                        arguments.Add(ctx);
                        arguments.Add(term);
                        arguments.Add(names);
                        arguments.Add(listItems.Count());
                        arguments.Add(term2);
                        arguments.Add(cmbWeight.Text);
                        backgroundWorker1.WorkerReportsProgress = true;
                        backgroundWorker1.RunWorkerAsync(arguments);
                        break;
                    }
                }
            }
        }

        private void backgroundWoker1_DoWork(object sender, DoWorkEventArgs e)
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();

            List<object> arguments = e.Argument as List<object>;
            int totalCount = (int)arguments[5];
            int count = 0;
            List<string> names = (System.Collections.Generic.List<string>)arguments[4];
            List<string> links = new List<string>();
            Web web = (Web)arguments[1];
            string weight = (string)arguments[7];

            //Arrays holding data to be sent to cortical.io api
            string[] postData = new string[10];
            for (int i = 0; i < postData.Length; i++)
            {
                postData[i] = "[";
            }
            string[] postData2 = new string[10];
            for (int i = 0; i < postData2.Length; i++)
            {
                postData2[i] = "";
            }
            if (!String.IsNullOrEmpty((string)arguments[6]))
            {
                for (int i = 0; i < postData2.Length; i++)
                {
                    postData2[i] = "[";
                }
            }
            int postDataCount = 0;

            bool isUsingCortical = true;
            List<KeyValuePair<double, int>> matchScoreName = new List<KeyValuePair<double, int>>();
            List<KeyValuePair<double, string>> matchScoreLink = new List<KeyValuePair<double, string>>();
            int matchScoreCounter = 0;

            //Check if there are resumes
            if (totalCount <= 0)
            {
                MessageBox.Show("\nThere are no Resumes in the SharePoint");
                List<object> newArgs = new List<object>();
                newArgs.Add("Results:");
                newArgs.Add(true);
                e.Result = newArgs;
                return;
            }

            //0 = energy, 1 = infrastructure, 2 = mining, 3 = conecessions, 4 = other
            int whichLib = -1;
            int whichLib2 = -1;

            //Library of keywords
            string[] energyLib = { "Energy", "Bruce", "Cogeneration", "Fabrication", "Gas", "Module", "Modules", "Nuclear", "Oil", "OPG", "Ontario Power Generation", "Pipeline", "Pipelines", "Utilities" };
            string[] infrastructureLib = { "Infrastructure", "Airport", "Airports", "Asphalt", "Bridge", "Bridges", "Hydroelectric", "Rail", "Road", "Roads", "Transit", "Tunnel", "Tunnels", "Water Treatment" };
            string[] miningLib = { "Mining", "Fabrication", "Mechanical Works", "Mine Site Development", "Module", "Modules", "Overburden Removal", "Processing Facilities", "Reclamation" };
            string[] concessionsLib = { "Concessions", "Accounting", "Bank", "Banks", "Equity Investments", "Maintenance", "Operations", "Project Financing", "Project Development", "Public Private Partnership", "P3" };
            string[] otherLib = { "Advisor", "Boilermaker", "Buyer", "CAD", "Carpenter", "Concrete", "Contract", "Controller", "Controls", "Coordinator", "Counsel", "Craft Recruiter", "Customer Service Representative", "Designer", "Dockmaster", "Document Control", "Draftsperson", "E&I", "Electrical and Instrumentation", "EHS", "Environmental health and safety", "Electrician", "Engineer", "Environment", "Equipment", "Estimator", "Field Support", "Network Support", "Fitter", "Welder", "Foreperson", "Foreman", "Inspector", "Ironwork", "Labourer", "Lead", "Locator", "Material", "Operator", "Pavement", "PEng", "Professional Engineer", "Planner", "Plumber", "Project Design", "Purchaser", "Requisitioner", "Risk", "Scheduler", "Specialist", "Splicer", "Superintendent", "Supervisor", "Support", "Surveyor", "Technical Services", "Technician", "Turnover", "Vendor" };

            //Check if keyword matches any keywords in library and thus we are not using cortical.io
            if (energyLib.Contains((string)arguments[3], StringComparer.OrdinalIgnoreCase))
            {
                whichLib = 0;
                isUsingCortical = false;
            }
            else if (infrastructureLib.Contains((string)arguments[3], StringComparer.OrdinalIgnoreCase))
            {
                whichLib = 1;
                isUsingCortical = false;
            }
            else if (miningLib.Contains((string)arguments[3], StringComparer.OrdinalIgnoreCase))
            {
                whichLib = 2;
                isUsingCortical = false;
            }
            else if (concessionsLib.Contains((string)arguments[3], StringComparer.OrdinalIgnoreCase))
            {
                whichLib = 3;
                isUsingCortical = false;
            }
            else if (otherLib.Contains((string)arguments[3], StringComparer.OrdinalIgnoreCase))
            {
                whichLib = 4;
                isUsingCortical = false;
            }

            if (!String.IsNullOrEmpty(postData2[0]))
            {
                if (energyLib.Contains((string)arguments[6], StringComparer.OrdinalIgnoreCase))
                {
                    whichLib2 = 0;
                }
                else if (infrastructureLib.Contains((string)arguments[6], StringComparer.OrdinalIgnoreCase))
                {
                    whichLib2 = 1;
                }
                else if (miningLib.Contains((string)arguments[6], StringComparer.OrdinalIgnoreCase))
                {
                    whichLib2 = 2;
                }
                else if (concessionsLib.Contains((string)arguments[6], StringComparer.OrdinalIgnoreCase))
                {
                    whichLib2 = 3;
                }
                else if (otherLib.Contains((string)arguments[6], StringComparer.OrdinalIgnoreCase))
                {
                    whichLib2 = 4;
                }
            }

            FileCollection fc = (FileCollection)arguments[0];
            int numFiles = fc.Count;

            //Loops through each file
            foreach (var item in (FileCollection)arguments[0])
            {
                count++;
                string fileName = item.Name;
                string newConvText = "";

                if (System.IO.File.Exists("textResumes/" + fileName + ".txt") == false)
                {
                    var filePath = web.ServerRelativeUrl + "/Shared%20Documents/Original/" + fileName;
                    //var filePathTxt = web.ServerRelativeUrl + "/Shared%20Documents/Text/" + fileName + ".txt";
                    FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect((ClientContext)arguments[2], filePath);
                    string ext = System.IO.Path.GetExtension(fileName);
                    string convText = "";

                    //Convert file into text
                    try
                    {
                        if (ext == ".pdf")
                        {
                            //Using ITextSharp pdf library
                            using (PdfReader reader = new PdfReader(fileInformation.Stream))
                            {
                                StringBuilder textBuild = new StringBuilder();
                                for (int i = 1; i <= reader.NumberOfPages; i++)
                                {
                                    textBuild.Append(PdfTextExtractor.GetTextFromPage(reader, i));
                                }
                                convText = textBuild.ToString();
                            }
                        }
                        else
                        {
                            //Using Spire office library instead of interop because interop is slow and Microsoft does not currently recommend,
                            //and does not support, Automation of Microsoft Office applications from any unattended non-interactive client application or component
                            using (var stream1 = new MemoryStream())
                            {
                                MemoryStream txtStream = new MemoryStream();
                                Document document = new Document();
                                fileInformation.Stream.CopyTo(stream1);
                                document.LoadFromStream(stream1, FileFormat.Auto);
                                document.SaveToStream(txtStream, FileFormat.Txt);
                                txtStream.Position = 0;

                                StreamReader reader = new StreamReader(txtStream);
                                string readText = reader.ReadToEnd();

                                //Remove watermark for spire
                                readText = readText.Replace("Evaluation Warning: The document was created with Spire.Doc for .NET.", "");
                                convText = readText;
                            }
                        }
                    }
                    catch
                    {
                        MessageBox.Show(fileName + " cannot be opened! Skipping this file.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        continue;
                    }

                    List<Char> builder = new List<char>();
                    //Used to fix if there are multiple newlines in a row
                    bool isNewLine = true;

                    //Remove special characters which would need to be escaped for JSON and creates new string using builder var
                    for (int i = 0; i < convText.Length; i++)
                    {
                        if (convText[i] == '\t')
                        {
                            builder.Add(' ');
                        }
                        else if (convText[i] == char.MinValue)
                        {
                            builder.Add(' ');
                        }
                        else if (convText[i] == '\\')
                        {
                            builder.Add('\\');
                            builder.Add('\\');
                        }
                        else if ((convText[i] == '\n' || convText[i] == '\r') && isNewLine == false)
                        {
                            if (convText[i - 1] == '.' || convText[i - 1] == ':' || convText[i - 1] == ',')
                            {
                                builder.Add(' ');
                            }
                            else if (convText[i - 1] != ' ')
                            {
                                builder.Add('.');
                                builder.Add(' ');
                            }
                            isNewLine = true;
                        }
                        else if (convText[i] != '\n' && convText[i] != '\r')
                        {
                            isNewLine = false;
                            //If '"' is already escaped ignore
                            if (convText[i] == '"' && convText[i - 1] != '\\')
                            {
                                //Adds a single '\' before the '"'
                                builder.Add('\\');
                                builder.Add('"');
                            }
                            else
                            {
                                builder.Add(convText[i]);
                            }
                        }
                    }
                    newConvText = new string(builder.ToArray());
                    if (Directory.Exists("textResumes") == false)
                    {
                        Directory.CreateDirectory("textResumes");
                    }
                    System.IO.File.WriteAllText("textResumes/" + fileName + ".txt", newConvText);
                }
                else
                {
                    newConvText = System.IO.File.ReadAllText("textResumes/" + fileName + ".txt");
                }

                //Calculate years of experience
                bool inExperience = false;
                int lowestYear = -1;
                int tempYear = 0;
                //Loop through doc word by word
                foreach (string word in newConvText.Split(' '))
                {
                    //Check if in experience section and came across a number 
                    if (inExperience == true && int.TryParse(word, out tempYear))
                    {
                        //If the number is greater then 1960 and less then the current year, then if this is the first number found or the number is less then the current smallest number, store it
                        if (tempYear > 1960 && tempYear < DateTime.Now.Year && (lowestYear == -1 || lowestYear > tempYear))
                        {
                            lowestYear = tempYear;
                        }
                    }
                    //If come across education section while searching in experience section stop searching
                    else if (inExperience == true && String.Equals(word, "Education", StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                    //If come across education sections or employment section start searching for years of experience
                    else if (String.Equals(word, "Experience", StringComparison.OrdinalIgnoreCase) || String.Equals(word, "Employment", StringComparison.OrdinalIgnoreCase) || String.Equals(word, "Employ", StringComparison.OrdinalIgnoreCase))
                    {
                        inExperience = true;
                    }
                }
                int experienceYears = 0;
                //If a lowest year was found, calculate years of experience
                if (lowestYear != -1)
                {
                    experienceYears = DateTime.Now.Year - lowestYear;
                }
                int txtBoxOutExperience;
                int.TryParse(txtBoxExperience.Text, out txtBoxOutExperience);

                //Only use candidates with the necessary years of experience in the final results
                if (experienceYears >= txtBoxOutExperience)
                {
                    links.Add(item.LinkingUri);
                    names.Add(fileName.Replace(".txt", ""));

                    //Use own library or cortical.io
                    if (isUsingCortical == false)
                    {
                        int numExactMatches = 0;
                        int numCategoryMatches = 0;

                        int numExactMatchesSecond = 0;
                        int numCategoryMatchesSecond = 0;
                        //Check occurances of keywords in resume
                        foreach (string word in newConvText.Split(' '))
                        {
                            if (whichLib == 0)
                            {
                                if (energyLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                {
                                    if (String.Equals((string)arguments[3], word, StringComparison.OrdinalIgnoreCase))
                                    {
                                        numExactMatches++;
                                    }
                                    else
                                    {
                                        numCategoryMatches++;
                                    }
                                }
                            }
                            else if (whichLib == 1)
                            {
                                if (infrastructureLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                {
                                    if (String.Equals((string)arguments[3], word, StringComparison.OrdinalIgnoreCase))
                                    {
                                        numExactMatches++;
                                    }
                                    else
                                    {
                                        numCategoryMatches++;
                                    }
                                }
                            }
                            else if (whichLib == 2)
                            {
                                if (miningLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                {
                                    if (String.Equals((string)arguments[3], word, StringComparison.OrdinalIgnoreCase))
                                    {
                                        numExactMatches++;
                                    }
                                    else
                                    {
                                        numCategoryMatches++;
                                    }
                                }
                            }
                            else if (whichLib == 3)
                            {
                                if (concessionsLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                {
                                    if (String.Equals((string)arguments[3], word, StringComparison.OrdinalIgnoreCase))
                                    {
                                        numExactMatches++;
                                    }
                                    else
                                    {
                                        numCategoryMatches++;
                                    }
                                }
                            }
                            else if (whichLib == 4)
                            {
                                if (otherLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                {
                                    if (String.Equals((string)arguments[3], word, StringComparison.OrdinalIgnoreCase))
                                    {
                                        numExactMatches++;
                                    }
                                }
                            }

                            //Do it again for second keyword if the field exists
                            if (!String.IsNullOrEmpty(postData2[0]))
                            {
                                if (whichLib2 == 0)
                                {
                                    if (energyLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                    {
                                        if (String.Equals((string)arguments[6], word, StringComparison.OrdinalIgnoreCase))
                                        {
                                            numExactMatchesSecond++;
                                        }
                                        else
                                        {
                                            numCategoryMatchesSecond++;
                                        }
                                    }
                                }
                                else if (whichLib2 == 1)
                                {
                                    if (infrastructureLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                    {
                                        if (String.Equals((string)arguments[6], word, StringComparison.OrdinalIgnoreCase))
                                        {
                                            numExactMatchesSecond++;
                                        }
                                        else
                                        {
                                            numCategoryMatchesSecond++;
                                        }
                                    }
                                }
                                else if (whichLib2 == 2)
                                {
                                    if (miningLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                    {
                                        if (String.Equals((string)arguments[6], word, StringComparison.OrdinalIgnoreCase))
                                        {
                                            numExactMatchesSecond++;
                                        }
                                        else
                                        {
                                            numCategoryMatchesSecond++;
                                        }
                                    }
                                }
                                else if (whichLib2 == 3)
                                {
                                    if (concessionsLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                    {
                                        if (String.Equals((string)arguments[6], word, StringComparison.OrdinalIgnoreCase))
                                        {
                                            numExactMatchesSecond++;
                                        }
                                        else
                                        {
                                            numCategoryMatchesSecond++;
                                        }
                                    }
                                }
                                else if (whichLib2 == 4)
                                {
                                    if (otherLib.Contains(word, StringComparer.OrdinalIgnoreCase))
                                    {
                                        if (String.Equals((string)arguments[6], word, StringComparison.OrdinalIgnoreCase))
                                        {
                                            numExactMatchesSecond++;
                                        }
                                    }
                                }
                            }
                        }
                        double totalMatchScore = 0;
                        if (!String.IsNullOrEmpty(postData2[0]))
                        {
                            double firstWeight = (double)(Int32.Parse(weight.Replace("%", ""))) / 100;
                            double secondWeight = 1 - firstWeight;

                            double matchPercent = (numExactMatches + ((double)numCategoryMatches / 10)) * 10;
                            double matchPercent2 = (numExactMatchesSecond + ((double)numCategoryMatchesSecond / 10)) * 10;
                            totalMatchScore = (((matchPercent / 100) * firstWeight) + ((matchPercent2 / 100) * secondWeight)) * 100;  
                        }
                        else
                        {
                            totalMatchScore = (numExactMatches + ((double)numCategoryMatches / 10)) * 10;
                        }
                        matchScoreLink.Add(new KeyValuePair<double, string>(totalMatchScore, links[matchScoreCounter]));
                        matchScoreName.Add(new KeyValuePair<double, int>(totalMatchScore, matchScoreCounter++));
                    }
                    else
                    {
                        //Build JSON request string each loop
                        postData[postDataCount] += "[{\"term\": \"" + (string)arguments[3] + "\"},{\"text\": \"" + newConvText + "\"}],";
                        if (!String.IsNullOrEmpty(postData2[0]))
                        {
                            postData2[postDataCount] += "[{\"term\": \"" + (string)arguments[6] + "\"},{\"text\": \"" + newConvText + "\"}],";
                        }
                    }
                }

                //Incriment postDataCount if number of files is past limits
                if (postDataCount == 0 && count > 200)
                {
                    postDataCount = 1;
                }
                else if (postDataCount == 1 && count > 400)
                {
                    postDataCount = 2;
                }
                else if (postDataCount == 2 && count > 600)
                {
                    postDataCount = 3;
                }
                else if (postDataCount == 3 && count > 800)
                {
                    postDataCount = 4;
                }
                else if (postDataCount == 4 && count > 1000)
                {
                    postDataCount = 5;
                }
                else if (postDataCount == 5 && count > 1200)
                {
                    postDataCount = 6;
                }

                //Send new progress bar value to backgroundWorker1_ProgressChanged as fields cannot be updated in backgroundWorker thread
                double progressPercent = ((double)count / totalCount) * 100;
                progressPercent = Math.Round(progressPercent, 0);
                //Leave 2 percent of progress for time it takes to make api call
                if (progressPercent <= 98)
                {
                    backgroundWorker1.ReportProgress((int)progressPercent);
                }
            }

            if (isUsingCortical == true)
            {
                //Removes trailing ',' and replaces with ']' to close JSON object
                for (int i = 0; i <= postDataCount; i++)
                {
                    postData[i] = postData[i].Remove(postData[i].Length - 1, 1) + "]";
                    if (!String.IsNullOrEmpty(postData2[0]))
                    {
                        postData2[i] = postData2[i].Remove(postData2[i].Length - 1, 1) + "]";
                    }

                    //No Data found
                    if (postData[i] == "]")
                    {
                        MessageBox.Show("No data could be obtained using the search parameters!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        List<object> newArgs = new List<object>();
                        newArgs.Add("Results:");
                        newArgs.Add(true);
                        e.Result = newArgs;
                        return;
                    }
                }

                //System.IO.File.WriteAllText(@"C:\\Users\\brahamj\\Downloads\\jsonPost.txt", postData);
                //postData = System.IO.File.ReadAllText(@"C:\\Users\\brahamj\\Downloads\\jsonPost.txt");

                //API Request to cortical.io to compare text taken from SharePoint with a keyword the user provided
                List<KeyValuePair<double, int>> percentName = new List<KeyValuePair<double, int>>();
                List<KeyValuePair<double, string>> percentLink = new List<KeyValuePair<double, string>>();
                backgroundWorker1.ReportProgress(99);
                for (int k = 0; k <= postDataCount; k++)
                {
                    WebRequest webRequest = WebRequest.Create("http://api.cortical.io:80/rest/compare/bulk?retina_name=en_associative");
                    webRequest.Method = "POST";
                    webRequest.Headers["api-key"] = "bb355cc0-5873-11e8-9172-3ff24e827f76";
                    webRequest.ContentType = "application/json";
                    //Send request with postData string as the body
                    using (var streamWriter = new StreamWriter(webRequest.GetRequestStream()))
                    {
                        streamWriter.Write(postData[k]);
                        streamWriter.Flush();
                        streamWriter.Close();
                    }
                    string result = "";
                    string result2 = "";
                    //Recieve response from cortical.io API
                    try
                    {
                        WebResponse webResp = webRequest.GetResponse();
                        using (var streamReader = new StreamReader(webResp.GetResponseStream()))
                        {
                            result = streamReader.ReadToEnd();
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("\nCannot connect to cortical.io API. Aborting!\n\nError: " + ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        List<object> newArgs = new List<object>();
                        newArgs.Add("Results:");
                        newArgs.Add(true);
                        e.Result = newArgs;
                        return;
                    }

                    if (!String.IsNullOrEmpty(postData2[0]))
                    {
                        webRequest = WebRequest.Create("http://api.cortical.io:80/rest/compare/bulk?retina_name=en_associative");
                        webRequest.Method = "POST";
                        webRequest.Headers["api-key"] = "bb355cc0-5873-11e8-9172-3ff24e827f76";
                        webRequest.ContentType = "application/json";
                        //Send request with postData string as the body
                        using (var streamWriter = new StreamWriter(webRequest.GetRequestStream()))
                        {
                            streamWriter.Write(postData2[k]);
                            streamWriter.Flush();
                            streamWriter.Close();
                        }
                        //Recieve response from cortical.io API
                        try
                        {
                            WebResponse webResp = webRequest.GetResponse();
                            using (var streamReader = new StreamReader(webResp.GetResponseStream()))
                            {
                                result2 = streamReader.ReadToEnd();
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("\nCannot connect to cortical.io API. Aborting!\n\nError: " + ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            List<object> newArgs = new List<object>();
                            newArgs.Add("Results:");
                            newArgs.Add(true);
                            e.Result = newArgs;
                            return;
                        }
                    }

                    //Formats return string as JSON
                    dynamic jsonObj = JsonConvert.DeserializeObject<dynamic>(result);
                    dynamic jsonObj2 = null;
                    if (!String.IsNullOrEmpty(postData2[0]))
                    {
                        jsonObj2 = JsonConvert.DeserializeObject<dynamic>(result2);
                    }

                    //Calculates match percent for each return object which correlates to each resume
                    for (int i = 0; i < jsonObj.Count; i++)
                    {
                        double matchPercent = Math.Round((double)jsonObj[i].cosineSimilarity, 3);
                        double matchPercent2 = 0;
                        if (!String.IsNullOrEmpty(postData2[0]))
                        {
                            matchPercent2 = Math.Round((double)jsonObj2[i].cosineSimilarity, 3);

                            if (matchPercent2 <= 0.1)
                            {
                                matchPercent2 = 0;
                            }
                            else
                            {
                                matchPercent2 = Math.Round(((Math.Pow(Math.Log10(1 / matchPercent2), 3.55) * -1) + 1) * 100, 2);
                            }
                        }

                        if (matchPercent <= 0.1)
                        {
                            matchPercent = 0;
                        }
                        else
                        {
                            matchPercent = Math.Round(((Math.Pow(Math.Log10(1 / matchPercent), 3.55) * -1) + 1) * 100, 2);
                        }

                        //If multiple keywords, get weighted percent
                        if (!String.IsNullOrEmpty(postData2[0]))
                        {
                            double firstWeight = (double)(Int32.Parse(weight.Replace("%", ""))) / 100;
                            double secondWeight = 1 - firstWeight;

                            matchPercent = (((matchPercent / 100) * firstWeight) + ((matchPercent2 / 100) * secondWeight)) * 100;
                        }
                        percentLink.Add(new KeyValuePair<double, string>(matchPercent, links[i]));
                        percentName.Add(new KeyValuePair<double, int>(matchPercent, i));
                    }
                }

                //Order from greatest to least match percent
                percentName = percentName.OrderByDescending(x => x.Key).ToList();
                percentLink = percentLink.OrderByDescending(x => x.Key).ToList();

                List<string> keyList = new List<string>();
                //Generates response to populate gridView
                for (int i = 0; i < percentName.Count(); i++)
                {
                    namesOrdered.Add(names[percentName[i].Value]);
                    linksOrdered.Add(percentLink[i].Value);
                    keyList.Add(percentName[i].Key + "%");
                }

                backgroundWorker1.ReportProgress(100);
                watch.Stop();
                Console.WriteLine("Run Time: " + (double)watch.ElapsedMilliseconds/1000 + " seconds");

                //Sends finished data to e.Result so when backgroundWorker1 is completed it can access the data and correctly update the fields
                //This has to be done as you cannot update the fields inside backgroundWorker thread
                List<object> returnArgs = new List<object>();
                if (!String.IsNullOrEmpty(postData2[0]))
                {
                    returnArgs.Add("Results for \"" + (string)arguments[3] + "\" and \"" + (string)arguments[6] + "\":\n(You can double click any row to view the resume)");
                }
                else
                {
                    returnArgs.Add("Results for \"" + (string)arguments[3] + "\":\n(You can double click any row to view the resume)");
                }
                returnArgs.Add(false);
                returnArgs.Add(keyList);
                e.Result = returnArgs;
            }
            else
            {
                backgroundWorker1.ReportProgress(99);

                //Order from greatest to least match percent
                matchScoreName = matchScoreName.OrderByDescending(x => x.Key).ToList();
                matchScoreLink = matchScoreLink.OrderByDescending(x => x.Key).ToList();

                List<string> keyList = new List<string>();
                //Generates response to populate gridView
                for (int i = 0; i < matchScoreName.Count(); i++)
                {
                    namesOrdered.Add(names[matchScoreName[i].Value]);
                    linksOrdered.Add(matchScoreLink[i].Value);
                    keyList.Add(matchScoreName[i].Key + "%");
                }

                backgroundWorker1.ReportProgress(100);
                watch.Stop();
                Console.WriteLine("Run Time: " + (double)watch.ElapsedMilliseconds/1000 + " seconds");

                //Sends finished data to e.Result so when backgroundWorker1 is completed it can access the data and correctly update the fields
                //This has to be done as you cannot update the fields inside backgroundWorker thread
                List<object> returnArgs = new List<object>();
                if (!String.IsNullOrEmpty(postData2[0]))
                {
                    returnArgs.Add("Results for \"" + (string)arguments[3] + "\" and \"" + (string)arguments[6] + "\":\n(You can double click any row to view the resume)");
                }
                else
                {
                    returnArgs.Add("Results for \"" + (string)arguments[3] + "\":\n(You can double click any row to view the resume)");
                }
                returnArgs.Add(false);
                returnArgs.Add(keyList);
                e.Result = returnArgs;
            }
        }

        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Update progress bar
            progressBar1.Value = e.ProgressPercentage;
        }

        void backgroundWorker1_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            //Update fields
            List<object> arguments = e.Result as List<object>;
            //Argument[2] will only be true if system could not connect to cortical.io service which in that case no results are available
            if ((Boolean)arguments[1] == false)
            {
                List<string> keyList = (List<string>)arguments[2];

                lblResults.Text = (string)arguments[0];
                progressBar1.Visible = false;
                for (int i = 0; i < namesOrdered.Count(); i++)
                {
                    resultsView.Rows.Add(namesOrdered[i], keyList[i]);
                    resultsView.Rows[i].HeaderCell.Value = String.Format("{0}", resultsView.Rows[i].Index + 1);
                }
                resultsView.Focus();
            }
            else
            {
                progressBar1.Visible = true;
            }
            btnKeywordSubmit.Enabled = true;
            txtBoxKeyword.Enabled = true;
            btnLogout.Enabled = true;
            if (lblAddTextBox.Visible == true)
            {
                lblAddTextBox.Enabled = true;
            }
            else
            {
                lblMinusTextBox.Enabled = true;
                txtBoxSecondKeyword.Enabled = true;
                cmbWeight.Enabled = true;
            }
            txtBoxExperience.Enabled = true;

            progressBar1.Value = 0;
        }

        private void resultsView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (System.IO.Path.GetExtension(namesOrdered[e.RowIndex]) == ".pdf")
            {
                System.Diagnostics.Process.Start("https://aecon1.sharepoint.com/sites/bd/resume/Shared%20Documents/Original/" + namesOrdered[e.RowIndex]);
            }
            else
            {
                System.Diagnostics.Process.Start(linksOrdered[e.RowIndex]);
            }
        }

        private void Encrypt()
        {
            string text = System.IO.File.ReadAllText("creds.xml");
            byte[] key = getKey();

            SymmetricAlgorithm algorithm = DES.Create();
            ICryptoTransform transform = algorithm.CreateEncryptor(key, key);
            byte[] inputbuffer = Encoding.Unicode.GetBytes(text);
            byte[] outputBuffer = transform.TransformFinalBlock(inputbuffer, 0, inputbuffer.Length);

            System.IO.File.WriteAllText(@"creds.xml", Convert.ToBase64String(outputBuffer));
        }

        private void Decrypt()
        {
            string text = System.IO.File.ReadAllText("creds.xml");
            byte[] key = getKey();

            SymmetricAlgorithm algorithm = DES.Create();
            ICryptoTransform transform = algorithm.CreateDecryptor(key, key);
            byte[] inputbuffer = Convert.FromBase64String(text);
            byte[] outputBuffer = transform.TransformFinalBlock(inputbuffer, 0, inputbuffer.Length);

            System.IO.File.WriteAllText(@"creds.xml", Encoding.Unicode.GetString(outputBuffer));
        }

        private byte[] getKey()
        {
            try
            {
                var macAddr =
                    (
                        from nic in NetworkInterface.GetAllNetworkInterfaces()
                        where nic.OperationalStatus == OperationalStatus.Up
                        select nic.GetPhysicalAddress().ToString()
                    ).FirstOrDefault();
                macAddr = macAddr.Substring(0, 8);
                byte[] macByte = new UTF8Encoding(true).GetBytes(macAddr);
                return macByte;
            }
            catch
            {
                //Mac Address could not be found, use default key
                byte[] key = new byte[8] { 3, 8, 6, 1, 5, 7, 9, 2 };
                return key;
            }
        }

        private void lblAddTextBox_Click(object sender, EventArgs e)
        {
            lblAddTextBox.Visible = false;

            lblMinusTextBox = new Label();
            txtBoxSecondKeyword = new TextBox();

            lblMinusTextBox.Name = "lblMinusTextBox";
            lblMinusTextBox.Text = "-";
            lblMinusTextBox.Location = new System.Drawing.Point(9, 91);
            lblMinusTextBox.Size = new System.Drawing.Size(13, 13);
            lblMinusTextBox.Click += new EventHandler(this.lblMinusTextBox_Click);
            lblMinusTextBox.TabIndex = 3;
            ToolTip tt = new ToolTip();
            tt.SetToolTip(lblMinusTextBox, "Click to remove the other keyword field");
            this.Controls.Add(lblMinusTextBox);
            
            txtBoxSecondKeyword.Name = "txtBoxSecondKeyword";
            txtBoxSecondKeyword.Location = new System.Drawing.Point(30, 91);
            txtBoxSecondKeyword.Size = new System.Drawing.Size(180, 20);
            txtBoxSecondKeyword.TabIndex = 4;
            this.Controls.Add(txtBoxSecondKeyword);

            cmbWeight.Items.RemoveAt(9);

            cmbWeight.SelectedIndex = 4;
            cmbWeight.Enabled = true;
        }

        private void lblMinusTextBox_Click(object sender, EventArgs e)
        {
            lblAddTextBox.Visible = true;
            if (this.Controls.Contains(lblMinusTextBox) && this.Controls.Contains(txtBoxSecondKeyword))
            {
                this.Controls.Remove(lblMinusTextBox);
                this.Controls.Remove(txtBoxSecondKeyword);

                lblMinusTextBox.Dispose();
                txtBoxSecondKeyword.Dispose();

                cmbWeight.Items.Insert(9, "100%");
                cmbWeight.ResetText();
                cmbWeight.SelectedIndex = 9;
                cmbWeight.Enabled = false;
            }
            else
            {
                MessageBox.Show("Failed remove extra keyword field", "Error! Field not found!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        private void textBoxUsername_Enter(object sender, EventArgs e)
        {
            if (textBoxUsername.Text == "Example: jbraham@aecon.com")
            {
                textBoxUsername.Text = "";
                textBoxUsername.ForeColor = SystemColors.WindowText;
            }
        }

        private void textBoxUsername_Leave(object sender, EventArgs e)
        {
            if (textBoxUsername.Text == "")
            {
                textBoxUsername.Text = "Example: jbraham@aecon.com";
                textBoxUsername.ForeColor = SystemColors.ButtonShadow;
            }
        }
    }
}
