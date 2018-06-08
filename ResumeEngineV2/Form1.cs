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

namespace ResumeEngineV2
{
    public partial class Form1 : System.Windows.Forms.Form
    {
        public Form1()
        {
            InitializeComponent();
            //Overlays progress bar ontop of gridview where results are displayed
            progressBar1.BringToFront();

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
                Decrypt();
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
                this.AcceptButton = btnKeywordSubmit;
                this.Text = "Resume Search Engine - Logged in as " + textBoxUsername.Text;
            }
            //Bad credentials, get user to try and login again
            catch
            {
                MessageBox.Show("Failed to authenticate username or password! Please try again.", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBoxUsername.Text = "";
                textBoxPassword.Text = "";
            }
        }

        private void btnLogout_Click(object sender, EventArgs e)
        {
            //Load xml and set credentials to ***
            Decrypt();
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
            textBoxUsername.Text = "";
            textBoxPassword.Text = "";
            lblResults.Text = "Results:";

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
        }

        private void btnKeywordSubmit_Click(object sender, EventArgs e)
        {
            //User must enter something for search to work
            if (txtBoxKeyword.Text == "")
            {
                MessageBox.Show("Please enter a valid keyword!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
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

                btnKeywordSubmit.Enabled = false;
                txtBoxKeyword.Enabled = false;
                btnLogout.Enabled = false;

                string targetSiteURL = @"https://aecon1.sharepoint.com/sites/bd/resume/";

                //Read credentials from creds.xml
                Decrypt();
                XmlDocument doc = new XmlDocument();
                doc.Load("creds.xml");
                Encrypt();

                var login = doc.DocumentElement.SelectSingleNode("/credentials/username").InnerText;
                var password = doc.DocumentElement.SelectSingleNode("/credentials/password").InnerText;
                string term = txtBoxKeyword.Text;
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
                        backgroundWorker1.WorkerReportsProgress = true;
                        backgroundWorker1.RunWorkerAsync(arguments);
                        break;
                    }
                }
            }
        }

        private void backgroundWoker1_DoWork(object sender, DoWorkEventArgs e)
        {
            List<object> arguments = e.Argument as List<object>;
            int totalCount = (int)arguments[5];
            int count = 0;
            List<string> names = (System.Collections.Generic.List<string>)arguments[4];
            Web web = (Web)arguments[1];
            string postData = "[";
            
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

            //Loops through each file
            foreach (var item in (FileCollection)arguments[0])
            {
                count++;
                string fileName = item.Name;
                names.Add(fileName.Replace(".txt",""));

                var filePath = web.ServerRelativeUrl + "/Shared%20Documents/Original/" + fileName;
                //var filePathTxt = web.ServerRelativeUrl + "/Shared%20Documents/Text/" + fileName + ".txt";
                FileInformation fileInformation = Microsoft.SharePoint.Client.File.OpenBinaryDirect((ClientContext)arguments[2], filePath);
                string ext = System.IO.Path.GetExtension(fileName);
                string convText = "";

                //Convert file into text
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
                string newConvText = new string(builder.ToArray());

                //System.IO.File.WriteAllText(@"C:\\Users\\brahamj\\Downloads\\newFormatTempText" + count + ".txt", newConvText);

                //Build JSON request string each loop
                postData += "[{\"term\": \"" + (string)arguments[3] + "\"},{\"text\": \"" + newConvText + "\"}],";

                //Send new progress bar value to backgroundWorker1_ProgressChanged as fields cannot be updated in backgroundWorker thread
                double progressPercent = ((double)count / totalCount) * 100;
                progressPercent = Math.Round(progressPercent, 0);
                backgroundWorker1.ReportProgress((int)progressPercent);
            }

            //Removes trailing ',' and replaces with ']' to close JSON object
            postData = postData.Remove(postData.Length - 1, 1) + "]";

            //System.IO.File.WriteAllText(@"C:\\Users\\brahamj\\Downloads\\jsonPost.txt", postData);
            //postData = System.IO.File.ReadAllText(@"C:\\Users\\brahamj\\Downloads\\jsonPost.txt");

            //API Request to cortical.io to compare text taken from SharePoint with a keyword the user provided
            WebRequest webRequest = WebRequest.Create("http://api.cortical.io:80/rest/compare/bulk?retina_name=en_associative");
            webRequest.Method = "POST";
            webRequest.Headers["api-key"] = "bb355cc0-5873-11e8-9172-3ff24e827f76";
            webRequest.ContentType = "application/json";
            //Send request with postData string as the body
            using (var streamWriter = new StreamWriter(webRequest.GetRequestStream()))
            {
                streamWriter.Write(postData);
                streamWriter.Flush();
                streamWriter.Close();
            }
            string result = "";
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
                MessageBox.Show("\nCannot connect to cortical.io API. Aborting!\n\nError: " + ex.Message);
                List<object> newArgs = new List<object>();
                newArgs.Add("Results:");
                newArgs.Add(true);
                e.Result = newArgs;
                return;
            }

            //Formats return string as JSON
            dynamic jsonObj = JsonConvert.DeserializeObject<dynamic>(result);

            //Calculates match percent for each return object which correlates to each resume
            List<KeyValuePair<double, int>> percentName = new List<KeyValuePair<double, int>>();
            for (int i = 0; i < jsonObj.Count; i++)
            {
                double matchPercent = Math.Round((double)jsonObj[i].cosineSimilarity, 3);
                if (matchPercent <= 0.1)
                {
                    matchPercent = 0;
                }
                else
                {
                    matchPercent = Math.Round(((Math.Pow(Math.Log10(1 / matchPercent), 3.55) * -1) + 1) * 100, 2);
                }

                percentName.Add(new KeyValuePair<double, int>(matchPercent, i));

                percentName = percentName.OrderByDescending(x => x.Key).ToList();
            }

            List<string> namesList = new List<string>();
            List<string> keyList = new List<string>();
            //Generates response to populate gridView
            for (int i = 0; i < percentName.Count(); i++)
            {
                namesList.Add(names[percentName[i].Value]);
                keyList.Add(percentName[i].Key + "%");
            }
            
            //Sends finished data to e.Result so when backgroundWorker1 is completed it can access the data and correctly update the fields
            //This has to be done as you cannot update the fields inside backgroundWorker thread
            List<object> returnArgs = new List<object>();
            returnArgs.Add("Results for \"" + (string)arguments[3] + "\":");
            returnArgs.Add(false);
            returnArgs.Add(namesList);
            returnArgs.Add(keyList);
            e.Result = returnArgs;
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
                List<string> namesList = (List<string>)arguments[2];
                List<string> keyList = (List<string>)arguments[3];

                lblResults.Text = (string)arguments[0];
                progressBar1.Visible = false;
                for (int i = 0; i < namesList.Count(); i++)
                {
                    resultsView.Rows.Add(namesList[i], keyList[i]);
                    resultsView.Rows[i].HeaderCell.Value = String.Format("{0}", resultsView.Rows[i].Index + 1);
                }
            }
            else
            {
                progressBar1.Visible = true;
            }
            btnKeywordSubmit.Enabled = true;
            txtBoxKeyword.Enabled = true;
            btnLogout.Enabled = true;
            progressBar1.Value = 0;
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
    }
}
