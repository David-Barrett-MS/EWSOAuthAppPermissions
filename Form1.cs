using System;
using System.Xml;
using System.Net.Http;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;
using Microsoft.Identity.Client;

namespace EWSOAuthAppPermissions
{
    public partial class Form1 : Form
    {
        HttpClient _httpClient = new HttpClient();

        public Form1()
        {
            InitializeComponent();
        }

        private void WriteToResults(string data)
        {
            // Add the given data to the results textbox

            Action action = new Action(() => {
                textBoxResults.AppendText($"{data}{Environment.NewLine}");
            });
            if (textBoxResults.InvokeRequired)
                textBoxResults.Invoke(action);
            else
                action();
        }

        private void buttonFindFolders_Click(object sender, EventArgs e)
        {
            Action action = new Action(async () =>
            {
                // Configure the MSAL client to get tokens
                var ewsScopes = new string[] { "https://outlook.office.com/.default" };

                var app = ConfidentialClientApplicationBuilder.Create(textBoxAppId.Text)
                    .WithAuthority(AzureCloudInstance.AzurePublic, textBoxTenantId.Text)
                    .WithClientSecret(textBoxClientSecret.Text)
                    .Build();

                AuthenticationResult result = null;

                try
                {
                    // Make the interactive token request
                    result = await app.AcquireTokenForClient(ewsScopes)
                        .ExecuteAsync();

                    // Configure the ExchangeService with the access token
                    var ewsClient = new ExchangeService(ExchangeVersion.Exchange2016);
                    ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                    ewsClient.Credentials = new OAuthCredentials(result.AccessToken);
                    ewsClient.TraceListener = new TraceListener("trace.log");
                    ewsClient.TraceFlags = TraceFlags.All;
                    ewsClient.TraceEnabled = true;

                    _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {result.AccessToken}");

                    //Impersonate the mailbox you'd like to access.
                    ewsClient.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, textBoxMailboxSMTPAddress.Text);

                    // Make an EWS call
                    if (checkBoxGetPublicFolders.Checked)
                    {
                        SetPublicFolderHeaders(ewsClient, textBoxMailboxSMTPAddress.Text);
                        var folders = ewsClient.FindFolders(WellKnownFolderName.PublicFoldersRoot, new FolderView(10));
                        foreach (var folder in folders)
                        {
                            WriteToResults($"Folder: {folder.DisplayName}");
                            ReadItemsFromFolder(folder);
                        }
                    }
                    else
                    {
                        ewsClient.HttpHeaders.Add("X-AnchorMailbox", textBoxMailboxSMTPAddress.Text);
                        var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
                        foreach (var folder in folders)
                        {
                            WriteToResults($"Folder: {folder.DisplayName}");
                        }
                    }
                }
                catch (MsalException ex)
                {
                    WriteToResults($"Error acquiring access token: {ex}");
                }
                catch (Exception ex)
                {
                    WriteToResults($"Error: {ex}");
                }


            });
            System.Threading.Tasks.Task.Run(action);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void SetPublicFolderHeaders(ExchangeService exchangeService, string mailbox)
        {
            // For public folders, we first need to get X-AnchorMailbox

            // Perform autodiscover for the mailbox and store the information
            AutodiscoverService autodiscover = new AutodiscoverService(new Uri("https://autodiscover-s.outlook.com/autodiscover/autodiscover.svc"), ExchangeVersion.Exchange2016);
            autodiscover.Credentials = exchangeService.Credentials;
            autodiscover.TraceListener = exchangeService.TraceListener;
            autodiscover.TraceFlags = TraceFlags.All;
            autodiscover.TraceEnabled = true;

            // Retrieve the autodiscover information
            GetUserSettingsResponse userSettings = null;
            try
            {
                userSettings = autodiscover.GetUserSettings(mailbox, UserSettingName.PublicFolderInformation);
            }
            catch (Exception ex)
            {
                WriteToResults($"Error: {ex}");
                return;
            }
            string xAnchor = userSettings.Settings[UserSettingName.PublicFolderInformation].ToString();
            WriteToResults($"Set X-AnchorMailbox to {xAnchor}");
            exchangeService.HttpHeaders.Add("X-AnchorMailbox", xAnchor);


            // Per the Exchange 2013 docs, we need to do a further AutoDiscover request to get X-PublicFolderMailbox
            // For Office 365, this will return an email address not found.
            // It doesn't seem to be required to set this header for Office 365, access to the public folder succeeds without it.

            /*
            string autodiscoverXml = System.IO.File.ReadAllText("Autodiscover.xml").Replace("<!--EMAILADDRESS-->", xAnchor);
            
            HttpResponseMessage response = _httpClient.PostAsync("https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml",
                new StringContent(autodiscoverXml, System.Text.Encoding.UTF8, "text/xml")).Result;
            if (!response.IsSuccessStatusCode)
            {
                WriteToResults("Failed to get X-PublicFolderMailbox");
                return;
            }

            string autodiscoverResponse = response.Content.ReadAsStringAsync().Result;
            XmlDocument autodiscoverXmlResponse = new XmlDocument();
            autodiscoverXmlResponse.LoadXml(autodiscoverResponse);

            WriteToResults(autodiscoverXmlResponse.ToString());
            //WriteToResults($"Set X-PublicFolderMailbox to {userSettings.Settings[UserSettingName.ExternalMailboxServer]}");
            */
        }

        private void ReadItemsFromFolder(Folder folder)
        {
            ItemView itemView = new ItemView(10, 0, OffsetBasePoint.Beginning);
            FindItemsResults<Item> results = folder.FindItems(itemView);
            foreach (var item in results.Items)
                WriteToResults($"Item: {item.Subject}");
        }
    }
}
