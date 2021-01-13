using System;
using System.Windows.Forms;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;

namespace EWSOAuthAppPermissions
{
    public partial class Form1 : Form
    {
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
                    var ewsClient = new ExchangeService();
                    ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                    ewsClient.Credentials = new OAuthCredentials(result.AccessToken);

                    //Impersonate the mailbox you'd like to access.
                    ewsClient.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, textBoxMailboxSMTPAddress.Text);

                    //Include x-anchormailbox header
                    ewsClient.HttpHeaders.Add("X-AnchorMailbox", textBoxMailboxSMTPAddress.Text);

                    // Make an EWS call
                    var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(10));
                    foreach (var folder in folders)
                    {
                        WriteToResults($"Folder: {folder.DisplayName}");
                    }
                }
                catch (MsalException ex)
                {
                    WriteToResults($"Error acquiring access token: {ex.ToString()}");
                }
                catch (Exception ex)
                {
                    WriteToResults($"Error: {ex.ToString()}");
                }


            });
            System.Threading.Tasks.Task.Run(action);
        }

        private void buttonClose_Click(object sender, EventArgs e)
        {
            Close();
        }


    }
}
