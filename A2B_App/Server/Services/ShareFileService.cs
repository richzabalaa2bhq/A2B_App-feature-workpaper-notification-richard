using A2B_App.Shared.Sox;
using Microsoft.Extensions.Configuration;
using ShareFile.Api.Client;
using ShareFile.Api.Client.Extensions;
using ShareFile.Api.Client.Logging;
using ShareFile.Api.Client.Models;
using ShareFile.Api.Client.Security.Authentication.OAuth2;
using ShareFile.Api.Client.Transfers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace A2B_App.Server.Services
{
    public class ShareFileService
    {
        private readonly IConfiguration _config;
        public ShareFileService(IConfiguration config)
        {
            _config = config;
        }

        public async Task<string> UploadWithUrlReturn(SharefileItem sharefileItem)
        {
            string url = string.Empty;
            try
            {
                Session session = null;
                ShareFileClient sfClient = null;
                SharefileUser user = new SharefileUser
                {
                    ControlPlane = _config.GetSection("SharefileApi").GetSection("ControlPane").Value,
                    Username = _config.GetSection("SharefileApi").GetSection("Username").Value,
                    Password = _config.GetSection("SharefileApi").GetSection("Password").Value,
                    Subdomain = _config.GetSection("SharefileApi").GetSection("SubDomain").Value
                };

                string oauthClientId = _config.GetSection("SharefileApi").GetSection("ClientId").Value;
                string oauthClientSecret = _config.GetSection("SharefileApi").GetSection("ClientSecret").Value;

                // Authenticate with username/password
                sfClient = await PasswordAuthentication(user, oauthClientId, oauthClientSecret);

                //Start session
                session = await sfClient.Sessions.Login().Expand("Principal").ExecuteAsync();
                //string filePath = sharefileItem.FilePath;
                var fileExtension = sharefileItem.FileName.Split('.').Last();
                var fileNameOnly = sharefileItem.FileName.Split('.').First();

                string sfDirectory = _config.GetSection("SharefileApi").GetSection(sharefileItem.Directory).GetSection("Path").Value;
                string sfLink = _config.GetSection("SharefileApi").GetSection(sharefileItem.Directory).GetSection("Link").Value;

                Folder folder = (Folder) await sfClient.Items.ByPath(sfDirectory).ExecuteAsync();

                var uploadedFileId = await UploadFile(sfClient, folder, sharefileItem.FilePath, fileExtension, null, fileNameOnly);

                var itemUri = sfClient.Items.GetAlias(uploadedFileId);
                var uploadedFile = await sfClient.Items.Get(itemUri).ExecuteAsync();
                Debug.WriteLine($"Successfully uploaded {uploadedFile}");
                url = sfLink;
            }
            catch (Exception ex)
            {
                FileLog.Write(ex.ToString(), "ErrorSharefileUploadWithUrlReturn");
                AdminService adminService = new AdminService(_config);
                adminService.SendAlert(true, true, ex.ToString(), "SharefileUploadWithUrlReturn");
            }


            return url;
        }


        private async Task<ShareFileClient> PasswordAuthentication(SharefileUser SFUser, string ClientID, string ClientSecret)
        {
            // Initialize ShareFileClient.
            var configuration = Configuration.Default();
            configuration.Logger = new DefaultLoggingProvider();

            var sfClient = new ShareFileClient("https://secure.sf-api.com/sf/v3/", configuration);
            var oauthService = new OAuthService(sfClient, ClientID, ClientSecret);

            // Perform a password grant request.  Will give us an OAuthToken
            var oauthToken = await oauthService.PasswordGrantAsync(SFUser.Username, SFUser.Password, SFUser.Subdomain, SFUser.ControlPlane);

            // Add credentials and update sfClient with new BaseUri
            sfClient.AddOAuthCredentials(oauthToken);
            sfClient.BaseUri = oauthToken.GetUri();
            Console.WriteLine("Sharefile authenticated");

            return sfClient;
        }


        private async Task<string> UploadFile(
            ShareFileClient sfClient,
            Folder destinationFolder,
            string FilePath,
            string FileExtension,
            string FileDetails,
            string RecordName)
        {
            string result = string.Empty;
            var file = System.IO.File.Open(FilePath, FileMode.OpenOrCreate);
            var uploadRequest = new UploadSpecificationRequest
            {
                FileName = RecordName + @"." + FileExtension,
                FileSize = file.Length,
                Details = FileDetails,
                Parent = destinationFolder.url
            };

            var uploader = sfClient.GetAsyncFileUploader(uploadRequest, file);
            decimal percent = 0;
            var uploadResponse = await uploader.UploadAsync();

            uploader.OnTransferProgress += (sender, e) =>
            {
                //Download progress
                percent = (((decimal)(e.Progress.BytesTransferred) / (decimal)(e.Progress.TotalBytes)) * 100);
                Debug.WriteLine("Uploading " + e.Progress.BytesTransferred + "/" + e.Progress.TotalBytes + " - " + percent.ToString("0.##") + "%");

                //Download completed
                if (e.Progress.Complete)
                {
                    Debug.WriteLine("Upload complete - " + file.Name);
                    file.Dispose();
                }
            };

            result = uploadResponse.First().Id;
            return result;

        }

    }
}
