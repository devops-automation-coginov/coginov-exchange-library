using Coginov.Exchange.Library.Enums;
using Coginov.Exchange.Library.Helpers;
using Coginov.Exchange.Library.Models;
using MailBee;
using MailBee.EwsMail;
using MailBee.Mime;
using MailBee.Outlook;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using MailBeeAttachment = MailBee.Mime.Attachment;

namespace Coginov.Exchange.Library.Services
{
    public class ExchangeService : IExchangeService
    {
        private readonly string mailBeeLicenseKey = "MN120-824AB5884BF84AC94ABFADC0512C-AB72";
        private readonly ILogger logger;

        private Inbox inbox;
        private Ews ewsClient;
        private ExchangeVersion exchangeVersion;

        private bool isEwsClientInitialized = false;
        private bool isServiceConnectedToInbox = false;
        private int tagFlagRetryDelay = 10; // 10 Milliseconds delay by default
        private int tagFlagRetryCount = 5; // 5 retry by default

        public ExchangeService(ILogger logger)
        {
            // Adding support for Windows-1252 encoding. Nuget Package "System.Text.Encoding.CodePages"
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            this.logger = logger;

            // Initializing EWS Client
            TryInitEws();
        }

        public async Task<bool> InitializeInbox(Inbox inbox,
                                                int tagFlagRetryDelay = 10,
                                                int tagFlagRetryCount = 5,
                                                bool forceReconnect = false,
                                                string accountToImpersonate = null)
        {
            this.inbox = inbox;
            this.tagFlagRetryDelay = tagFlagRetryDelay;
            this.tagFlagRetryCount = tagFlagRetryCount;

            if (!await TryToConnectToInbox(forceReconnect, accountToImpersonate))
            {
                return false;
            }

            return true;
        }

        public async Task<EwsItemList> GetEmailsFromFolder(string folder,
                                                           int startIndex = 0,
                                                           int emailCount = 5,
                                                           List<string> errors = null,
                                                           bool unreadOnly = false,
                                                           EwsItemParts itemParts = EwsItemParts.MailMessageFull)
        {
            if (errors == null)
            {
                errors = new List<string>();
            }

            if (!IsClientInitializedAndConnectedToInbox())
            {
                errors.Add($"{Resource.EwsClientNotInitialized} {Resource.Or} {Resource.UnableToConnectToExchange}");
                return null;
            }

            EwsItemList ewsItemList;

            try
            {
                // Get FolderId from folder
                var folderId = await ewsClient.FindFolderIdByFullNameAsync(folder);
                if (folderId == null)
                {
                    var errorMsg = $"{Resource.ErrorReadingEmail}. {Resource.InvalidFolder}: {folder}";
                    errors.Add(errorMsg);
                    logger.LogError(errorMsg);
                    return null;
                }

                // Get list of emails in given folder and the attachments
                ewsItemList = await ewsClient.DownloadItemsAsync(folderId, startIndex, emailCount, unreadOnly, itemParts);
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorReadingEmail}: {folder}. {ex.Message}{innerException}";
                errors.Add(errorMsg);
                logger.LogError(errorMsg);
                return null;
            }

            return ewsItemList;
        }

        public async Task<bool> MoveEmailToFolder(string messageId,
                                                  string srcFolder,
                                                  string destFolder)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            try
            {
                var ewsItem = await FindEmailInFolder(srcFolder, messageId);
                if (ewsItem != null)
                {
                    var folderId = await ewsClient.FindFolderIdByFullNameAsync(destFolder);
                    if (folderId == null)
                    {
                        logger.LogError($"{Resource.ErrorMovingEmail}. {Resource.InvalidFolder}: {destFolder}");
                        return false;
                    }

                    var item = await ewsItem.NativeItem.Move(folderId);
                    return item != null;
                }
                else
                {
                    logger.LogError($"{Resource.ErrorMovingEmail}. {Resource.EmailToMoveNotFound} {Resource.Or} {Resource.InvalidFolder}: {srcFolder}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorMovingEmail}: {srcFolder} -> {destFolder}. {ex.Message}{innerException}";
                logger.LogError(errorMsg);
                return false;
            }
        }

        public async Task<bool> MoveEmailsToFolder(EwsItemList ewsItemList,
                                                   string srcFolder,
                                                   string destFolder)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(srcFolder) || string.IsNullOrWhiteSpace(destFolder))
            {
                logger.LogError($"{Resource.ErrorMovingEmail}. {Resource.EmailsFolderNotDefined}");
                return false;
            }

            var result = true;
            foreach (var mail in ewsItemList)
            {
                if (await MoveEmailToFolder(mail.MailBeeMessage.MessageID, srcFolder, destFolder))
                {
                    logger.LogInformation($"{Resource.EmailMovedToFolder}: {destFolder}. {Resource.Id}: {mail.MailBeeMessage.MessageID}");
                }
                else
                {
                    logger.LogError($"{Resource.ErrorMovingEmail}: {srcFolder} -> {destFolder}");
                    result = false;
                }
            }

            return result;
        }

        public async Task<bool> CopyEmailToFolder(string messageId,
                                                  string srcFolder,
                                                  string destFolder)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            try
            {
                var ewsItem = await FindEmailInFolder(srcFolder, messageId);
                if (ewsItem != null)
                {
                    var folderId = await ewsClient.FindFolderIdByFullNameAsync(destFolder);
                    if (folderId == null)
                    {
                        logger.LogError($"{Resource.ErrorCopyingEmail}. {Resource.InvalidFolder}: {destFolder}");
                        return false;
                    }

                    await ewsItem.NativeItem.Copy(folderId);
                }
                else
                {
                    logger.LogError($"{Resource.ErrorCopyingEmail}. {Resource.EmailToCopyNotFound} {Resource.Or} {Resource.InvalidFolder}: {srcFolder}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorCopyingEmail}: {srcFolder} -> {destFolder}. {ex.Message}{innerException}";
                logger.LogError(errorMsg);
                return false;
            }

            return true;
        }

        public async Task<bool> UploadEmailToFolder(MailMessage message,
                                                    string folder)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            try
            {
                var folderId = await ewsClient.FindFolderIdByFullNameAsync(folder);
                if (folderId == null)
                {
                    logger.LogError($"{Resource.ErrorUploadingEmail}. {Resource.InvalidFolder}: {folder}");
                    return false;
                }

                return await ewsClient.UploadMessageAsync(folderId, message, false);
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorUploadingEmail}: {folder}. {ex.Message}{innerException}";
                logger.LogError(errorMsg);
                return false;
            }
        }

        public async Task<bool> SendEmail(MailMessage email,
                                          List<string> to,
                                          bool saveCopySentFolder = false,
                                          bool includeCc = false)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            bool success = false;
            try
            {
                // Set From to user / email that owns inbox
                email.From = new MailBee.Mime.EmailAddress(inbox.User);

                // WARNING: Please make sure to use correctly parameter includeCc to avoid sending
                // unintentional emails to addresses on Cc and Bcc.
                // By default we are setting this parameter to false to avoid mistakes and to clear
                // Cc and Bcc. You can use it as soon as you are sure that the email should also be
                // sent to addresses on fields Cc and Bcc
                if (!includeCc)
                {
                    email.Cc = new MailBee.Mime.EmailAddressCollection();
                    email.Bcc = new MailBee.Mime.EmailAddressCollection();
                }

                // Set To to values received as parameter
                email.To = MailBee.Mime.EmailAddressCollection.Parse(string.Join(';', to));

                // Update DateSent to Now
                email.DateSent = DateTime.Now;

                // Try to send email
                success = saveCopySentFolder
                    ? await ewsClient.SendMessageAndSaveCopyAsync(email, null)
                    : await ewsClient.SendMessageAsync(email);

                if (success)
                {
                    logger.LogInformation($"{Resource.SuccessSendingEmail}: {email.To} | {email.Subject}");
                }
                else
                {
                    logger.LogError($"{Resource.ErrorSendingEmail}: {email.To} | {email.Subject}");
                }
            }
            catch (Exception ex)
            {
                success = false;

                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorSendingEmail}: {email.To} | {email.Subject}. {Resource.ErrorDetails} - {ex.Message}{innerException}";
                logger.LogError(errorMsg);
            }

            return success;
        }

        public async Task<bool> TagEmail(string messageId,
                                         string folder,
                                         List<string> categoryList)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            try
            {
                // Always Tag emails on clientInbox
                var retryCount = 0;
                EwsItem ewsItem = null;

                while (ewsItem == null && retryCount < tagFlagRetryCount)
                {
                    ewsItem = await FindEmailInFolder(folder, messageId);
                    retryCount++;

                    if (ewsItem == null)
                    {
                        Thread.Sleep(tagFlagRetryDelay);
                    }
                }

                if (ewsItem != null)
                {
                    ewsItem.NativeItem.Categories.AddRange(categoryList);
                    ewsItem.IsRead = false;
                    return await ewsClient.UpdateItemAsync(ewsItem);
                }
                else
                {
                    logger.LogWarning(Resource.EmailToTagNotFound);
                    return false;
                }
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorTaggingEmail}: {messageId}. {ex.Message}{innerException}";
                logger.LogError(errorMsg);
                return false;
            }
        }

        public async Task<bool> FlagEmail(string messageId,
                                          string folder,
                                          ItemFlagStatus status = ItemFlagStatus.Flagged,
                                          DateTime? startDate = null,
                                          DateTime? dueDate = null)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            // Exchange 2013 is the minimun version required for flagging
            if (exchangeVersion < ExchangeVersion.Exchange2013)
            {
                logger.LogWarning(Resource.ErrorFlaggingExchange2013Required);
                return false;
            }

            try
            {
                // Always Flag emails on clientInbox
                var retryCount = 0;
                EwsItem ewsItem = null;

                while (ewsItem == null && retryCount < tagFlagRetryCount)
                {
                    ewsItem = await FindEmailInFolder(folder, messageId);
                    retryCount++;

                    if (ewsItem == null)
                    {
                        Thread.Sleep(tagFlagRetryDelay);
                    }
                }

                if (ewsItem != null)
                {
                    ewsItem.NativeItem.Flag = new Flag
                    {
                        FlagStatus = status,
                        CompleteDate = status == ItemFlagStatus.Complete ? DateTime.UtcNow : new DateTime(),
                        StartDate = startDate == null ? new DateTime() : startDate.Value,
                        DueDate = dueDate == null ? new DateTime() : dueDate.Value
                    };

                    return await ewsClient.UpdateItemAsync(ewsItem);
                }
                else
                {
                    logger.LogWarning(Resource.EmailToFlagNotFound);
                    return false;
                }
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorFlaggingEmail}: {messageId}. {ex.Message}{innerException}";
                logger.LogError(errorMsg);
                return false;
            }
        }

        public async Task<bool> DeleteEmail(string messageId)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            return await ewsClient.DeleteItemAsync(messageId);
        }

        public async Task<bool> SendTemplate(string templatePath,
                                             List<string> to)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            try
            {
                var msgReader = new MsgConvert();
                var template = msgReader.MsgToMailMessage(templatePath);

                return await SendEmail(template, to);
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorSendingTemplate}: {ex.Message}{innerException}";
                logger.LogError(errorMsg);
                return false;
            }
        }

        public async Task<bool> ForwardEmailPrefixedByTemplate(MailMessage email,
                                                               string templatePath,
                                                               List<string> to)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            try
            {
                var msgReader = new MsgConvert();
                var template = msgReader.MsgToMailMessage(templatePath);

                email.DateSent = DateTime.Now;
                var fullTo = email.To[0].DisplayName.Equals(email.To[0].Email, StringComparison.InvariantCultureIgnoreCase) ? email.To[0].Email : $"{email.To[0].DisplayName} {email.To[0].Email}";
                var forwardHeader = $"<hr><p><strong>From:</strong> {email.From.DisplayName} {email.From.Email} </p><p><strong>Sent:</strong> {email.DateSent}</p><p><strong>To:</strong> {fullTo}</p><p><strong>Subject:</strong> {email.Subject}</p><br/><br/>";

                var body = string.IsNullOrWhiteSpace(email.BodyHtmlText)
                    ? email.BodyPlainText
                    : email.BodyHtmlText;

                email.BodyHtmlText = HtmlHelper.AppendHtmlToBody(template.BodyHtmlText, forwardHeader, body);
                email.MakePlainBodyFromHtmlBody();
                for (int i = 0; i < template.Attachments.Count; i++)
                {
                    email.Attachments.Add(template.Attachments[i]);
                }

                return await SendEmail(email, to);
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorSendingTemplate}: {ex.Message}{innerException}";
                logger.LogError(errorMsg);
                return false;
            }
        }

        public async Task<bool> ForwardEmailUsingTemplate(string templatePath,
                                                          MailMessage email,
                                                          List<string> to,
                                                          Dictionary<string, string> parameters = null)
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                return false;
            }

            try
            {
                var msgReader = new MsgConvert();
                MailMessage template = msgReader.MsgToMailMessage(templatePath);
                if (parameters != null)
                {
                    template.BodyHtmlText = HtmlHelper.HtmlBodyReplaceParams(template.BodyHtmlText, parameters);
                }

                template.Attachments.Add(email, $"{email.Subject}.eml", "", null, null, NewAttachmentOptions.None, MailTransferEncoding.None);
                return await SendEmail(template, to);
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorSendingTemplate}: {ex.Message}{innerException}";
                logger.LogError(errorMsg);
                return false;
            }
        }

        public async Task<bool> SaveEmailAsMsg(MailMessage email,
                                               string destinationFolder,
                                               string msgFileName,
                                               bool saveAttachmentsAsFiles,
                                               List<string> errors)
        {
            if (errors == null)
            {
                errors = new List<string>();
            }

            if (!IsClientInitializedAndConnectedToInbox())
            {
                errors.Add($"{Resource.EwsClientNotInitialized} {Resource.Or} {Resource.UnableToConnectToExchange}");
                return false;
            }

            string errorMsg;
            var error = false;
            try
            {
                // Try saving the message as Msg file
                var fileName = Path.Combine(destinationFolder, msgFileName);
                if (File.Exists(fileName))
                {
                    logger.LogWarning($"{Resource.MessageFileAlreadyExists}: {fileName}");
                }

                var msgReader = new MsgConvert();
                msgReader.MailMessageToMsg(email, Path.Combine(destinationFolder, msgFileName));
            }
            catch (NotSupportedException ex)
            {
                //Support for UTF-7 is disabled. Default to saving as .txt file
                using StreamWriter textFile = new(Path.Combine(destinationFolder, msgFileName));
                await textFile.WriteLineAsync($"From: {email.From}");
                await textFile.WriteLineAsync($"To: {email.To}");
                await textFile.WriteLineAsync($"Subject: {email.Subject}");
                await textFile.WriteLineAsync($"Body: {email.BodyPlainText ?? email.BodyHtmlText}");
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                errorMsg = $"{Resource.ErrorSavingEmailAsMsg}. {Resource.MessageId}: {email.MessageID}. {Resource.DestinationFolder}: {destinationFolder}. {Resource.MessageFileName}: {msgFileName}. {Resource.ErrorMessage}: {ex.Message}{innerException}";
                errors.Add(errorMsg);
                logger.LogError(errorMsg);

                // If an exception happens while trying to save the Email as Msg
                // we don't try to save attachments and return false
                return false;
            }

            if (saveAttachmentsAsFiles)
            {
                foreach (MailBeeAttachment attachment in email.Attachments)
                {
                    try
                    {
                        // Try saving each message attachement as a file
                        if (!await attachment.SaveToFolderAsync(destinationFolder, true))
                        {
                            error = true;
                            errorMsg = $"{Resource.ErrorSavingEmailAttachmentToFileSystem}, {Resource.MessageId}: {email.MessageID}. {Resource.DestinationFolder}: {destinationFolder}. {Resource.AttachmentFileName}: {attachment.Filename}";
                            errors.Add(errorMsg);
                            logger.LogError(errorMsg);
                        }
                    }
                    catch (Exception ex)
                    {
                        error = true;
                        var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                        errorMsg = $"{Resource.ErrorSavingEmailAttachmentToFileSystem}. {Resource.MessageId}: {email.MessageID}. {Resource.DestinationFolder}: {destinationFolder}. {Resource.AttachmentFileName}: {attachment.Filename}. {Resource.ErrorMessage}{innerException}: {ex.Message}";
                        errors.Add(errorMsg);
                        logger.LogError(errorMsg);
                    }
                }
            }

            return !error;
        }

        public async Task<List<EwsFolder>> GetFolders(bool includeSubfolders = true)
        {
            try
            {
                return await ewsClient.DownloadFoldersAsync(includeSubfolders);
            }
            catch (MailBeeEwsException ex)
            {
                var errorMsg = $"{Resource.ErrorReadingEmail}: {ex.Message}";
                logger.LogError(errorMsg);
                if (ex.InnerException is ServerBusyException busyException)
                {
                    logger.LogError(string.Format(Resource.ExchangeServerBusy, (int)busyException.BackOffMilliseconds / 1000));
                    Thread.Sleep(busyException.BackOffMilliseconds);
                }
                throw;
            }
        }

        public async Task<EwsFolder> GetFolder(string uniqueId)
        {
            try
            {
                return await ewsClient.DownloadFolderByIdAsync(new FolderId(uniqueId));
            }
            catch (MailBeeEwsException ex)
            {
                var errorMsg = $"{Resource.ErrorReadingEmail}: {ex.Message}";
                logger.LogError(errorMsg);
                if (ex.InnerException is ServerBusyException busyException)
                {
                    logger.LogError(string.Format(Resource.ExchangeServerBusy, (int)busyException.BackOffMilliseconds / 1000));
                    Thread.Sleep(busyException.BackOffMilliseconds);
                }
                throw;
            }
        }

        public async Task<EwsFolder> GetAllItemsFolder()
        {
            if (!IsClientInitializedAndConnectedToInbox())
            {
                logger.LogError($"{Resource.EwsClientNotInitialized} {Resource.Or} {Resource.UnableToConnectToExchange}");
                return null;
            }

            ExtendedPropertyDefinition allFoldersType = new ExtendedPropertyDefinition(13825, MapiPropertyType.Integer);
            SearchFilter searchFilter1 = new SearchFilter.IsEqualTo(allFoldersType, "2");
            SearchFilter searchFilter2 = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "allitems");
            var searchFilterCollection = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilter1, searchFilter2);

            var allFolders = await ewsClient.DownloadFoldersAsync(new FolderId(WellKnownFolderName.Root), new FolderView(10), searchFilterCollection, false);

            return allFolders.Any() ? allFolders[0] : null;
        }

        // References:
        // https://newbedev.com/exchange-web-services-ews-finditems-within-all-folders
        // https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-perform-grouped-searches-by-using-ews-in-exchange
        // We recommend adding a millisecond to the afterDate parameter due to current issues in EWS Api SearchFilter.IsGreaterThan. See link below
        // https://github.com/OfficeDev/ews-managed-api/issues/139
        public async Task<List<EwsItem>> GetEmailsFromFolderAfterDate(EwsFolder folder,
                                                                    DateTime afterDate,
                                                                    int startIndex = 0,
                                                                    int emailCount = 10,
                                                                    List<string> errors = null,
                                                                    bool unreadOnly = false,
                                                                    EwsItemParts itemParts = EwsItemParts.MailMessageFull)
        {
            if (errors == null)
            {
                errors = new List<string>();
            }

            if (!IsClientInitializedAndConnectedToInbox())
            {
                errors.Add($"{Resource.EwsClientNotInitialized} {Resource.Or} {Resource.UnableToConnectToExchange}");
                return null;
            }

            try
            {
                var view = new ItemView(emailCount, startIndex);
                view.OrderBy.Add(ItemSchema.DateTimeCreated, SortDirection.Ascending);

                var filter = new SearchFilter.SearchFilterCollection(LogicalOperator.And);
                filter.Add(new SearchFilter.IsGreaterThan(ItemSchema.DateTimeCreated, afterDate));
                filter.Add(new SearchFilter.IsEqualTo(ItemSchema.ItemClass, "IPM.Note"));
                var searchItemList = await ewsClient.SearchAsync(folder.Id, filter, view);

                try
                {
                    var ewsItemList = await ewsClient.DownloadItemsAsync(searchItemList.ToList(), itemParts);
                    return ewsItemList;
                }
                catch(System.Xml.XmlException)
                {
                    // Server may be busy and return a malformed xml
                    // Fall back to download all items one by one
                    var itemList = new List<EwsItem>();
                    foreach(var item in searchItemList)
                    {
                        try
                        {
                            itemList.Add(await ewsClient.DownloadItemAsync(item.Id, itemParts));
                        }
                        catch (MailBeeInvalidArgumentException)
                        {
                            continue;
                        }
                    }
                    return itemList;
                }
            }
            catch (MailBeeEwsException ex)
            {
                var errorMsg = $"{Resource.ErrorReadingEmail}: {folder.FullName}. {ex.Message}";
                errors.Add(errorMsg);
                logger.LogError(errorMsg);
                if (ex.InnerException is ServerBusyException busyException)
                {
                    logger.LogError(string.Format(Resource.ExchangeServerBusy, (int)busyException.BackOffMilliseconds/1000));
                    Thread.Sleep(busyException.BackOffMilliseconds);
                }

                throw;
            }
            catch (Exception ex)
            {
                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorReadingEmail}: {folder.FullName}. {ex.Message}{innerException}";
                errors.Add(errorMsg);
                logger.LogError(errorMsg);

                // Throw same exception so the client implements other recovery mechanisms
                throw;
            }
        }

        #region Private Methods

        private bool TryInitEws()
        {
            if (isEwsClientInitialized)
            {
                return true;
            }

            Global.LicenseKey = mailBeeLicenseKey;
            ewsClient = new Ews();
            isEwsClientInitialized = false;

            var minVersionExchange = Enum.GetValues(typeof(ExchangeVersion)).Cast<int>().Min();
            var maxVersionExchange = Enum.GetValues(typeof(ExchangeVersion)).Cast<int>().Max();
            for (var version = maxVersionExchange; version >= minVersionExchange; version--)
            {
                try
                {
                    ewsClient.InitEwsClient((ExchangeVersion)version, TimeZoneInfo.Utc);
                    ewsClient.Service.DateTimePrecision = DateTimePrecision.Milliseconds;
                    exchangeVersion = (ExchangeVersion)version;
                    isEwsClientInitialized = true;
                    break;
                }
                catch (Exception ex)
                {
                    // Fallback for older versions of Exchange.
                    // Continue to next iteration for lower version
                    var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                    var errorMsg = $"{Resource.ErrorInitEwsClientForVersion}: {version}. {ex.Message}{innerException}";
                    logger.LogError(errorMsg);
                    isEwsClientInitialized = false;
                }
            }

            ewsClient.Service.AcceptGzipEncoding = false;
            return isEwsClientInitialized;
        }

        private async Task<bool> TryToConnectToInbox(bool forceReconnect, string accountToImpersonate)
        {
            try
            {
                if (!isServiceConnectedToInbox || forceReconnect)
                {
                    switch (inbox.AuthenticationMethod)
                    {
                        case AuthenticationMethod.Basic:
                            isServiceConnectedToInbox = await ConnectBasic(inbox, accountToImpersonate);
                            break;
                        case AuthenticationMethod.OAuthAppPermissions:
                            isServiceConnectedToInbox = await ConnectOAuthAppPermissions(inbox, accountToImpersonate);
                            break;
                        case AuthenticationMethod.OAuthDelegatedPermissions:
                            isServiceConnectedToInbox = await ConnectOAuthDelegatedPermissions(inbox);
                            break;
                        default:
                            break;
                    }

                    if (!isServiceConnectedToInbox)
                    {
                        logger.LogError($"{Resource.UnableToConnectToExchange} - {Resource.InMailInbox} - {inbox.ServerUrl}");
                    }
                }
            }
            catch (Exception ex)
            {
                isServiceConnectedToInbox = false;

                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorConnectingToExchange} - {Resource.InMailInbox} - {inbox.ServerUrl}: {ex.Message}{innerException}";
                logger.LogError(errorMsg);
            }

            return isServiceConnectedToInbox;
        }

        private bool IsClientInitializedAndConnectedToInbox()
        {
            if (!isEwsClientInitialized)
            {
                logger.LogError(Resource.EwsClientNotInitialized);
                return false;
            }

            if (!isServiceConnectedToInbox)
            {
                var url = inbox != null ? $"- {inbox.ServerUrl}" : string.Empty;
                logger.LogError($"{Resource.UnableToConnectToExchange}: {Resource.InMailInbox} {url}");
                return false;
            }

            return true;
        }

        private async Task<EwsItem> FindEmailInFolder(string folder,
                                                      string messageId)
        {
            var folderId = await ewsClient.FindFolderIdByFullNameAsync(folder);
            if (folderId == null)
            {
                logger.LogError($"{Resource.ErrorSearchingForEmail}. {Resource.InvalidFolder}: {folder}");
                return null;
            }

            var filter = new SearchFilter.IsEqualTo(EmailMessageSchema.InternetMessageId, messageId);
            return (await ewsClient.SearchAsync(folderId, filter)).FirstOrDefault();
        }

        private async Task<bool> ConnectBasic(Inbox inbox, string accountToImpersonate)
        {
            // Basic Authentication = User / Password
            // Less secure authentication mechanism, avoid it at all cost
            bool success = false;
            try
            {
                ewsClient.SetCredentials(inbox.User, inbox.Password);
                if (!string.IsNullOrWhiteSpace(accountToImpersonate))
                {
                    ewsClient.Service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, accountToImpersonate);
                }

                if (!string.IsNullOrWhiteSpace(inbox.ServerUrl))
                {
                    ewsClient.SetServerUrl(inbox.ServerUrl);
                    success = await ewsClient.TestConnectionAsync();
                }
                else if (ewsClient.Autodiscover(inbox.ReplyFrom))
                {
                    success = await ewsClient.TestConnectionAsync();
                }

                if (success)
                {
                    logger.LogInformation($"{Resource.SuccessConnectingToExchange}: {inbox.ServerUrl}");
                }
                else
                {
                    logger.LogError($"{Resource.ErrorConnectingToExchange}: {inbox.ServerUrl}");
                }
            }
            catch (Exception ex)
            {
                success = false;

                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorConnectingToExchange}: {ex.Message}{innerException}";
                logger.LogError(errorMsg);
            }

            return success;
        }

        private async Task<bool> ConnectOAuthAppPermissions(Inbox inbox, string accountToImpersonate)
        {
            // Application Permissions = Excess of priviledges
            // Some organizations forbid the use of this mechanism for security reasons
            // Non-Interactive scenarios, user presence not required. Used for services or deamon applications
            // https://www.enowsoftware.com/solutions-engine/accessing-exchange-online-objects-without-legacy-auth
            // https://mailbeenet.wordpress.com/2021/04/26/using-mailbee-net-ews-to-access-office-365-mailbox-in-non-interactive-case/
            bool success = false;
            try
            {
                // Using Microsoft.Identity.Client 4.22.0 (MSAL)
                IConfidentialClientApplication app = ConfidentialClientApplicationBuilder
                    .Create(inbox.ClientId)
                    .WithClientSecret(inbox.ClientSecret)
                    .WithTenantId(inbox.TenantId)
                    .Build();

                // The permission scope required for EWS access
                var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

                // Make the token request
                AuthenticationResult authResult = null;
                authResult = await app.AcquireTokenForClient(ewsScopes).ExecuteAsync();
                OAuthCredentials authCreds = new OAuthCredentials(authResult.AccessToken);

                ewsClient.SetCredentials(authCreds);
                if (!string.IsNullOrWhiteSpace(accountToImpersonate))
                {
                    ewsClient.Service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, accountToImpersonate);
                }
                else
                {
                    ewsClient.Service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, inbox.User);
                }

                if (!string.IsNullOrWhiteSpace(inbox.ServerUrl))
                {
                    ewsClient.SetServerUrl(inbox.ServerUrl);
                    success = await ewsClient.TestConnectionAsync();
                }
                else if (ewsClient.Autodiscover(inbox.ReplyFrom))
                {
                    success = await ewsClient.TestConnectionAsync();
                }

                if (success)
                {
                    logger.LogInformation($"{Resource.SuccessConnectingToExchange}: {inbox.ServerUrl}");
                }
                else
                {
                    logger.LogError($"{Resource.ErrorConnectingToExchange}: {inbox.ServerUrl}");
                }
            }
            catch (Exception ex)
            {
                success = false;

                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorConnectingToExchange}: {ex.Message}{innerException}";
                logger.LogError(errorMsg);
            }

            return success;
        }

        private async Task<bool> ConnectOAuthDelegatedPermissions(Inbox inbox)
        {
            // Delegated permissions = Interception between user and App permissions
            // If user has less permissions than App then App will only be granted with user permissions 
            // Interactive scenarios, user presence required
            // https://www.enowsoftware.com/solutions-engine/accessing-exchange-online-objects-without-legacy-auth
            // https://docs.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth#delegated-permissions
            bool success = false;
            try
            {
                // Configure the MSAL client to get tokens
                var pcaOptions = new PublicClientApplicationOptions
                {
                    ClientId = inbox.ClientId,
                    TenantId = inbox.TenantId,
                    RedirectUri = "http://localhost"
                };

                var pca = PublicClientApplicationBuilder
                    .CreateWithApplicationOptions(pcaOptions)
                    // We could add this if ADFS is used by the client and we will need ClientADFSAuthorityUri
                    //.WithAdfsAuthority(ClientADFSAuthorityUri)
                    .Build();

                // The permission scope required for EWS access
                var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };

                // Make the interactive token request
                var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();
                OAuthCredentials authCreds = new OAuthCredentials(authResult.AccessToken);

                ewsClient.SetCredentials(authCreds);

                if (!string.IsNullOrWhiteSpace(inbox.ServerUrl))
                {
                    ewsClient.SetServerUrl(inbox.ServerUrl);
                    success = await ewsClient.TestConnectionAsync();
                }
                else if (ewsClient.Autodiscover(inbox.ReplyFrom))
                {
                    success = await ewsClient.TestConnectionAsync();
                }

                if (success)
                {
                    logger.LogInformation($"{Resource.SuccessConnectingToExchange}: {inbox.ServerUrl}");
                }
                else
                {
                    logger.LogError($"{Resource.ErrorConnectingToExchange}: {inbox.ServerUrl}");
                }
            }
            catch (Exception ex)
            {
                success = false;

                var innerException = ex.InnerException != null ? $" | {ex.InnerException.Message}" : string.Empty;
                var errorMsg = $"{Resource.ErrorConnectingToExchange}: {ex.Message}{innerException}";
                logger.LogError(errorMsg);
            }

            return success;
        }

        #endregion
    }
}