using Coginov.Exchange.Library.Models;
using MailBee.EwsMail;
using MailBee.Mime;
using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Coginov.Exchange.Library.Services
{
    public interface IExchangeService
    {
        Task<bool> InitializeInbox(Inbox inbox, int tagFlagRetryDelay = 10, int tagFlagRetryCount = 5, bool forceReconnect = false, string accountToImpersonate = null);

        Task<EwsItemList> GetEmailsFromFolder(string folder, int startIndex = 0, int emailCount = 10, List<string> errors = null, bool unreadOnly = false, EwsItemParts itemParts = EwsItemParts.MailMessageFull);

        Task<bool> MoveEmailToFolder(string messageId, string srcFolder, string destFolder);

        Task<bool> MoveEmailsToFolder(EwsItemList ewsItemList, string srcFolder, string destFolder);

        Task<bool> CopyEmailToFolder(string messageId, string srcFolder, string destFolder);

        Task<bool> UploadEmailToFolder(MailMessage message, string folder);

        Task<bool> SendEmail(MailMessage email, List<string> to, bool saveCopySentFolder = false, bool includeCc = false);

        Task<bool> TagEmail(string messageId, string folder, List<string> categoryList);

        Task<bool> FlagEmail(string messageId, string folder, ItemFlagStatus status = ItemFlagStatus.Flagged, DateTime? startDate = null, DateTime? dueDate = null);

        Task<bool> DeleteEmail(string messageId);

        Task<bool> SendTemplate(string templatePath, List<string> to);

        Task<bool> ForwardEmailPrefixedByTemplate(MailMessage email, string filePath, List<string> to);

        Task<bool> ForwardEmailUsingTemplate(string templatePath, MailMessage email, List<string> to, Dictionary<string, string> parameters = null);

        Task<bool> SaveEmailAsMsg(MailMessage email, string destinationFolder, string msgFileName, bool saveAttachmentsAsFiles, List<string> errors);

        Task<List<EwsFolder>> GetFolders(bool includeSubfolders = true);

        Task<EwsFolder> GetFolder(string uniqueId);

        Task<EwsFolder> GetAllItemsFolder();

        Task<EwsItemList> GetEmailsFromFolderAfterDate(EwsFolder folder, DateTime afterDate, int startIndex = 0, int emailCount = 10, List<string> errors = null, bool unreadOnly = false, EwsItemParts itemParts = EwsItemParts.MailMessageFull);
    }
}