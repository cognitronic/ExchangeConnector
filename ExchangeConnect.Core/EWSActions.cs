using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExchangeConnect;
using Microsoft.Exchange.WebServices.Data;

namespace ExchangeConnect.Core
{
    public class EWSActions
    {
        public IList<EmailMessage> GetSupportMessages(int numberOfViewItems)
        {
            ExchangeService service = EWSServerInfo.GetExchangeProxy();

            ItemView view = new ItemView(numberOfViewItems);
            IList<EmailMessage> messages = new List<EmailMessage>();
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, view);
            foreach (Item item in findResults.Items)
            {
                EmailMessage message = EmailMessage.Bind(service, item.Id);
                messages.Add(message);
            }
            return messages;
        }

        public void MoveProcessedSupportMessage(EmailMessage message)
        {
            ExchangeService service = EWSServerInfo.GetExchangeProxy();
            FindFoldersResults folderResults = service.FindFolders(WellKnownFolderName.Inbox, new FolderView(int.MaxValue));
            message.IsRead = true;
            message.Update(ConflictResolutionMode.AlwaysOverwrite);
            message.Move(folderResults.Folders[0].Id);
        }

        public void MoveProcessedMessage(EmailMessage message)
        {
            ExchangeService service = EWSServerInfo.GetExchangeProxy();
            FindFoldersResults folderResults = service.FindFolders(WellKnownFolderName.Inbox, new FolderView(int.MaxValue));
            message.IsRead = true;
            message.Update(ConflictResolutionMode.AlwaysOverwrite);
            message.Move(folderResults.Folders[0].Id);
        }

        public void MoveProcessedMessage(EmailMessage message, ExchangeService serviceInfo)
        {
            ExchangeService service = serviceInfo;
            FindFoldersResults folderResults = service.FindFolders(WellKnownFolderName.Inbox, new FolderView(int.MaxValue));
            message.IsRead = true;
            message.Update(ConflictResolutionMode.AlwaysOverwrite);
            message.Move(folderResults.Folders[0].Id);
        }

        public IList<EmailMessage> GetEmployeeSubmittedTaskMessages(int numberOfViewItems)
        {
            ExchangeService service = EWSServerInfo.GetExchangeProxy();

            ItemView view = new ItemView(numberOfViewItems);
            IList<EmailMessage> messages = new List<EmailMessage>();
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, view);
            foreach (Item item in findResults.Items)
            {
                EmailMessage message = EmailMessage.Bind(service, item.Id);
                messages.Add(message);
            }
            return messages;
        }

        public IList<EmailMessage> GetEmployeeSubmittedTaskMessages(int numberOfViewItems, string user, string password, string domain, string autoDiscoverURL)
        {
            ExchangeService service = EWSServerInfo.GetExchangeProxy(user, password, domain, autoDiscoverURL);

            ItemView view = new ItemView(numberOfViewItems);
            IList<EmailMessage> messages = new List<EmailMessage>();
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, view);
            foreach (Item item in findResults.Items)
            {
                EmailMessage message = EmailMessage.Bind(service, item.Id);
                messages.Add(message);
            }
            return messages;
        }
    }
}
