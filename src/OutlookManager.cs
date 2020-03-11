using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Extensions.Logging;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace SaveOutlookAttachments
{
    public sealed class OutlookManager : IDisposable
    {
        private readonly ILogger<OutlookManager> m_logger;
        private Outlook.Application m_application;
        private Outlook.NameSpace m_session;
        private Outlook.Store m_store;
        private bool m_removeStore;
        private Outlook.MAPIFolder m_rootFolder;

        public Action<AppContext, object> ProcessItem { private get; set; }

        public OutlookManager(ILogger<OutlookManager> logger)
        {
            m_logger = logger;
        }

        public void Initialise()
        {
            m_application = new Outlook.Application();

            m_session = m_application.GetNamespace(Constants.MapiNamespace);
        }

        public Outlook.Store GetCurrentStore()
        {
            return m_store;
        }

        public IEnumerable<Outlook.Store> GetStores()
        {
            foreach (var store in m_session.GetStores())
            {
                yield return store;
            }
        }

        public bool TrySetStore(string storeDescriptor)
        {
            if (string.IsNullOrWhiteSpace(storeDescriptor))
            {
                m_store = m_session.GetStores().FirstOrDefault();
            }
            else
            {
                if (m_session.TryGetStore(storeDescriptor, out m_store) == false)
                {
                    m_session.AddStore(storeDescriptor);

                    m_removeStore = m_session.TryGetStore(storeDescriptor, out m_store);
                }
            }

            return m_store != null;
        }

        public bool TrySetFolder()
        {
            m_rootFolder = m_store.GetRootFolder();

            return m_rootFolder != null;
        }

        public void ForEachAttachment(AppContext ctx)
        {
            ProcessFolder(ctx, m_rootFolder);
        }

        public object OpenSharedItem(string path)
        {
            return m_session.OpenSharedItem(path);
        }

        private void ProcessFolder(AppContext ctx, Outlook.MAPIFolder folder)
        {
            ProcessFolderItems(ctx, folder);

            ProcessFolders(ctx, folder.Folders);
        }

        private void ProcessFolders(AppContext ctx, Outlook.Folders folders)
        {
            foreach (Outlook.MAPIFolder folder in folders)
            {
                ProcessFolder(ctx, folder);
            }
        }

        private void ProcessFolderItems(AppContext ctx, Outlook.MAPIFolder folder)
        {
            foreach (var item in folder.Items)
            {
                ProcessItem?.Invoke(ctx, item);
            }
        }

        public void Dispose()
        {
            if (m_session != null && m_removeStore && m_rootFolder != null)
            {
                m_session.RemoveStore(m_rootFolder);
            }

            m_session = null;

            m_application?.Quit();
            m_application = null;
        }
    }
}
