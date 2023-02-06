using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using Windows.Services.Store;

using GlobalStructures;
using System.Xml.Linq;
using System.Reflection;

namespace WinUI3_SendMail
{
    internal class Outlook
    {
        [ComImport]
        [Guid("00020400-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IDispatch
        {
            int GetTypeInfoCount();
            [return: MarshalAs(UnmanagedType.Interface)]
            IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
            [PreserveSig]
            HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
                [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            [PreserveSig]
            HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
                [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
        }

        [ComImport]
        [Guid("00063001-0000-0000-c000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface _Application
        {
            #region <IDispatch>
            int GetTypeInfoCount();
            [return: MarshalAs(UnmanagedType.Interface)]
            IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
            [PreserveSig]
            HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
                [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            [PreserveSig]
            HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
                [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
            #endregion

            HRESULT get_Application(out _Application Application);
            HRESULT get_Class(out OlObjectClass Class);
            HRESULT get_Session(out _NameSpace Session);
            HRESULT get_Parent(out IDispatch Parent);
            //HRESULT get_Assistant(out Office::Assistant Assistant);
            HRESULT get_Assistant(out IntPtr Assistant);
            HRESULT get_Name(out string Name);
            HRESULT get_Version(out string Version);
            //HRESULT ActiveExplorer(out _Explorer ActiveExplorer);
            HRESULT ActiveExplorer(out IntPtr ActiveExplorer);
            //HRESULT ActiveInspector(out _Inspector ActiveInspector);
            HRESULT ActiveInspector(out IntPtr ActiveInspector);
            HRESULT CreateItem(OlItemType ItemType, out IDispatch Item);
            HRESULT CreateItemFromTemplate(string TemplatePath, PROPVARIANT InFolder, out IDispatch Item);
            HRESULT CreateObject(string ObjectName, out IDispatch Object);
            HRESULT GetNamespace(string Type, out _NameSpace NameSpace);
            HRESULT Quit();
        }

        public enum OlObjectClass
        {
            olApplication = 0,
            olNamespace = 1,
            olFolder = 2,
            olRecipient = 4,
            olAttachment = 5,
            olAddressList = 7,
            olAddressEntry = 8,
            olFolders = 15,
            olItems = 16,
            olRecipients = 17,
            olAttachments = 18,
            olAddressLists = 20,
            olAddressEntries = 21,
            olAppointment = 26,
            olMeetingRequest = 53,
            olMeetingCancellation = 54,
            olMeetingResponseNegative = 55,
            olMeetingResponsePositive = 56,
            olMeetingResponseTentative = 57,
            olRecurrencePattern = 28,
            olExceptions = 29,
            olException = 30,
            olAction = 32,
            olActions = 33,
            olExplorer = 34,
            olInspector = 35,
            olPages = 36,
            olFormDescription = 37,
            olUserProperties = 38,
            olUserProperty = 39,
            olContact = 40,
            olDocument = 41,
            olJournal = 42,
            olMail = 43,
            olNote = 44,
            olPost = 45,
            olReport = 46,
            olRemote = 47,
            olTask = 48,
            olTaskRequest = 49,
            olTaskRequestUpdate = 50,
            olTaskRequestAccept = 51,
            olTaskRequestDecline = 52,
            olExplorers = 60,
            olInspectors = 61,
            olPanes = 62,
            olOutlookBarPane = 63,
            olOutlookBarStorage = 64,
            olOutlookBarGroups = 65,
            olOutlookBarGroup = 66,
            olOutlookBarShortcuts = 67,
            olOutlookBarShortcut = 68,
            olDistributionList = 69,
            olPropertyPageSite = 70,
            olPropertyPages = 71,
            olSyncObject = 72,
            olSyncObjects = 73,
            olSelection = 74,
            olLink = 75,
            olLinks = 76,
            olSearch = 77,
            olResults = 78,
            olViews = 79,
            olView = 80,
            olItemProperties = 98,
            olItemProperty = 99,
            olReminders = 100,
            olReminder = 101,
            olConflict = 102,
            olConflicts = 103,
            olSharing = 104,
            olAccount = 105,
            olAccounts = 106,
            olStore = 107,
            olStores = 108,
            olSelectNamesDialog = 109,
            olExchangeUser = 110,
            olExchangeDistributionList = 111,
            olPropertyAccessor = 112,
            olStorageItem = 113,
            olRules = 114,
            olRule = 115,
            olRuleActions = 116,
            olRuleAction = 117,
            olMoveOrCopyRuleAction = 118,
            olSendRuleAction = 119,
            olTable = 120,
            olRow = 121,
            olAssignToCategoryRuleAction = 122,
            olPlaySoundRuleAction = 123,
            olMarkAsTaskRuleAction = 124,
            olNewItemAlertRuleAction = 125,
            olRuleConditions = 126,
            olRuleCondition = 127,
            olImportanceRuleCondition = 128,
            olFormRegion = 129,
            olCategoryRuleCondition = 130,
            olFormNameRuleCondition = 131,
            olFromRuleCondition = 132,
            olSenderInAddressListRuleCondition = 133,
            olTextRuleCondition = 134,
            olAccountRuleCondition = 135,
            olClassTableView = 136,
            olClassIconView = 137,
            olClassCardView = 138,
            olClassCalendarView = 139,
            olClassTimeLineView = 140,
            olViewFields = 141,
            olViewField = 142,
            olOrderField = 144,
            olOrderFields = 145,
            olViewFont = 146,
            olAutoFormatRule = 147,
            olAutoFormatRules = 148,
            olColumnFormat = 149,
            olColumns = 150,
            olCalendarSharing = 151,
            olCategory = 152,
            olCategories = 153,
            olColumn = 154,
            olClassNavigationPane = 155,
            olNavigationModules = 156,
            olNavigationModule = 157,
            olMailModule = 158,
            olCalendarModule = 159,
            olContactsModule = 160,
            olTasksModule = 161,
            olJournalModule = 162,
            olNotesModule = 163,
            olNavigationGroups = 164,
            olNavigationGroup = 165,
            olNavigationFolders = 166,
            olNavigationFolder = 167,
            olClassBusinessCardView = 168,
            olAttachmentSelection = 169,
            olAddressRuleCondition = 170,
            olUserDefinedProperty = 171,
            olUserDefinedProperties = 172,
            olFromRssFeedRuleCondition = 173,
            olClassTimeZone = 174,
            olClassTimeZones = 175,
            olMobile = 176,
            olSolutionsModule = 177,
            olConversation = 178,
            olSimpleItems = 179,
            olOutspace = 180,
            olMeetingForwardNotification = 181,
            olConversationHeader = 182,
            olClassPeopleView = 183,
            olClassThreadView = 184,
            olPreviewPane = 185,
            olSensitivityRuleCondition = 186,
            olClassMessageListView = 187,
            olClassSearchView = 188
        }

        public enum OlItemType
        {
            olMailItem = 0,
            olAppointmentItem = 1,
            olContactItem = 2,
            olTaskItem = 3,
            olJournalItem = 4,
            olNoteItem = 5,
            olPostItem = 6,
            olDistributionListItem = 7,
            olMobileItemSMS = 11,
            olMobileItemMMS = 12
        }

        [ComImport]
        [Guid("00063002-0000-0000-c000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface _NameSpace
        {
            #region <IDispatch>
            int GetTypeInfoCount();
            [return: MarshalAs(UnmanagedType.Interface)]
            IntPtr GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
            [PreserveSig]
            HRESULT GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames,
                [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
            [PreserveSig]
            HRESULT Invoke(int dispIdMember, [In] ref Guid riid, [In, MarshalAs(UnmanagedType.U4)] int lcid, [In, MarshalAs(UnmanagedType.U4)] int dwFlags,
                [Out, In] System.Runtime.InteropServices.ComTypes.DISPPARAMS pDispParams, [Out] out object pVarResult, [Out, In] System.Runtime.InteropServices.ComTypes.EXCEPINFO pExcepInfo, [Out, MarshalAs(UnmanagedType.LPArray)] IntPtr[] pArgErr);
            #endregion

            HRESULT get_Application(out _Application Application);
            HRESULT get_Class(out OlObjectClass Class);
            HRESULT get_Session(out _NameSpace Session);
            HRESULT get_Parent(out IDispatch Parent);
            //HRESULT get_CurrentUser(out Recipient CurrentUser);
            HRESULT get_CurrentUser(out IntPtr CurrentUser);
            //HRESULT get_Folders(out _Folders Folders);
            HRESULT get_Folders(out IntPtr Folders);
            HRESULT get_Type(out string Type);
            //HRESULT get_AddressLists(out AddressLists AddressLists);
            HRESULT get_AddressLists(out IntPtr AddressLists);
            //HRESULT CreateRecipient(string RecipientName, out Recipient Recipient);
            HRESULT CreateRecipient(string RecipientName, out IntPtr Recipient);
            //HRESULT GetDefaultFolder(OlDefaultFolders FolderType, out MAPIFolder Folder);
            HRESULT GetDefaultFolder(OlDefaultFolders FolderType, out IntPtr Folder);
            //HRESULT GetFolderFromID(string EntryIDFolder, PROPVARIANT EntryIDStore, out MAPIFolder Folder);
            HRESULT GetFolderFromID(string EntryIDFolder, PROPVARIANT EntryIDStore, out IntPtr Folder);
            HRESULT GetItemFromID(string EntryIDItem, PROPVARIANT EntryIDStore, out IDispatch Item);
            //HRESULT GetRecipientFromID(string EntryID, out Recipient Recipient);
            HRESULT GetRecipientFromID(string EntryID, out IntPtr Recipient);
            //HRESULT GetSharedDefaultFolder(Recipient Recipient, OlDefaultFolders FolderType, out MAPIFolder Folder);
            HRESULT GetSharedDefaultFolder(IntPtr Recipient, OlDefaultFolders FolderType, out IntPtr Folder);
            HRESULT Logoff();
            //HRESULT Logon(PROPVARIANT Profile, PROPVARIANT Password, PROPVARIANT ShowDialog, PROPVARIANT NewSession);
            HRESULT Logon(object Profile, object Password, object ShowDialog, object NewSession);
            //HRESULT PickFolder(out MAPIFolder Folder);
            HRESULT PickFolder(out IntPtr Folder);
            HRESULT RefreshRemoteHeaders();
            //HRESULT get_SyncObjects(out SyncObjects SyncObjects);
            HRESULT get_SyncObjects(out IntPtr SyncObjects);
            HRESULT AddStore(PROPVARIANT Store);
            //HRESULT RemoveStore(MAPIFolder Folder);
            HRESULT RemoveStore(IntPtr Folder);
            HRESULT get_Offline(out bool Offline);
            HRESULT Dial(PROPVARIANT ContactItem);
            HRESULT get_MAPIOBJECT(out IntPtr MAPIOBJECT);
            HRESULT get_ExchangeConnectionMode(out OlExchangeConnectionMode ExchangeConnectionMode);
            HRESULT AddStoreEx(PROPVARIANT Store, OlStoreType Type);
            //HRESULT get_Accounts(out _Accounts Accounts);
            HRESULT get_Accounts(out IntPtr Accounts);
            HRESULT get_CurrentProfileName(out string CurrentProfileName);
            //HRESULT get_Stores(out _Stores Stores);
            HRESULT get_Stores(out IntPtr Stores);
            //HRESULT GetSelectNamesDialog(out _SelectNamesDialog SelectNamesDialog);
            HRESULT GetSelectNamesDialog(out IntPtr SelectNamesDialog);
            HRESULT SendAndReceive(bool showProgressDialog);
            //HRESULT get_DefaultStore(out _Store DefaultStore);
            HRESULT get_DefaultStore(out IntPtr DefaultStore);
            //HRESULT GetAddressEntryFromID(string ID, out AddressEntry AddressEntry);
            HRESULT GetAddressEntryFromID(string ID, out IntPtr AddressEntry);
            //HRESULT GetGlobalAddressList(out AddressList globalAddressList);
            HRESULT GetGlobalAddressList(out IntPtr globalAddressList);
            //HRESULT GetStoreFromID(string ID, out _Store Store);
            HRESULT GetStoreFromID(string ID, out IntPtr Store);
            //HRESULT get_Categories(out _Categories Categories);
            HRESULT get_Categories(out IntPtr Categories);
            //HRESULT OpenSharedFolder(string Path, PROPVARIANT Name, PROPVARIANT DownloadAttachments, PROPVARIANT UseTTL, out MAPIFolder ret);
            HRESULT OpenSharedFolder(string Path, PROPVARIANT Name, PROPVARIANT DownloadAttachments, PROPVARIANT UseTTL, out IntPtr ret);
            HRESULT OpenSharedItem(string Path, out IDispatch Item);
            //HRESULT CreateSharingItem( PROPVARIANT Context, PROPVARIANT Provider,  out _SharingItem Item);
            HRESULT CreateSharingItem(PROPVARIANT Context, PROPVARIANT Provider, out IntPtr Item);
            HRESULT get_ExchangeMailboxServerName(out string ExchangeMailboxServerName);
            HRESULT get_ExchangeMailboxServerVersion(out string ExchangeMailboxServerVersion);
            HRESULT CompareEntryIDs(string FirstEntryID, string SecondEntryID, out bool Result);
            HRESULT get_AutoDiscoverXml(out string AutoDiscoverXml);
            HRESULT get_AutoDiscoverConnectionMode(out OlAutoDiscoverConnectionMode AutoDiscoverConnectionMode);
            //HRESULT CreateContactCard(AddressEntry AddressEntry, out Office::ContactCard Card);
            HRESULT CreateContactCard(IntPtr AddressEntry, out IntPtr Card);

        }

        public enum OlDefaultFolders
        {
            olFolderDeletedItems = 3,
            olFolderOutbox = 4,
            olFolderSentMail = 5,
            olFolderInbox = 6,
            olFolderCalendar = 9,
            olFolderContacts = 10,
            olFolderJournal = 11,
            olFolderNotes = 12,
            olFolderTasks = 13,
            olFolderDrafts = 16,
            olPublicFoldersAllPublicFolders = 18,
            olFolderConflicts = 19,
            olFolderSyncIssues = 20,
            olFolderLocalFailures = 21,
            olFolderServerFailures = 22,
            olFolderJunk = 23,
            olFolderRssFeeds = 25,
            olFolderToDo = 28,
            olFolderManagedEmail = 29,
            olFolderSuggestedContacts = 30
        }

        public enum OlExchangeConnectionMode
        {
            olNoExchange = 0,
            olOffline = 100,
            olCachedOffline = 200,
            olDisconnected = 300,
            olCachedDisconnected = 400,
            olCachedConnectedHeaders = 500,
            olCachedConnectedDrizzle = 600,
            olCachedConnectedFull = 700,
            olOnline = 800
        }

        public enum OlAutoDiscoverConnectionMode
        {
            olAutoDiscoverConnectionUnknown = 0,
            olAutoDiscoverConnectionExternal = 1,
            olAutoDiscoverConnectionInternal = 2,
            olAutoDiscoverConnectionInternalDomain = 3
        }

        public enum OlStoreType
        {
            olStoreDefault = 1,
            olStoreUnicode = 2,
            olStoreANSI = 3
        };

    }
}
