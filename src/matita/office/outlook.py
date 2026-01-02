import win32com.client

class Account:

    def __init__(self, account=None):
        self.account = account

    @property
    def AccountType(self):
        return OlAccountType(self.account.AccountType)

    @property
    def Application(self):
        return Application(self.account.Application)

    @property
    def AutoDiscoverConnectionMode(self):
        return OlAutoDiscoverConnectionMode(self.account.AutoDiscoverConnectionMode)

    @property
    def AutoDiscoverXml(self):
        return self.account.AutoDiscoverXml

    @property
    def Class(self):
        return OlObjectClass(self.account.Class)

    @property
    def CurrentUser(self):
        return Recipient(self.account.CurrentUser)

    @property
    def DeliveryStore(self):
        return Store(self.account.DeliveryStore)

    @property
    def DisplayName(self):
        return Account(self.account.DisplayName)

    @property
    def ExchangeConnectionMode(self):
        return OlExchangeConnectionMode(self.account.ExchangeConnectionMode)

    @property
    def ExchangeMailboxServerName(self):
        return self.account.ExchangeMailboxServerName

    @property
    def ExchangeMailboxServerVersion(self):
        return self.account.ExchangeMailboxServerVersion

    @property
    def Parent(self):
        return self.account.Parent

    @property
    def Session(self):
        return NameSpace(self.account.Session)

    @property
    def SmtpAddress(self):
        return Account(self.account.SmtpAddress)

    @property
    def UserName(self):
        return Account(self.account.UserName)

    def GetAddressEntryFromID(self, *args, ID=None):
        arguments = {"ID": ID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ID(self.account.GetAddressEntryFromID(*args, **arguments))

    def GetRecipientFromID(self, *args, EntryID=None):
        arguments = {"EntryID": EntryID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.account.GetRecipientFromID(*args, **arguments)


class AccountRuleCondition:

    def __init__(self, accountrulecondition=None):
        self.accountrulecondition = accountrulecondition

    @property
    def Account(self):
        return Account(self.accountrulecondition.Account)

    @Account.setter
    def Account(self, value):
        self.accountrulecondition.Account = value

    @property
    def Application(self):
        return Application(self.accountrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.accountrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.accountrulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.accountrulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.accountrulecondition.Enabled = value

    @property
    def Parent(self):
        return self.accountrulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.accountrulecondition.Session)


class Accounts:

    def __init__(self, accounts=None):
        self.accounts = accounts

    @property
    def Application(self):
        return Application(self.accounts.Application)

    @property
    def Class(self):
        return OlObjectClass(self.accounts.Class)

    @property
    def Count(self):
        return self.accounts.Count

    @property
    def Parent(self):
        return self.accounts.Parent

    @property
    def Session(self):
        return NameSpace(self.accounts.Session)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.accounts.Item(*args, **arguments)


class AccountSelector:

    def __init__(self, accountselector=None):
        self.accountselector = accountselector

    @property
    def Application(self):
        return Application(self.accountselector.Application)

    @property
    def Class(self):
        return OlObjectClass(self.accountselector.Class)

    @property
    def Parent(self):
        return self.accountselector.Parent

    @property
    def SelectedAccount(self):
        return Account(self.accountselector.SelectedAccount)

    @property
    def Session(self):
        return NameSpace(self.accountselector.Session)


class Action:

    def __init__(self, action=None):
        self.action = action

    @property
    def Application(self):
        return Application(self.action.Application)

    @property
    def Class(self):
        return OlObjectClass(self.action.Class)

    @property
    def CopyLike(self):
        return OlActionCopyLike(self.action.CopyLike)

    @CopyLike.setter
    def CopyLike(self, value):
        self.action.CopyLike = value

    @property
    def Enabled(self):
        return self.action.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.action.Enabled = value

    @property
    def MessageClass(self):
        return Action(self.action.MessageClass)

    @MessageClass.setter
    def MessageClass(self, value):
        self.action.MessageClass = value

    @property
    def Name(self):
        return self.action.Name

    @Name.setter
    def Name(self, value):
        self.action.Name = value

    @property
    def Parent(self):
        return self.action.Parent

    @property
    def Prefix(self):
        return self.action.Prefix

    @Prefix.setter
    def Prefix(self, value):
        self.action.Prefix = value

    @property
    def ReplyStyle(self):
        return OlActionReplyStyle(self.action.ReplyStyle)

    @ReplyStyle.setter
    def ReplyStyle(self, value):
        self.action.ReplyStyle = value

    @property
    def ResponseStyle(self):
        return OlActionResponseStyle(self.action.ResponseStyle)

    @ResponseStyle.setter
    def ResponseStyle(self, value):
        self.action.ResponseStyle = value

    @property
    def Session(self):
        return NameSpace(self.action.Session)

    @property
    def ShowOn(self):
        return OlActionShowOn(self.action.ShowOn)

    @ShowOn.setter
    def ShowOn(self, value):
        self.action.ShowOn = value

    def Delete(self):
        self.action.Delete()

    def Execute(self):
        return self.action.Execute()


class Actions:

    def __init__(self, actions=None):
        self.actions = actions

    @property
    def Application(self):
        return Application(self.actions.Application)

    @property
    def Class(self):
        return OlObjectClass(self.actions.Class)

    @property
    def Count(self):
        return self.actions.Count

    @property
    def Parent(self):
        return self.actions.Parent

    @property
    def Session(self):
        return NameSpace(self.actions.Session)

    def Add(self):
        return Action(self.actions.Add())

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.actions.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.actions.Remove(*args, **arguments)


class AddressEntries:

    def __init__(self, addressentries=None):
        self.addressentries = addressentries

    @property
    def Application(self):
        return Application(self.addressentries.Application)

    @property
    def Class(self):
        return OlObjectClass(self.addressentries.Class)

    @property
    def Count(self):
        return self.addressentries.Count

    @property
    def Parent(self):
        return self.addressentries.Parent

    @property
    def Session(self):
        return NameSpace(self.addressentries.Session)

    def Add(self, *args, Type=None, Name=None, Address=None):
        arguments = {"Type": Type, "Name": Name, "Address": Address}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return AddressEntry(self.addressentries.Add(*args, **arguments))

    def GetFirst(self):
        return AddressEntry(self.addressentries.GetFirst())

    def GetLast(self):
        return AddressEntry(self.addressentries.GetLast())

    def GetNext(self):
        return AddressEntry(self.addressentries.GetNext())

    def GetPrevious(self):
        return AddressEntry(self.addressentries.GetPrevious())

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.addressentries.Item(*args, **arguments)

    def Sort(self, *args, Property=None, Order=None):
        arguments = {"Property": Property, "Order": Order}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.addressentries.Sort(*args, **arguments)


class AddressEntry:

    def __init__(self, addressentry=None):
        self.addressentry = addressentry

    @property
    def Address(self):
        return AddressEntry(self.addressentry.Address)

    @Address.setter
    def Address(self, value):
        self.addressentry.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.addressentry.AddressEntryUserType)

    @property
    def Application(self):
        return Application(self.addressentry.Application)

    @property
    def Class(self):
        return OlObjectClass(self.addressentry.Class)

    @property
    def DisplayType(self):
        return OlDisplayType(self.addressentry.DisplayType)

    @property
    def ID(self):
        return self.addressentry.ID

    @property
    def Name(self):
        return self.addressentry.Name

    @Name.setter
    def Name(self, value):
        self.addressentry.Name = value

    @property
    def Parent(self):
        return self.addressentry.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.addressentry.PropertyAccessor)

    @property
    def Session(self):
        return NameSpace(self.addressentry.Session)

    @property
    def Type(self):
        return self.addressentry.Type

    @Type.setter
    def Type(self, value):
        self.addressentry.Type = value

    def Delete(self):
        self.addressentry.Delete()

    def Details(self, *args, HWnd=None):
        arguments = {"HWnd": HWnd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.addressentry.Details(*args, **arguments)

    def GetContact(self):
        return self.addressentry.GetContact()

    def GetExchangeDistributionList(self):
        return self.addressentry.GetExchangeDistributionList()

    def GetExchangeUser(self):
        return self.addressentry.GetExchangeUser()

    def GetFreeBusy(self, *args, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = {"Start": Start, "MinPerChar": MinPerChar, "CompleteFormat": CompleteFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.addressentry.GetFreeBusy(*args, **arguments)

    def Update(self, *args, MakePermanent=None, Refresh=None):
        arguments = {"MakePermanent": MakePermanent, "Refresh": Refresh}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.addressentry.Update(*args, **arguments)


class AddressList:

    def __init__(self, addresslist=None):
        self.addresslist = addresslist

    @property
    def AddressEntries(self):
        return AddressEntries(self.addresslist.AddressEntries)

    @property
    def AddressListType(self):
        return OlAddressListType(self.addresslist.AddressListType)

    @property
    def Application(self):
        return Application(self.addresslist.Application)

    @property
    def Class(self):
        return OlObjectClass(self.addresslist.Class)

    @property
    def ID(self):
        return self.addresslist.ID

    @property
    def Index(self):
        return self.addresslist.Index

    @property
    def IsInitialAddressList(self):
        return AddressList(self.addresslist.IsInitialAddressList)

    @property
    def IsReadOnly(self):
        return AddressList(self.addresslist.IsReadOnly)

    @property
    def Name(self):
        return self.addresslist.Name

    @property
    def Parent(self):
        return self.addresslist.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.addresslist.PropertyAccessor)

    @property
    def ResolutionOrder(self):
        return AddressList(self.addresslist.ResolutionOrder)

    @property
    def Session(self):
        return NameSpace(self.addresslist.Session)

    def GetContactsFolder(self):
        return self.addresslist.GetContactsFolder()


class AddressLists:

    def __init__(self, addresslists=None):
        self.addresslists = addresslists

    @property
    def Application(self):
        return Application(self.addresslists.Application)

    @property
    def Class(self):
        return OlObjectClass(self.addresslists.Class)

    @property
    def Count(self):
        return self.addresslists.Count

    @property
    def Parent(self):
        return self.addresslists.Parent

    @property
    def Session(self):
        return NameSpace(self.addresslists.Session)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.addresslists.Item(*args, **arguments)


class AddressRuleCondition:

    def __init__(self, addressrulecondition=None):
        self.addressrulecondition = addressrulecondition

    @property
    def Address(self):
        return self.addressrulecondition.Address

    @Address.setter
    def Address(self, value):
        self.addressrulecondition.Address = value

    @property
    def Application(self):
        return Application(self.addressrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.addressrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.addressrulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.addressrulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.addressrulecondition.Enabled = value

    @property
    def Parent(self):
        return self.addressrulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.addressrulecondition.Session)


class Application:

    def __init__(self, application=None):
        self.application = application

    def new(self):
        self.application = win32com.client.Dispatch("Outlook.Application")
        return self

    @property
    def Application(self):
        return Application(self.application.Application)

    @property
    def Assistance(self):
        return self.application.Assistance

    @property
    def Class(self):
        return OlObjectClass(self.application.Class)

    @property
    def COMAddIns(self):
        return self.application.COMAddIns

    @property
    def DefaultProfileName(self):
        return self.application.DefaultProfileName

    @property
    def Explorers(self):
        return Explorers(self.application.Explorers)

    @property
    def Inspectors(self):
        return Inspectors(self.application.Inspectors)

    @property
    def IsTrusted(self):
        return self.application.IsTrusted

    @property
    def LanguageSettings(self):
        return self.application.LanguageSettings

    @property
    def Name(self):
        return self.application.Name

    @property
    def Parent(self):
        return self.application.Parent

    @property
    def PickerDialog(self):
        return self.application.PickerDialog

    @property
    def ProductCode(self):
        return self.application.ProductCode

    @property
    def Reminders(self):
        return Reminders(self.application.Reminders)

    @property
    def Session(self):
        return NameSpace(self.application.Session)

    @property
    def TimeZones(self):
        return TimeZones(self.application.TimeZones)

    @property
    def Version(self):
        return self.application.Version

    @Version.setter
    def Version(self, value):
        self.application.Version = value

    def ActiveExplorer(self):
        return self.application.ActiveExplorer()

    def ActiveInspector(self):
        return self.application.ActiveInspector()

    def ActiveWindow(self):
        return self.application.ActiveWindow()

    def AdvancedSearch(self, *args, Scope=None, Filter=None, SearchSubFolders=None, Tag=None):
        arguments = {"Scope": Scope, "Filter": Filter, "SearchSubFolders": SearchSubFolders, "Tag": Tag}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Search(self.application.AdvancedSearch(*args, **arguments))

    def CopyFile(self, *args, FilePath=None, DestFolderPath=None):
        arguments = {"FilePath": FilePath, "DestFolderPath": DestFolderPath}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CopyFile(*args, **arguments)

    def CreateItem(self, *args, ItemType=None):
        arguments = {"ItemType": ItemType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateItem(*args, **arguments)

    def CreateItemFromTemplate(self, *args, TemplatePath=None, InFolder=None):
        arguments = {"TemplatePath": TemplatePath, "InFolder": InFolder}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateItemFromTemplate(*args, **arguments)

    def CreateObject(self, *args, ObjectName=None):
        arguments = {"ObjectName": ObjectName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CreateObject(*args, **arguments)

    def GetNamespace(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetNamespace(*args, **arguments)

    def GetObjectReference(self, *args, Item=None, ReferenceType=None):
        arguments = {"Item": Item, "ReferenceType": ReferenceType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetObjectReference(*args, **arguments)

    def IsSearchSynchronous(self, *args, LookInFolders=None):
        arguments = {"LookInFolders": LookInFolders}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.IsSearchSynchronous(*args, **arguments)

    def Quit(self):
        self.application.Quit()

    def RefreshFormRegionDefinition(self, *args, RegionName=None):
        arguments = {"RegionName": RegionName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.RefreshFormRegionDefinition(*args, **arguments)


class AppointmentItem:

    def __init__(self, appointmentitem=None):
        self.appointmentitem = appointmentitem

    @property
    def Actions(self):
        return Actions(self.appointmentitem.Actions)

    @property
    def AllDayEvent(self):
        return self.appointmentitem.AllDayEvent

    @AllDayEvent.setter
    def AllDayEvent(self, value):
        self.appointmentitem.AllDayEvent = value

    @property
    def Application(self):
        return Application(self.appointmentitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.appointmentitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.appointmentitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.appointmentitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.appointmentitem.BillingInformation = value

    @property
    def Body(self):
        return self.appointmentitem.Body

    @Body.setter
    def Body(self, value):
        self.appointmentitem.Body = value

    @property
    def BusyStatus(self):
        return OlBusyStatus(self.appointmentitem.BusyStatus)

    @BusyStatus.setter
    def BusyStatus(self, value):
        self.appointmentitem.BusyStatus = value

    @property
    def Categories(self):
        return self.appointmentitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.appointmentitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.appointmentitem.Class)

    @property
    def Companies(self):
        return self.appointmentitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.appointmentitem.Companies = value

    @property
    def Conflicts(self):
        return self.appointmentitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.appointmentitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.appointmentitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.appointmentitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.appointmentitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.appointmentitem.DownloadState)

    @property
    def Duration(self):
        return AppointmentItem(self.appointmentitem.Duration)

    @Duration.setter
    def Duration(self, value):
        self.appointmentitem.Duration = value

    @property
    def End(self):
        return AppointmentItem(self.appointmentitem.End)

    @End.setter
    def End(self, value):
        self.appointmentitem.End = value

    @property
    def EndInEndTimeZone(self):
        return AppointmentItem.EndTimeZone(self.appointmentitem.EndInEndTimeZone)

    @EndInEndTimeZone.setter
    def EndInEndTimeZone(self, value):
        self.appointmentitem.EndInEndTimeZone = value

    @property
    def EndTimeZone(self):
        return TimeZone(self.appointmentitem.EndTimeZone)

    @EndTimeZone.setter
    def EndTimeZone(self, value):
        self.appointmentitem.EndTimeZone = value

    @property
    def EndUTC(self):
        return self.appointmentitem.EndUTC

    @EndUTC.setter
    def EndUTC(self, value):
        self.appointmentitem.EndUTC = value

    @property
    def EntryID(self):
        return self.appointmentitem.EntryID

    @property
    def ForceUpdateToAllAttendees(self):
        return self.appointmentitem.ForceUpdateToAllAttendees

    @ForceUpdateToAllAttendees.setter
    def ForceUpdateToAllAttendees(self, value):
        self.appointmentitem.ForceUpdateToAllAttendees = value

    @property
    def FormDescription(self):
        return FormDescription(self.appointmentitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.appointmentitem.GetInspector)

    @property
    def GlobalAppointmentID(self):
        return AppointmentItem(self.appointmentitem.GlobalAppointmentID)

    @property
    def Importance(self):
        return OlImportance(self.appointmentitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.appointmentitem.Importance = value

    @property
    def InternetCodepage(self):
        return self.appointmentitem.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.appointmentitem.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.appointmentitem.IsConflict

    @property
    def IsRecurring(self):
        return self.appointmentitem.IsRecurring

    @property
    def ItemProperties(self):
        return ItemProperties(self.appointmentitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.appointmentitem.LastModificationTime

    @property
    def Location(self):
        return self.appointmentitem.Location

    @Location.setter
    def Location(self, value):
        self.appointmentitem.Location = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.appointmentitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.appointmentitem.MarkForDownload = value

    @property
    def MeetingStatus(self):
        return OlMeetingStatus(self.appointmentitem.MeetingStatus)

    @MeetingStatus.setter
    def MeetingStatus(self, value):
        self.appointmentitem.MeetingStatus = value

    @property
    def MeetingWorkspaceURL(self):
        return self.appointmentitem.MeetingWorkspaceURL

    @property
    def MessageClass(self):
        return self.appointmentitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.appointmentitem.MessageClass = value

    @property
    def Mileage(self):
        return self.appointmentitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.appointmentitem.Mileage = value

    @property
    def NoAging(self):
        return self.appointmentitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.appointmentitem.NoAging = value

    @property
    def OptionalAttendees(self):
        return self.appointmentitem.OptionalAttendees

    @OptionalAttendees.setter
    def OptionalAttendees(self, value):
        self.appointmentitem.OptionalAttendees = value

    @property
    def Organizer(self):
        return self.appointmentitem.Organizer

    @property
    def OutlookInternalVersion(self):
        return self.appointmentitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.appointmentitem.OutlookVersion

    @property
    def Parent(self):
        return self.appointmentitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.appointmentitem.PropertyAccessor)

    @property
    def Recipients(self):
        return Recipients(self.appointmentitem.Recipients)

    @property
    def RecurrenceState(self):
        return OlRecurrenceState(self.appointmentitem.RecurrenceState)

    @property
    def ReminderMinutesBeforeStart(self):
        return self.appointmentitem.ReminderMinutesBeforeStart

    @ReminderMinutesBeforeStart.setter
    def ReminderMinutesBeforeStart(self, value):
        self.appointmentitem.ReminderMinutesBeforeStart = value

    @property
    def ReminderOverrideDefault(self):
        return self.appointmentitem.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.appointmentitem.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.appointmentitem.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.appointmentitem.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.appointmentitem.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.appointmentitem.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.appointmentitem.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.appointmentitem.ReminderSoundFile = value

    @property
    def ReplyTime(self):
        return self.appointmentitem.ReplyTime

    @ReplyTime.setter
    def ReplyTime(self, value):
        self.appointmentitem.ReplyTime = value

    @property
    def RequiredAttendees(self):
        return self.appointmentitem.RequiredAttendees

    @RequiredAttendees.setter
    def RequiredAttendees(self, value):
        self.appointmentitem.RequiredAttendees = value

    @property
    def Resources(self):
        return self.appointmentitem.Resources

    @Resources.setter
    def Resources(self, value):
        self.appointmentitem.Resources = value

    @property
    def ResponseRequested(self):
        return self.appointmentitem.ResponseRequested

    @ResponseRequested.setter
    def ResponseRequested(self, value):
        self.appointmentitem.ResponseRequested = value

    @property
    def ResponseStatus(self):
        return OlResponseStatus(self.appointmentitem.ResponseStatus)

    @property
    def RTFBody(self):
        return self.appointmentitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.appointmentitem.RTFBody = value

    @property
    def Saved(self):
        return self.appointmentitem.Saved

    @property
    def SendUsingAccount(self):
        return Account(self.appointmentitem.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.appointmentitem.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.appointmentitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.appointmentitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.appointmentitem.Session)

    @property
    def Size(self):
        return self.appointmentitem.Size

    @property
    def Start(self):
        return self.appointmentitem.Start

    @Start.setter
    def Start(self, value):
        self.appointmentitem.Start = value

    @property
    def StartInStartTimeZone(self):
        return AppointmentItem.StartTimeZone(self.appointmentitem.StartInStartTimeZone)

    @StartInStartTimeZone.setter
    def StartInStartTimeZone(self, value):
        self.appointmentitem.StartInStartTimeZone = value

    @property
    def StartTimeZone(self):
        return TimeZone(self.appointmentitem.StartTimeZone)

    @StartTimeZone.setter
    def StartTimeZone(self, value):
        self.appointmentitem.StartTimeZone = value

    @property
    def StartUTC(self):
        return self.appointmentitem.StartUTC

    @StartUTC.setter
    def StartUTC(self, value):
        self.appointmentitem.StartUTC = value

    @property
    def Subject(self):
        return self.appointmentitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.appointmentitem.Subject = value

    @property
    def UnRead(self):
        return self.appointmentitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.appointmentitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.appointmentitem.UserProperties)

    def ClearRecurrencePattern(self):
        self.appointmentitem.ClearRecurrencePattern()

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.appointmentitem.Close(*args, **arguments)

    def Copy(self):
        self.appointmentitem.Copy()

    def CopyTo(self, *args, DestinationFolder=None, CopyOptions=None):
        arguments = {"DestinationFolder": DestinationFolder, "CopyOptions": CopyOptions}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.appointmentitem.CopyTo(*args, **arguments)

    def Delete(self):
        self.appointmentitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.appointmentitem.Display(*args, **arguments)

    def ForwardAsVcal(self):
        return MailItem(self.appointmentitem.ForwardAsVcal())

    def GetConversation(self):
        return self.appointmentitem.GetConversation()

    def GetOrganizer(self):
        return self.appointmentitem.GetOrganizer()

    def GetRecurrencePattern(self):
        return self.appointmentitem.GetRecurrencePattern()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.appointmentitem.Move(*args, **arguments)

    def PrintOut(self):
        self.appointmentitem.PrintOut()

    def Respond(self, *args, Response=None, fNoUI=None, fAdditionalTextDialog=None):
        arguments = {"Response": Response, "fNoUI": fNoUI, "fAdditionalTextDialog": fAdditionalTextDialog}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return MeetingItem(self.appointmentitem.Respond(*args, **arguments))

    def Save(self):
        self.appointmentitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.appointmentitem.SaveAs(*args, **arguments)

    def Send(self):
        self.appointmentitem.Send()

    def ShowCategoriesDialog(self):
        self.appointmentitem.ShowCategoriesDialog()


class AssignToCategoryRuleAction:

    def __init__(self, assigntocategoryruleaction=None):
        self.assigntocategoryruleaction = assigntocategoryruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.assigntocategoryruleaction.ActionType)

    @property
    def Application(self):
        return Application(self.assigntocategoryruleaction.Application)

    @property
    def Categories(self):
        return self.assigntocategoryruleaction.Categories

    @Categories.setter
    def Categories(self, value):
        self.assigntocategoryruleaction.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.assigntocategoryruleaction.Class)

    @property
    def Enabled(self):
        return self.assigntocategoryruleaction.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.assigntocategoryruleaction.Enabled = value

    @property
    def Parent(self):
        return self.assigntocategoryruleaction.Parent

    @property
    def Session(self):
        return NameSpace(self.assigntocategoryruleaction.Session)


class Attachment:

    def __init__(self, attachment=None):
        self.attachment = attachment

    @property
    def Application(self):
        return Application(self.attachment.Application)

    @property
    def BlockLevel(self):
        return OlAttachmentBlockLevel(self.attachment.BlockLevel)

    @property
    def Class(self):
        return OlObjectClass(self.attachment.Class)

    @property
    def DisplayName(self):
        return self.attachment.DisplayName

    @DisplayName.setter
    def DisplayName(self, value):
        self.attachment.DisplayName = value

    @property
    def FileName(self):
        return self.attachment.FileName

    @property
    def Index(self):
        return self.attachment.Index

    @property
    def Parent(self):
        return self.attachment.Parent

    @property
    def PathName(self):
        return self.attachment.PathName

    @property
    def Position(self):
        return self.attachment.Position

    @Position.setter
    def Position(self, value):
        self.attachment.Position = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.attachment.PropertyAccessor)

    @property
    def Session(self):
        return NameSpace(self.attachment.Session)

    @property
    def Size(self):
        return self.attachment.Size

    @property
    def Type(self):
        return OlAttachmentType(self.attachment.Type)

    def Delete(self):
        self.attachment.Delete()

    def GetTemporaryFilePath(self):
        return self.attachment.GetTemporaryFilePath()

    def SaveAsFile(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.attachment.SaveAsFile(*args, **arguments)


class Attachments:

    def __init__(self, attachments=None):
        self.attachments = attachments

    @property
    def Application(self):
        return Application(self.attachments.Application)

    @property
    def Class(self):
        return OlObjectClass(self.attachments.Class)

    @property
    def Count(self):
        return self.attachments.Count

    @property
    def Parent(self):
        return self.attachments.Parent

    @property
    def Session(self):
        return NameSpace(self.attachments.Session)

    def Add(self, *args, Source=None, Type=None, Position=None, DisplayName=None):
        arguments = {"Source": Source, "Type": Type, "Position": Position, "DisplayName": DisplayName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Attachment(self.attachments.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.attachments.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.attachments.Remove(*args, **arguments)


class AttachmentSelection:

    def __init__(self, attachmentselection=None):
        self.attachmentselection = attachmentselection

    @property
    def Application(self):
        return Application(self.attachmentselection.Application)

    @property
    def Class(self):
        return OlObjectClass(self.attachmentselection.Class)

    @property
    def Count(self):
        return self.attachmentselection.Count

    @property
    def Location(self):
        return OlSelectionLocation(self.attachmentselection.Location)

    @property
    def Parent(self):
        return self.attachmentselection.Parent

    @property
    def Session(self):
        return NameSpace(self.attachmentselection.Session)

    def GetSelection(self, *args, SelectionContents=None):
        arguments = {"SelectionContents": SelectionContents}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.attachmentselection.GetSelection(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.attachmentselection.Item(*args, **arguments)


class AutoFormatRule:

    def __init__(self, autoformatrule=None):
        self.autoformatrule = autoformatrule

    @property
    def Application(self):
        return Application(self.autoformatrule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.autoformatrule.Class)

    @property
    def Enabled(self):
        return AutoFormatRule(self.autoformatrule.Enabled)

    @Enabled.setter
    def Enabled(self, value):
        self.autoformatrule.Enabled = value

    @property
    def Filter(self):
        return self.autoformatrule.Filter

    @Filter.setter
    def Filter(self, value):
        self.autoformatrule.Filter = value

    @property
    def Font(self):
        return ViewFont(self.autoformatrule.Font)

    @property
    def Name(self):
        return self.autoformatrule.Name

    @Name.setter
    def Name(self, value):
        self.autoformatrule.Name = value

    @property
    def Parent(self):
        return self.autoformatrule.Parent

    @property
    def Session(self):
        return NameSpace(self.autoformatrule.Session)

    @property
    def Standard(self):
        return AutoFormatRule(self.autoformatrule.Standard)


class AutoFormatRules:

    def __init__(self, autoformatrules=None):
        self.autoformatrules = autoformatrules

    @property
    def Application(self):
        return Application(self.autoformatrules.Application)

    @property
    def Class(self):
        return OlObjectClass(self.autoformatrules.Class)

    @property
    def Count(self):
        return AutoFormatRule(self.autoformatrules.Count)

    @property
    def Parent(self):
        return self.autoformatrules.Parent

    @property
    def Session(self):
        return NameSpace(self.autoformatrules.Session)

    def Add(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.autoformatrules.Add(*args, **arguments)

    def Insert(self, *args, Name=None, Index=None):
        arguments = {"Name": Name, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.autoformatrules.Insert(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.autoformatrules.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.autoformatrules.Remove(*args, **arguments)

    def RemoveAll(self):
        self.autoformatrules.RemoveAll()

    def Save(self):
        self.autoformatrules.Save()


class BusinessCardView:

    def __init__(self, businesscardview=None):
        self.businesscardview = businesscardview

    @property
    def Application(self):
        return Application(self.businesscardview.Application)

    @property
    def CardSize(self):
        return self.businesscardview.CardSize

    @CardSize.setter
    def CardSize(self, value):
        self.businesscardview.CardSize = value

    @property
    def Class(self):
        return OlObjectClass(self.businesscardview.Class)

    @property
    def Filter(self):
        return self.businesscardview.Filter

    @Filter.setter
    def Filter(self, value):
        self.businesscardview.Filter = value

    @property
    def HeadingsFont(self):
        return ViewFont(self.businesscardview.HeadingsFont)

    @property
    def Language(self):
        return self.businesscardview.Language

    @Language.setter
    def Language(self, value):
        self.businesscardview.Language = value

    @property
    def LockUserChanges(self):
        return self.businesscardview.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.businesscardview.LockUserChanges = value

    @property
    def Name(self):
        return self.businesscardview.Name

    @Name.setter
    def Name(self, value):
        self.businesscardview.Name = value

    @property
    def Parent(self):
        return self.businesscardview.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.businesscardview.SaveOption)

    @property
    def Session(self):
        return NameSpace(self.businesscardview.Session)

    @property
    def SortFields(self):
        return OrderFields(self.businesscardview.SortFields)

    @property
    def Standard(self):
        return BusinessCardView(self.businesscardview.Standard)

    @property
    def ViewType(self):
        return OlViewType(self.businesscardview.ViewType)

    @property
    def XML(self):
        return self.businesscardview.XML

    @XML.setter
    def XML(self, value):
        self.businesscardview.XML = value

    def Apply(self):
        self.businesscardview.Apply()

    def Copy(self, *args, Name=None, SaveOption=None):
        arguments = {"Name": Name, "SaveOption": SaveOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.businesscardview.Copy(*args, **arguments)

    def Delete(self):
        self.businesscardview.Delete()

    def GoToDate(self, *args, Date=None):
        arguments = {"Date": Date}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.businesscardview.GoToDate(*args, **arguments)

    def Reset(self):
        self.businesscardview.Reset()

    def Save(self):
        self.businesscardview.Save()


class CalendarModule:

    def __init__(self, calendarmodule=None):
        self.calendarmodule = calendarmodule

    @property
    def Application(self):
        return Application(self.calendarmodule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.calendarmodule.Class)

    @property
    def Name(self):
        return CalendarModule(self.calendarmodule.Name)

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.calendarmodule.NavigationGroups)

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.calendarmodule.NavigationModuleType)

    @property
    def Parent(self):
        return self.calendarmodule.Parent

    @property
    def Position(self):
        return CalendarModule(self.calendarmodule.Position)

    @Position.setter
    def Position(self, value):
        self.calendarmodule.Position = value

    @property
    def Session(self):
        return NameSpace(self.calendarmodule.Session)

    @property
    def Visible(self):
        return CalendarModule(self.calendarmodule.Visible)

    @Visible.setter
    def Visible(self, value):
        self.calendarmodule.Visible = value


class CalendarSharing:

    def __init__(self, calendarsharing=None):
        self.calendarsharing = calendarsharing

    @property
    def Application(self):
        return Application(self.calendarsharing.Application)

    @property
    def CalendarDetail(self):
        return OlCalendarDetail(self.calendarsharing.CalendarDetail)

    @CalendarDetail.setter
    def CalendarDetail(self, value):
        self.calendarsharing.CalendarDetail = value

    @property
    def Class(self):
        return OlObjectClass(self.calendarsharing.Class)

    @property
    def EndDate(self):
        return CalendarSharing(self.calendarsharing.EndDate)

    @EndDate.setter
    def EndDate(self, value):
        self.calendarsharing.EndDate = value

    @property
    def Folder(self):
        return Folder(self.calendarsharing.Folder)

    @property
    def IncludeAttachments(self):
        return self.calendarsharing.IncludeAttachments

    @IncludeAttachments.setter
    def IncludeAttachments(self, value):
        self.calendarsharing.IncludeAttachments = value

    @property
    def IncludePrivateDetails(self):
        return self.calendarsharing.IncludePrivateDetails

    @IncludePrivateDetails.setter
    def IncludePrivateDetails(self, value):
        self.calendarsharing.IncludePrivateDetails = value

    @property
    def IncludeWholeCalendar(self):
        return self.calendarsharing.IncludeWholeCalendar

    @IncludeWholeCalendar.setter
    def IncludeWholeCalendar(self, value):
        self.calendarsharing.IncludeWholeCalendar = value

    @property
    def Parent(self):
        return self.calendarsharing.Parent

    @property
    def RestrictToWorkingHours(self):
        return self.calendarsharing.RestrictToWorkingHours

    @RestrictToWorkingHours.setter
    def RestrictToWorkingHours(self, value):
        self.calendarsharing.RestrictToWorkingHours = value

    @property
    def Session(self):
        return NameSpace(self.calendarsharing.Session)

    @property
    def StartDate(self):
        return CalendarSharing(self.calendarsharing.StartDate)

    @StartDate.setter
    def StartDate(self, value):
        self.calendarsharing.StartDate = value

    def ForwardAsICal(self, *args, MailFormat=None):
        arguments = {"MailFormat": MailFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.calendarsharing.ForwardAsICal(*args, **arguments)

    def SaveAsICal(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.calendarsharing.SaveAsICal(*args, **arguments)


class CalendarView:

    def __init__(self, calendarview=None):
        self.calendarview = calendarview

    @property
    def Application(self):
        return Application(self.calendarview.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.calendarview.AutoFormatRules)

    @property
    def BoldDatesWithItems(self):
        return CalendarView(self.calendarview.BoldDatesWithItems)

    @BoldDatesWithItems.setter
    def BoldDatesWithItems(self, value):
        self.calendarview.BoldDatesWithItems = value

    @property
    def BoldSubjects(self):
        return CalendarView(self.calendarview.BoldSubjects)

    @BoldSubjects.setter
    def BoldSubjects(self, value):
        self.calendarview.BoldSubjects = value

    @property
    def CalendarViewMode(self):
        return OlCalendarViewMode(self.calendarview.CalendarViewMode)

    @CalendarViewMode.setter
    def CalendarViewMode(self, value):
        self.calendarview.CalendarViewMode = value

    @property
    def Class(self):
        return OlObjectClass(self.calendarview.Class)

    @property
    def DaysInMultiDayMode(self):
        return CalendarView(self.calendarview.DaysInMultiDayMode)

    @DaysInMultiDayMode.setter
    def DaysInMultiDayMode(self, value):
        self.calendarview.DaysInMultiDayMode = value

    @property
    def DayWeekTimeScale(self):
        return OlDayWeekTimeScale(self.calendarview.DayWeekTimeScale)

    @DayWeekTimeScale.setter
    def DayWeekTimeScale(self, value):
        self.calendarview.DayWeekTimeScale = value

    @property
    def DisplayedDates(self):
        return CalendarView(self.calendarview.DisplayedDates)

    @property
    def EndField(self):
        return CalendarView(self.calendarview.EndField)

    @EndField.setter
    def EndField(self, value):
        self.calendarview.EndField = value

    @property
    def Filter(self):
        return self.calendarview.Filter

    @Filter.setter
    def Filter(self, value):
        self.calendarview.Filter = value

    @property
    def Language(self):
        return self.calendarview.Language

    @Language.setter
    def Language(self, value):
        self.calendarview.Language = value

    @property
    def LockUserChanges(self):
        return self.calendarview.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.calendarview.LockUserChanges = value

    @property
    def MonthShowEndTime(self):
        return CalendarView(self.calendarview.MonthShowEndTime)

    @MonthShowEndTime.setter
    def MonthShowEndTime(self, value):
        self.calendarview.MonthShowEndTime = value

    @property
    def Name(self):
        return self.calendarview.Name

    @Name.setter
    def Name(self, value):
        self.calendarview.Name = value

    @property
    def Parent(self):
        return self.calendarview.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.calendarview.SaveOption)

    @property
    def SelectedEndTime(self):
        return CalendarView(self.calendarview.SelectedEndTime)

    @property
    def SelectedStartTime(self):
        return CalendarView(self.calendarview.SelectedStartTime)

    @property
    def Session(self):
        return NameSpace(self.calendarview.Session)

    @property
    def Standard(self):
        return CalendarView(self.calendarview.Standard)

    @property
    def StartField(self):
        return CalendarView(self.calendarview.StartField)

    @StartField.setter
    def StartField(self, value):
        self.calendarview.StartField = value

    @property
    def ViewType(self):
        return OlViewType(self.calendarview.ViewType)

    @property
    def XML(self):
        return self.calendarview.XML

    @XML.setter
    def XML(self, value):
        self.calendarview.XML = value

    def Apply(self):
        self.calendarview.Apply()

    def Copy(self, *args, Name=None, SaveOption=None):
        arguments = {"Name": Name, "SaveOption": SaveOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.calendarview.Copy(*args, **arguments)

    def Delete(self):
        self.calendarview.Delete()

    def GoToDate(self, *args, Date=None):
        arguments = {"Date": Date}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.calendarview.GoToDate(*args, **arguments)

    def Reset(self):
        self.calendarview.Reset()

    def Save(self):
        self.calendarview.Save()


class CardView:

    def __init__(self, cardview=None):
        self.cardview = cardview

    @property
    def AllowInCellEditing(self):
        return CardView(self.cardview.AllowInCellEditing)

    @AllowInCellEditing.setter
    def AllowInCellEditing(self, value):
        self.cardview.AllowInCellEditing = value

    @property
    def Application(self):
        return Application(self.cardview.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.cardview.AutoFormatRules)

    @property
    def BodyFont(self):
        return ViewFont(self.cardview.BodyFont)

    @property
    def Class(self):
        return OlObjectClass(self.cardview.Class)

    @property
    def Filter(self):
        return self.cardview.Filter

    @Filter.setter
    def Filter(self, value):
        self.cardview.Filter = value

    @property
    def HeadingsFont(self):
        return ViewFont(self.cardview.HeadingsFont)

    @property
    def Language(self):
        return self.cardview.Language

    @Language.setter
    def Language(self, value):
        self.cardview.Language = value

    @property
    def LockUserChanges(self):
        return self.cardview.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.cardview.LockUserChanges = value

    @property
    def MultiLineFieldHeight(self):
        return CardView(self.cardview.MultiLineFieldHeight)

    @MultiLineFieldHeight.setter
    def MultiLineFieldHeight(self, value):
        self.cardview.MultiLineFieldHeight = value

    @property
    def Name(self):
        return self.cardview.Name

    @Name.setter
    def Name(self, value):
        self.cardview.Name = value

    @property
    def Parent(self):
        return self.cardview.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.cardview.SaveOption)

    @property
    def Session(self):
        return NameSpace(self.cardview.Session)

    @property
    def ShowEmptyFields(self):
        return CardView(self.cardview.ShowEmptyFields)

    @ShowEmptyFields.setter
    def ShowEmptyFields(self, value):
        self.cardview.ShowEmptyFields = value

    @property
    def SortFields(self):
        return OrderFields(self.cardview.SortFields)

    @property
    def Standard(self):
        return CardView(self.cardview.Standard)

    @property
    def ViewFields(self):
        return ViewFields(self.cardview.ViewFields)

    @property
    def ViewType(self):
        return OlViewType(self.cardview.ViewType)

    @property
    def Width(self):
        return CardView(self.cardview.Width)

    @Width.setter
    def Width(self, value):
        self.cardview.Width = value

    @property
    def XML(self):
        return self.cardview.XML

    @XML.setter
    def XML(self, value):
        self.cardview.XML = value

    def Apply(self):
        self.cardview.Apply()

    def Copy(self, *args, Name=None, SaveOption=None):
        arguments = {"Name": Name, "SaveOption": SaveOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.cardview.Copy(*args, **arguments)

    def Delete(self):
        self.cardview.Delete()

    def GoToDate(self, *args, Date=None):
        arguments = {"Date": Date}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cardview.GoToDate(*args, **arguments)

    def Reset(self):
        self.cardview.Reset()

    def Save(self):
        self.cardview.Save()


class Categories:

    def __init__(self, categories=None):
        self.categories = categories

    @property
    def Application(self):
        return Application(self.categories.Application)

    @property
    def Class(self):
        return OlObjectClass(self.categories.Class)

    @property
    def Count(self):
        return Category(self.categories.Count)

    @property
    def Parent(self):
        return self.categories.Parent

    @property
    def Session(self):
        return NameSpace(self.categories.Session)

    def Add(self, *args, Name=None, Color=None, ShortcutKey=None):
        arguments = {"Name": Name, "Color": Color, "ShortcutKey": ShortcutKey}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.categories.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.categories.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.categories.Remove(*args, **arguments)


class Category:

    def __init__(self, category=None):
        self.category = category

    @property
    def Application(self):
        return Application(self.category.Application)

    @property
    def CategoryBorderColor(self):
        return Category(self.category.CategoryBorderColor)

    @property
    def CategoryGradientBottomColor(self):
        return Category(self.category.CategoryGradientBottomColor)

    @property
    def CategoryGradientTopColor(self):
        return Category(self.category.CategoryGradientTopColor)

    @property
    def CategoryID(self):
        return Category(self.category.CategoryID)

    @property
    def Class(self):
        return OlObjectClass(self.category.Class)

    @property
    def Color(self):
        return OlCategoryColor(self.category.Color)

    @Color.setter
    def Color(self, value):
        self.category.Color = value

    @property
    def Name(self):
        return self.category.Name

    @Name.setter
    def Name(self, value):
        self.category.Name = value

    @property
    def Parent(self):
        return self.category.Parent

    @property
    def Session(self):
        return NameSpace(self.category.Session)

    @property
    def ShortcutKey(self):
        return OlCategoryShortcutKey(self.category.ShortcutKey)

    @ShortcutKey.setter
    def ShortcutKey(self, value):
        self.category.ShortcutKey = value


class CategoryRuleCondition:

    def __init__(self, categoryrulecondition=None):
        self.categoryrulecondition = categoryrulecondition

    @property
    def Application(self):
        return Application(self.categoryrulecondition.Application)

    @property
    def Categories(self):
        return self.categoryrulecondition.Categories

    @Categories.setter
    def Categories(self, value):
        self.categoryrulecondition.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.categoryrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.categoryrulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.categoryrulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.categoryrulecondition.Enabled = value

    @property
    def Parent(self):
        return self.categoryrulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.categoryrulecondition.Session)


class Column:

    def __init__(self, column=None):
        self.column = column

    @property
    def Application(self):
        return Application(self.column.Application)

    @property
    def Class(self):
        return OlObjectClass(self.column.Class)

    @property
    def Name(self):
        return Column(self.column.Name)

    @property
    def Parent(self):
        return Column(self.column.Parent)

    @property
    def Session(self):
        return NameSpace(self.column.Session)


class ColumnFormat:

    def __init__(self, columnformat=None):
        self.columnformat = columnformat

    @property
    def Align(self):
        return OlAlign(self.columnformat.Align)

    @Align.setter
    def Align(self, value):
        self.columnformat.Align = value

    @property
    def Application(self):
        return Application(self.columnformat.Application)

    @property
    def Class(self):
        return OlObjectClass(self.columnformat.Class)

    @property
    def FieldFormat(self):
        return ColumnFormat(self.columnformat.FieldFormat)

    @FieldFormat.setter
    def FieldFormat(self, value):
        self.columnformat.FieldFormat = value

    @property
    def FieldType(self):
        return OlUserPropertyType(self.columnformat.FieldType)

    @property
    def Label(self):
        return ColumnFormat(self.columnformat.Label)

    @Label.setter
    def Label(self, value):
        self.columnformat.Label = value

    @property
    def Parent(self):
        return self.columnformat.Parent

    @property
    def Session(self):
        return NameSpace(self.columnformat.Session)

    @property
    def Width(self):
        return self.columnformat.Width

    @Width.setter
    def Width(self, value):
        self.columnformat.Width = value


class Columns:

    def __init__(self, columns=None):
        self.columns = columns

    @property
    def Application(self):
        return Application(self.columns.Application)

    @property
    def Class(self):
        return OlObjectClass(self.columns.Class)

    @property
    def Count(self):
        return Column(self.columns.Count)

    @property
    def Parent(self):
        return Columns(self.columns.Parent)

    @property
    def Session(self):
        return NameSpace(self.columns.Session)

    def Add(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.columns.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Table(self.columns.Item(*args, **arguments))

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.columns.Remove(*args, **arguments)

    def RemoveAll(self):
        self.columns.RemoveAll()


class Conflict:

    def __init__(self, conflict=None):
        self.conflict = conflict

    @property
    def Application(self):
        return Application(self.conflict.Application)

    @property
    def Class(self):
        return OlObjectClass(self.conflict.Class)

    @property
    def Item(self):
        return self.conflict.Item

    @property
    def Name(self):
        return self.conflict.Name

    @property
    def Parent(self):
        return self.conflict.Parent

    @property
    def Session(self):
        return NameSpace(self.conflict.Session)

    @property
    def Type(self):
        return OlObjectClass(self.conflict.Type)


class Conflicts:

    def __init__(self, conflicts=None):
        self.conflicts = conflicts

    @property
    def Application(self):
        return Application(self.conflicts.Application)

    @property
    def Class(self):
        return OlObjectClass(self.conflicts.Class)

    @property
    def Count(self):
        return self.conflicts.Count

    @property
    def Parent(self):
        return self.conflicts.Parent

    @property
    def Session(self):
        return NameSpace(self.conflicts.Session)

    def GetFirst(self):
        return Conflict(self.conflicts.GetFirst())

    def GetLast(self):
        return Conflict(self.conflicts.GetLast())

    def GetNext(self):
        return Conflict(self.conflicts.GetNext())

    def GetPrevious(self):
        return Conflict(self.conflicts.GetPrevious())

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.conflicts.Item(*args, **arguments)


class ContactItem:

    def __init__(self, contactitem=None):
        self.contactitem = contactitem

    @property
    def Account(self):
        return self.contactitem.Account

    @Account.setter
    def Account(self, value):
        self.contactitem.Account = value

    @property
    def Actions(self):
        return Actions(self.contactitem.Actions)

    @property
    def Anniversary(self):
        return self.contactitem.Anniversary

    @Anniversary.setter
    def Anniversary(self, value):
        self.contactitem.Anniversary = value

    @property
    def Application(self):
        return Application(self.contactitem.Application)

    @property
    def AssistantName(self):
        return self.contactitem.AssistantName

    @AssistantName.setter
    def AssistantName(self, value):
        self.contactitem.AssistantName = value

    @property
    def AssistantTelephoneNumber(self):
        return self.contactitem.AssistantTelephoneNumber

    @AssistantTelephoneNumber.setter
    def AssistantTelephoneNumber(self, value):
        self.contactitem.AssistantTelephoneNumber = value

    @property
    def Attachments(self):
        return Attachments(self.contactitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.contactitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.contactitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.contactitem.BillingInformation = value

    @property
    def Birthday(self):
        return self.contactitem.Birthday

    @Birthday.setter
    def Birthday(self, value):
        self.contactitem.Birthday = value

    @property
    def Body(self):
        return self.contactitem.Body

    @Body.setter
    def Body(self, value):
        self.contactitem.Body = value

    @property
    def Business2TelephoneNumber(self):
        return self.contactitem.Business2TelephoneNumber

    @Business2TelephoneNumber.setter
    def Business2TelephoneNumber(self, value):
        self.contactitem.Business2TelephoneNumber = value

    @property
    def BusinessAddress(self):
        return self.contactitem.BusinessAddress

    @BusinessAddress.setter
    def BusinessAddress(self, value):
        self.contactitem.BusinessAddress = value

    @property
    def BusinessAddressCity(self):
        return self.contactitem.BusinessAddressCity

    @BusinessAddressCity.setter
    def BusinessAddressCity(self, value):
        self.contactitem.BusinessAddressCity = value

    @property
    def BusinessAddressCountry(self):
        return self.contactitem.BusinessAddressCountry

    @BusinessAddressCountry.setter
    def BusinessAddressCountry(self, value):
        self.contactitem.BusinessAddressCountry = value

    @property
    def BusinessAddressPostalCode(self):
        return self.contactitem.BusinessAddressPostalCode

    @BusinessAddressPostalCode.setter
    def BusinessAddressPostalCode(self, value):
        self.contactitem.BusinessAddressPostalCode = value

    @property
    def BusinessAddressPostOfficeBox(self):
        return self.contactitem.BusinessAddressPostOfficeBox

    @BusinessAddressPostOfficeBox.setter
    def BusinessAddressPostOfficeBox(self, value):
        self.contactitem.BusinessAddressPostOfficeBox = value

    @property
    def BusinessAddressState(self):
        return self.contactitem.BusinessAddressState

    @BusinessAddressState.setter
    def BusinessAddressState(self, value):
        self.contactitem.BusinessAddressState = value

    @property
    def BusinessAddressStreet(self):
        return self.contactitem.BusinessAddressStreet

    @BusinessAddressStreet.setter
    def BusinessAddressStreet(self, value):
        self.contactitem.BusinessAddressStreet = value

    @property
    def BusinessCardLayoutXml(self):
        return self.contactitem.BusinessCardLayoutXml

    @BusinessCardLayoutXml.setter
    def BusinessCardLayoutXml(self, value):
        self.contactitem.BusinessCardLayoutXml = value

    @property
    def BusinessCardType(self):
        return OlBusinessCardType(self.contactitem.BusinessCardType)

    @property
    def BusinessFaxNumber(self):
        return self.contactitem.BusinessFaxNumber

    @BusinessFaxNumber.setter
    def BusinessFaxNumber(self, value):
        self.contactitem.BusinessFaxNumber = value

    @property
    def BusinessHomePage(self):
        return self.contactitem.BusinessHomePage

    @BusinessHomePage.setter
    def BusinessHomePage(self, value):
        self.contactitem.BusinessHomePage = value

    @property
    def BusinessTelephoneNumber(self):
        return self.contactitem.BusinessTelephoneNumber

    @BusinessTelephoneNumber.setter
    def BusinessTelephoneNumber(self, value):
        self.contactitem.BusinessTelephoneNumber = value

    @property
    def CallbackTelephoneNumber(self):
        return self.contactitem.CallbackTelephoneNumber

    @CallbackTelephoneNumber.setter
    def CallbackTelephoneNumber(self, value):
        self.contactitem.CallbackTelephoneNumber = value

    @property
    def CarTelephoneNumber(self):
        return self.contactitem.CarTelephoneNumber

    @CarTelephoneNumber.setter
    def CarTelephoneNumber(self, value):
        self.contactitem.CarTelephoneNumber = value

    @property
    def Categories(self):
        return self.contactitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.contactitem.Categories = value

    @property
    def Children(self):
        return self.contactitem.Children

    @Children.setter
    def Children(self, value):
        self.contactitem.Children = value

    @property
    def Class(self):
        return OlObjectClass(self.contactitem.Class)

    @property
    def Companies(self):
        return self.contactitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.contactitem.Companies = value

    @property
    def CompanyAndFullName(self):
        return self.contactitem.CompanyAndFullName

    @property
    def CompanyLastFirstNoSpace(self):
        return self.contactitem.CompanyLastFirstNoSpace

    @property
    def CompanyLastFirstSpaceOnly(self):
        return self.contactitem.CompanyLastFirstSpaceOnly

    @property
    def CompanyMainTelephoneNumber(self):
        return self.contactitem.CompanyMainTelephoneNumber

    @CompanyMainTelephoneNumber.setter
    def CompanyMainTelephoneNumber(self, value):
        self.contactitem.CompanyMainTelephoneNumber = value

    @property
    def CompanyName(self):
        return self.contactitem.CompanyName

    @CompanyName.setter
    def CompanyName(self, value):
        self.contactitem.CompanyName = value

    @property
    def ComputerNetworkName(self):
        return self.contactitem.ComputerNetworkName

    @ComputerNetworkName.setter
    def ComputerNetworkName(self, value):
        self.contactitem.ComputerNetworkName = value

    @property
    def Conflicts(self):
        return self.contactitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.contactitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.contactitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.contactitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.contactitem.CreationTime

    @property
    def CustomerID(self):
        return self.contactitem.CustomerID

    @CustomerID.setter
    def CustomerID(self, value):
        self.contactitem.CustomerID = value

    @property
    def Department(self):
        return self.contactitem.Department

    @Department.setter
    def Department(self, value):
        self.contactitem.Department = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.contactitem.DownloadState)

    @property
    def Email1Address(self):
        return self.contactitem.Email1Address

    @Email1Address.setter
    def Email1Address(self, value):
        self.contactitem.Email1Address = value

    @property
    def Email1AddressType(self):
        return self.contactitem.Email1AddressType

    @Email1AddressType.setter
    def Email1AddressType(self, value):
        self.contactitem.Email1AddressType = value

    @property
    def Email1DisplayName(self):
        return self.contactitem.Email1DisplayName

    @Email1DisplayName.setter
    def Email1DisplayName(self, value):
        self.contactitem.Email1DisplayName = value

    @property
    def Email1EntryID(self):
        return self.contactitem.Email1EntryID

    @property
    def Email2Address(self):
        return self.contactitem.Email2Address

    @Email2Address.setter
    def Email2Address(self, value):
        self.contactitem.Email2Address = value

    @property
    def Email2AddressType(self):
        return self.contactitem.Email2AddressType

    @Email2AddressType.setter
    def Email2AddressType(self, value):
        self.contactitem.Email2AddressType = value

    @property
    def Email2DisplayName(self):
        return self.contactitem.Email2DisplayName

    @Email2DisplayName.setter
    def Email2DisplayName(self, value):
        self.contactitem.Email2DisplayName = value

    @property
    def Email2EntryID(self):
        return self.contactitem.Email2EntryID

    @property
    def Email3Address(self):
        return self.contactitem.Email3Address

    @Email3Address.setter
    def Email3Address(self, value):
        self.contactitem.Email3Address = value

    @property
    def Email3AddressType(self):
        return self.contactitem.Email3AddressType

    @Email3AddressType.setter
    def Email3AddressType(self, value):
        self.contactitem.Email3AddressType = value

    @property
    def Email3DisplayName(self):
        return self.contactitem.Email3DisplayName

    @Email3DisplayName.setter
    def Email3DisplayName(self, value):
        self.contactitem.Email3DisplayName = value

    @property
    def Email3EntryID(self):
        return self.contactitem.Email3EntryID

    @property
    def EntryID(self):
        return self.contactitem.EntryID

    @property
    def FileAs(self):
        return self.contactitem.FileAs

    @FileAs.setter
    def FileAs(self, value):
        self.contactitem.FileAs = value

    @property
    def FirstName(self):
        return self.contactitem.FirstName

    @FirstName.setter
    def FirstName(self, value):
        self.contactitem.FirstName = value

    @property
    def FormDescription(self):
        return FormDescription(self.contactitem.FormDescription)

    @property
    def FTPSite(self):
        return self.contactitem.FTPSite

    @FTPSite.setter
    def FTPSite(self, value):
        self.contactitem.FTPSite = value

    @property
    def FullName(self):
        return self.contactitem.FullName

    @FullName.setter
    def FullName(self, value):
        self.contactitem.FullName = value

    @property
    def FullNameAndCompany(self):
        return self.contactitem.FullNameAndCompany

    @property
    def Gender(self):
        return OlGender(self.contactitem.Gender)

    @Gender.setter
    def Gender(self, value):
        self.contactitem.Gender = value

    @property
    def GetInspector(self):
        return Inspector(self.contactitem.GetInspector)

    @property
    def GovernmentIDNumber(self):
        return self.contactitem.GovernmentIDNumber

    @GovernmentIDNumber.setter
    def GovernmentIDNumber(self, value):
        self.contactitem.GovernmentIDNumber = value

    @property
    def HasPicture(self):
        return self.contactitem.HasPicture

    @property
    def Hobby(self):
        return self.contactitem.Hobby

    @Hobby.setter
    def Hobby(self, value):
        self.contactitem.Hobby = value

    @property
    def Home2TelephoneNumber(self):
        return self.contactitem.Home2TelephoneNumber

    @Home2TelephoneNumber.setter
    def Home2TelephoneNumber(self, value):
        self.contactitem.Home2TelephoneNumber = value

    @property
    def HomeAddress(self):
        return self.contactitem.HomeAddress

    @HomeAddress.setter
    def HomeAddress(self, value):
        self.contactitem.HomeAddress = value

    @property
    def HomeAddressCity(self):
        return self.contactitem.HomeAddressCity

    @HomeAddressCity.setter
    def HomeAddressCity(self, value):
        self.contactitem.HomeAddressCity = value

    @property
    def HomeAddressCountry(self):
        return self.contactitem.HomeAddressCountry

    @HomeAddressCountry.setter
    def HomeAddressCountry(self, value):
        self.contactitem.HomeAddressCountry = value

    @property
    def HomeAddressPostalCode(self):
        return self.contactitem.HomeAddressPostalCode

    @HomeAddressPostalCode.setter
    def HomeAddressPostalCode(self, value):
        self.contactitem.HomeAddressPostalCode = value

    @property
    def HomeAddressPostOfficeBox(self):
        return self.contactitem.HomeAddressPostOfficeBox

    @HomeAddressPostOfficeBox.setter
    def HomeAddressPostOfficeBox(self, value):
        self.contactitem.HomeAddressPostOfficeBox = value

    @property
    def HomeAddressState(self):
        return self.contactitem.HomeAddressState

    @HomeAddressState.setter
    def HomeAddressState(self, value):
        self.contactitem.HomeAddressState = value

    @property
    def HomeAddressStreet(self):
        return self.contactitem.HomeAddressStreet

    @HomeAddressStreet.setter
    def HomeAddressStreet(self, value):
        self.contactitem.HomeAddressStreet = value

    @property
    def HomeFaxNumber(self):
        return self.contactitem.HomeFaxNumber

    @HomeFaxNumber.setter
    def HomeFaxNumber(self, value):
        self.contactitem.HomeFaxNumber = value

    @property
    def HomeTelephoneNumber(self):
        return self.contactitem.HomeTelephoneNumber

    @HomeTelephoneNumber.setter
    def HomeTelephoneNumber(self, value):
        self.contactitem.HomeTelephoneNumber = value

    @property
    def IMAddress(self):
        return self.contactitem.IMAddress

    @IMAddress.setter
    def IMAddress(self, value):
        self.contactitem.IMAddress = value

    @property
    def Importance(self):
        return OlImportance(self.contactitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.contactitem.Importance = value

    @property
    def Initials(self):
        return self.contactitem.Initials

    @Initials.setter
    def Initials(self, value):
        self.contactitem.Initials = value

    @property
    def InternetFreeBusyAddress(self):
        return self.contactitem.InternetFreeBusyAddress

    @InternetFreeBusyAddress.setter
    def InternetFreeBusyAddress(self, value):
        self.contactitem.InternetFreeBusyAddress = value

    @property
    def IsConflict(self):
        return self.contactitem.IsConflict

    @property
    def ISDNNumber(self):
        return self.contactitem.ISDNNumber

    @ISDNNumber.setter
    def ISDNNumber(self, value):
        self.contactitem.ISDNNumber = value

    @property
    def IsMarkedAsTask(self):
        return ContactItem(self.contactitem.IsMarkedAsTask)

    @property
    def ItemProperties(self):
        return ItemProperties(self.contactitem.ItemProperties)

    @property
    def JobTitle(self):
        return self.contactitem.JobTitle

    @JobTitle.setter
    def JobTitle(self, value):
        self.contactitem.JobTitle = value

    @property
    def Journal(self):
        return self.contactitem.Journal

    @Journal.setter
    def Journal(self, value):
        self.contactitem.Journal = value

    @property
    def Language(self):
        return self.contactitem.Language

    @Language.setter
    def Language(self, value):
        self.contactitem.Language = value

    @property
    def LastFirstAndSuffix(self):
        return self.contactitem.LastFirstAndSuffix

    @property
    def LastFirstNoSpace(self):
        return self.contactitem.LastFirstNoSpace

    @property
    def LastFirstNoSpaceAndSuffix(self):
        return self.contactitem.LastFirstNoSpaceAndSuffix

    @property
    def LastFirstNoSpaceCompany(self):
        return self.contactitem.LastFirstNoSpaceCompany

    @property
    def LastFirstSpaceOnly(self):
        return self.contactitem.LastFirstSpaceOnly

    @property
    def LastFirstSpaceOnlyCompany(self):
        return self.contactitem.LastFirstSpaceOnlyCompany

    @property
    def LastModificationTime(self):
        return self.contactitem.LastModificationTime

    @property
    def LastName(self):
        return self.contactitem.LastName

    @LastName.setter
    def LastName(self, value):
        self.contactitem.LastName = value

    @property
    def LastNameAndFirstName(self):
        return self.contactitem.LastNameAndFirstName

    @property
    def MailingAddress(self):
        return self.contactitem.MailingAddress

    @MailingAddress.setter
    def MailingAddress(self, value):
        self.contactitem.MailingAddress = value

    @property
    def MailingAddressCity(self):
        return self.contactitem.MailingAddressCity

    @MailingAddressCity.setter
    def MailingAddressCity(self, value):
        self.contactitem.MailingAddressCity = value

    @property
    def MailingAddressCountry(self):
        return self.contactitem.MailingAddressCountry

    @MailingAddressCountry.setter
    def MailingAddressCountry(self, value):
        self.contactitem.MailingAddressCountry = value

    @property
    def MailingAddressPostalCode(self):
        return self.contactitem.MailingAddressPostalCode

    @MailingAddressPostalCode.setter
    def MailingAddressPostalCode(self, value):
        self.contactitem.MailingAddressPostalCode = value

    @property
    def MailingAddressPostOfficeBox(self):
        return self.contactitem.MailingAddressPostOfficeBox

    @MailingAddressPostOfficeBox.setter
    def MailingAddressPostOfficeBox(self, value):
        self.contactitem.MailingAddressPostOfficeBox = value

    @property
    def MailingAddressState(self):
        return self.contactitem.MailingAddressState

    @MailingAddressState.setter
    def MailingAddressState(self, value):
        self.contactitem.MailingAddressState = value

    @property
    def MailingAddressStreet(self):
        return self.contactitem.MailingAddressStreet

    @MailingAddressStreet.setter
    def MailingAddressStreet(self, value):
        self.contactitem.MailingAddressStreet = value

    @property
    def ManagerName(self):
        return self.contactitem.ManagerName

    @ManagerName.setter
    def ManagerName(self, value):
        self.contactitem.ManagerName = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.contactitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.contactitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.contactitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.contactitem.MessageClass = value

    @property
    def MiddleName(self):
        return self.contactitem.MiddleName

    @MiddleName.setter
    def MiddleName(self, value):
        self.contactitem.MiddleName = value

    @property
    def Mileage(self):
        return self.contactitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.contactitem.Mileage = value

    @property
    def MobileTelephoneNumber(self):
        return self.contactitem.MobileTelephoneNumber

    @MobileTelephoneNumber.setter
    def MobileTelephoneNumber(self, value):
        self.contactitem.MobileTelephoneNumber = value

    @property
    def NetMeetingAlias(self):
        return self.contactitem.NetMeetingAlias

    @NetMeetingAlias.setter
    def NetMeetingAlias(self, value):
        self.contactitem.NetMeetingAlias = value

    @property
    def NetMeetingServer(self):
        return self.contactitem.NetMeetingServer

    @NetMeetingServer.setter
    def NetMeetingServer(self, value):
        self.contactitem.NetMeetingServer = value

    @property
    def NickName(self):
        return self.contactitem.NickName

    @NickName.setter
    def NickName(self, value):
        self.contactitem.NickName = value

    @property
    def NoAging(self):
        return self.contactitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.contactitem.NoAging = value

    @property
    def OfficeLocation(self):
        return self.contactitem.OfficeLocation

    @OfficeLocation.setter
    def OfficeLocation(self, value):
        self.contactitem.OfficeLocation = value

    @property
    def OrganizationalIDNumber(self):
        return self.contactitem.OrganizationalIDNumber

    @OrganizationalIDNumber.setter
    def OrganizationalIDNumber(self, value):
        self.contactitem.OrganizationalIDNumber = value

    @property
    def OtherAddress(self):
        return self.contactitem.OtherAddress

    @OtherAddress.setter
    def OtherAddress(self, value):
        self.contactitem.OtherAddress = value

    @property
    def OtherAddressCity(self):
        return self.contactitem.OtherAddressCity

    @OtherAddressCity.setter
    def OtherAddressCity(self, value):
        self.contactitem.OtherAddressCity = value

    @property
    def OtherAddressCountry(self):
        return self.contactitem.OtherAddressCountry

    @OtherAddressCountry.setter
    def OtherAddressCountry(self, value):
        self.contactitem.OtherAddressCountry = value

    @property
    def OtherAddressPostalCode(self):
        return self.contactitem.OtherAddressPostalCode

    @OtherAddressPostalCode.setter
    def OtherAddressPostalCode(self, value):
        self.contactitem.OtherAddressPostalCode = value

    @property
    def OtherAddressPostOfficeBox(self):
        return self.contactitem.OtherAddressPostOfficeBox

    @OtherAddressPostOfficeBox.setter
    def OtherAddressPostOfficeBox(self, value):
        self.contactitem.OtherAddressPostOfficeBox = value

    @property
    def OtherAddressState(self):
        return self.contactitem.OtherAddressState

    @OtherAddressState.setter
    def OtherAddressState(self, value):
        self.contactitem.OtherAddressState = value

    @property
    def OtherAddressStreet(self):
        return self.contactitem.OtherAddressStreet

    @OtherAddressStreet.setter
    def OtherAddressStreet(self, value):
        self.contactitem.OtherAddressStreet = value

    @property
    def OtherFaxNumber(self):
        return self.contactitem.OtherFaxNumber

    @OtherFaxNumber.setter
    def OtherFaxNumber(self, value):
        self.contactitem.OtherFaxNumber = value

    @property
    def OtherTelephoneNumber(self):
        return self.contactitem.OtherTelephoneNumber

    @OtherTelephoneNumber.setter
    def OtherTelephoneNumber(self, value):
        self.contactitem.OtherTelephoneNumber = value

    @property
    def OutlookInternalVersion(self):
        return self.contactitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.contactitem.OutlookVersion

    @property
    def PagerNumber(self):
        return self.contactitem.PagerNumber

    @PagerNumber.setter
    def PagerNumber(self, value):
        self.contactitem.PagerNumber = value

    @property
    def Parent(self):
        return self.contactitem.Parent

    @property
    def PersonalHomePage(self):
        return self.contactitem.PersonalHomePage

    @PersonalHomePage.setter
    def PersonalHomePage(self, value):
        self.contactitem.PersonalHomePage = value

    @property
    def PrimaryTelephoneNumber(self):
        return self.contactitem.PrimaryTelephoneNumber

    @PrimaryTelephoneNumber.setter
    def PrimaryTelephoneNumber(self, value):
        self.contactitem.PrimaryTelephoneNumber = value

    @property
    def Profession(self):
        return self.contactitem.Profession

    @Profession.setter
    def Profession(self, value):
        self.contactitem.Profession = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.contactitem.PropertyAccessor)

    @property
    def RadioTelephoneNumber(self):
        return self.contactitem.RadioTelephoneNumber

    @RadioTelephoneNumber.setter
    def RadioTelephoneNumber(self, value):
        self.contactitem.RadioTelephoneNumber = value

    @property
    def ReferredBy(self):
        return self.contactitem.ReferredBy

    @ReferredBy.setter
    def ReferredBy(self, value):
        self.contactitem.ReferredBy = value

    @property
    def ReminderOverrideDefault(self):
        return self.contactitem.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.contactitem.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.contactitem.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.contactitem.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.contactitem.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.contactitem.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.contactitem.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.contactitem.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.contactitem.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.contactitem.ReminderTime = value

    @property
    def RTFBody(self):
        return self.contactitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.contactitem.RTFBody = value

    @property
    def Saved(self):
        return self.contactitem.Saved

    @property
    def SelectedMailingAddress(self):
        return OlMailingAddress(self.contactitem.SelectedMailingAddress)

    @SelectedMailingAddress.setter
    def SelectedMailingAddress(self, value):
        self.contactitem.SelectedMailingAddress = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.contactitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.contactitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.contactitem.Session)

    @property
    def Size(self):
        return self.contactitem.Size

    @property
    def Spouse(self):
        return self.contactitem.Spouse

    @Spouse.setter
    def Spouse(self, value):
        self.contactitem.Spouse = value

    @property
    def Subject(self):
        return self.contactitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.contactitem.Subject = value

    @property
    def Suffix(self):
        return self.contactitem.Suffix

    @Suffix.setter
    def Suffix(self, value):
        self.contactitem.Suffix = value

    @property
    def TaskCompletedDate(self):
        return ContactItem(self.contactitem.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.contactitem.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return ContactItem(self.contactitem.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.contactitem.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return ContactItem(self.contactitem.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.contactitem.TaskStartDate = value

    @property
    def TaskSubject(self):
        return ContactItem(self.contactitem.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.contactitem.TaskSubject = value

    @property
    def TelexNumber(self):
        return self.contactitem.TelexNumber

    @TelexNumber.setter
    def TelexNumber(self, value):
        self.contactitem.TelexNumber = value

    @property
    def Title(self):
        return self.contactitem.Title

    @Title.setter
    def Title(self, value):
        self.contactitem.Title = value

    @property
    def ToDoTaskOrdinal(self):
        return ContactItem(self.contactitem.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.contactitem.ToDoTaskOrdinal = value

    @property
    def TTYTDDTelephoneNumber(self):
        return self.contactitem.TTYTDDTelephoneNumber

    @TTYTDDTelephoneNumber.setter
    def TTYTDDTelephoneNumber(self, value):
        self.contactitem.TTYTDDTelephoneNumber = value

    @property
    def UnRead(self):
        return self.contactitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.contactitem.UnRead = value

    @property
    def User1(self):
        return self.contactitem.User1

    @User1.setter
    def User1(self, value):
        self.contactitem.User1 = value

    @property
    def User2(self):
        return self.contactitem.User2

    @User2.setter
    def User2(self, value):
        self.contactitem.User2 = value

    @property
    def User3(self):
        return self.contactitem.User3

    @User3.setter
    def User3(self, value):
        self.contactitem.User3 = value

    @property
    def User4(self):
        return self.contactitem.User4

    @User4.setter
    def User4(self, value):
        self.contactitem.User4 = value

    @property
    def UserProperties(self):
        return UserProperties(self.contactitem.UserProperties)

    @property
    def WebPage(self):
        return self.contactitem.WebPage

    @WebPage.setter
    def WebPage(self, value):
        self.contactitem.WebPage = value

    @property
    def YomiCompanyName(self):
        return self.contactitem.YomiCompanyName

    @YomiCompanyName.setter
    def YomiCompanyName(self, value):
        self.contactitem.YomiCompanyName = value

    @property
    def YomiFirstName(self):
        return self.contactitem.YomiFirstName

    @YomiFirstName.setter
    def YomiFirstName(self, value):
        self.contactitem.YomiFirstName = value

    @property
    def YomiLastName(self):
        return self.contactitem.YomiLastName

    @YomiLastName.setter
    def YomiLastName(self, value):
        self.contactitem.YomiLastName = value

    def AddBusinessCardLogoPicture(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contactitem.AddBusinessCardLogoPicture(*args, **arguments)

    def AddPicture(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contactitem.AddPicture(*args, **arguments)

    def ClearTaskFlag(self):
        self.contactitem.ClearTaskFlag()

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contactitem.Close(*args, **arguments)

    def Copy(self):
        self.contactitem.Copy()

    def Delete(self):
        self.contactitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contactitem.Display(*args, **arguments)

    def ForwardAsBusinessCard(self):
        return self.contactitem.ForwardAsBusinessCard()

    def ForwardAsVcard(self):
        return self.contactitem.ForwardAsVcard()

    def GetConversation(self):
        return self.contactitem.GetConversation()

    def MarkAsTask(self, *args, MarkInterval=None):
        arguments = {"MarkInterval": MarkInterval}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contactitem.MarkAsTask(*args, **arguments)

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.contactitem.Move(*args, **arguments)

    def PrintOut(self):
        self.contactitem.PrintOut()

    def RemovePicture(self):
        self.contactitem.RemovePicture()

    def ResetBusinessCard(self):
        self.contactitem.ResetBusinessCard()

    def Save(self):
        self.contactitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contactitem.SaveAs(*args, **arguments)

    def SaveBusinessCardImage(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contactitem.SaveBusinessCardImage(*args, **arguments)

    def ShowBusinessCardEditor(self):
        self.contactitem.ShowBusinessCardEditor()

    def ShowCategoriesDialog(self):
        self.contactitem.ShowCategoriesDialog()

    def ShowCheckPhoneDialog(self, *args, PhoneNumber=None):
        arguments = {"PhoneNumber": PhoneNumber}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contactitem.ShowCheckPhoneDialog(*args, **arguments)


class ContactsModule:

    def __init__(self, contactsmodule=None):
        self.contactsmodule = contactsmodule

    @property
    def Application(self):
        return Application(self.contactsmodule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.contactsmodule.Class)

    @property
    def Name(self):
        return ContactsModule(self.contactsmodule.Name)

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.contactsmodule.NavigationGroups)

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.contactsmodule.NavigationModuleType)

    @property
    def Parent(self):
        return self.contactsmodule.Parent

    @property
    def Position(self):
        return ContactsModule(self.contactsmodule.Position)

    @Position.setter
    def Position(self, value):
        self.contactsmodule.Position = value

    @property
    def Session(self):
        return NameSpace(self.contactsmodule.Session)

    @property
    def Visible(self):
        return ContactsModule(self.contactsmodule.Visible)

    @Visible.setter
    def Visible(self, value):
        self.contactsmodule.Visible = value


class Conversation:

    def __init__(self, conversation=None):
        self.conversation = conversation

    @property
    def Application(self):
        return Application(self.conversation.Application)

    @property
    def Class(self):
        return OlObjectClass(self.conversation.Class)

    @property
    def ConversationID(self):
        return Conversation(self.conversation.ConversationID)

    @property
    def Parent(self):
        return Conversation(self.conversation.Parent)

    @property
    def Session(self):
        return NameSpace(self.conversation.Session)

    def ClearAlwaysAssignCategories(self, *args, Store=None):
        arguments = {"Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.conversation.ClearAlwaysAssignCategories(*args, **arguments)

    def GetAlwaysAssignCategories(self, *args, Store=None):
        arguments = {"Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.conversation.GetAlwaysAssignCategories(*args, **arguments)

    def GetAlwaysDelete(self, *args, Store=None):
        arguments = {"Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.conversation.GetAlwaysDelete(*args, **arguments)

    def GetAlwaysMoveToFolder(self, *args, Store=None):
        arguments = {"Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.conversation.GetAlwaysMoveToFolder(*args, **arguments)

    def GetChildren(self, *args, Item=None):
        arguments = {"Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.conversation.GetChildren(*args, **arguments)

    def GetParent(self, *args, Item=None):
        arguments = {"Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.conversation.GetParent(*args, **arguments)

    def GetRootItems(self):
        return self.conversation.GetRootItems()

    def GetTable(self):
        return self.conversation.GetTable()

    def MarkAsRead(self):
        self.conversation.MarkAsRead()

    def MarkAsUnread(self):
        self.conversation.MarkAsUnread()

    def SetAlwaysAssignCategories(self, *args, Categories=None, Store=None):
        arguments = {"Categories": Categories, "Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.conversation.SetAlwaysAssignCategories(*args, **arguments)

    def SetAlwaysDelete(self, *args, AlwaysDelete=None, Store=None):
        arguments = {"AlwaysDelete": AlwaysDelete, "Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.conversation.SetAlwaysDelete(*args, **arguments)

    def SetAlwaysMoveToFolder(self, *args, MoveToFolder=None, Store=None):
        arguments = {"MoveToFolder": MoveToFolder, "Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.conversation.SetAlwaysMoveToFolder(*args, **arguments)

    def StopAlwaysDelete(self, *args, Store=None):
        arguments = {"Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.conversation.StopAlwaysDelete(*args, **arguments)

    def StopAlwaysMoveToFolder(self, *args, Store=None):
        arguments = {"Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.conversation.StopAlwaysMoveToFolder(*args, **arguments)


class ConversationHeader:

    def __init__(self, conversationheader=None):
        self.conversationheader = conversationheader

    @property
    def Application(self):
        return Application(self.conversationheader.Application)

    @property
    def Class(self):
        return OlObjectClass(self.conversationheader.Class)

    @property
    def ConversationID(self):
        return Conversation(self.conversationheader.ConversationID)

    @property
    def ConversationTopic(self):
        return self.conversationheader.ConversationTopic

    @property
    def Parent(self):
        return self.conversationheader.Parent

    @property
    def Session(self):
        return NameSpace(self.conversationheader.Session)

    def GetConversation(self):
        return self.conversationheader.GetConversation()

    def GetItems(self):
        return self.conversationheader.GetItems()


class DistListItem:

    def __init__(self, distlistitem=None):
        self.distlistitem = distlistitem

    @property
    def Actions(self):
        return Actions(self.distlistitem.Actions)

    @property
    def Application(self):
        return Application(self.distlistitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.distlistitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.distlistitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.distlistitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.distlistitem.BillingInformation = value

    @property
    def Body(self):
        return self.distlistitem.Body

    @Body.setter
    def Body(self, value):
        self.distlistitem.Body = value

    @property
    def Categories(self):
        return self.distlistitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.distlistitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.distlistitem.Class)

    @property
    def Companies(self):
        return self.distlistitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.distlistitem.Companies = value

    @property
    def Conflicts(self):
        return self.distlistitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.distlistitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.distlistitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.distlistitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.distlistitem.CreationTime

    @property
    def DLName(self):
        return self.distlistitem.DLName

    @DLName.setter
    def DLName(self, value):
        self.distlistitem.DLName = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.distlistitem.DownloadState)

    @property
    def EntryID(self):
        return self.distlistitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.distlistitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.distlistitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.distlistitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.distlistitem.Importance = value

    @property
    def IsConflict(self):
        return self.distlistitem.IsConflict

    @property
    def IsMarkedAsTask(self):
        return DistListItem(self.distlistitem.IsMarkedAsTask)

    @property
    def ItemProperties(self):
        return ItemProperties(self.distlistitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.distlistitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.distlistitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.distlistitem.MarkForDownload = value

    @property
    def MemberCount(self):
        return self.distlistitem.MemberCount

    @property
    def MessageClass(self):
        return self.distlistitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.distlistitem.MessageClass = value

    @property
    def Mileage(self):
        return self.distlistitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.distlistitem.Mileage = value

    @property
    def NoAging(self):
        return self.distlistitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.distlistitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.distlistitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.distlistitem.OutlookVersion

    @property
    def Parent(self):
        return self.distlistitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.distlistitem.PropertyAccessor)

    @property
    def ReminderOverrideDefault(self):
        return self.distlistitem.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.distlistitem.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.distlistitem.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.distlistitem.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.distlistitem.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.distlistitem.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.distlistitem.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.distlistitem.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.distlistitem.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.distlistitem.ReminderTime = value

    @property
    def RTFBody(self):
        return self.distlistitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.distlistitem.RTFBody = value

    @property
    def Saved(self):
        return self.distlistitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.distlistitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.distlistitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.distlistitem.Session)

    @property
    def Size(self):
        return self.distlistitem.Size

    @property
    def Subject(self):
        return self.distlistitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.distlistitem.Subject = value

    @property
    def TaskCompletedDate(self):
        return DistListItem(self.distlistitem.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.distlistitem.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return DistListItem(self.distlistitem.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.distlistitem.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return DistListItem(self.distlistitem.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.distlistitem.TaskStartDate = value

    @property
    def TaskSubject(self):
        return DistListItem(self.distlistitem.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.distlistitem.TaskSubject = value

    @property
    def ToDoTaskOrdinal(self):
        return DistListItem(self.distlistitem.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.distlistitem.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.distlistitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.distlistitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.distlistitem.UserProperties)

    def AddMember(self, *args, Recipient=None):
        arguments = {"Recipient": Recipient}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.distlistitem.AddMember(*args, **arguments)

    def AddMembers(self, *args, Recipients=None):
        arguments = {"Recipients": Recipients}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.distlistitem.AddMembers(*args, **arguments)

    def ClearTaskFlag(self):
        self.distlistitem.ClearTaskFlag()

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.distlistitem.Close(*args, **arguments)

    def Copy(self):
        self.distlistitem.Copy()

    def Delete(self):
        self.distlistitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.distlistitem.Display(*args, **arguments)

    def GetConversation(self):
        return self.distlistitem.GetConversation()

    def GetMember(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.distlistitem.GetMember(*args, **arguments)

    def MarkAsTask(self, *args, MarkInterval=None):
        arguments = {"MarkInterval": MarkInterval}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.distlistitem.MarkAsTask(*args, **arguments)

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.distlistitem.Move(*args, **arguments)

    def PrintOut(self):
        self.distlistitem.PrintOut()

    def RemoveMember(self, *args, Recipient=None):
        arguments = {"Recipient": Recipient}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.distlistitem.RemoveMember(*args, **arguments)

    def RemoveMembers(self, *args, Recipients=None):
        arguments = {"Recipients": Recipients}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.distlistitem.RemoveMembers(*args, **arguments)

    def Save(self):
        self.distlistitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.distlistitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.distlistitem.ShowCategoriesDialog()


class DocumentItem:

    def __init__(self, documentitem=None):
        self.documentitem = documentitem

    @property
    def Actions(self):
        return Actions(self.documentitem.Actions)

    @property
    def Application(self):
        return Application(self.documentitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.documentitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.documentitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.documentitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.documentitem.BillingInformation = value

    @property
    def Body(self):
        return self.documentitem.Body

    @property
    def Categories(self):
        return self.documentitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.documentitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.documentitem.Class)

    @property
    def Companies(self):
        return self.documentitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.documentitem.Companies = value

    @property
    def Conflicts(self):
        return self.documentitem.Conflicts

    @property
    def ConversationIndex(self):
        return self.documentitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.documentitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.documentitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.documentitem.DownloadState)

    @property
    def EntryID(self):
        return self.documentitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.documentitem.FormDescription)

    @property
    def GetInspector(self):
        return self.documentitem.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.documentitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.documentitem.Importance = value

    @property
    def IsConflict(self):
        return self.documentitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.documentitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.documentitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return self.documentitem.MarkForDownload

    @property
    def MessageClass(self):
        return self.documentitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.documentitem.MessageClass = value

    @property
    def Mileage(self):
        return self.documentitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.documentitem.Mileage = value

    @property
    def NoAging(self):
        return self.documentitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.documentitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.documentitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.documentitem.OutlookVersion

    @property
    def Parent(self):
        return self.documentitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.documentitem.PropertyAccessor)

    @property
    def Saved(self):
        return self.documentitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.documentitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.documentitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.documentitem.Session)

    @property
    def Size(self):
        return self.documentitem.Size

    @property
    def Subject(self):
        return self.documentitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.documentitem.Subject = value

    @property
    def UnRead(self):
        return self.documentitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.documentitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.documentitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.documentitem.Close(*args, **arguments)

    def Copy(self):
        self.documentitem.Copy()

    def Delete(self):
        self.documentitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.documentitem.Display(*args, **arguments)

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.documentitem.Move(*args, **arguments)

    def PrintOut(self):
        self.documentitem.PrintOut()

    def Save(self):
        self.documentitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.documentitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.documentitem.ShowCategoriesDialog()


class Exception:

    def __init__(self, exception=None):
        self.exception = exception

    @property
    def Application(self):
        return Application(self.exception.Application)

    @property
    def AppointmentItem(self):
        return AppointmentItem(self.exception.AppointmentItem)

    @property
    def Class(self):
        return OlObjectClass(self.exception.Class)

    @property
    def Deleted(self):
        return AppointmentItem(self.exception.Deleted)

    @property
    def OriginalDate(self):
        return AppointmentItem(self.exception.OriginalDate)

    @property
    def Parent(self):
        return self.exception.Parent

    @property
    def Session(self):
        return NameSpace(self.exception.Session)


class Exceptions:

    def __init__(self, exceptions=None):
        self.exceptions = exceptions

    @property
    def Application(self):
        return Application(self.exceptions.Application)

    @property
    def Class(self):
        return OlObjectClass(self.exceptions.Class)

    @property
    def Count(self):
        return self.exceptions.Count

    @property
    def Parent(self):
        return self.exceptions.Parent

    @property
    def Session(self):
        return NameSpace(self.exceptions.Session)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.exceptions.Item(*args, **arguments)


class ExchangeDistributionList:

    def __init__(self, exchangedistributionlist=None):
        self.exchangedistributionlist = exchangedistributionlist

    @property
    def Address(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Address)

    @Address.setter
    def Address(self, value):
        self.exchangedistributionlist.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.exchangedistributionlist.AddressEntryUserType)

    @property
    def Alias(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Alias)

    @property
    def Application(self):
        return Application(self.exchangedistributionlist.Application)

    @property
    def Class(self):
        return OlObjectClass(self.exchangedistributionlist.Class)

    @property
    def Comments(self):
        return self.exchangedistributionlist.Comments

    @Comments.setter
    def Comments(self, value):
        self.exchangedistributionlist.Comments = value

    @property
    def DisplayType(self):
        return OlDisplayType(self.exchangedistributionlist.DisplayType)

    @property
    def ID(self):
        return ExchangeDistributionList(self.exchangedistributionlist.ID)

    @property
    def Name(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Name)

    @Name.setter
    def Name(self, value):
        self.exchangedistributionlist.Name = value

    @property
    def Parent(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Parent)

    @property
    def PrimarySmtpAddress(self):
        return ExchangeDistributionList(self.exchangedistributionlist.PrimarySmtpAddress)

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.exchangedistributionlist.PropertyAccessor)

    @property
    def Session(self):
        return NameSpace(self.exchangedistributionlist.Session)

    @property
    def Type(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Type)

    @Type.setter
    def Type(self, value):
        self.exchangedistributionlist.Type = value

    def Delete(self):
        self.exchangedistributionlist.Delete()

    def Details(self, *args, HWnd=None):
        arguments = {"HWnd": HWnd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.exchangedistributionlist.Details(*args, **arguments)

    def GetContact(self):
        return self.exchangedistributionlist.GetContact()

    def GetExchangeDistributionList(self):
        return self.exchangedistributionlist.GetExchangeDistributionList()

    def GetExchangeDistributionListMembers(self):
        return AddressEntry(self.exchangedistributionlist.GetExchangeDistributionListMembers())

    def GetExchangeUser(self):
        return self.exchangedistributionlist.GetExchangeUser()

    def GetFreeBusy(self, *args, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = {"Start": Start, "MinPerChar": MinPerChar, "CompleteFormat": CompleteFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.exchangedistributionlist.GetFreeBusy(*args, **arguments)

    def GetMemberOfList(self):
        return self.exchangedistributionlist.GetMemberOfList()

    def GetOwners(self):
        return AddressEntry(self.exchangedistributionlist.GetOwners())

    def Update(self, *args, MakePermanent=None, Refresh=None):
        arguments = {"MakePermanent": MakePermanent, "Refresh": Refresh}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.exchangedistributionlist.Update(*args, **arguments)


class ExchangeUser:

    def __init__(self, exchangeuser=None):
        self.exchangeuser = exchangeuser

    @property
    def Address(self):
        return ExchangeUser(self.exchangeuser.Address)

    @Address.setter
    def Address(self, value):
        self.exchangeuser.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.exchangeuser.AddressEntryUserType)

    @property
    def Alias(self):
        return ExchangeUser(self.exchangeuser.Alias)

    @property
    def Application(self):
        return Application(self.exchangeuser.Application)

    @property
    def AssistantName(self):
        return ExchangeUser(self.exchangeuser.AssistantName)

    @AssistantName.setter
    def AssistantName(self, value):
        self.exchangeuser.AssistantName = value

    @property
    def BusinessTelephoneNumber(self):
        return ExchangeUser(self.exchangeuser.BusinessTelephoneNumber)

    @BusinessTelephoneNumber.setter
    def BusinessTelephoneNumber(self, value):
        self.exchangeuser.BusinessTelephoneNumber = value

    @property
    def City(self):
        return ExchangeUser(self.exchangeuser.City)

    @City.setter
    def City(self, value):
        self.exchangeuser.City = value

    @property
    def Class(self):
        return OlObjectClass(self.exchangeuser.Class)

    @property
    def Comments(self):
        return self.exchangeuser.Comments

    @Comments.setter
    def Comments(self, value):
        self.exchangeuser.Comments = value

    @property
    def CompanyName(self):
        return ExchangeUser(self.exchangeuser.CompanyName)

    @CompanyName.setter
    def CompanyName(self, value):
        self.exchangeuser.CompanyName = value

    @property
    def Department(self):
        return ExchangeUser(self.exchangeuser.Department)

    @Department.setter
    def Department(self, value):
        self.exchangeuser.Department = value

    @property
    def DisplayType(self):
        return OlDisplayType(self.exchangeuser.DisplayType)

    @property
    def FirstName(self):
        return ExchangeUser(self.exchangeuser.FirstName)

    @FirstName.setter
    def FirstName(self, value):
        self.exchangeuser.FirstName = value

    @property
    def ID(self):
        return ExchangeUser(self.exchangeuser.ID)

    @property
    def JobTitle(self):
        return ExchangeUser(self.exchangeuser.JobTitle)

    @JobTitle.setter
    def JobTitle(self, value):
        self.exchangeuser.JobTitle = value

    @property
    def LastName(self):
        return ExchangeUser(self.exchangeuser.LastName)

    @LastName.setter
    def LastName(self, value):
        self.exchangeuser.LastName = value

    @property
    def MobileTelephoneNumber(self):
        return ExchangeUser(self.exchangeuser.MobileTelephoneNumber)

    @MobileTelephoneNumber.setter
    def MobileTelephoneNumber(self, value):
        self.exchangeuser.MobileTelephoneNumber = value

    @property
    def Name(self):
        return ExchangeUser(self.exchangeuser.Name)

    @Name.setter
    def Name(self, value):
        self.exchangeuser.Name = value

    @property
    def OfficeLocation(self):
        return ExchangeUser(self.exchangeuser.OfficeLocation)

    @OfficeLocation.setter
    def OfficeLocation(self, value):
        self.exchangeuser.OfficeLocation = value

    @property
    def Parent(self):
        return ExchangeUser(self.exchangeuser.Parent)

    @property
    def PostalCode(self):
        return ExchangeUser(self.exchangeuser.PostalCode)

    @PostalCode.setter
    def PostalCode(self, value):
        self.exchangeuser.PostalCode = value

    @property
    def PrimarySmtpAddress(self):
        return ExchangeUser(self.exchangeuser.PrimarySmtpAddress)

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.exchangeuser.PropertyAccessor)

    @property
    def Session(self):
        return NameSpace(self.exchangeuser.Session)

    @property
    def StateOrProvince(self):
        return ExchangeUser(self.exchangeuser.StateOrProvince)

    @StateOrProvince.setter
    def StateOrProvince(self, value):
        self.exchangeuser.StateOrProvince = value

    @property
    def StreetAddress(self):
        return ExchangeUser(self.exchangeuser.StreetAddress)

    @StreetAddress.setter
    def StreetAddress(self, value):
        self.exchangeuser.StreetAddress = value

    @property
    def Type(self):
        return ExchangeUser(self.exchangeuser.Type)

    @Type.setter
    def Type(self, value):
        self.exchangeuser.Type = value

    @property
    def YomiCompanyName(self):
        return ExchangeUser(self.exchangeuser.YomiCompanyName)

    @YomiCompanyName.setter
    def YomiCompanyName(self, value):
        self.exchangeuser.YomiCompanyName = value

    @property
    def YomiDepartment(self):
        return ExchangeUser(self.exchangeuser.YomiDepartment)

    @YomiDepartment.setter
    def YomiDepartment(self, value):
        self.exchangeuser.YomiDepartment = value

    @property
    def YomiDisplayName(self):
        return ExchangeUser(self.exchangeuser.YomiDisplayName)

    @YomiDisplayName.setter
    def YomiDisplayName(self, value):
        self.exchangeuser.YomiDisplayName = value

    @property
    def YomiFirstName(self):
        return ExchangeUser(self.exchangeuser.YomiFirstName)

    @YomiFirstName.setter
    def YomiFirstName(self, value):
        self.exchangeuser.YomiFirstName = value

    @property
    def YomiLastName(self):
        return ExchangeUser(self.exchangeuser.YomiLastName)

    @YomiLastName.setter
    def YomiLastName(self, value):
        self.exchangeuser.YomiLastName = value

    def Delete(self):
        self.exchangeuser.Delete()

    def Details(self, *args, HWnd=None):
        arguments = {"HWnd": HWnd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.exchangeuser.Details(*args, **arguments)

    def GetContact(self):
        return self.exchangeuser.GetContact()

    def GetDirectReports(self):
        return AddressEntry(self.exchangeuser.GetDirectReports())

    def GetExchangeDistributionList(self):
        return self.exchangeuser.GetExchangeDistributionList()

    def GetExchangeUser(self):
        return self.exchangeuser.GetExchangeUser()

    def GetExchangeUserManager(self):
        return self.exchangeuser.GetExchangeUserManager()

    def GetFreeBusy(self, *args, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = {"Start": Start, "MinPerChar": MinPerChar, "CompleteFormat": CompleteFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.exchangeuser.GetFreeBusy(*args, **arguments)

    def GetMemberOfList(self):
        return ExchangeUser(self.exchangeuser.GetMemberOfList())

    def GetPicture(self):
        return self.exchangeuser.GetPicture()

    def Update(self, *args, MakePermanent=None, Refresh=None):
        arguments = {"MakePermanent": MakePermanent, "Refresh": Refresh}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.exchangeuser.Update(*args, **arguments)


class Explorer:

    def __init__(self, explorer=None):
        self.explorer = explorer

    @property
    def AccountSelector(self):
        return AccountSelector(self.explorer.AccountSelector)

    @property
    def Application(self):
        return Application(self.explorer.Application)

    @property
    def AttachmentSelection(self):
        return AttachmentSelection(self.explorer.AttachmentSelection)

    @property
    def Caption(self):
        return self.explorer.Caption

    @property
    def Class(self):
        return OlObjectClass(self.explorer.Class)

    @property
    def CurrentFolder(self):
        return Folder(self.explorer.CurrentFolder)

    @CurrentFolder.setter
    def CurrentFolder(self, value):
        self.explorer.CurrentFolder = value

    @property
    def CurrentView(self):
        return self.explorer.CurrentView

    @CurrentView.setter
    def CurrentView(self, value):
        self.explorer.CurrentView = value

    @property
    def Height(self):
        return self.explorer.Height

    @Height.setter
    def Height(self, value):
        self.explorer.Height = value

    @property
    def HTMLDocument(self):
        return self.explorer.HTMLDocument

    @property
    def Left(self):
        return self.explorer.Left

    @Left.setter
    def Left(self, value):
        self.explorer.Left = value

    @property
    def NavigationPane(self):
        return NavigationPane(self.explorer.NavigationPane)

    @property
    def Panes(self):
        return Panes(self.explorer.Panes)

    @property
    def Parent(self):
        return self.explorer.Parent

    @property
    def Selection(self):
        return Selection(self.explorer.Selection)

    @property
    def Session(self):
        return NameSpace(self.explorer.Session)

    @property
    def Top(self):
        return self.explorer.Top

    @Top.setter
    def Top(self, value):
        self.explorer.Top = value

    @property
    def Width(self):
        return self.explorer.Width

    @Width.setter
    def Width(self, value):
        self.explorer.Width = value

    @property
    def WindowState(self):
        return OlWindowState(self.explorer.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.explorer.WindowState = value

    def Activate(self):
        self.explorer.Activate()

    def AddToSelection(self, *args, Item=None):
        arguments = {"Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.explorer.AddToSelection(*args, **arguments)

    def ClearSearch(self):
        self.explorer.ClearSearch()

    def ClearSelection(self):
        self.explorer.ClearSelection()

    def Close(self):
        self.explorer.Close()

    def Display(self):
        self.explorer.Display()

    def IsItemSelectableInView(self, *args, Item=None):
        arguments = {"Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.explorer.IsItemSelectableInView(*args, **arguments)

    def IsPaneVisible(self, *args, Pane=None):
        arguments = {"Pane": Pane}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.explorer.IsPaneVisible(*args, **arguments)

    def RemoveFromSelection(self, *args, Item=None):
        arguments = {"Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.explorer.RemoveFromSelection(*args, **arguments)

    def Search(self, *args, Query=None, SearchScope=None):
        arguments = {"Query": Query, "SearchScope": SearchScope}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.explorer.Search(*args, **arguments)

    def SelectAllItems(self):
        self.explorer.SelectAllItems()

    def ShowPane(self, *args, Pane=None, Visible=None):
        arguments = {"Pane": Pane, "Visible": Visible}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.explorer.ShowPane(*args, **arguments)


class Explorers:

    def __init__(self, explorers=None):
        self.explorers = explorers

    @property
    def Application(self):
        return Application(self.explorers.Application)

    @property
    def Class(self):
        return OlObjectClass(self.explorers.Class)

    @property
    def Count(self):
        return self.explorers.Count

    @property
    def Parent(self):
        return self.explorers.Parent

    @property
    def Session(self):
        return NameSpace(self.explorers.Session)

    def Add(self, *args, Folder=None, DisplayMode=None):
        arguments = {"Folder": Folder, "DisplayMode": DisplayMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Explorer(self.explorers.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.explorers.Item(*args, **arguments)


class Folder:

    def __init__(self, folder=None):
        self.folder = folder

    @property
    def AddressBookName(self):
        return Folder(self.folder.AddressBookName)

    @AddressBookName.setter
    def AddressBookName(self, value):
        self.folder.AddressBookName = value

    @property
    def Application(self):
        return Application(self.folder.Application)

    @property
    def Class(self):
        return OlObjectClass(self.folder.Class)

    @property
    def CurrentView(self):
        return View(self.folder.CurrentView)

    @property
    def CustomViewsOnly(self):
        return self.folder.CustomViewsOnly

    @CustomViewsOnly.setter
    def CustomViewsOnly(self, value):
        self.folder.CustomViewsOnly = value

    @property
    def DefaultItemType(self):
        return OlItemType(self.folder.DefaultItemType)

    @property
    def DefaultMessageClass(self):
        return self.folder.DefaultMessageClass

    @property
    def Description(self):
        return self.folder.Description

    @Description.setter
    def Description(self, value):
        self.folder.Description = value

    @property
    def EntryID(self):
        return self.folder.EntryID

    @property
    def FolderPath(self):
        return self.folder.FolderPath

    @property
    def Folders(self):
        return Folders(self.folder.Folders)

    @property
    def InAppFolderSyncObject(self):
        return self.folder.InAppFolderSyncObject

    @InAppFolderSyncObject.setter
    def InAppFolderSyncObject(self, value):
        self.folder.InAppFolderSyncObject = value

    @property
    def IsSharePointFolder(self):
        return self.folder.IsSharePointFolder

    @property
    def Items(self):
        return Items(self.folder.Items)

    @property
    def Name(self):
        return self.folder.Name

    @Name.setter
    def Name(self, value):
        self.folder.Name = value

    @property
    def Parent(self):
        return self.folder.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.folder.PropertyAccessor)

    @property
    def Session(self):
        return NameSpace(self.folder.Session)

    @property
    def ShowAsOutlookAB(self):
        return self.folder.ShowAsOutlookAB

    @ShowAsOutlookAB.setter
    def ShowAsOutlookAB(self, value):
        self.folder.ShowAsOutlookAB = value

    @property
    def ShowItemCount(self):
        return self.folder.ShowItemCount

    @ShowItemCount.setter
    def ShowItemCount(self, value):
        self.folder.ShowItemCount = value

    @property
    def Store(self):
        return Store(self.folder.Store)

    @property
    def StoreID(self):
        return self.folder.StoreID

    @property
    def UnReadItemCount(self):
        return self.folder.UnReadItemCount

    @property
    def UserDefinedProperties(self):
        return UserDefinedProperties(self.folder.UserDefinedProperties)

    @property
    def Views(self):
        return Views(self.folder.Views)

    @property
    def WebViewOn(self):
        return self.folder.WebViewOn

    @WebViewOn.setter
    def WebViewOn(self, value):
        self.folder.WebViewOn = value

    @property
    def WebViewURL(self):
        return self.folder.WebViewURL

    @WebViewURL.setter
    def WebViewURL(self, value):
        self.folder.WebViewURL = value

    def AddToPFFavorites(self):
        self.folder.AddToPFFavorites()

    def CopyTo(self, *args, DestinationFolder=None):
        arguments = {"DestinationFolder": DestinationFolder}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.folder.CopyTo(*args, **arguments)

    def Delete(self):
        self.folder.Delete()

    def Display(self):
        self.folder.Display()

    def GetCalendarExporter(self):
        return self.folder.GetCalendarExporter()

    def GetCustomIcon(self):
        return self.folder.GetCustomIcon()

    def GetExplorer(self, *args, DisplayMode=None):
        arguments = {"DisplayMode": DisplayMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.folder.GetExplorer(*args, **arguments)

    def GetOrganizer(self):
        return self.folder.GetOrganizer()

    def GetStorage(self, *args, StorageIdentifier=None, StorageIdentifierType=None):
        arguments = {"StorageIdentifier": StorageIdentifier, "StorageIdentifierType": StorageIdentifierType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.folder.GetStorage(*args, **arguments)

    def GetTable(self, *args, Filter=None, TableContents=None):
        arguments = {"Filter": Filter, "TableContents": TableContents}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Folder(self.folder.GetTable(*args, **arguments))

    def MoveTo(self, *args, DestinationFolder=None):
        arguments = {"DestinationFolder": DestinationFolder}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.folder.MoveTo(*args, **arguments)

    def SetCustomIcon(self, *args, Picture=None):
        arguments = {"Picture": Picture}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.folder.SetCustomIcon(*args, **arguments)


class Folders:

    def __init__(self, folders=None):
        self.folders = folders

    @property
    def Application(self):
        return Application(self.folders.Application)

    @property
    def Class(self):
        return OlObjectClass(self.folders.Class)

    @property
    def Count(self):
        return self.folders.Count

    @property
    def Parent(self):
        return self.folders.Parent

    @property
    def Session(self):
        return NameSpace(self.folders.Session)

    def Add(self, *args, Name=None, Type=None):
        arguments = {"Name": Name, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Folder(self.folders.Add(*args, **arguments))

    def GetFirst(self):
        return Folder(self.folders.GetFirst())

    def GetLast(self):
        return Folder(self.folders.GetLast())

    def GetNext(self):
        return Folder(self.folders.GetNext())

    def GetPrevious(self):
        return Folder(self.folders.GetPrevious())

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.folders.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.folders.Remove(*args, **arguments)


class FormDescription:

    def __init__(self, formdescription=None):
        self.formdescription = formdescription

    @property
    def Application(self):
        return Application(self.formdescription.Application)

    @property
    def Category(self):
        return self.formdescription.Category

    @Category.setter
    def Category(self, value):
        self.formdescription.Category = value

    @property
    def CategorySub(self):
        return self.formdescription.CategorySub

    @CategorySub.setter
    def CategorySub(self, value):
        self.formdescription.CategorySub = value

    @property
    def Class(self):
        return OlObjectClass(self.formdescription.Class)

    @property
    def Comment(self):
        return self.formdescription.Comment

    @Comment.setter
    def Comment(self, value):
        self.formdescription.Comment = value

    @property
    def ContactName(self):
        return FormDescription(self.formdescription.ContactName)

    @ContactName.setter
    def ContactName(self, value):
        self.formdescription.ContactName = value

    @property
    def DisplayName(self):
        return self.formdescription.DisplayName

    @DisplayName.setter
    def DisplayName(self, value):
        self.formdescription.DisplayName = value

    @property
    def Hidden(self):
        return self.formdescription.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.formdescription.Hidden = value

    @property
    def Icon(self):
        return self.formdescription.Icon

    @Icon.setter
    def Icon(self, value):
        self.formdescription.Icon = value

    @property
    def Locked(self):
        return self.formdescription.Locked

    @Locked.setter
    def Locked(self, value):
        self.formdescription.Locked = value

    @property
    def MessageClass(self):
        return FormDescription(self.formdescription.MessageClass)

    @property
    def MiniIcon(self):
        return self.formdescription.MiniIcon

    @MiniIcon.setter
    def MiniIcon(self, value):
        self.formdescription.MiniIcon = value

    @property
    def Name(self):
        return self.formdescription.Name

    @Name.setter
    def Name(self, value):
        self.formdescription.Name = value

    @property
    def Number(self):
        return self.formdescription.Number

    @Number.setter
    def Number(self, value):
        self.formdescription.Number = value

    @property
    def OneOff(self):
        return self.formdescription.OneOff

    @OneOff.setter
    def OneOff(self, value):
        self.formdescription.OneOff = value

    @property
    def Parent(self):
        return self.formdescription.Parent

    @property
    def ScriptText(self):
        return self.formdescription.ScriptText

    @property
    def Session(self):
        return NameSpace(self.formdescription.Session)

    @property
    def Template(self):
        return self.formdescription.Template

    @Template.setter
    def Template(self, value):
        self.formdescription.Template = value

    @property
    def UseWordMail(self):
        return self.formdescription.UseWordMail

    @UseWordMail.setter
    def UseWordMail(self, value):
        self.formdescription.UseWordMail = value

    @property
    def Version(self):
        return self.formdescription.Version

    @Version.setter
    def Version(self, value):
        self.formdescription.Version = value

    def PublishForm(self, *args, Registry=None, Folder=None):
        arguments = {"Registry": Registry, "Folder": Folder}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.formdescription.PublishForm(*args, **arguments)


class FormNameRuleCondition:

    def __init__(self, formnamerulecondition=None):
        self.formnamerulecondition = formnamerulecondition

    @property
    def Application(self):
        return Application(self.formnamerulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.formnamerulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.formnamerulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.formnamerulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.formnamerulecondition.Enabled = value

    @property
    def FormName(self):
        return self.formnamerulecondition.FormName

    @FormName.setter
    def FormName(self, value):
        self.formnamerulecondition.FormName = value

    @property
    def Parent(self):
        return self.formnamerulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.formnamerulecondition.Session)


class FormRegion:

    def __init__(self, formregion=None):
        self.formregion = formregion

    @property
    def Application(self):
        return Application(self.formregion.Application)

    @property
    def Class(self):
        return OlObjectClass(self.formregion.Class)

    @property
    def Detail(self):
        return self.formregion.Detail

    @Detail.setter
    def Detail(self, value):
        self.formregion.Detail = value

    @property
    def DisplayName(self):
        return self.formregion.DisplayName

    @property
    def EnableAutoLayout(self):
        return self.formregion.EnableAutoLayout

    @EnableAutoLayout.setter
    def EnableAutoLayout(self, value):
        self.formregion.EnableAutoLayout = value

    @property
    def Form(self):
        return self.formregion.Form

    @property
    def FormRegionMode(self):
        return OlFormRegionMode(self.formregion.FormRegionMode)

    @property
    def Inspector(self):
        return Inspector(self.formregion.Inspector)

    @property
    def InternalName(self):
        return self.formregion.InternalName

    @property
    def IsExpanded(self):
        return self.formregion.IsExpanded

    @property
    def Item(self):
        return self.formregion.Item

    @property
    def Language(self):
        return self.formregion.Language

    @property
    def Parent(self):
        return self.formregion.Parent

    @property
    def Session(self):
        return NameSpace(self.formregion.Session)

    @property
    def SuppressControlReplacement(self):
        return self.formregion.SuppressControlReplacement

    @SuppressControlReplacement.setter
    def SuppressControlReplacement(self, value):
        self.formregion.SuppressControlReplacement = value

    @property
    def Visible(self):
        return self.formregion.Visible

    @Visible.setter
    def Visible(self, value):
        self.formregion.Visible = value

    def Reflow(self):
        self.formregion.Reflow()

    def Select(self):
        self.formregion.Select()

    def SetControlItemProperty(self, *args, Control=None, PropertyName=None):
        arguments = {"Control": Control, "PropertyName": PropertyName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.formregion.SetControlItemProperty(*args, **arguments)


class FromRssFeedRuleCondition:

    def __init__(self, fromrssfeedrulecondition=None):
        self.fromrssfeedrulecondition = fromrssfeedrulecondition

    @property
    def Application(self):
        return Application(self.fromrssfeedrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.fromrssfeedrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.fromrssfeedrulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.fromrssfeedrulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.fromrssfeedrulecondition.Enabled = value

    @property
    def FromRssFeed(self):
        return self.fromrssfeedrulecondition.FromRssFeed

    @FromRssFeed.setter
    def FromRssFeed(self, value):
        self.fromrssfeedrulecondition.FromRssFeed = value

    @property
    def Parent(self):
        return self.fromrssfeedrulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.fromrssfeedrulecondition.Session)


class IconView:

    def __init__(self, iconview=None):
        self.iconview = iconview

    @property
    def Application(self):
        return Application(self.iconview.Application)

    @property
    def Class(self):
        return OlObjectClass(self.iconview.Class)

    @property
    def Filter(self):
        return self.iconview.Filter

    @Filter.setter
    def Filter(self, value):
        self.iconview.Filter = value

    @property
    def IconPlacement(self):
        return OlIconViewPlacement(self.iconview.IconPlacement)

    @IconPlacement.setter
    def IconPlacement(self, value):
        self.iconview.IconPlacement = value

    @property
    def IconViewType(self):
        return OlIconViewType(self.iconview.IconViewType)

    @IconViewType.setter
    def IconViewType(self, value):
        self.iconview.IconViewType = value

    @property
    def Language(self):
        return self.iconview.Language

    @Language.setter
    def Language(self, value):
        self.iconview.Language = value

    @property
    def LockUserChanges(self):
        return self.iconview.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.iconview.LockUserChanges = value

    @property
    def Name(self):
        return self.iconview.Name

    @Name.setter
    def Name(self, value):
        self.iconview.Name = value

    @property
    def Parent(self):
        return self.iconview.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.iconview.SaveOption)

    @property
    def Session(self):
        return NameSpace(self.iconview.Session)

    @property
    def SortFields(self):
        return OrderFields(self.iconview.SortFields)

    @property
    def Standard(self):
        return IconView(self.iconview.Standard)

    @property
    def ViewType(self):
        return OlViewType(self.iconview.ViewType)

    @property
    def XML(self):
        return self.iconview.XML

    @XML.setter
    def XML(self, value):
        self.iconview.XML = value

    def Apply(self):
        self.iconview.Apply()

    def Copy(self, *args, Name=None, SaveOption=None):
        arguments = {"Name": Name, "SaveOption": SaveOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.iconview.Copy(*args, **arguments)

    def Delete(self):
        self.iconview.Delete()

    def GoToDate(self, *args, Date=None):
        arguments = {"Date": Date}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.iconview.GoToDate(*args, **arguments)

    def Reset(self):
        self.iconview.Reset()

    def Save(self):
        self.iconview.Save()


class ImportanceRuleCondition:

    def __init__(self, importancerulecondition=None):
        self.importancerulecondition = importancerulecondition

    @property
    def Application(self):
        return Application(self.importancerulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.importancerulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.importancerulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.importancerulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.importancerulecondition.Enabled = value

    @property
    def Importance(self):
        return OlImportance(self.importancerulecondition.Importance)

    @Importance.setter
    def Importance(self, value):
        self.importancerulecondition.Importance = value

    @property
    def Parent(self):
        return self.importancerulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.importancerulecondition.Session)


class Inspector:

    def __init__(self, inspector=None):
        self.inspector = inspector

    @property
    def Application(self):
        return Application(self.inspector.Application)

    @property
    def AttachmentSelection(self):
        return AttachmentSelection(self.inspector.AttachmentSelection)

    @property
    def Caption(self):
        return self.inspector.Caption

    @property
    def Class(self):
        return OlObjectClass(self.inspector.Class)

    @property
    def CurrentItem(self):
        return self.inspector.CurrentItem

    @property
    def EditorType(self):
        return OlEditorType(self.inspector.EditorType)

    @property
    def Height(self):
        return self.inspector.Height

    @Height.setter
    def Height(self, value):
        self.inspector.Height = value

    @property
    def Left(self):
        return self.inspector.Left

    @Left.setter
    def Left(self, value):
        self.inspector.Left = value

    @property
    def ModifiedFormPages(self):
        return Pages(self.inspector.ModifiedFormPages)

    @property
    def Parent(self):
        return self.inspector.Parent

    @property
    def Session(self):
        return NameSpace(self.inspector.Session)

    @property
    def Top(self):
        return self.inspector.Top

    @Top.setter
    def Top(self, value):
        self.inspector.Top = value

    @property
    def Width(self):
        return self.inspector.Width

    @Width.setter
    def Width(self, value):
        self.inspector.Width = value

    @property
    def WindowState(self):
        return OlWindowState(self.inspector.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.inspector.WindowState = value

    @property
    def WordEditor(self):
        return self.inspector.WordEditor

    def Activate(self):
        self.inspector.Activate()

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.inspector.Close(*args, **arguments)

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.inspector.Display(*args, **arguments)

    def HideFormPage(self, *args, PageName=None):
        arguments = {"PageName": PageName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.inspector.HideFormPage(*args, **arguments)

    def IsWordMail(self):
        return self.inspector.IsWordMail()

    def NewFormRegion(self):
        return self.inspector.NewFormRegion()

    def OpenFormRegion(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.inspector.OpenFormRegion(*args, **arguments)

    def SaveFormRegion(self, *args, Page=None, FileName=None):
        arguments = {"Page": Page, "FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.inspector.SaveFormRegion(*args, **arguments)

    def SetControlItemProperty(self, *args, Control=None, PropertyName=None):
        arguments = {"Control": Control, "PropertyName": PropertyName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.inspector.SetControlItemProperty(*args, **arguments)

    def SetCurrentFormPage(self, *args, PageName=None):
        arguments = {"PageName": PageName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.inspector.SetCurrentFormPage(*args, **arguments)

    def SetSchedulingStartTime(self, *args, Start=None):
        arguments = {"Start": Start}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.inspector.SetSchedulingStartTime(*args, **arguments)

    def ShowFormPage(self, *args, PageName=None):
        arguments = {"PageName": PageName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.inspector.ShowFormPage(*args, **arguments)


class Inspectors:

    def __init__(self, inspectors=None):
        self.inspectors = inspectors

    @property
    def Application(self):
        return Application(self.inspectors.Application)

    @property
    def Class(self):
        return OlObjectClass(self.inspectors.Class)

    @property
    def Count(self):
        return self.inspectors.Count

    @property
    def Parent(self):
        return self.inspectors.Parent

    @property
    def Session(self):
        return NameSpace(self.inspectors.Session)

    def Add(self):
        return Inspector(self.inspectors.Add())

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.inspectors.Item(*args, **arguments)


class ItemProperties:

    def __init__(self, itemproperties=None):
        self.itemproperties = itemproperties

    def __call__(self, item):
        return ItemPropertie(self.itemproperties(item))

    @property
    def Application(self):
        return Application(self.itemproperties.Application)

    @property
    def Class(self):
        return OlObjectClass(self.itemproperties.Class)

    @property
    def Count(self):
        return self.itemproperties.Count

    @property
    def Parent(self):
        return self.itemproperties.Parent

    @property
    def Session(self):
        return NameSpace(self.itemproperties.Session)

    def Add(self, *args, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = {"Name": Name, "Type": Type, "AddToFolderFields": AddToFolderFields, "DisplayFormat": DisplayFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.itemproperties.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.itemproperties.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.itemproperties.Remove(*args, **arguments)


class ItemProperty:

    def __init__(self, itemproperty=None):
        self.itemproperty = itemproperty

    @property
    def Application(self):
        return Application(self.itemproperty.Application)

    @property
    def Class(self):
        return OlObjectClass(self.itemproperty.Class)

    @property
    def IsUserProperty(self):
        return self.itemproperty.IsUserProperty

    @property
    def Name(self):
        return self.itemproperty.Name

    @Name.setter
    def Name(self, value):
        self.itemproperty.Name = value

    @property
    def Parent(self):
        return self.itemproperty.Parent

    @property
    def Session(self):
        return NameSpace(self.itemproperty.Session)

    @property
    def Type(self):
        return OlUserPropertyType(self.itemproperty.Type)

    @property
    def Value(self):
        return self.itemproperty.Value

    @Value.setter
    def Value(self, value):
        self.itemproperty.Value = value


class Items:

    def __init__(self, items=None):
        self.items = items

    @property
    def Application(self):
        return Application(self.items.Application)

    @property
    def Class(self):
        return OlObjectClass(self.items.Class)

    @property
    def Count(self):
        return self.items.Count

    @property
    def IncludeRecurrences(self):
        return Items(self.items.IncludeRecurrences)

    @IncludeRecurrences.setter
    def IncludeRecurrences(self, value):
        self.items.IncludeRecurrences = value

    @property
    def Parent(self):
        return self.items.Parent

    @property
    def Session(self):
        return NameSpace(self.items.Session)

    def Add(self):
        return self.items.Add()

    def Find(self, *args, Filter=None):
        arguments = {"Filter": Filter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.items.Find(*args, **arguments)

    def FindNext(self):
        return self.items.FindNext()

    def GetFirst(self):
        return self.items.GetFirst()

    def GetLast(self):
        return self.items.GetLast()

    def GetNext(self):
        return self.items.GetNext()

    def GetPrevious(self):
        return self.items.GetPrevious()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.items.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.items.Remove(*args, **arguments)

    def ResetColumns(self):
        self.items.ResetColumns()

    def Restrict(self, *args, Filter=None):
        arguments = {"Filter": Filter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.items.Restrict(*args, **arguments)

    def SetColumns(self, *args, Columns=None):
        arguments = {"Columns": Columns}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.items.SetColumns(*args, **arguments)

    def Sort(self, *args, Property=None, Descending=None):
        arguments = {"Property": Property, "Descending": Descending}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.items.Sort(*args, **arguments)


class JournalItem:

    def __init__(self, journalitem=None):
        self.journalitem = journalitem

    @property
    def Actions(self):
        return Actions(self.journalitem.Actions)

    @property
    def Application(self):
        return Application(self.journalitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.journalitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.journalitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.journalitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.journalitem.BillingInformation = value

    @property
    def Body(self):
        return self.journalitem.Body

    @Body.setter
    def Body(self, value):
        self.journalitem.Body = value

    @property
    def Categories(self):
        return self.journalitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.journalitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.journalitem.Class)

    @property
    def Companies(self):
        return self.journalitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.journalitem.Companies = value

    @property
    def Conflicts(self):
        return Conflicts(self.journalitem.Conflicts)

    @property
    def ContactNames(self):
        return self.journalitem.ContactNames

    @ContactNames.setter
    def ContactNames(self, value):
        self.journalitem.ContactNames = value

    @property
    def ConversationID(self):
        return Conversation(self.journalitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.journalitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.journalitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.journalitem.CreationTime

    @property
    def DocPosted(self):
        return self.journalitem.DocPosted

    @DocPosted.setter
    def DocPosted(self, value):
        self.journalitem.DocPosted = value

    @property
    def DocPrinted(self):
        return self.journalitem.DocPrinted

    @DocPrinted.setter
    def DocPrinted(self, value):
        self.journalitem.DocPrinted = value

    @property
    def DocRouted(self):
        return self.journalitem.DocRouted

    @DocRouted.setter
    def DocRouted(self, value):
        self.journalitem.DocRouted = value

    @property
    def DocSaved(self):
        return self.journalitem.DocSaved

    @DocSaved.setter
    def DocSaved(self, value):
        self.journalitem.DocSaved = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.journalitem.DownloadState)

    @property
    def Duration(self):
        return JournalItem(self.journalitem.Duration)

    @Duration.setter
    def Duration(self, value):
        self.journalitem.Duration = value

    @property
    def End(self):
        return self.journalitem.End

    @End.setter
    def End(self, value):
        self.journalitem.End = value

    @property
    def EntryID(self):
        return self.journalitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.journalitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.journalitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.journalitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.journalitem.Importance = value

    @property
    def IsConflict(self):
        return self.journalitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.journalitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.journalitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.journalitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.journalitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.journalitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.journalitem.MessageClass = value

    @property
    def Mileage(self):
        return self.journalitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.journalitem.Mileage = value

    @property
    def NoAging(self):
        return self.journalitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.journalitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.journalitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.journalitem.OutlookVersion

    @property
    def Parent(self):
        return self.journalitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.journalitem.PropertyAccessor)

    @property
    def Recipients(self):
        return Recipients(self.journalitem.Recipients)

    @property
    def Saved(self):
        return self.journalitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.journalitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.journalitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.journalitem.Session)

    @property
    def Size(self):
        return self.journalitem.Size

    @property
    def Start(self):
        return self.journalitem.Start

    @Start.setter
    def Start(self, value):
        self.journalitem.Start = value

    @property
    def Subject(self):
        return self.journalitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.journalitem.Subject = value

    @property
    def Type(self):
        return self.journalitem.Type

    @Type.setter
    def Type(self, value):
        self.journalitem.Type = value

    @property
    def UnRead(self):
        return self.journalitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.journalitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.journalitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.journalitem.Close(*args, **arguments)

    def Copy(self):
        self.journalitem.Copy()

    def Delete(self):
        self.journalitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.journalitem.Display(*args, **arguments)

    def Forward(self):
        return self.journalitem.Forward()

    def GetConversation(self):
        return self.journalitem.GetConversation()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.journalitem.Move(*args, **arguments)

    def PrintOut(self):
        self.journalitem.PrintOut()

    def Reply(self):
        return MailItem(self.journalitem.Reply())

    def ReplyAll(self):
        return MailItem(self.journalitem.ReplyAll())

    def Save(self):
        self.journalitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.journalitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.journalitem.ShowCategoriesDialog()

    def StartTimer(self):
        self.journalitem.StartTimer()

    def StopTimer(self):
        self.journalitem.StopTimer()


class JournalModule:

    def __init__(self, journalmodule=None):
        self.journalmodule = journalmodule

    @property
    def Application(self):
        return Application(self.journalmodule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.journalmodule.Class)

    @property
    def Name(self):
        return JournalModule(self.journalmodule.Name)

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.journalmodule.NavigationGroups)

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.journalmodule.NavigationModuleType)

    @property
    def Parent(self):
        return self.journalmodule.Parent

    @property
    def Position(self):
        return JournalModule(self.journalmodule.Position)

    @Position.setter
    def Position(self, value):
        self.journalmodule.Position = value

    @property
    def Session(self):
        return NameSpace(self.journalmodule.Session)

    @property
    def Visible(self):
        return JournalModule(self.journalmodule.Visible)

    @Visible.setter
    def Visible(self, value):
        self.journalmodule.Visible = value


class MailItem:

    def __init__(self, mailitem=None):
        self.mailitem = mailitem

    @property
    def Actions(self):
        return Actions(self.mailitem.Actions)

    @property
    def AlternateRecipientAllowed(self):
        return self.mailitem.AlternateRecipientAllowed

    @AlternateRecipientAllowed.setter
    def AlternateRecipientAllowed(self, value):
        self.mailitem.AlternateRecipientAllowed = value

    @property
    def Application(self):
        return Application(self.mailitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.mailitem.Attachments)

    @property
    def AutoForwarded(self):
        return self.mailitem.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.mailitem.AutoForwarded = value

    @property
    def AutoResolvedWinner(self):
        return self.mailitem.AutoResolvedWinner

    @property
    def BCC(self):
        return MailItem(self.mailitem.BCC)

    @BCC.setter
    def BCC(self, value):
        self.mailitem.BCC = value

    @property
    def BillingInformation(self):
        return self.mailitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.mailitem.BillingInformation = value

    @property
    def Body(self):
        return self.mailitem.Body

    @Body.setter
    def Body(self, value):
        self.mailitem.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.mailitem.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.mailitem.BodyFormat = value

    @property
    def Categories(self):
        return self.mailitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.mailitem.Categories = value

    @property
    def CC(self):
        return MailItem(self.mailitem.CC)

    @CC.setter
    def CC(self, value):
        self.mailitem.CC = value

    @property
    def Class(self):
        return OlObjectClass(self.mailitem.Class)

    @property
    def Companies(self):
        return self.mailitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.mailitem.Companies = value

    @property
    def Conflicts(self):
        return self.mailitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.mailitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.mailitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.mailitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.mailitem.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return self.mailitem.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.mailitem.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.mailitem.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.mailitem.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.mailitem.DownloadState)

    @property
    def EntryID(self):
        return self.mailitem.EntryID

    @property
    def ExpiryTime(self):
        return self.mailitem.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.mailitem.ExpiryTime = value

    @property
    def FlagRequest(self):
        return self.mailitem.FlagRequest

    @FlagRequest.setter
    def FlagRequest(self, value):
        self.mailitem.FlagRequest = value

    @property
    def FormDescription(self):
        return FormDescription(self.mailitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.mailitem.GetInspector)

    @property
    def HTMLBody(self):
        return self.mailitem.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.mailitem.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.mailitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.mailitem.Importance = value

    @property
    def InternetCodepage(self):
        return self.mailitem.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.mailitem.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.mailitem.IsConflict

    @property
    def IsMarkedAsTask(self):
        return MailItem(self.mailitem.IsMarkedAsTask)

    @property
    def ItemProperties(self):
        return ItemProperties(self.mailitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.mailitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.mailitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.mailitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.mailitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.mailitem.MessageClass = value

    @property
    def Mileage(self):
        return self.mailitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.mailitem.Mileage = value

    @property
    def NoAging(self):
        return self.mailitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.mailitem.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return self.mailitem.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.mailitem.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return self.mailitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.mailitem.OutlookVersion

    @property
    def Parent(self):
        return self.mailitem.Parent

    @property
    def Permission(self):
        return self.mailitem.Permission

    @Permission.setter
    def Permission(self, value):
        self.mailitem.Permission = value

    @property
    def PermissionService(self):
        return self.mailitem.PermissionService

    @PermissionService.setter
    def PermissionService(self, value):
        self.mailitem.PermissionService = value

    @property
    def PermissionTemplateGuid(self):
        return MailItem(self.mailitem.PermissionTemplateGuid)

    @PermissionTemplateGuid.setter
    def PermissionTemplateGuid(self, value):
        self.mailitem.PermissionTemplateGuid = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.mailitem.PropertyAccessor)

    @property
    def ReadReceiptRequested(self):
        return self.mailitem.ReadReceiptRequested

    @property
    def ReceivedByEntryID(self):
        return self.mailitem.ReceivedByEntryID

    @property
    def ReceivedByName(self):
        return self.mailitem.ReceivedByName

    @property
    def ReceivedOnBehalfOfEntryID(self):
        return self.mailitem.ReceivedOnBehalfOfEntryID

    @property
    def ReceivedOnBehalfOfName(self):
        return self.mailitem.ReceivedOnBehalfOfName

    @property
    def ReceivedTime(self):
        return self.mailitem.ReceivedTime

    @property
    def RecipientReassignmentProhibited(self):
        return self.mailitem.RecipientReassignmentProhibited

    @RecipientReassignmentProhibited.setter
    def RecipientReassignmentProhibited(self, value):
        self.mailitem.RecipientReassignmentProhibited = value

    @property
    def Recipients(self):
        return Recipients(self.mailitem.Recipients)

    @property
    def ReminderOverrideDefault(self):
        return self.mailitem.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.mailitem.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.mailitem.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.mailitem.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.mailitem.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.mailitem.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.mailitem.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.mailitem.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.mailitem.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.mailitem.ReminderTime = value

    @property
    def RemoteStatus(self):
        return OlRemoteStatus(self.mailitem.RemoteStatus)

    @RemoteStatus.setter
    def RemoteStatus(self, value):
        self.mailitem.RemoteStatus = value

    @property
    def ReplyRecipientNames(self):
        return self.mailitem.ReplyRecipientNames

    @property
    def ReplyRecipients(self):
        return Recipients(self.mailitem.ReplyRecipients)

    @property
    def RetentionExpirationDate(self):
        return MailItem(self.mailitem.RetentionExpirationDate)

    @property
    def RetentionPolicyName(self):
        return self.mailitem.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.mailitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.mailitem.RTFBody = value

    @property
    def Saved(self):
        return self.mailitem.Saved

    @property
    def SaveSentMessageFolder(self):
        return Folder(self.mailitem.SaveSentMessageFolder)

    @SaveSentMessageFolder.setter
    def SaveSentMessageFolder(self, value):
        self.mailitem.SaveSentMessageFolder = value

    @property
    def Sender(self):
        return self.mailitem.Sender

    @Sender.setter
    def Sender(self, value):
        self.mailitem.Sender = value

    @property
    def SenderEmailAddress(self):
        return self.mailitem.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.mailitem.SenderEmailType

    @property
    def SenderName(self):
        return self.mailitem.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.mailitem.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.mailitem.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.mailitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.mailitem.Sensitivity = value

    @property
    def Sent(self):
        return self.mailitem.Sent

    @property
    def SentOn(self):
        return self.mailitem.SentOn

    @property
    def SentOnBehalfOfName(self):
        return self.mailitem.SentOnBehalfOfName

    @SentOnBehalfOfName.setter
    def SentOnBehalfOfName(self, value):
        self.mailitem.SentOnBehalfOfName = value

    @property
    def Session(self):
        return NameSpace(self.mailitem.Session)

    @property
    def Size(self):
        return self.mailitem.Size

    @property
    def Subject(self):
        return self.mailitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.mailitem.Subject = value

    @property
    def Submitted(self):
        return self.mailitem.Submitted

    @property
    def TaskCompletedDate(self):
        return MailItem(self.mailitem.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.mailitem.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return MailItem(self.mailitem.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.mailitem.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return MailItem(self.mailitem.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.mailitem.TaskStartDate = value

    @property
    def TaskSubject(self):
        return MailItem(self.mailitem.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.mailitem.TaskSubject = value

    @property
    def To(self):
        return self.mailitem.To

    @To.setter
    def To(self, value):
        self.mailitem.To = value

    @property
    def ToDoTaskOrdinal(self):
        return MailItem(self.mailitem.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.mailitem.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.mailitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.mailitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.mailitem.UserProperties)

    @property
    def VotingOptions(self):
        return self.mailitem.VotingOptions

    @VotingOptions.setter
    def VotingOptions(self, value):
        self.mailitem.VotingOptions = value

    @property
    def VotingResponse(self):
        return self.mailitem.VotingResponse

    @VotingResponse.setter
    def VotingResponse(self, value):
        self.mailitem.VotingResponse = value

    def AddBusinessCard(self, *args, contact=None):
        arguments = {"contact": contact}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailitem.AddBusinessCard(*args, **arguments)

    def ClearConversationIndex(self):
        self.mailitem.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.mailitem.ClearTaskFlag()

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailitem.Close(*args, **arguments)

    def Copy(self):
        self.mailitem.Copy()

    def Delete(self):
        self.mailitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailitem.Display(*args, **arguments)

    def Forward(self):
        return self.mailitem.Forward()

    def GetConversation(self):
        return self.mailitem.GetConversation()

    def MarkAsTask(self, *args, MarkInterval=None):
        arguments = {"MarkInterval": MarkInterval}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailitem.MarkAsTask(*args, **arguments)

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mailitem.Move(*args, **arguments)

    def PrintOut(self):
        self.mailitem.PrintOut()

    def Reply(self):
        return MailItem(self.mailitem.Reply())

    def ReplyAll(self):
        return MailItem(self.mailitem.ReplyAll())

    def Save(self):
        self.mailitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailitem.SaveAs(*args, **arguments)

    def Send(self):
        self.mailitem.Send()

    def ShowCategoriesDialog(self):
        self.mailitem.ShowCategoriesDialog()


class MailModule:

    def __init__(self, mailmodule=None):
        self.mailmodule = mailmodule

    @property
    def Application(self):
        return Application(self.mailmodule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.mailmodule.Class)

    @property
    def Name(self):
        return MailModule(self.mailmodule.Name)

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.mailmodule.NavigationGroups)

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.mailmodule.NavigationModuleType)

    @property
    def Parent(self):
        return self.mailmodule.Parent

    @property
    def Position(self):
        return MailModule(self.mailmodule.Position)

    @Position.setter
    def Position(self, value):
        self.mailmodule.Position = value

    @property
    def Session(self):
        return NameSpace(self.mailmodule.Session)

    @property
    def Visible(self):
        return MailModule(self.mailmodule.Visible)

    @Visible.setter
    def Visible(self, value):
        self.mailmodule.Visible = value


class MarkAsTaskRuleAction:

    def __init__(self, markastaskruleaction=None):
        self.markastaskruleaction = markastaskruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.markastaskruleaction.ActionType)

    @property
    def Application(self):
        return Application(self.markastaskruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.markastaskruleaction.Class)

    @property
    def Enabled(self):
        return self.markastaskruleaction.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.markastaskruleaction.Enabled = value

    @property
    def FlagTo(self):
        return self.markastaskruleaction.FlagTo

    @FlagTo.setter
    def FlagTo(self, value):
        self.markastaskruleaction.FlagTo = value

    @property
    def MarkInterval(self):
        return OlMarkInterval(self.markastaskruleaction.MarkInterval)

    @MarkInterval.setter
    def MarkInterval(self, value):
        self.markastaskruleaction.MarkInterval = value

    @property
    def Parent(self):
        return self.markastaskruleaction.Parent

    @property
    def Session(self):
        return NameSpace(self.markastaskruleaction.Session)


class MeetingItem:

    def __init__(self, meetingitem=None):
        self.meetingitem = meetingitem

    @property
    def Actions(self):
        return Actions(self.meetingitem.Actions)

    @property
    def Application(self):
        return Application(self.meetingitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.meetingitem.Attachments)

    @property
    def AutoForwarded(self):
        return self.meetingitem.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.meetingitem.AutoForwarded = value

    @property
    def AutoResolvedWinner(self):
        return self.meetingitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.meetingitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.meetingitem.BillingInformation = value

    @property
    def Body(self):
        return self.meetingitem.Body

    @Body.setter
    def Body(self, value):
        self.meetingitem.Body = value

    @property
    def Categories(self):
        return self.meetingitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.meetingitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.meetingitem.Class)

    @property
    def Companies(self):
        return self.meetingitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.meetingitem.Companies = value

    @property
    def Conflicts(self):
        return self.meetingitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.meetingitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.meetingitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.meetingitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.meetingitem.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return self.meetingitem.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.meetingitem.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.meetingitem.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.meetingitem.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.meetingitem.DownloadState)

    @property
    def EntryID(self):
        return self.meetingitem.EntryID

    @property
    def ExpiryTime(self):
        return self.meetingitem.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.meetingitem.ExpiryTime = value

    @property
    def FormDescription(self):
        return FormDescription(self.meetingitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.meetingitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.meetingitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.meetingitem.Importance = value

    @property
    def IsConflict(self):
        return self.meetingitem.IsConflict

    @property
    def IsLatestVersion(self):
        return MeetingItem(self.meetingitem.IsLatestVersion)

    @property
    def ItemProperties(self):
        return ItemProperties(self.meetingitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.meetingitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.meetingitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.meetingitem.MarkForDownload = value

    @property
    def MeetingWorkspaceURL(self):
        return self.meetingitem.MeetingWorkspaceURL

    @property
    def MessageClass(self):
        return self.meetingitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.meetingitem.MessageClass = value

    @property
    def Mileage(self):
        return self.meetingitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.meetingitem.Mileage = value

    @property
    def NoAging(self):
        return self.meetingitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.meetingitem.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return self.meetingitem.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.meetingitem.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return self.meetingitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.meetingitem.OutlookVersion

    @property
    def Parent(self):
        return self.meetingitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.meetingitem.PropertyAccessor)

    @property
    def ReceivedTime(self):
        return self.meetingitem.ReceivedTime

    @ReceivedTime.setter
    def ReceivedTime(self, value):
        self.meetingitem.ReceivedTime = value

    @property
    def Recipients(self):
        return Recipients(self.meetingitem.Recipients)

    @property
    def ReminderSet(self):
        return self.meetingitem.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.meetingitem.ReminderSet = value

    @property
    def ReminderTime(self):
        return self.meetingitem.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.meetingitem.ReminderTime = value

    @property
    def ReplyRecipients(self):
        return Recipients(self.meetingitem.ReplyRecipients)

    @property
    def RetentionExpirationDate(self):
        return MeetingItem(self.meetingitem.RetentionExpirationDate)

    @property
    def RetentionPolicyName(self):
        return self.meetingitem.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.meetingitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.meetingitem.RTFBody = value

    @property
    def Saved(self):
        return self.meetingitem.Saved

    @property
    def SaveSentMessageFolder(self):
        return self.meetingitem.SaveSentMessageFolder

    @property
    def SenderEmailAddress(self):
        return self.meetingitem.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.meetingitem.SenderEmailType

    @property
    def SenderName(self):
        return self.meetingitem.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.meetingitem.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.meetingitem.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.meetingitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.meetingitem.Sensitivity = value

    @property
    def Sent(self):
        return self.meetingitem.Sent

    @property
    def SentOn(self):
        return self.meetingitem.SentOn

    @property
    def Session(self):
        return NameSpace(self.meetingitem.Session)

    @property
    def Size(self):
        return self.meetingitem.Size

    @property
    def Subject(self):
        return self.meetingitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.meetingitem.Subject = value

    @property
    def Submitted(self):
        return self.meetingitem.Submitted

    @property
    def UnRead(self):
        return self.meetingitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.meetingitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.meetingitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.meetingitem.Close(*args, **arguments)

    def Copy(self):
        self.meetingitem.Copy()

    def Delete(self):
        self.meetingitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.meetingitem.Display(*args, **arguments)

    def Forward(self):
        return self.meetingitem.Forward()

    def GetAssociatedAppointment(self, *args, AddToCalendar=None):
        arguments = {"AddToCalendar": AddToCalendar}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.meetingitem.GetAssociatedAppointment(*args, **arguments)

    def GetConversation(self):
        return self.meetingitem.GetConversation()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.meetingitem.Move(*args, **arguments)

    def PrintOut(self):
        self.meetingitem.PrintOut()

    def Reply(self):
        return MailItem(self.meetingitem.Reply())

    def ReplyAll(self):
        return MailItem(self.meetingitem.ReplyAll())

    def Save(self):
        self.meetingitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.meetingitem.SaveAs(*args, **arguments)

    def Send(self):
        self.meetingitem.Send()

    def ShowCategoriesDialog(self):
        self.meetingitem.ShowCategoriesDialog()


class MoveOrCopyRuleAction:

    def __init__(self, moveorcopyruleaction=None):
        self.moveorcopyruleaction = moveorcopyruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.moveorcopyruleaction.ActionType)

    @property
    def Application(self):
        return Application(self.moveorcopyruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.moveorcopyruleaction.Class)

    @property
    def Enabled(self):
        return self.moveorcopyruleaction.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.moveorcopyruleaction.Enabled = value

    @property
    def Folder(self):
        return Folder(self.moveorcopyruleaction.Folder)

    @Folder.setter
    def Folder(self, value):
        self.moveorcopyruleaction.Folder = value

    @property
    def Parent(self):
        return self.moveorcopyruleaction.Parent

    @property
    def Session(self):
        return NameSpace(self.moveorcopyruleaction.Session)


class NameSpace:

    def __init__(self, namespace=None):
        self.namespace = namespace

    @property
    def Accounts(self):
        return Accounts(self.namespace.Accounts)

    @property
    def AddressLists(self):
        return AddressLists(self.namespace.AddressLists)

    @property
    def Application(self):
        return Application(self.namespace.Application)

    @property
    def AutoDiscoverConnectionMode(self):
        return OlAutoDiscoverConnectionMode(self.namespace.AutoDiscoverConnectionMode)

    @property
    def AutoDiscoverXml(self):
        return self.namespace.AutoDiscoverXml

    @property
    def Categories(self):
        return Categories(self.namespace.Categories)

    @Categories.setter
    def Categories(self, value):
        self.namespace.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.namespace.Class)

    @property
    def CurrentProfileName(self):
        return self.namespace.CurrentProfileName

    @property
    def CurrentUser(self):
        return Recipient(self.namespace.CurrentUser)

    @property
    def DefaultStore(self):
        return Store(self.namespace.DefaultStore)

    @property
    def ExchangeConnectionMode(self):
        return OlExchangeConnectionMode(self.namespace.ExchangeConnectionMode)

    @property
    def ExchangeMailboxServerName(self):
        return self.namespace.ExchangeMailboxServerName

    @property
    def ExchangeMailboxServerVersion(self):
        return self.namespace.ExchangeMailboxServerVersion

    @property
    def Folders(self):
        return Folders(self.namespace.Folders)

    @property
    def Offline(self):
        return self.namespace.Offline

    @property
    def Parent(self):
        return self.namespace.Parent

    @property
    def Session(self):
        return NameSpace(self.namespace.Session)

    @property
    def Stores(self):
        return Stores(self.namespace.Stores)

    @property
    def SyncObjects(self):
        return SyncObjects(self.namespace.SyncObjects)

    @property
    def Type(self):
        return self.namespace.Type

    def AddStore(self, *args, Store=None):
        arguments = {"Store": Store}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.namespace.AddStore(*args, **arguments)

    def AddStoreEx(self, *args, Store=None, Type=None):
        arguments = {"Store": Store, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.namespace.AddStoreEx(*args, **arguments)

    def CompareEntryIDs(self, *args, FirstEntryID=None, SecondEntryID=None):
        arguments = {"FirstEntryID": FirstEntryID, "SecondEntryID": SecondEntryID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.CompareEntryIDs(*args, **arguments)

    def CreateContactCard(self, *args, Address=None):
        arguments = {"Address": Address}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.CreateContactCard(*args, **arguments)

    def CreateRecipient(self, *args, RecipientName=None):
        arguments = {"RecipientName": RecipientName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.CreateRecipient(*args, **arguments)

    def CreateSharingItem(self, *args, Context=None, Provider=None):
        arguments = {"Context": Context, "Provider": Provider}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.CreateSharingItem(*args, **arguments)

    def Dial(self, *args, ContactItem=None):
        arguments = {"ContactItem": ContactItem}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.namespace.Dial(*args, **arguments)

    def GetAddressEntryFromID(self, *args, ID=None):
        arguments = {"ID": ID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ID(self.namespace.GetAddressEntryFromID(*args, **arguments))

    def GetDefaultFolder(self, *args, FolderType=None):
        arguments = {"FolderType": FolderType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.GetDefaultFolder(*args, **arguments)

    def GetFolderFromID(self, *args, EntryIDFolder=None, EntryIDStore=None):
        arguments = {"EntryIDFolder": EntryIDFolder, "EntryIDStore": EntryIDStore}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.GetFolderFromID(*args, **arguments)

    def GetGlobalAddressList(self):
        return self.namespace.GetGlobalAddressList()

    def GetItemFromID(self, *args, EntryIDItem=None, EntryIDStore=None):
        arguments = {"EntryIDItem": EntryIDItem, "EntryIDStore": EntryIDStore}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.GetItemFromID(*args, **arguments)

    def GetRecipientFromID(self, *args, EntryID=None):
        arguments = {"EntryID": EntryID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.GetRecipientFromID(*args, **arguments)

    def GetSelectNamesDialog(self):
        return self.namespace.GetSelectNamesDialog()

    def GetSharedDefaultFolder(self, *args, Recipient=None, FolderType=None):
        arguments = {"Recipient": Recipient, "FolderType": FolderType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.GetSharedDefaultFolder(*args, **arguments)

    def GetStoreFromID(self, *args, ID=None):
        arguments = {"ID": ID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return StoreID(self.namespace.GetStoreFromID(*args, **arguments))

    def Logoff(self):
        self.namespace.Logoff()

    def Logon(self, *args, Profile=None, Password=None, ShowDialog=None, NewSession=None):
        arguments = {"Profile": Profile, "Password": Password, "ShowDialog": ShowDialog, "NewSession": NewSession}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.namespace.Logon(*args, **arguments)

    def OpenSharedFolder(self, *args, Path=None, Name=None, DownloadAttachments=None, UseTTL=None):
        arguments = {"Path": Path, "Name": Name, "DownloadAttachments": DownloadAttachments, "UseTTL": UseTTL}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Folder(self.namespace.OpenSharedFolder(*args, **arguments))

    def OpenSharedItem(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.namespace.OpenSharedItem(*args, **arguments)

    def PickFolder(self):
        return Folder(self.namespace.PickFolder())

    def RemoveStore(self, *args, Folder=None):
        arguments = {"Folder": Folder}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.namespace.RemoveStore(*args, **arguments)

    def SendAndReceive(self, *args, showProgressDialog=None):
        arguments = {"showProgressDialog": showProgressDialog}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.namespace.SendAndReceive(*args, **arguments)


class NavigationFolder:

    def __init__(self, navigationfolder=None):
        self.navigationfolder = navigationfolder

    @property
    def Application(self):
        return Application(self.navigationfolder.Application)

    @property
    def Class(self):
        return OlObjectClass(self.navigationfolder.Class)

    @property
    def DisplayName(self):
        return NavigationFolder(self.navigationfolder.DisplayName)

    @property
    def Folder(self):
        return Folder(self.navigationfolder.Folder)

    @property
    def IsRemovable(self):
        return NavigationFolder(self.navigationfolder.IsRemovable)

    @property
    def IsSelected(self):
        return NavigationFolder(self.navigationfolder.IsSelected)

    @IsSelected.setter
    def IsSelected(self, value):
        self.navigationfolder.IsSelected = value

    @property
    def IsSideBySide(self):
        return NavigationFolder(self.navigationfolder.IsSideBySide)

    @IsSideBySide.setter
    def IsSideBySide(self, value):
        self.navigationfolder.IsSideBySide = value

    @property
    def Parent(self):
        return self.navigationfolder.Parent

    @property
    def Position(self):
        return NavigationFolder(self.navigationfolder.Position)

    @Position.setter
    def Position(self, value):
        self.navigationfolder.Position = value

    @property
    def Session(self):
        return NameSpace(self.navigationfolder.Session)


class NavigationFolders:

    def __init__(self, navigationfolders=None):
        self.navigationfolders = navigationfolders

    @property
    def Application(self):
        return Application(self.navigationfolders.Application)

    @property
    def Class(self):
        return OlObjectClass(self.navigationfolders.Class)

    @property
    def Count(self):
        return self.navigationfolders.Count

    @property
    def Parent(self):
        return self.navigationfolders.Parent

    @property
    def Session(self):
        return NameSpace(self.navigationfolders.Session)

    def Add(self, *args, Folder=None):
        arguments = {"Folder": Folder}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.navigationfolders.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.navigationfolders.Item(*args, **arguments)

    def Remove(self, *args, RemovableFolder=None):
        arguments = {"RemovableFolder": RemovableFolder}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.navigationfolders.Remove(*args, **arguments)


class NavigationGroup:

    def __init__(self, navigationgroup=None):
        self.navigationgroup = navigationgroup

    @property
    def Application(self):
        return Application(self.navigationgroup.Application)

    @property
    def Class(self):
        return OlObjectClass(self.navigationgroup.Class)

    @property
    def GroupType(self):
        return OlGroupType(self.navigationgroup.GroupType)

    @property
    def Name(self):
        return NavigationGroup(self.navigationgroup.Name)

    @Name.setter
    def Name(self, value):
        self.navigationgroup.Name = value

    @property
    def NavigationFolders(self):
        return NavigationFolders(self.navigationgroup.NavigationFolders)

    @property
    def Parent(self):
        return self.navigationgroup.Parent

    @property
    def Position(self):
        return NavigationGroup(self.navigationgroup.Position)

    @Position.setter
    def Position(self, value):
        self.navigationgroup.Position = value

    @property
    def Session(self):
        return NameSpace(self.navigationgroup.Session)


class NavigationGroups:

    def __init__(self, navigationgroups=None):
        self.navigationgroups = navigationgroups

    @property
    def Application(self):
        return Application(self.navigationgroups.Application)

    @property
    def Class(self):
        return OlObjectClass(self.navigationgroups.Class)

    @property
    def Count(self):
        return self.navigationgroups.Count

    @property
    def Parent(self):
        return self.navigationgroups.Parent

    @property
    def Session(self):
        return NameSpace(self.navigationgroups.Session)

    def Create(self, *args, GroupDisplayName=None):
        arguments = {"GroupDisplayName": GroupDisplayName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.navigationgroups.Create(*args, **arguments)

    def Delete(self, *args, Group=None):
        arguments = {"Group": Group}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.navigationgroups.Delete(*args, **arguments)

    def GetDefaultNavigationGroup(self, *args, DefaultFolderGroup=None):
        arguments = {"DefaultFolderGroup": DefaultFolderGroup}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.navigationgroups.GetDefaultNavigationGroup(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.navigationgroups.Item(*args, **arguments)


class NavigationModule:

    def __init__(self, navigationmodule=None):
        self.navigationmodule = navigationmodule

    @property
    def Application(self):
        return Application(self.navigationmodule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.navigationmodule.Class)

    @property
    def Name(self):
        return NavigationModule(self.navigationmodule.Name)

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.navigationmodule.NavigationModuleType)

    @property
    def Parent(self):
        return self.navigationmodule.Parent

    @property
    def Position(self):
        return NavigationModule(self.navigationmodule.Position)

    @Position.setter
    def Position(self, value):
        self.navigationmodule.Position = value

    @property
    def Session(self):
        return NameSpace(self.navigationmodule.Session)

    @property
    def Visible(self):
        return NavigationModule(self.navigationmodule.Visible)

    @Visible.setter
    def Visible(self, value):
        self.navigationmodule.Visible = value


class NavigationModules:

    def __init__(self, navigationmodules=None):
        self.navigationmodules = navigationmodules

    @property
    def Application(self):
        return Application(self.navigationmodules.Application)

    @property
    def Class(self):
        return OlObjectClass(self.navigationmodules.Class)

    @property
    def Count(self):
        return self.navigationmodules.Count

    @property
    def Parent(self):
        return self.navigationmodules.Parent

    @property
    def Session(self):
        return NameSpace(self.navigationmodules.Session)

    def GetNavigationModule(self, *args, ModuleType=None):
        arguments = {"ModuleType": ModuleType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.navigationmodules.GetNavigationModule(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.navigationmodules.Item(*args, **arguments)


class NavigationPane:

    def __init__(self, navigationpane=None):
        self.navigationpane = navigationpane

    @property
    def Application(self):
        return Application(self.navigationpane.Application)

    @property
    def Class(self):
        return OlObjectClass(self.navigationpane.Class)

    @property
    def CurrentModule(self):
        return NavigationModule(self.navigationpane.CurrentModule)

    @CurrentModule.setter
    def CurrentModule(self, value):
        self.navigationpane.CurrentModule = value

    @property
    def DisplayedModuleCount(self):
        return NavigationModule(self.navigationpane.DisplayedModuleCount)

    @DisplayedModuleCount.setter
    def DisplayedModuleCount(self, value):
        self.navigationpane.DisplayedModuleCount = value

    @property
    def IsCollapsed(self):
        return self.navigationpane.IsCollapsed

    @IsCollapsed.setter
    def IsCollapsed(self, value):
        self.navigationpane.IsCollapsed = value

    @property
    def Modules(self):
        return NavigationModules(self.navigationpane.Modules)

    @property
    def Parent(self):
        return self.navigationpane.Parent

    @property
    def Session(self):
        return NameSpace(self.navigationpane.Session)


class NewItemAlertRuleAction:

    def __init__(self, newitemalertruleaction=None):
        self.newitemalertruleaction = newitemalertruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.newitemalertruleaction.ActionType)

    @property
    def Application(self):
        return Application(self.newitemalertruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.newitemalertruleaction.Class)

    @property
    def Enabled(self):
        return self.newitemalertruleaction.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.newitemalertruleaction.Enabled = value

    @property
    def Parent(self):
        return self.newitemalertruleaction.Parent

    @property
    def Session(self):
        return NameSpace(self.newitemalertruleaction.Session)

    @property
    def Text(self):
        return self.newitemalertruleaction.Text

    @Text.setter
    def Text(self, value):
        self.newitemalertruleaction.Text = value


class NoteItem:

    def __init__(self, noteitem=None):
        self.noteitem = noteitem

    @property
    def Application(self):
        return Application(self.noteitem.Application)

    @property
    def AutoResolvedWinner(self):
        return self.noteitem.AutoResolvedWinner

    @property
    def Body(self):
        return self.noteitem.Body

    @Body.setter
    def Body(self, value):
        self.noteitem.Body = value

    @property
    def Categories(self):
        return self.noteitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.noteitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.noteitem.Class)

    @property
    def Conflicts(self):
        return self.noteitem.Conflicts

    @property
    def CreationTime(self):
        return self.noteitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.noteitem.DownloadState)

    @property
    def EntryID(self):
        return self.noteitem.EntryID

    @property
    def GetInspector(self):
        return Inspector(self.noteitem.GetInspector)

    @property
    def Height(self):
        return self.noteitem.Height

    @Height.setter
    def Height(self, value):
        self.noteitem.Height = value

    @property
    def IsConflict(self):
        return self.noteitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.noteitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.noteitem.LastModificationTime

    @property
    def Left(self):
        return self.noteitem.Left

    @Left.setter
    def Left(self, value):
        self.noteitem.Left = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.noteitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.noteitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.noteitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.noteitem.MessageClass = value

    @property
    def Parent(self):
        return self.noteitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.noteitem.PropertyAccessor)

    @property
    def Saved(self):
        return self.noteitem.Saved

    @property
    def Session(self):
        return NameSpace(self.noteitem.Session)

    @property
    def Size(self):
        return self.noteitem.Size

    @property
    def Subject(self):
        return self.noteitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.noteitem.Subject = value

    @property
    def Top(self):
        return self.noteitem.Top

    @Top.setter
    def Top(self, value):
        self.noteitem.Top = value

    @property
    def Width(self):
        return self.noteitem.Width

    @Width.setter
    def Width(self, value):
        self.noteitem.Width = value

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.noteitem.Close(*args, **arguments)

    def Copy(self):
        return NoteItem(self.noteitem.Copy())

    def Delete(self):
        self.noteitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.noteitem.Display(*args, **arguments)

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.noteitem.Move(*args, **arguments)

    def PrintOut(self):
        self.noteitem.PrintOut()

    def Save(self):
        self.noteitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.noteitem.SaveAs(*args, **arguments)


class NotesModule:

    def __init__(self, notesmodule=None):
        self.notesmodule = notesmodule

    @property
    def Application(self):
        return Application(self.notesmodule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.notesmodule.Class)

    @property
    def Name(self):
        return NotesModule(self.notesmodule.Name)

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.notesmodule.NavigationGroups)

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.notesmodule.NavigationModuleType)

    @property
    def Parent(self):
        return self.notesmodule.Parent

    @property
    def Position(self):
        return NotesModule(self.notesmodule.Position)

    @Position.setter
    def Position(self, value):
        self.notesmodule.Position = value

    @property
    def Session(self):
        return NameSpace(self.notesmodule.Session)

    @property
    def Visible(self):
        return NotesModule(self.notesmodule.Visible)

    @Visible.setter
    def Visible(self, value):
        self.notesmodule.Visible = value


# OlAccountType enumeration
olEas = 4
olExchange = 0
olHttp = 3
olImap = 1
olOtherAccount = 5
olPop3 = 2

# OlActionCopyLike enumeration
olForward = 2
olReply = 0
olReplyAll = 1
olReplyFolder = 3
olRespond = 4

# OlActionReplyStyle enumeration
olEmbedOriginalItem = 1
olIncludeOriginalText = 2
olIndentOriginalText = 3
olLinkOriginalItem = 4
olOmitOriginalText = 0
olReplyTickOriginalText = 1000
olUserPreference = 5

# OlActionResponseStyle enumeration
olOpen = 0
olPrompt = 2
olSend = 1

# OlActionShowOn enumeration
olDontShow = 0
olMenu = 1
olMenuAndToolbar = 2

# OlAddressEntryUserType enumeration
olExchangeAgentAddressEntry = 3
olExchangeDistributionListAddressEntry = 1
olExchangeOrganizationAddressEntry = 4
olExchangePublicFolderAddressEntry = 2
olExchangeRemoteUserAddressEntry = 5
olExchangeUserAddressEntry = 0
olLdapAddressEntry = 20
olOtherAddressEntry = 40
olOutlookContactAddressEntry = 10
olOutlookDistributionListAddressEntry = 11
olSmtpAddressEntry = 30

# OlAddressListType enumeration
olCustomAddressList = 4
olExchangeContainer = 1
olExchangeGlobalAddressList = 0
olOutlookAddressList = 2
olOutlookLdapAddressList = 3

# OlAlign enumeration
olAlignCenter = 1
olAlignLeft = 0
olAlignRight = 2

# OlAlignment enumeration
olAlignmentLeft = 0
olAlignmentRight = 1

# OlAlwaysDeleteConversation enumeration
olAlwaysDelete = 1
olAlwaysDeleteUnsupported = 2
olDoNotDelete = 0

# OlAppointmentCopyOptions enumeration
olCopyAsAccept = 2
olCreateAppointment = 1
olPromptUser = 0

# OlAppointmentTimeField enumeration
olAppointmentTimeFieldEnd = 3
olAppointmentTimeFieldNone = 1
olAppointmentTimeFieldStart = 2

# OlAttachmentBlockLevel enumeration
olAttachmentBlockLevelNone = 0
olAttachmentBlockLevelOpen = 1

# OlAttachmentType enumeration
olByReference = 4
olByValue = 1
olEmbeddeditem = 5
olOLE = 6

# OlAutoDiscoverConnectionMode enumeration
olAutoDiscoverConnectionExternal = 1
olAutoDiscoverConnectionInternal = 2
olAutoDiscoverConnectionInternalDomain = 3
olAutoDiscoverConnectionUnknown = 0

# OlAutoPreview enumeration
olAutoPreviewAll = 0
olAutoPreviewNone = 2
olAutoPreviewUnread = 1

# OlBackStyle enumeration
olBackStyleOpaque = 1
olBackStyleTransparent = 0

# OlBodyFormat enumeration
olFormatHTML = 2
olFormatPlain = 1
olFormatRichText = 3
olFormatUnspecified = 0

# OlBorderStyle enumeration
olBorderStyleNone = 0
olBorderStyleSingle = 1

# OlBusinessCardType enumeration
olBusinessCardTypeInterConnect = 1
olBusinessCardTypeOutlook = 0

# OlBusyStatus enumeration
olBusy = 2
olFree = 0
olOutOfOffice = 3
olTentative = 1
olWorkingElsewhere = 4

# OlCalendarDetail enumeration
olFreeBusyAndSubject = 1
olFreeBusyOnly = 0
olFullDetails = 2

# OlCalendarMailFormat enumeration
olCalendarMailFormatDailySchedule = 0
olCalendarMailFormatEventList = 1

# OlCalendarViewMode enumeration
olCalendarView5DayWeek = 4
olCalendarViewDay = 0
olCalendarViewMonth = 2
olCalendarViewMultiDay = 3
olCalendarViewWeek = 1

# OlCategoryColor enumeration
olCategoryColorBlack = 15
olCategoryColorBlue = 8
olCategoryColorDarkBlue = 23
olCategoryColorDarkGray = 14
olCategoryColorDarkGreen = 20
olCategoryColorDarkMaroon = 25
olCategoryColorDarkOlive = 22
olCategoryColorDarkOrange = 17
olCategoryColorDarkPeach = 18
olCategoryColorDarkPurple = 24
olCategoryColorDarkRed = 16
olCategoryColorDarkSteel = 12
olCategoryColorDarkTeal = 21
olCategoryColorDarkYellow = 19
olCategoryColorGray = 13
olCategoryColorGreen = 5
olCategoryColorMaroon = 10
olCategoryColorNone = 0
olCategoryColorOlive = 7
olCategoryColorOrange = 2
olCategoryColorPeach = 3
olCategoryColorPurple = 9
olCategoryColorRed = 1
olCategoryColorSteel = 11
olCategoryColorTeal = 6
olCategoryColorYellow = 4

# OlCategoryShortcutKey enumeration
olCategoryShortcutKeyCtrlF10 = 10
olCategoryShortcutKeyCtrlF11 = 11
olCategoryShortcutKeyCtrlF12 = 12
olCategoryShortcutKeyCtrlF2 = 2
olCategoryShortcutKeyCtrlF3 = 3
olCategoryShortcutKeyCtrlF4 = 4
olCategoryShortcutKeyCtrlF5 = 5
olCategoryShortcutKeyCtrlF6 = 6
olCategoryShortcutKeyCtrlF7 = 7
olCategoryShortcutKeyCtrlF8 = 8
olCategoryShortcutKeyCtrlF9 = 9
olCategoryShortcutKeyNone = 0

# OlColor enumeration
olAutoColor = 0
olColorAqua = 15
olColorBlack = 1
olColorBlue = 13
olColorFuchsia = 14
olColorGray = 8
olColorGreen = 3
olColorLime = 11
olColorMaroon = 2
olColorNavy = 5
olColorOlive = 4
olColorPurple = 6
olColorRed = 10
olColorSilver = 9
olColorTeal = 7
olColorWhite = 16
olColorYellow = 12

# OlComboBoxStyle enumeration
olComboBoxStyleCombo = 0
olComboBoxStyleList = 1

# OlContactPhoneNumber enumeration
olContactPhoneAssistant = 0
olContactPhoneBusiness = 1
olContactPhoneBusiness2 = 2
olContactPhoneBusinessFax = 3
olContactPhoneCallback = 4
olContactPhoneCar = 5
olContactPhoneCompany = 6
olContactPhoneHome = 7
olContactPhoneHome2 = 8
olContactPhoneHomeFax = 9
olContactPhoneISDN = 10
olContactPhoneMobile = 11
olContactPhoneOther = 12
olContactPhoneOtherFax = 13
olContactPhonePager = 14
olContactPhonePrimary = 15
olContactPhoneRadio = 16
olContactPhoneTelex = 17
olContactPhoneTTYTTD = 18

# OlDaysOfWeek enumeration
olSunday = 1
olMonday = 2
olTuesday = 4
olWednesday = 8
olThursday = 16
olFriday = 32
olSaturday = 64

# OlDayWeekTimeScale enumeration
olTimeScale10Minutes = 2
olTimeScale15Minutes = 3
olTimeScale30Minutes = 4
olTimeScale5Minutes = 0
olTimeScale60Minutes = 5
olTimeScale6Minutes = 1

# OlDefaultExpandCollapseSetting enumeration
olAllCollapsed = 1
olAllExpanded = 0
olLastViewed = 2

# OlDefaultFolders enumeration
olFolderCalendar = 9
olFolderConflicts = 19
olFolderContacts = 10
olFolderDeletedItems = 3
olFolderDrafts = 16
olFolderInbox = 6
olFolderJournal = 11
olFolderJunk = 23
olFolderLocalFailures = 21
olFolderManagedEmail = 29
olFolderNotes = 12
olFolderOutbox = 4
olFolderSentMail = 5
olFolderServerFailures = 22
olFolderSuggestedContacts = 30
olFolderSyncIssues = 20
olFolderTasks = 13
olFolderToDo = 28
olPublicFoldersAllPublicFolders = 18
olFolderRssFeeds = 25

# OlDefaultSelectNamesDisplayMode enumeration
olDefaultDelegates = 6
olDefaultMail = 1
olDefaultMeeting = 2
olDefaultMembers = 5
olDefaultPickRooms = 8
olDefaultSharingRequest = 4
olDefaultSingleName = 7
olDefaultTask = 3

# OlDisplayType enumeration
olAgent = 3
olDistList = 1
olForum = 2
olOrganization = 4
olPrivateDistList = 5
olRemoteUser = 6
olUser = 0

# OlDownloadState enumeration
olFullItem = 1
olHeaderOnly = 0

# OlDragBehavior enumeration
olDragBehaviorDisabled = 0
olDragBehaviorEnabled = 1

# OlEditorType enumeration
olEditorHTML = 2
olEditorRTF = 3
olEditorText = 1
olEditorWord = 4

# OlEnterFieldBehavior enumeration
olEnterFieldBehaviorRecallSelection = 1
olEnterFieldBehaviorSelectAll = 0

# OlExchangeConnectionMode enumeration
olCachedConnectedDrizzle = 600
olCachedConnectedFull = 700
olCachedConnectedHeaders = 500
olCachedDisconnected = 400
olCachedOffline = 200
olDisconnected = 300
olNoExchange = 0
olOffline = 100
olOnline = 800

# OlExchangeStoreType enumeration
olExchangeMailbox = 1
olExchangePublicFolder = 2
olNotExchange = 3
olPrimaryExchangeMailbox = 0
olAdditionalExchangeMailbox = 4

# OlFolderDisplayMode enumeration
olFolderDisplayFolderOnly = 1
olFolderDisplayNoNavigation = 2
olFolderDisplayNormal = 0

# OlFormatCurrency enumeration
olFormatCurrencyDecimal = 1
olFormatCurrencyNonDecimal = 2

# OlFormatDateTime enumeration
olFormatDateTimeBestFit = 17
olFormatDateTimeLongDate = 6
olFormatDateTimeLongDateReversed = 7
OlFormatDateTimeLongDayDate = 5
olFormatDateTimeLongDayDateTime = 1
olFormatDateTimeLongTime = 15
olFormatDateTimeShortDate = 8
olFormatDateTimeShortDateNumOnly = 9
olFormatDateTimeShortDateTime = 2
olFormatDateTimeShortDayDate = 13
olFormatDateTimeShortDayDateTime = 3
olFormatDateTimeShortDayMonth = 10
olFormatDateTimeShortDayMonthDateTime = 4
olFormatDateTimeShortMonthYear = 11
olFormatDateTimeShortMonthYearNumOnly = 12
olFormatDateTimeShortTime = 16

# OlFormatDuration enumeration
olFormatDurationLong = 2
olFormatDurationLongBusiness = 4
olFormatDurationShort = 1
olFormatDurationShortBusiness = 3

# OlFormatEnumeration enumeration
olFormatEnumBitmap = 1
olFormatEnumText = 2

# OlFormatInteger enumeration
olFormatIntegerComputer1 = 2
olFormatIntegerComputer2 = 3
olFormatIntegerComputer3 = 4
olFormatIntegerPlain = 1

# OlFormatKeywords enumeration
olFormatKeywordsText = 1

# OlFormatNumber enumeration
olFormatNumber1Decimal = 3
olFormatNumber2Decimal = 4
olFormatNumberAllDigits = 1
olFormatNumberComputer1 = 6
olFormatNumberComputer2 = 7
olFormatNumberComputer3 = 8
olFormatNumberRaw = 9
olFormatNumberScientific = 5
olFormatNumberTruncated = 2

# OlFormatPercent enumeration
olFormatPercent1Decimal = 2
olFormatPercent2Decimal = 3
olFormatPercentAllDigits = 4
olFormatPercentRounded = 1

# OlFormatSmartFrom enumeration
olFormatSmartFromFromOnly = 2
olFormatSmartFromFromTo = 1

# OlFormatText enumeration
olFormatTextText = 1

# OlFormatYesNo enumeration
olFormatYesNoIcon = 4
olFormatYesNoOnOff = 2
olFormatYesNoTrueFalse = 3
olFormatYesNoYesNo = 1

# OlFormRegionIcon enumeration
olFormRegionIconDefault = 1
olFormRegionIconEncrypted = 9
olFormRegionIconForwarded = 5
olFormRegionIconPage = 11
olFormRegionIconRead = 3
olFormRegionIconRecurring = 12
olFormRegionIconReplied = 4
olFormRegionIconSigned = 8
olFormRegionIconSubmitted = 7
olFormRegionIconUnread = 2
olFormRegionIconUnsent = 6
olFormRegionIconWindow = 10

# OlFormRegionMode enumeration
olFormRegionCompose = 1
olFormRegionPreview = 2
olFormRegionRead = 0

# OlFormRegionSize enumeration
olFormRegionTypeAdjoining = 1
olFormRegionTypeSeparate = 0

# OlFormRegistry enumeration
olDefaultRegistry = 0
olFolderRegistry = 3
olOrganizationRegistry = 4
olPersonalRegistry = 2

# OlGender enumeration
olFemale = 1
olMale = 2
olUnspecified = 0

# OlGridLineStyle enumeration
olGridLineDashes = 3
olGridLineLargeDots = 2
olGridLineNone = 0
olGridLineSmallDots = 1
olGridLineSolid = 4

# OlGroupType enumeration
olCustomFoldersGroup = 0
olFavoriteFoldersGroup = 4
olMyFoldersGroup = 1
olOtherFoldersGroup = 3
olPeopleFoldersGroup = 2
olReadOnlyGroup = 6
olRoomsGroup = 5

# OlHorizontalLayout enumeration
olHorizontalLayoutAlignCenter = 1
olHorizontalLayoutGrow = 3
olHorizontalLayoutAlignLeft = 0
olHorizontalLayoutAlignRight = 2

# OlIconViewPlacement enumeration
olIconAutoArrange = 2
olIconDoNotArrange = 0
olIconLineUp = 1
olIconSortAndAutoArrange = 3

# OlIconViewType enumeration
olIconViewLarge = 0
olIconViewList = 2
olIconViewSmall = 1

# OlImportance enumeration
olImportanceHigh = 2
olImportanceLow = 0
olImportanceNormal = 1

# OlInspectorClose enumeration
olDiscard = 1
olPromptForSave = 2
olSave = 0

# OlItemType enumeration
olAppointmentItem = 1
olContactItem = 2
olDistributionListItem = 7
olJournalItem = 4
olMailItem = 0
olNoteItem = 5
olPostItem = 6
olTaskItem = 3

# OlJournalRecipientType enumeration
olAssociatedContact = 1

class OlkBusinessCardControl:

    def __init__(self, olkbusinesscardcontrol=None):
        self.olkbusinesscardcontrol = olkbusinesscardcontrol

    @property
    def MouseIcon(self):
        return self.olkbusinesscardcontrol.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkbusinesscardcontrol.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkbusinesscardcontrol.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkbusinesscardcontrol.MousePointer = value


class OlkCategory:

    def __init__(self, olkcategory=None):
        self.olkcategory = olkcategory

    @property
    def AutoSize(self):
        return self.olkcategory.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olkcategory.AutoSize = value

    @property
    def BackColor(self):
        return self.olkcategory.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkcategory.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olkcategory.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.olkcategory.BackStyle = value

    @property
    def Enabled(self):
        return self.olkcategory.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcategory.Enabled = value

    @property
    def ForeColor(self):
        return self.olkcategory.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkcategory.ForeColor = value

    @property
    def MouseIcon(self):
        return self.olkcategory.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcategory.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcategory.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcategory.MousePointer = value


class OlkCheckBox:

    def __init__(self, olkcheckbox=None):
        self.olkcheckbox = olkcheckbox

    @property
    def Accelerator(self):
        return self.olkcheckbox.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.olkcheckbox.Accelerator = value

    @property
    def Alignment(self):
        return olAlignment(self.olkcheckbox.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.olkcheckbox.Alignment = value

    @property
    def BackColor(self):
        return self.olkcheckbox.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkcheckbox.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olkcheckbox.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.olkcheckbox.BackStyle = value

    @property
    def Caption(self):
        return self.olkcheckbox.Caption

    @Caption.setter
    def Caption(self, value):
        self.olkcheckbox.Caption = value

    @property
    def Enabled(self):
        return self.olkcheckbox.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcheckbox.Enabled = value

    @property
    def Font(self):
        return self.olkcheckbox.Font

    @property
    def ForeColor(self):
        return self.olkcheckbox.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkcheckbox.ForeColor = value

    @property
    def MouseIcon(self):
        return self.olkcheckbox.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcheckbox.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcheckbox.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcheckbox.MousePointer = value

    @property
    def TripleState(self):
        return self.olkcheckbox.TripleState

    @TripleState.setter
    def TripleState(self, value):
        self.olkcheckbox.TripleState = value

    @property
    def Value(self):
        return self.olkcheckbox.Value

    @Value.setter
    def Value(self, value):
        self.olkcheckbox.Value = value

    @property
    def WordWrap(self):
        return self.olkcheckbox.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olkcheckbox.WordWrap = value


class OlkComboBox:

    def __init__(self, olkcombobox=None):
        self.olkcombobox = olkcombobox

    @property
    def AutoSize(self):
        return self.olkcombobox.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olkcombobox.AutoSize = value

    @property
    def AutoTab(self):
        return self.olkcombobox.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.olkcombobox.AutoTab = value

    @property
    def AutoWordSelect(self):
        return self.olkcombobox.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.olkcombobox.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.olkcombobox.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkcombobox.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olkcombobox.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olkcombobox.BorderStyle = value

    @property
    def DragBehavior(self):
        return self.olkcombobox.DragBehavior

    @DragBehavior.setter
    def DragBehavior(self, value):
        self.olkcombobox.DragBehavior = value

    @property
    def Enabled(self):
        return self.olkcombobox.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcombobox.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.olkcombobox.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.olkcombobox.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.olkcombobox.Font

    @property
    def ForeColor(self):
        return self.olkcombobox.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkcombobox.ForeColor = value

    @property
    def HideSelection(self):
        return self.olkcombobox.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.olkcombobox.HideSelection = value

    @property
    def ListCount(self):
        return self.olkcombobox.ListCount

    @property
    def ListIndex(self):
        return self.olkcombobox.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.olkcombobox.ListIndex = value

    @property
    def Locked(self):
        return self.olkcombobox.Locked

    @Locked.setter
    def Locked(self, value):
        self.olkcombobox.Locked = value

    @property
    def MaxLength(self):
        return self.olkcombobox.MaxLength

    @MaxLength.setter
    def MaxLength(self, value):
        self.olkcombobox.MaxLength = value

    @property
    def MouseIcon(self):
        return self.olkcombobox.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcombobox.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcombobox.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcombobox.MousePointer = value

    @property
    def SelectionMargin(self):
        return self.olkcombobox.SelectionMargin

    @SelectionMargin.setter
    def SelectionMargin(self, value):
        self.olkcombobox.SelectionMargin = value

    @property
    def SelLength(self):
        return self.olkcombobox.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.olkcombobox.SelLength = value

    @property
    def SelStart(self):
        return self.olkcombobox.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.olkcombobox.SelStart = value

    @property
    def SelText(self):
        return self.olkcombobox.SelText

    @property
    def Style(self):
        return OlComboBoxStyle(self.olkcombobox.Style)

    @Style.setter
    def Style(self, value):
        self.olkcombobox.Style = value

    @property
    def Text(self):
        return self.olkcombobox.Text

    @Text.setter
    def Text(self, value):
        self.olkcombobox.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olkcombobox.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.olkcombobox.TextAlign = value

    @property
    def TopIndex(self):
        return self.olkcombobox.TopIndex

    @TopIndex.setter
    def TopIndex(self, value):
        self.olkcombobox.TopIndex = value

    @property
    def Value(self):
        return self.olkcombobox.Value

    @Value.setter
    def Value(self, value):
        self.olkcombobox.Value = value

    def AddItem(self, *args, ItemText=None, Index=None):
        arguments = {"ItemText": ItemText, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.olkcombobox.AddItem(*args, **arguments)

    def Clear(self):
        self.olkcombobox.Clear()

    def Copy(self):
        self.olkcombobox.Copy()

    def Cut(self):
        self.olkcombobox.Cut()

    def DropDown(self):
        self.olkcombobox.DropDown()

    def GetItem(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.olkcombobox.GetItem(*args, **arguments)

    def Paste(self):
        self.olkcombobox.Paste()

    def RemoveItem(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.olkcombobox.RemoveItem(*args, **arguments)

    def SetItem(self, *args, Index=None, Item=None):
        arguments = {"Index": Index, "Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.olkcombobox.SetItem(*args, **arguments)


class OlkCommandButton:

    def __init__(self, olkcommandbutton=None):
        self.olkcommandbutton = olkcommandbutton

    @property
    def Accelerator(self):
        return self.olkcommandbutton.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.olkcommandbutton.Accelerator = value

    @property
    def AutoSize(self):
        return self.olkcommandbutton.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olkcommandbutton.AutoSize = value

    @property
    def Caption(self):
        return self.olkcommandbutton.Caption

    @Caption.setter
    def Caption(self, value):
        self.olkcommandbutton.Caption = value

    @property
    def DisplayDropArrow(self):
        return self.olkcommandbutton.DisplayDropArrow

    @DisplayDropArrow.setter
    def DisplayDropArrow(self, value):
        self.olkcommandbutton.DisplayDropArrow = value

    @property
    def Enabled(self):
        return self.olkcommandbutton.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcommandbutton.Enabled = value

    @property
    def Font(self):
        return self.olkcommandbutton.Font

    @property
    def MouseIcon(self):
        return self.olkcommandbutton.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcommandbutton.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcommandbutton.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcommandbutton.MousePointer = value

    @property
    def Picture(self):
        return self.olkcommandbutton.Picture

    @Picture.setter
    def Picture(self, value):
        self.olkcommandbutton.Picture = value

    @property
    def PictureAlignment(self):
        return OlPictureAlignment(self.olkcommandbutton.PictureAlignment)

    @PictureAlignment.setter
    def PictureAlignment(self, value):
        self.olkcommandbutton.PictureAlignment = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olkcommandbutton.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.olkcommandbutton.TextAlign = value

    @property
    def WordWrap(self):
        return self.olkcommandbutton.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olkcommandbutton.WordWrap = value


class OlkContactPhoto:

    def __init__(self, olkcontactphoto=None):
        self.olkcontactphoto = olkcontactphoto

    @property
    def Enabled(self):
        return self.olkcontactphoto.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcontactphoto.Enabled = value

    @property
    def MouseIcon(self):
        return self.olkcontactphoto.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcontactphoto.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcontactphoto.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcontactphoto.MousePointer = value


class OlkDateControl:

    def __init__(self, olkdatecontrol=None):
        self.olkdatecontrol = olkdatecontrol

    @property
    def AutoSize(self):
        return self.olkdatecontrol.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olkdatecontrol.AutoSize = value

    @property
    def AutoWordSelect(self):
        return self.olkdatecontrol.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.olkdatecontrol.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.olkdatecontrol.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkdatecontrol.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olkdatecontrol.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.olkdatecontrol.BackStyle = value

    @property
    def Date(self):
        return self.olkdatecontrol.Date

    @Date.setter
    def Date(self, value):
        self.olkdatecontrol.Date = value

    @property
    def Enabled(self):
        return self.olkdatecontrol.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkdatecontrol.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.olkdatecontrol.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.olkdatecontrol.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.olkdatecontrol.Font

    @property
    def ForeColor(self):
        return self.olkdatecontrol.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkdatecontrol.ForeColor = value

    @property
    def HideSelection(self):
        return self.olkdatecontrol.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.olkdatecontrol.HideSelection = value

    @property
    def Locked(self):
        return self.olkdatecontrol.Locked

    @Locked.setter
    def Locked(self, value):
        self.olkdatecontrol.Locked = value

    @property
    def MouseIcon(self):
        return self.olkdatecontrol.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkdatecontrol.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkdatecontrol.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkdatecontrol.MousePointer = value

    @property
    def ShowNoneButton(self):
        return self.olkdatecontrol.ShowNoneButton

    @ShowNoneButton.setter
    def ShowNoneButton(self, value):
        self.olkdatecontrol.ShowNoneButton = value

    @property
    def Text(self):
        return self.olkdatecontrol.Text

    @Text.setter
    def Text(self, value):
        self.olkdatecontrol.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olkdatecontrol.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.olkdatecontrol.TextAlign = value

    @property
    def Value(self):
        return self.olkdatecontrol.Value

    @Value.setter
    def Value(self, value):
        self.olkdatecontrol.Value = value

    def DropDown(self):
        self.olkdatecontrol.DropDown()


class OlkFrameHeader:

    def __init__(self, olkframeheader=None):
        self.olkframeheader = olkframeheader

    @property
    def Alignment(self):
        return olAlignment(self.olkframeheader.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.olkframeheader.Alignment = value

    @property
    def Caption(self):
        return self.olkframeheader.Caption

    @Caption.setter
    def Caption(self, value):
        self.olkframeheader.Caption = value

    @property
    def Enabled(self):
        return self.olkframeheader.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkframeheader.Enabled = value

    @property
    def Font(self):
        return self.olkframeheader.Font

    @property
    def ForeColor(self):
        return self.olkframeheader.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkframeheader.ForeColor = value

    @property
    def MouseIcon(self):
        return self.olkframeheader.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkframeheader.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkframeheader.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkframeheader.MousePointer = value


class OlkInfoBar:

    def __init__(self, olkinfobar=None):
        self.olkinfobar = olkinfobar

    @property
    def MouseIcon(self):
        return self.olkinfobar.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkinfobar.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkinfobar.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkinfobar.MousePointer = value


class OlkLabel:

    def __init__(self, olklabel=None):
        self.olklabel = olklabel

    @property
    def Accelerator(self):
        return self.olklabel.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.olklabel.Accelerator = value

    @property
    def AutoSize(self):
        return self.olklabel.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olklabel.AutoSize = value

    @property
    def BackColor(self):
        return self.olklabel.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olklabel.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olklabel.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.olklabel.BackStyle = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olklabel.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olklabel.BorderStyle = value

    @property
    def Caption(self):
        return self.olklabel.Caption

    @Caption.setter
    def Caption(self, value):
        self.olklabel.Caption = value

    @property
    def Enabled(self):
        return self.olklabel.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olklabel.Enabled = value

    @property
    def Font(self):
        return self.olklabel.Font

    @property
    def ForeColor(self):
        return self.olklabel.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olklabel.ForeColor = value

    @property
    def MouseIcon(self):
        return self.olklabel.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olklabel.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olklabel.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olklabel.MousePointer = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olklabel.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.olklabel.TextAlign = value

    @property
    def UseHeaderColor(self):
        return self.olklabel.UseHeaderColor

    @UseHeaderColor.setter
    def UseHeaderColor(self, value):
        self.olklabel.UseHeaderColor = value

    @property
    def WordWrap(self):
        return self.olklabel.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olklabel.WordWrap = value


class OlkListBox:

    def __init__(self, olklistbox=None):
        self.olklistbox = olklistbox

    @property
    def BackColor(self):
        return self.olklistbox.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olklistbox.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olklistbox.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olklistbox.BorderStyle = value

    @property
    def Enabled(self):
        return self.olklistbox.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olklistbox.Enabled = value

    @property
    def Font(self):
        return self.olklistbox.Font

    @property
    def ForeColor(self):
        return self.olklistbox.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olklistbox.ForeColor = value

    @property
    def ListCount(self):
        return self.olklistbox.ListCount

    @property
    def ListIndex(self):
        return self.olklistbox.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.olklistbox.ListIndex = value

    @property
    def Locked(self):
        return self.olklistbox.Locked

    @Locked.setter
    def Locked(self, value):
        self.olklistbox.Locked = value

    @property
    def MatchEntry(self):
        return olMatchEntry(self.olklistbox.MatchEntry)

    @MatchEntry.setter
    def MatchEntry(self, value):
        self.olklistbox.MatchEntry = value

    @property
    def MouseIcon(self):
        return self.olklistbox.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olklistbox.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olklistbox.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olklistbox.MousePointer = value

    @property
    def MultiSelect(self):
        return OlMultiSelect(self.olklistbox.MultiSelect)

    @MultiSelect.setter
    def MultiSelect(self, value):
        self.olklistbox.MultiSelect = value

    @property
    def Text(self):
        return self.olklistbox.Text

    @Text.setter
    def Text(self, value):
        self.olklistbox.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olklistbox.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.olklistbox.TextAlign = value

    @property
    def TopIndex(self):
        return self.olklistbox.TopIndex

    @TopIndex.setter
    def TopIndex(self, value):
        self.olklistbox.TopIndex = value

    @property
    def Value(self):
        return self.olklistbox.Value

    @Value.setter
    def Value(self, value):
        self.olklistbox.Value = value

    def AddItem(self, *args, ItemText=None, Index=None):
        arguments = {"ItemText": ItemText, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.olklistbox.AddItem(*args, **arguments)

    def Clear(self):
        self.olklistbox.Clear()

    def Copy(self):
        self.olklistbox.Copy()

    def GetItem(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.olklistbox.GetItem(*args, **arguments)

    def GetSelected(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.olklistbox.GetSelected(*args, **arguments)

    def RemoveItem(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.olklistbox.RemoveItem(*args, **arguments)

    def SetItem(self, *args, Index=None, Item=None):
        arguments = {"Index": Index, "Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.olklistbox.SetItem(*args, **arguments)

    def SetSelected(self, *args, Index=None, Selected=None):
        arguments = {"Index": Index, "Selected": Selected}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.olklistbox.SetSelected(*args, **arguments)


class OlkOptionButton:

    def __init__(self, olkoptionbutton=None):
        self.olkoptionbutton = olkoptionbutton

    @property
    def Accelerator(self):
        return self.olkoptionbutton.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.olkoptionbutton.Accelerator = value

    @property
    def Alignment(self):
        return olAlignment(self.olkoptionbutton.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.olkoptionbutton.Alignment = value

    @property
    def BackColor(self):
        return self.olkoptionbutton.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkoptionbutton.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olkoptionbutton.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.olkoptionbutton.BackStyle = value

    @property
    def Caption(self):
        return self.olkoptionbutton.Caption

    @Caption.setter
    def Caption(self, value):
        self.olkoptionbutton.Caption = value

    @property
    def Enabled(self):
        return self.olkoptionbutton.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkoptionbutton.Enabled = value

    @property
    def Font(self):
        return self.olkoptionbutton.Font

    @property
    def ForeColor(self):
        return self.olkoptionbutton.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkoptionbutton.ForeColor = value

    @property
    def GroupName(self):
        return self.olkoptionbutton.GroupName

    @GroupName.setter
    def GroupName(self, value):
        self.olkoptionbutton.GroupName = value

    @property
    def MouseIcon(self):
        return self.olkoptionbutton.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkoptionbutton.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkoptionbutton.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkoptionbutton.MousePointer = value

    @property
    def Value(self):
        return self.olkoptionbutton.Value

    @Value.setter
    def Value(self, value):
        self.olkoptionbutton.Value = value

    @property
    def WordWrap(self):
        return self.olkoptionbutton.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olkoptionbutton.WordWrap = value


class OlkPageControl:

    def __init__(self, olkpagecontrol=None):
        self.olkpagecontrol = olkpagecontrol

    @property
    def Page(self):
        return OlPageType(self.olkpagecontrol.Page)

    @Page.setter
    def Page(self, value):
        self.olkpagecontrol.Page = value


class OlkSenderPhoto:

    def __init__(self, olksenderphoto=None):
        self.olksenderphoto = olksenderphoto

    @property
    def Enabled(self):
        return self.olksenderphoto.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olksenderphoto.Enabled = value

    @property
    def MouseIcon(self):
        return self.olksenderphoto.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olksenderphoto.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olksenderphoto.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olksenderphoto.MousePointer = value

    @property
    def PreferredHeight(self):
        return self.olksenderphoto.PreferredHeight

    @property
    def PreferredWidth(self):
        return self.olksenderphoto.PreferredWidth


class OlkTextBox:

    def __init__(self, olktextbox=None):
        self.olktextbox = olktextbox

    @property
    def AutoSize(self):
        return self.olktextbox.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olktextbox.AutoSize = value

    @property
    def AutoTab(self):
        return self.olktextbox.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.olktextbox.AutoTab = value

    @property
    def AutoWordSelect(self):
        return self.olktextbox.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.olktextbox.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.olktextbox.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olktextbox.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olktextbox.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olktextbox.BorderStyle = value

    @property
    def DragBehavior(self):
        return self.olktextbox.DragBehavior

    @DragBehavior.setter
    def DragBehavior(self, value):
        self.olktextbox.DragBehavior = value

    @property
    def Enabled(self):
        return self.olktextbox.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olktextbox.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.olktextbox.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.olktextbox.EnterFieldBehavior = value

    @property
    def EnterKeyBehavior(self):
        return self.olktextbox.EnterKeyBehavior

    @EnterKeyBehavior.setter
    def EnterKeyBehavior(self, value):
        self.olktextbox.EnterKeyBehavior = value

    @property
    def Font(self):
        return self.olktextbox.Font

    @property
    def ForeColor(self):
        return self.olktextbox.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olktextbox.ForeColor = value

    @property
    def HideSelection(self):
        return self.olktextbox.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.olktextbox.HideSelection = value

    @property
    def IntegralHeight(self):
        return self.olktextbox.IntegralHeight

    @IntegralHeight.setter
    def IntegralHeight(self, value):
        self.olktextbox.IntegralHeight = value

    @property
    def Locked(self):
        return self.olktextbox.Locked

    @Locked.setter
    def Locked(self, value):
        self.olktextbox.Locked = value

    @property
    def MaxLength(self):
        return self.olktextbox.MaxLength

    @MaxLength.setter
    def MaxLength(self, value):
        self.olktextbox.MaxLength = value

    @property
    def MouseIcon(self):
        return self.olktextbox.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olktextbox.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olktextbox.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olktextbox.MousePointer = value

    @property
    def Multiline(self):
        return self.olktextbox.Multiline

    @Multiline.setter
    def Multiline(self, value):
        self.olktextbox.Multiline = value

    @property
    def PasswordChar(self):
        return self.olktextbox.PasswordChar

    @PasswordChar.setter
    def PasswordChar(self, value):
        self.olktextbox.PasswordChar = value

    @property
    def Scrollbars(self):
        return olScrollBars(self.olktextbox.Scrollbars)

    @Scrollbars.setter
    def Scrollbars(self, value):
        self.olktextbox.Scrollbars = value

    @property
    def SelectionMargin(self):
        return self.olktextbox.SelectionMargin

    @SelectionMargin.setter
    def SelectionMargin(self, value):
        self.olktextbox.SelectionMargin = value

    @property
    def SelLength(self):
        return self.olktextbox.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.olktextbox.SelLength = value

    @property
    def SelStart(self):
        return self.olktextbox.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.olktextbox.SelStart = value

    @property
    def SelText(self):
        return self.olktextbox.SelText

    @property
    def TabKeyBehavior(self):
        return self.olktextbox.TabKeyBehavior

    @TabKeyBehavior.setter
    def TabKeyBehavior(self, value):
        self.olktextbox.TabKeyBehavior = value

    @property
    def Text(self):
        return self.olktextbox.Text

    @Text.setter
    def Text(self, value):
        self.olktextbox.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olktextbox.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.olktextbox.TextAlign = value

    @property
    def Value(self):
        return self.olktextbox.Value

    @Value.setter
    def Value(self, value):
        self.olktextbox.Value = value

    @property
    def WordWrap(self):
        return self.olktextbox.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olktextbox.WordWrap = value

    def Clear(self):
        self.olktextbox.Clear()

    def Copy(self):
        self.olktextbox.Copy()

    def Cut(self):
        self.olktextbox.Cut()

    def Paste(self):
        self.olktextbox.Paste()


class OlkTimeControl:

    def __init__(self, olktimecontrol=None):
        self.olktimecontrol = olktimecontrol

    @property
    def AutoSize(self):
        return self.olktimecontrol.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olktimecontrol.AutoSize = value

    @property
    def AutoWordSelect(self):
        return self.olktimecontrol.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.olktimecontrol.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.olktimecontrol.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olktimecontrol.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olktimecontrol.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.olktimecontrol.BackStyle = value

    @property
    def Enabled(self):
        return self.olktimecontrol.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olktimecontrol.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.olktimecontrol.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.olktimecontrol.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.olktimecontrol.Font

    @property
    def ForeColor(self):
        return self.olktimecontrol.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olktimecontrol.ForeColor = value

    @property
    def HideSelection(self):
        return self.olktimecontrol.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.olktimecontrol.HideSelection = value

    @property
    def IntervalTime(self):
        return self.olktimecontrol.IntervalTime

    @IntervalTime.setter
    def IntervalTime(self, value):
        self.olktimecontrol.IntervalTime = value

    @property
    def Locked(self):
        return self.olktimecontrol.Locked

    @Locked.setter
    def Locked(self, value):
        self.olktimecontrol.Locked = value

    @property
    def MouseIcon(self):
        return self.olktimecontrol.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olktimecontrol.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olktimecontrol.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olktimecontrol.MousePointer = value

    @property
    def ReferenceTime(self):
        return self.olktimecontrol.ReferenceTime

    @ReferenceTime.setter
    def ReferenceTime(self, value):
        self.olktimecontrol.ReferenceTime = value

    @property
    def Style(self):
        return OlTimeStyle(self.olktimecontrol.Style)

    @Style.setter
    def Style(self, value):
        self.olktimecontrol.Style = value

    @property
    def Text(self):
        return self.olktimecontrol.Text

    @Text.setter
    def Text(self, value):
        self.olktimecontrol.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olktimecontrol.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.olktimecontrol.TextAlign = value

    @property
    def Time(self):
        return self.olktimecontrol.Time

    @Time.setter
    def Time(self, value):
        self.olktimecontrol.Time = value

    @property
    def Value(self):
        return self.olktimecontrol.Value

    @Value.setter
    def Value(self, value):
        self.olktimecontrol.Value = value

    def DropDown(self):
        self.olktimecontrol.DropDown()


class OlkTimeZoneControl:

    def __init__(self, olktimezonecontrol=None):
        self.olktimezonecontrol = olktimezonecontrol

    @property
    def AppointmentTimeField(self):
        return OlAppointmentTimeField(self.olktimezonecontrol.AppointmentTimeField)

    @AppointmentTimeField.setter
    def AppointmentTimeField(self, value):
        self.olktimezonecontrol.AppointmentTimeField = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olktimezonecontrol.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olktimezonecontrol.BorderStyle = value

    @property
    def Enabled(self):
        return self.olktimezonecontrol.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olktimezonecontrol.Enabled = value

    @property
    def Locked(self):
        return self.olktimezonecontrol.Locked

    @Locked.setter
    def Locked(self, value):
        self.olktimezonecontrol.Locked = value

    @property
    def MouseIcon(self):
        return self.olktimezonecontrol.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olktimezonecontrol.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olktimezonecontrol.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.olktimezonecontrol.MousePointer = value

    @property
    def SelectedTimeZoneIndex(self):
        return Application.TimeZones(self.olktimezonecontrol.SelectedTimeZoneIndex)

    @SelectedTimeZoneIndex.setter
    def SelectedTimeZoneIndex(self, value):
        self.olktimezonecontrol.SelectedTimeZoneIndex = value

    @property
    def Value(self):
        return self.olktimezonecontrol.Value

    @Value.setter
    def Value(self, value):
        self.olktimezonecontrol.Value = value

    def DropDown(self):
        self.olktimezonecontrol.DropDown()


# OlMailingAddress enumeration
olBusiness = 2
olHome = 1
olNone = 0
olOther = 3

# OlMailRecipientType enumeration
olBCC = 3
olCC = 2
olOriginator = 0
olTo = 1

# OlMarkInterval enumeration
olMarkComplete = 5
olMarkNextWeek = 3
olMarkNoDate = 4
olMarkThisWeek = 2
olMarkToday = 0
olMarkTomorrow = 1

# OlMatchEntry enumeration
olMatchEntryComplete = 1
olMatchEntryFirstLetter = 0
olMatchEntryNone = 2

# OlMeetingRecipientType enumeration
olOptional = 2
olOrganizer = 0
olRequired = 1
olResource = 3

# OlMeetingResponse enumeration
olMeetingAccepted = 3
olMeetingDeclined = 4
olMeetingTentative = 2

# OlMeetingStatus enumeration
olMeeting = 1
olMeetingCanceled = 5
olMeetingReceived = 3
olMeetingReceivedAndCanceled = 7
olNonMeeting = 0

# OlMouseButton enumeration
olMouseButtonLeft = 1
olMouseButtonRight = 2
olMouseButtonMiddle = 4

# OlMousePointer enumeration
olMousePointerAppStarting = 13
olMousePointerArrow = 1
olMousePointerCross = 2
olMousePointerCustom = 99
olMousePointerDefault = 0
olMousePointerHelp = 14
olMousePointerHourGlass = 11
olMousePointerIBeam = 3
olMousePointerNoDrop = 12
olMousePointerSizeAll = 15
olMousePointerSizeNESW = 6
olMousePointerSizeNS = 7
olMousePointerSizeNWSE = 8
olMousePointerSizeWE = 9
olMousePointerUpArrow = 10

# OlMultiLine enumeration
olAlwaysMultiLine = 2
olAlwaysSingleLine = 1
olWidthMultiLine = 0

# OlMultiSelect enumeration
olMultiSelectExtended = 2
olMultiSelectMulti = 1
olMultiSelectSingle = 0

# OlNavigationModuleType enumeration
olModuleCalendar = 1
olModuleContacts = 2
olModuleFolderList = 6
olModuleJournal = 4
olModuleMail = 0
olModuleNotes = 5
olModuleShortcuts = 7
olModuleSolutions = 8
olModuleTasks = 3

# OlObjectClass enumeration
olAccount = 105
olAccountRuleCondition = 135
olAccounts = 106
olAction = 32
olActions = 33
olAddressEntries = 21
olAddressEntry = 8
olAddressList = 7
olAddressLists = 20
olAddressRuleCondition = 170
olApplication = 0
olAppointment = 26
olAssignToCategoryRuleAction = 122
olAttachment = 5
olAttachments = 18
olAttachmentSelection = 169
olAutoFormatRule = 147
olAutoFormatRules = 148
olCalendarModule = 159
olCalendarSharing = 151
olCategories = 153
olCategory = 152
olCategoryRuleCondition = 130
olClassBusinessCardView = 168
olClassCalendarView = 139
olClassCardView = 138
olClassIconView = 137
olClassNavigationPane = 155
olClassPeopleView = 183
olClassTableView = 136
olClassTimeLineView = 140
olClassTimeZone = 174
olClassTimeZones = 175
olColumn = 154
olColumnFormat = 149
olColumns = 150
olConflict = 102
olConflicts = 103
olContact = 40
olContactsModule = 160
olConversation = 178
olConversationHeader = 182
olDistributionList = 69
olDocument = 41
olException = 30
olExceptions = 29
olExchangeDistributionList = 111
olExchangeUser = 110
olExplorer = 34
olExplorers = 60
olFolder = 2
olFolders = 15
olFolderUserProperties = 172
olFolderUserProperty = 171
olFormDescription = 37
olFormNameRuleCondition = 131
olFormRegion = 129
olFromRssFeedRuleCondition = 173
olFromRuleCondition = 132
olImportanceRuleCondition = 128
olInspector = 35
olInspectors = 61
olItemProperties = 98
olItemProperty = 99
olItems = 16
olJournal = 42
olJournalModule = 162
olMail = 43
olMailModule = 158
olMarkAsTaskRuleAction = 124
olMeetingCancellation = 54
olMeetingForwardNotification = 181
olMeetingRequest = 53
olMeetingResponseNegative = 55
olMeetingResponsePositive = 56
olMeetingResponseTentative = 57
olMoveOrCopyRuleAction = 118
olNamespace = 1
olNavigationFolder = 167
olNavigationFolders = 166
olNavigationGroup = 165
olNavigationGroups = 164
olNavigationModule = 157
olNavigationModules = 156
olNewItemAlertRuleAction = 125
olNote = 44
olNotesModule = 163
olOrderField = 144
olOrderFields = 145
olOutlookBarGroup = 66
olOutlookBarGroups = 65
olOutlookBarPane = 63
olOutlookBarShortcut = 68
olOutlookBarShortcuts = 67
olOutlookBarStorage = 64
olOutspace = 180
olPages = 36
olPanes = 62
olPlaySoundRuleAction = 123
olPost = 45
olPropertyAccessor = 112
olPropertyPages = 71
olPropertyPageSite = 70
olRecipient = 4
olRecipients = 17
olRecurrencePattern = 28
olReminder = 101
olReminders = 100
olRemote = 47
olReport = 46
olResults = 78
olRow = 121
olRule = 115
olRuleAction = 117
olRuleActions = 116
olRuleCondition = 127
olRuleConditions = 126
olRules = 114
olSearch = 77
olSelection = 74
olSelectNamesDialog = 109
olSenderInAddressListRuleCondition = 133
olSendRuleAction = 119
olSharing = 104
olSimpleItems = 179
olSolutionsModule = 177
olStorageItem = 113
olStore = 107
olStores = 108
olSyncObject = 72
olSyncObjects = 73
olTable = 120
olTask = 48
olTaskRequest = 49
olTaskRequestAccept = 51
olTaskRequestDecline = 52
olTaskRequestUpdate = 50
olTasksModule = 161
olTextRuleCondition = 134
olUserDefinedProperties = 172
olUserDefinedProperty = 171
olUserProperties = 38
olUserProperty = 39
olView = 80
olViewField = 142
olViewFields = 141
olViewFont = 146
olViews = 79

# OlOutlookBarViewType enumeration
olLargeIcon = 0
olSmallIcon = 1

# OlPageType enumeration
olPageTypePlanner = 0
olPageTypeTracker = 1

# OlPane enumeration
olFolderList = 2
olNavigationPane = 4
olOutlookBar = 1
olPreview = 3

# OlPermission enumeration
olDoNotForward = 1
olPermissionTemplate = 2
olUnrestricted = 0

# OlPermissionService enumeration
olPassport = 2
olUnknown = 0
olWindows = 1

# OlPictureAlignment enumeration
olPictureAlignmentLeft = 0
olPictureAlignmentTop = 1

# OlRecipientSelectors enumeration
olShowNone = 0
olShowTo = 1
olShowToCc = 2
olShowToCcBcc = 3

# OlRecurrenceState enumeration
olApptException = 3
olApptMaster = 1
olApptNotRecurring = 0
olApptOccurrence = 2

# OlRecurrenceType enumeration
olRecursDaily = 0
olRecursMonthly = 2
olRecursMonthNth = 3
olRecursWeekly = 1
olRecursYearly = 5
olRecursYearNth = 6

# OlReferenceType enumeration
olStrong = 1
olWeak = 0

# OlRemoteStatus enumeration
olMarkedForCopy = 3
olMarkedForDelete = 4
olMarkedForDownload = 2
olRemoteStatusNone = 0
olUnMarked = 1

# OlResponseStatus enumeration
olResponseAccepted = 3
olResponseDeclined = 4
olResponseNone = 0
olResponseNotResponded = 5
olResponseOrganized = 1
olResponseTentative = 2

# OlRuleActionType enumeration
olRuleActionAssignToCategory = 2
olRuleActionCcMessage = 27
olRuleActionClearCategories = 30
olRuleActionCopyToFolder = 5
olRuleActionCustomAction = 22
olRuleActionDefer = 28
olRuleActionDelete = 3
olRuleActionDeletePermanently = 4
olRuleActionDesktopAlert = 24
olRuleActionFlagClear = 13
olRuleActionFlagColor = 12
olRuleActionFlagForActionInDays = 11
olRuleActionForward = 6
olRuleActionForwardAsAttachment = 7
olRuleActionImportance = 14
olRuleActionMarkAsTask = 41
olRuleActionMarkRead = 19
olRuleActionMoveToFolder = 1
olRuleActionNewItemAlert = 23
olRuleActionNotifyDelivery = 26
olRuleActionNotifyRead = 25
olRuleActionPlaySound = 17
olRuleActionPrint = 16
olRuleActionRedirect = 8
olRuleActionRunScript = 20
olRuleActionSensitivity = 15
olRuleActionServerReply = 9
olRuleActionStartApplication = 18
olRuleActionStop = 21
olRuleActionTemplate = 10
olRuleActionUnknown = 0

# OlRuleConditionType enumeration
olConditionAccount = 3
olConditionAnyCategory = 29
olConditionBody = 13
olConditionBodyOrSubject = 14
olConditionCategory = 18
olConditionCc = 9
olConditionDateRange = 22
olConditionFlaggedForAction = 8
olConditionFormName = 23
olConditionFrom = 1
olConditionFromAnyRssFeed = 31
olConditionFromRssFeed = 30
olConditionHasAttachment = 20
olConditionImportance = 6
olConditionLocalMachineOnly = 27
olConditionMeetingInviteOrUpdate = 26
olConditionMessageHeader = 15
olConditionNotTo = 11
olConditionOnlyToMe = 4
olConditionOOF = 19
olConditionOtherMachine = 28
olConditionProperty = 24
olConditionRecipientAddress = 16
olConditionSenderAddress = 17
olConditionSenderInAddressBook = 25
olConditionSensitivity = 7
olConditionSentTo = 12
olConditionSizeRange = 21
olConditionSubject = 2
olConditionTo = 5
olConditionToOrCc = 10
olConditionUnknown = 0

# OlRuleExecuteOption enumeration
olRuleExecuteAllMessages = 0
olRuleExecuteReadMessages = 1
olRuleExecuteUnreadMessages = 2

# OlRuleType enumeration
olRuleReceive = 0
olRuleSend = 1

# OlSaveAsType enumeration
olDoc = 4
olHTML = 5
olICal = 8
olMHTML = 10
olMSG = 3
olMSGUnicode = 9
olRTF = 1
olTemplate = 2
olTXT = 0
olVCal = 7
olVCard = 6

# OlScrollBars enumeration
olScrollBarsBoth = 3
olScrollBarsHorizontal = 1
olScrollBarsNone = 0
olScrollBarsVertical = 2

# OlSearchScope enumeration
olSearchScopeAllFolders = 1
olSearchScopeAllOutlookItems = 2
olSearchScopeCurrentFolder = 0
olSearchScopeCurrentStore = 4
olSearchScopeSubfolders = 3

# OlSelectionContents enumeration
olConversationHeaders = 1

# OlSelectionLocation enumeration
olAttachmentWell = 4
olDailyTaskList = 3
olToDoBarAppointmentList = 2
olToDoBarTaskList = 1
olViewList = 0

# OlSensitivity enumeration
olConfidential = 3
olNormal = 0
olPersonal = 1
olPrivate = 2

# OlSharingMsgType enumeration
olSharingMsgTypeInvite = 2
olSharingMsgTypeInviteAndRequest = 3
olSharingMsgTypeRequest = 1
olSharingMsgTypeResponseAllow = 4
olSharingMsgTypeResponseDeny = 5
olSharingMsgTypeUnknown = 0

# OlSharingProvider enumeration
olProviderExchange = 1
olProviderFederate = 7
olProviderICal = 4
olProviderPubCal = 3
olProviderRSS = 6
olProviderSharePoint = 5
olProviderUnknown = 0
olProviderWebCal = 2

# OlShiftState enumeration
olShiftStateShiftMask = 1
olShiftStateCtrlMask = 2
olShiftStateAltMask = 4

# OlShowItemCount enumeration
olNoItemCount = 0
olShowTotalItemCount = 2
olShowUnreadItemCount = 1

# OlSolutionScope enumeration
olHideInDefaultModules = 0
olShowInDefaultModules = 1

# OlSortOrder enumeration
olAscending = 1
olDescending = 2
olSortNone = 0

# OlSpecialFolders enumeration
olSpecialFolderAllTasks = 0
olSpecialFolderReminders = 1

# OlStorageIdentifierType enumeration
olIdentifyByEntryID = 1
olIdentifyByMessageClass = 2
olIdentifyBySubject = 0

# OlStoreType enumeration
olStoreANSI = 3
olStoreDefault = 1
olStoreUnicode = 2

# OlSyncState enumeration
olSyncStarted = 1
olSyncStopped = 0

# OlTableContents enumeration
olHiddenItems = 1
olUserItems = 0

# OlTaskDelegationState enumeration
olTaskDelegationAccepted = 2
olTaskDelegationDeclined = 3
olTaskDelegationUnknown = 1
olTaskNotDelegated = 0

# OlTaskOwnership enumeration
olDelegatedTask = 1
olNewTask = 0
olOwnTask = 2

# OlTaskRecipientType enumeration
olFinalStatus = 3
olUpdate = 2

# OlTaskResponse enumeration
olTaskAccept = 2
olTaskAssign = 1
olTaskDecline = 3
olTaskSimple = 0

# OlTaskStatus enumeration
olTaskComplete = 2
olTaskDeferred = 4
olTaskInProgress = 1
olTaskNotStarted = 0
olTaskWaiting = 3

# OlTextAlign enumeration
olTextAlignCenter = 2
olTextAlignLeft = 1
olTextAlignRight = 3

# OlTimelineViewMode enumeration
olTimelineViewDay = 0
olTimelineViewMonth = 2
olTimelineViewWeek = 1

# OlTimeStyle enumeration
olTimeStyleShortDuration = 4
olTimeStyleTimeDuration = 1
olTimeStyleTimeOnly = 0

# OlTrackingStatus enumeration
olTrackingDelivered = 1
olTrackingNone = 0
olTrackingNotDelivered = 2
olTrackingNotRead = 3
olTrackingRead = 6
olTrackingRecallFailure = 4
olTrackingRecallSuccess = 5
olTrackingReplied = 7

# OlUserPropertyType enumeration
olCombination = 19
olCurrency = 14
olDateTime = 5
olDuration = 7
olEnumeration = 21
olFormula = 18
olInteger = 20
olKeywords = 11
olNumber = 3
olOutlookInternal = 0
olPercent = 12
olSmartFrom = 22
olText = 1
olYesNo = 6

# OlVerticalLayout enumeration
olVerticalLayoutAlignBottom = 2
olVerticalLayoutAlignGrow = 3
olVerticalLayoutAlignMiddle = 1
olVerticalLayoutAlignTop = 0

# OlViewSaveOption enumeration
olViewSaveOptionAllFoldersOfType = 2
olViewSaveOptionThisFolderEveryone = 0
olViewSaveOptionThisFolderOnlyMe = 1

# OlViewType enumeration
olTableView = 0
olCardView = 1
olCalendarView = 2
olIconView = 3
olTimelineView = 4
olBusinessCardView = 5
olDailyTaskListView = 6
olPeopleView = 7

# OlWindowState enumeration
olMaximized = 0
olMinimized = 1
olNormalWindow = 2

class OrderField:

    def __init__(self, orderfield=None):
        self.orderfield = orderfield

    @property
    def Application(self):
        return Application(self.orderfield.Application)

    @property
    def Class(self):
        return OlObjectClass(self.orderfield.Class)

    @property
    def IsDescending(self):
        return OrderField(self.orderfield.IsDescending)

    @IsDescending.setter
    def IsDescending(self, value):
        self.orderfield.IsDescending = value

    @property
    def Parent(self):
        return self.orderfield.Parent

    @property
    def Session(self):
        return NameSpace(self.orderfield.Session)

    @property
    def ViewXMLSchemaName(self):
        return OrderField(self.orderfield.ViewXMLSchemaName)


class OrderFields:

    def __init__(self, orderfields=None):
        self.orderfields = orderfields

    @property
    def Application(self):
        return Application(self.orderfields.Application)

    @property
    def Class(self):
        return OlObjectClass(self.orderfields.Class)

    @property
    def Count(self):
        return OrderField(self.orderfields.Count)

    @property
    def Parent(self):
        return self.orderfields.Parent

    @property
    def Session(self):
        return NameSpace(self.orderfields.Session)

    def Add(self, *args, PropertyName=None, IsDescending=None):
        arguments = {"PropertyName": PropertyName, "IsDescending": IsDescending}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.orderfields.Add(*args, **arguments)

    def Insert(self, *args, PropertyName=None, Index=None, IsDescending=None):
        arguments = {"PropertyName": PropertyName, "Index": Index, "IsDescending": IsDescending}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.orderfields.Insert(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.orderfields.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.orderfields.Remove(*args, **arguments)

    def RemoveAll(self):
        self.orderfields.RemoveAll()


class OutlookBarGroup:

    def __init__(self, outlookbargroup=None):
        self.outlookbargroup = outlookbargroup

    @property
    def Application(self):
        return Application(self.outlookbargroup.Application)

    @property
    def Class(self):
        return OlObjectClass(self.outlookbargroup.Class)

    @property
    def Name(self):
        return self.outlookbargroup.Name

    @Name.setter
    def Name(self, value):
        self.outlookbargroup.Name = value

    @property
    def Parent(self):
        return self.outlookbargroup.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbargroup.Session)

    @property
    def Shortcuts(self):
        return OutlookBarShortcuts(self.outlookbargroup.Shortcuts)

    @property
    def ViewType(self):
        return OlOutlookBarViewType(self.outlookbargroup.ViewType)

    @ViewType.setter
    def ViewType(self, value):
        self.outlookbargroup.ViewType = value


class OutlookBarGroups:

    def __init__(self, outlookbargroups=None):
        self.outlookbargroups = outlookbargroups

    @property
    def Application(self):
        return Application(self.outlookbargroups.Application)

    @property
    def Class(self):
        return OlObjectClass(self.outlookbargroups.Class)

    @property
    def Count(self):
        return self.outlookbargroups.Count

    @property
    def Parent(self):
        return self.outlookbargroups.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbargroups.Session)

    def Add(self, *args, Name=None, Index=None):
        arguments = {"Name": Name, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return OutlookBarGroup(self.outlookbargroups.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.outlookbargroups.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.outlookbargroups.Remove(*args, **arguments)


class OutlookBarPane:

    def __init__(self, outlookbarpane=None):
        self.outlookbarpane = outlookbarpane

    @property
    def Application(self):
        return Application(self.outlookbarpane.Application)

    @property
    def Class(self):
        return OlObjectClass(self.outlookbarpane.Class)

    @property
    def Contents(self):
        return OutlookBarStorage(self.outlookbarpane.Contents)

    @property
    def Name(self):
        return self.outlookbarpane.Name

    @property
    def Parent(self):
        return self.outlookbarpane.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbarpane.Session)

    @property
    def Visible(self):
        return self.outlookbarpane.Visible

    @Visible.setter
    def Visible(self, value):
        self.outlookbarpane.Visible = value


class OutlookBarShortcut:

    def __init__(self, outlookbarshortcut=None):
        self.outlookbarshortcut = outlookbarshortcut

    @property
    def Application(self):
        return Application(self.outlookbarshortcut.Application)

    @property
    def Class(self):
        return OlObjectClass(self.outlookbarshortcut.Class)

    @property
    def Name(self):
        return self.outlookbarshortcut.Name

    @Name.setter
    def Name(self, value):
        self.outlookbarshortcut.Name = value

    @property
    def Parent(self):
        return self.outlookbarshortcut.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbarshortcut.Session)

    @property
    def Target(self):
        return self.outlookbarshortcut.Target

    def SetIcon(self, *args, Icon=None):
        arguments = {"Icon": Icon}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.outlookbarshortcut.SetIcon(*args, **arguments)


class OutlookBarShortcuts:

    def __init__(self, outlookbarshortcuts=None):
        self.outlookbarshortcuts = outlookbarshortcuts

    @property
    def Application(self):
        return Application(self.outlookbarshortcuts.Application)

    @property
    def Class(self):
        return OlObjectClass(self.outlookbarshortcuts.Class)

    @property
    def Count(self):
        return self.outlookbarshortcuts.Count

    @property
    def Parent(self):
        return self.outlookbarshortcuts.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbarshortcuts.Session)

    def Add(self, *args, Target=None, Name=None, Index=None):
        arguments = {"Target": Target, "Name": Name, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return OutlookBarShortcut(self.outlookbarshortcuts.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.outlookbarshortcuts.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.outlookbarshortcuts.Remove(*args, **arguments)


class OutlookBarStorage:

    def __init__(self, outlookbarstorage=None):
        self.outlookbarstorage = outlookbarstorage

    @property
    def Application(self):
        return Application(self.outlookbarstorage.Application)

    @property
    def Class(self):
        return OlObjectClass(self.outlookbarstorage.Class)

    @property
    def Groups(self):
        return OutlookBarGroups(self.outlookbarstorage.Groups)

    @property
    def Parent(self):
        return self.outlookbarstorage.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbarstorage.Session)


class Pages:

    def __init__(self, pages=None):
        self.pages = pages

    @property
    def Application(self):
        return Application(self.pages.Application)

    @property
    def Class(self):
        return OlObjectClass(self.pages.Class)

    @property
    def Count(self):
        return self.pages.Count

    @property
    def Parent(self):
        return self.pages.Parent

    @property
    def Session(self):
        return NameSpace(self.pages.Session)

    def Add(self):
        return Page(self.pages.Add())

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.pages.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pages.Remove(*args, **arguments)


class Panes:

    def __init__(self, panes=None):
        self.panes = panes

    @property
    def Application(self):
        return Application(self.panes.Application)

    @property
    def Class(self):
        return OlObjectClass(self.panes.Class)

    @property
    def Count(self):
        return self.panes.Count

    @property
    def Parent(self):
        return self.panes.Parent

    @property
    def Session(self):
        return NameSpace(self.panes.Session)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.panes.Item(*args, **arguments)


class PlaySoundRuleAction:

    def __init__(self, playsoundruleaction=None):
        self.playsoundruleaction = playsoundruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.playsoundruleaction.ActionType)

    @property
    def Application(self):
        return Application(self.playsoundruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.playsoundruleaction.Class)

    @property
    def Enabled(self):
        return self.playsoundruleaction.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.playsoundruleaction.Enabled = value

    @property
    def FilePath(self):
        return self.playsoundruleaction.FilePath

    @FilePath.setter
    def FilePath(self, value):
        self.playsoundruleaction.FilePath = value

    @property
    def Parent(self):
        return self.playsoundruleaction.Parent

    @property
    def Session(self):
        return NameSpace(self.playsoundruleaction.Session)


class PostItem:

    def __init__(self, postitem=None):
        self.postitem = postitem

    @property
    def Actions(self):
        return Actions(self.postitem.Actions)

    @property
    def Application(self):
        return Application(self.postitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.postitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.postitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.postitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.postitem.BillingInformation = value

    @property
    def Body(self):
        return self.postitem.Body

    @Body.setter
    def Body(self, value):
        self.postitem.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.postitem.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.postitem.BodyFormat = value

    @property
    def Categories(self):
        return self.postitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.postitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.postitem.Class)

    @property
    def Companies(self):
        return self.postitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.postitem.Companies = value

    @property
    def Conflicts(self):
        return self.postitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.postitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.postitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.postitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.postitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.postitem.DownloadState)

    @property
    def EntryID(self):
        return self.postitem.EntryID

    @property
    def ExpiryTime(self):
        return self.postitem.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.postitem.ExpiryTime = value

    @property
    def FormDescription(self):
        return FormDescription(self.postitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.postitem.GetInspector)

    @property
    def HTMLBody(self):
        return self.postitem.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.postitem.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.postitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.postitem.Importance = value

    @property
    def InternetCodepage(self):
        return self.postitem.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.postitem.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.postitem.IsConflict

    @property
    def IsMarkedAsTask(self):
        return PostItem(self.postitem.IsMarkedAsTask)

    @property
    def ItemProperties(self):
        return ItemProperties(self.postitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.postitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.postitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.postitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.postitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.postitem.MessageClass = value

    @property
    def Mileage(self):
        return self.postitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.postitem.Mileage = value

    @property
    def NoAging(self):
        return self.postitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.postitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.postitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.postitem.OutlookVersion

    @property
    def Parent(self):
        return self.postitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.postitem.PropertyAccessor)

    @property
    def ReceivedTime(self):
        return self.postitem.ReceivedTime

    @property
    def ReminderOverrideDefault(self):
        return self.postitem.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.postitem.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.postitem.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.postitem.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.postitem.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.postitem.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.postitem.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.postitem.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.postitem.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.postitem.ReminderTime = value

    @property
    def RTFBody(self):
        return self.postitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.postitem.RTFBody = value

    @property
    def Saved(self):
        return self.postitem.Saved

    @property
    def SenderEmailAddress(self):
        return self.postitem.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.postitem.SenderEmailType

    @property
    def SenderName(self):
        return self.postitem.SenderName

    @property
    def Sensitivity(self):
        return OlSensitivity(self.postitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.postitem.Sensitivity = value

    @property
    def SentOn(self):
        return self.postitem.SentOn

    @property
    def Session(self):
        return NameSpace(self.postitem.Session)

    @property
    def Size(self):
        return self.postitem.Size

    @property
    def Subject(self):
        return self.postitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.postitem.Subject = value

    @property
    def TaskCompletedDate(self):
        return PostItem(self.postitem.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.postitem.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return PostItem(self.postitem.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.postitem.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return PostItem(self.postitem.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.postitem.TaskStartDate = value

    @property
    def TaskSubject(self):
        return PostItem(self.postitem.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.postitem.TaskSubject = value

    @property
    def ToDoTaskOrdinal(self):
        return PostItem(self.postitem.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.postitem.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.postitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.postitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.postitem.UserProperties)

    def ClearConversationIndex(self):
        self.postitem.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.postitem.ClearTaskFlag()

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.postitem.Close(*args, **arguments)

    def Copy(self):
        self.postitem.Copy()

    def Delete(self):
        self.postitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.postitem.Display(*args, **arguments)

    def Forward(self):
        return self.postitem.Forward()

    def GetConversation(self):
        return self.postitem.GetConversation()

    def MarkAsTask(self, *args, MarkInterval=None):
        arguments = {"MarkInterval": MarkInterval}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.postitem.MarkAsTask(*args, **arguments)

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.postitem.Move(*args, **arguments)

    def Post(self):
        self.postitem.Post()

    def PrintOut(self):
        self.postitem.PrintOut()

    def Reply(self):
        return MailItem(self.postitem.Reply())

    def Save(self):
        self.postitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.postitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.postitem.ShowCategoriesDialog()


class PropertyAccessor:

    def __init__(self, propertyaccessor=None):
        self.propertyaccessor = propertyaccessor

    @property
    def Application(self):
        return Application(self.propertyaccessor.Application)

    @property
    def Class(self):
        return PropertyAccessor(self.propertyaccessor.Class)

    @property
    def Parent(self):
        return PropertyAccessor(self.propertyaccessor.Parent)

    @property
    def Session(self):
        return NameSpace(self.propertyaccessor.Session)

    def BinaryToString(self, *args, Value=None):
        arguments = {"Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertyaccessor.BinaryToString(*args, **arguments)

    def DeleteProperties(self, *args, SchemaNames=None):
        arguments = {"SchemaNames": SchemaNames}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Err(self.propertyaccessor.DeleteProperties(*args, **arguments))

    def DeleteProperty(self, *args, SchemaName=None):
        arguments = {"SchemaName": SchemaName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.propertyaccessor.DeleteProperty(*args, **arguments)

    def GetProperties(self, *args, SchemaNames=None):
        arguments = {"SchemaNames": SchemaNames}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertyaccessor.GetProperties(*args, **arguments)

    def GetProperty(self, *args, SchemaName=None):
        arguments = {"SchemaName": SchemaName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertyaccessor.GetProperty(*args, **arguments)

    def LocalTimeToUTC(self, *args, Value=None):
        arguments = {"Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertyaccessor.LocalTimeToUTC(*args, **arguments)

    def SetProperties(self, *args, SchemaNames=None, Values=None):
        arguments = {"SchemaNames": SchemaNames, "Values": Values}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertyaccessor.SetProperties(*args, **arguments)

    def SetProperty(self, *args, SchemaName=None, Value=None):
        arguments = {"SchemaName": SchemaName, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.propertyaccessor.SetProperty(*args, **arguments)

    def StringToBinary(self, *args, Value=None):
        arguments = {"Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertyaccessor.StringToBinary(*args, **arguments)

    def UTCToLocalTime(self, *args, Value=None):
        arguments = {"Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertyaccessor.UTCToLocalTime(*args, **arguments)


class PropertyPage:

    def __init__(self, propertypage=None):
        self.propertypage = propertypage

    def Dirty(self, *args, Dirty=None):
        arguments = {"Dirty": Dirty}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        if callable(self.propertypage.Dirty):
            return self.propertypage.Dirty(*args, **arguments)
        else:
            return self.propertypage.GetDirty(*args, **arguments)

    def Apply(self):
        return self.propertypage.Apply()

    def GetPageInfo(self, *args, HelpFile=None, HelpContext=None):
        arguments = {"HelpFile": HelpFile, "HelpContext": HelpContext}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertypage.GetPageInfo(*args, **arguments)


class PropertyPages:

    def __init__(self, propertypages=None):
        self.propertypages = propertypages

    @property
    def Application(self):
        return Application(self.propertypages.Application)

    @property
    def Class(self):
        return OlObjectClass(self.propertypages.Class)

    @property
    def Count(self):
        return self.propertypages.Count

    @property
    def Parent(self):
        return self.propertypages.Parent

    @property
    def Session(self):
        return NameSpace(self.propertypages.Session)

    def Add(self, *args, Page=None, Title=None):
        arguments = {"Page": Page, "Title": Title}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.propertypages.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.propertypages.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.propertypages.Remove(*args, **arguments)


class PropertyPageSite:

    def __init__(self, propertypagesite=None):
        self.propertypagesite = propertypagesite

    @property
    def Application(self):
        return Application(self.propertypagesite.Application)

    @property
    def Class(self):
        return OlObjectClass(self.propertypagesite.Class)

    @property
    def Parent(self):
        return self.propertypagesite.Parent

    @property
    def Session(self):
        return NameSpace(self.propertypagesite.Session)

    def OnStatusChange(self):
        self.propertypagesite.OnStatusChange()


class Recipient:

    def __init__(self, recipient=None):
        self.recipient = recipient

    @property
    def Address(self):
        return self.recipient.Address

    @property
    def AddressEntry(self):
        return AddressEntry(self.recipient.AddressEntry)

    @AddressEntry.setter
    def AddressEntry(self, value):
        self.recipient.AddressEntry = value

    @property
    def Application(self):
        return Application(self.recipient.Application)

    @property
    def AutoResponse(self):
        return Recipient(self.recipient.AutoResponse)

    @AutoResponse.setter
    def AutoResponse(self, value):
        self.recipient.AutoResponse = value

    @property
    def Class(self):
        return OlObjectClass(self.recipient.Class)

    @property
    def DisplayType(self):
        return OlDisplayType(self.recipient.DisplayType)

    @property
    def EntryID(self):
        return self.recipient.EntryID

    @property
    def Index(self):
        return self.recipient.Index

    @property
    def MeetingResponseStatus(self):
        return OlResponseStatus(self.recipient.MeetingResponseStatus)

    @property
    def Name(self):
        return self.recipient.Name

    @property
    def Parent(self):
        return self.recipient.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.recipient.PropertyAccessor)

    @property
    def Resolved(self):
        return self.recipient.Resolved

    @property
    def Sendable(self):
        return Recipient(self.recipient.Sendable)

    @Sendable.setter
    def Sendable(self, value):
        self.recipient.Sendable = value

    @property
    def Session(self):
        return NameSpace(self.recipient.Session)

    @property
    def TrackingStatus(self):
        return OlTrackingStatus(self.recipient.TrackingStatus)

    @TrackingStatus.setter
    def TrackingStatus(self, value):
        self.recipient.TrackingStatus = value

    @property
    def TrackingStatusTime(self):
        return self.recipient.TrackingStatusTime

    @TrackingStatusTime.setter
    def TrackingStatusTime(self, value):
        self.recipient.TrackingStatusTime = value

    @property
    def Type(self):
        return self.recipient.Type

    @Type.setter
    def Type(self, value):
        self.recipient.Type = value

    def Delete(self):
        self.recipient.Delete()

    def FreeBusy(self, *args, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = {"Start": Start, "MinPerChar": MinPerChar, "CompleteFormat": CompleteFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.recipient.FreeBusy(*args, **arguments)

    def Resolve(self):
        return self.recipient.Resolve()


class Recipients:

    def __init__(self, recipients=None):
        self.recipients = recipients

    @property
    def Application(self):
        return Application(self.recipients.Application)

    @property
    def Class(self):
        return OlObjectClass(self.recipients.Class)

    @property
    def Count(self):
        return self.recipients.Count

    @property
    def Parent(self):
        return self.recipients.Parent

    @property
    def Session(self):
        return NameSpace(self.recipients.Session)

    def Add(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Recipient(self.recipients.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.recipients.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.recipients.Remove(*args, **arguments)

    def ResolveAll(self):
        return self.recipients.ResolveAll()


class RecurrencePattern:

    def __init__(self, recurrencepattern=None):
        self.recurrencepattern = recurrencepattern

    @property
    def Application(self):
        return Application(self.recurrencepattern.Application)

    @property
    def Class(self):
        return OlObjectClass(self.recurrencepattern.Class)

    @property
    def DayOfMonth(self):
        return self.recurrencepattern.DayOfMonth

    @DayOfMonth.setter
    def DayOfMonth(self, value):
        self.recurrencepattern.DayOfMonth = value

    @property
    def DayOfWeekMask(self):
        return OlDaysOfWeek(self.recurrencepattern.DayOfWeekMask)

    @DayOfWeekMask.setter
    def DayOfWeekMask(self, value):
        self.recurrencepattern.DayOfWeekMask = value

    @property
    def Duration(self):
        return RecurrencePattern(self.recurrencepattern.Duration)

    @Duration.setter
    def Duration(self, value):
        self.recurrencepattern.Duration = value

    @property
    def EndTime(self):
        return self.recurrencepattern.EndTime

    @EndTime.setter
    def EndTime(self, value):
        self.recurrencepattern.EndTime = value

    @property
    def Exceptions(self):
        return Exceptions(self.recurrencepattern.Exceptions)

    @property
    def Instance(self):
        return self.recurrencepattern.Instance

    @Instance.setter
    def Instance(self, value):
        self.recurrencepattern.Instance = value

    @property
    def Interval(self):
        return self.recurrencepattern.Interval

    @Interval.setter
    def Interval(self, value):
        self.recurrencepattern.Interval = value

    @property
    def MonthOfYear(self):
        return self.recurrencepattern.MonthOfYear

    @MonthOfYear.setter
    def MonthOfYear(self, value):
        self.recurrencepattern.MonthOfYear = value

    @property
    def NoEndDate(self):
        return self.recurrencepattern.NoEndDate

    @NoEndDate.setter
    def NoEndDate(self, value):
        self.recurrencepattern.NoEndDate = value

    @property
    def Occurrences(self):
        return self.recurrencepattern.Occurrences

    @Occurrences.setter
    def Occurrences(self, value):
        self.recurrencepattern.Occurrences = value

    @property
    def Parent(self):
        return self.recurrencepattern.Parent

    @property
    def PatternEndDate(self):
        return self.recurrencepattern.PatternEndDate

    @PatternEndDate.setter
    def PatternEndDate(self, value):
        self.recurrencepattern.PatternEndDate = value

    @property
    def PatternStartDate(self):
        return self.recurrencepattern.PatternStartDate

    @PatternStartDate.setter
    def PatternStartDate(self, value):
        self.recurrencepattern.PatternStartDate = value

    @property
    def RecurrenceType(self):
        return OlRecurrenceType(self.recurrencepattern.RecurrenceType)

    @RecurrenceType.setter
    def RecurrenceType(self, value):
        self.recurrencepattern.RecurrenceType = value

    @property
    def Regenerate(self):
        return self.recurrencepattern.Regenerate

    @Regenerate.setter
    def Regenerate(self, value):
        self.recurrencepattern.Regenerate = value

    @property
    def Session(self):
        return NameSpace(self.recurrencepattern.Session)

    @property
    def StartTime(self):
        return self.recurrencepattern.StartTime

    @StartTime.setter
    def StartTime(self, value):
        self.recurrencepattern.StartTime = value

    def GetOccurrence(self, *args, StartDate=None):
        arguments = {"StartDate": StartDate}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.recurrencepattern.GetOccurrence(*args, **arguments)


class Reminder:

    def __init__(self, reminder=None):
        self.reminder = reminder

    @property
    def Application(self):
        return Application(self.reminder.Application)

    @property
    def Caption(self):
        return self.reminder.Caption

    @property
    def Class(self):
        return OlObjectClass(self.reminder.Class)

    @property
    def IsVisible(self):
        return self.reminder.IsVisible

    @property
    def Item(self):
        return self.reminder.Item

    @property
    def NextReminderDate(self):
        return self.reminder.NextReminderDate

    @property
    def OriginalReminderDate(self):
        return self.reminder.OriginalReminderDate

    @property
    def Parent(self):
        return self.reminder.Parent

    @property
    def Session(self):
        return NameSpace(self.reminder.Session)

    def Dismiss(self):
        self.reminder.Dismiss()

    def Snooze(self, *args, SnoozeTime=None):
        arguments = {"SnoozeTime": SnoozeTime}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.reminder.Snooze(*args, **arguments)


class Reminders:

    def __init__(self, reminders=None):
        self.reminders = reminders

    @property
    def Application(self):
        return Application(self.reminders.Application)

    @property
    def Class(self):
        return OlObjectClass(self.reminders.Class)

    @property
    def Count(self):
        return self.reminders.Count

    @property
    def Parent(self):
        return self.reminders.Parent

    @property
    def Session(self):
        return NameSpace(self.reminders.Session)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.reminders.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.reminders.Remove(*args, **arguments)


class RemoteItem:

    def __init__(self, remoteitem=None):
        self.remoteitem = remoteitem

    @property
    def Actions(self):
        return Actions(self.remoteitem.Actions)

    @property
    def Application(self):
        return Application(self.remoteitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.remoteitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.remoteitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.remoteitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.remoteitem.BillingInformation = value

    @property
    def Body(self):
        return self.remoteitem.Body

    @Body.setter
    def Body(self, value):
        self.remoteitem.Body = value

    @property
    def Categories(self):
        return self.remoteitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.remoteitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.remoteitem.Class)

    @property
    def Companies(self):
        return self.remoteitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.remoteitem.Companies = value

    @property
    def Conflicts(self):
        return self.remoteitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.remoteitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.remoteitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.remoteitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.remoteitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.remoteitem.DownloadState)

    @property
    def EntryID(self):
        return self.remoteitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.remoteitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.remoteitem.GetInspector)

    @property
    def HasAttachment(self):
        return self.remoteitem.HasAttachment

    @property
    def Importance(self):
        return OlImportance(self.remoteitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.remoteitem.Importance = value

    @property
    def IsConflict(self):
        return self.remoteitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.remoteitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.remoteitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.remoteitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.remoteitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.remoteitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.remoteitem.MessageClass = value

    @property
    def Mileage(self):
        return self.remoteitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.remoteitem.Mileage = value

    @property
    def NoAging(self):
        return self.remoteitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.remoteitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.remoteitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.remoteitem.OutlookVersion

    @property
    def Parent(self):
        return self.remoteitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.remoteitem.PropertyAccessor)

    @property
    def RemoteMessageClass(self):
        return self.remoteitem.RemoteMessageClass

    @property
    def Saved(self):
        return self.remoteitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.remoteitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.remoteitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.remoteitem.Session)

    @property
    def Size(self):
        return self.remoteitem.Size

    @property
    def Subject(self):
        return self.remoteitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.remoteitem.Subject = value

    @property
    def TransferSize(self):
        return self.remoteitem.TransferSize

    @property
    def TransferTime(self):
        return self.remoteitem.TransferTime

    @property
    def UnRead(self):
        return self.remoteitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.remoteitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.remoteitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.remoteitem.Close(*args, **arguments)

    def Copy(self):
        self.remoteitem.Copy()

    def Delete(self):
        self.remoteitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.remoteitem.Display(*args, **arguments)

    def GetConversation(self):
        return self.remoteitem.GetConversation()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.remoteitem.Move(*args, **arguments)

    def PrintOut(self):
        self.remoteitem.PrintOut()

    def Save(self):
        self.remoteitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.remoteitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.remoteitem.ShowCategoriesDialog()


class ReportItem:

    def __init__(self, reportitem=None):
        self.reportitem = reportitem

    @property
    def Actions(self):
        return Actions(self.reportitem.Actions)

    @property
    def Application(self):
        return Application(self.reportitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.reportitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.reportitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.reportitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.reportitem.BillingInformation = value

    @property
    def Body(self):
        return self.reportitem.Body

    @Body.setter
    def Body(self, value):
        self.reportitem.Body = value

    @property
    def Categories(self):
        return self.reportitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.reportitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.reportitem.Class)

    @property
    def Companies(self):
        return self.reportitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.reportitem.Companies = value

    @property
    def Conflicts(self):
        return self.reportitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.reportitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.reportitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.reportitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.reportitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.reportitem.DownloadState)

    @property
    def EntryID(self):
        return self.reportitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.reportitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.reportitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.reportitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.reportitem.Importance = value

    @property
    def IsConflict(self):
        return self.reportitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.reportitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.reportitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.reportitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.reportitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.reportitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.reportitem.MessageClass = value

    @property
    def Mileage(self):
        return self.reportitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.reportitem.Mileage = value

    @property
    def NoAging(self):
        return self.reportitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.reportitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.reportitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.reportitem.OutlookVersion

    @property
    def Parent(self):
        return self.reportitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.reportitem.PropertyAccessor)

    @property
    def RetentionExpirationDate(self):
        return ReportItem(self.reportitem.RetentionExpirationDate)

    @property
    def RetentionPolicyName(self):
        return self.reportitem.RetentionPolicyName

    @property
    def Saved(self):
        return self.reportitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.reportitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.reportitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.reportitem.Session)

    @property
    def Size(self):
        return self.reportitem.Size

    @property
    def Subject(self):
        return self.reportitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.reportitem.Subject = value

    @property
    def UnRead(self):
        return self.reportitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.reportitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.reportitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.reportitem.Close(*args, **arguments)

    def Copy(self):
        self.reportitem.Copy()

    def Delete(self):
        self.reportitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.reportitem.Display(*args, **arguments)

    def GetConversation(self):
        return self.reportitem.GetConversation()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.reportitem.Move(*args, **arguments)

    def PrintOut(self):
        self.reportitem.PrintOut()

    def Save(self):
        self.reportitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.reportitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.reportitem.ShowCategoriesDialog()


class Results:

    def __init__(self, results=None):
        self.results = results

    @property
    def Application(self):
        return Application(self.results.Application)

    @property
    def Class(self):
        return OlObjectClass(self.results.Class)

    @property
    def Count(self):
        return self.results.Count

    @property
    def DefaultItemType(self):
        return OlItemType(self.results.DefaultItemType)

    @DefaultItemType.setter
    def DefaultItemType(self, value):
        self.results.DefaultItemType = value

    @property
    def Parent(self):
        return self.results.Parent

    @property
    def Session(self):
        return NameSpace(self.results.Session)

    def GetFirst(self):
        return self.results.GetFirst()

    def GetLast(self):
        return self.results.GetLast()

    def GetNext(self):
        return self.results.GetNext()

    def GetPrevious(self):
        return self.results.GetPrevious()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.results.Item(*args, **arguments)

    def ResetColumns(self):
        self.results.ResetColumns()

    def SetColumns(self, *args, Columns=None):
        arguments = {"Columns": Columns}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.results.SetColumns(*args, **arguments)

    def Sort(self, *args, Property=None, Descending=None):
        arguments = {"Property": Property, "Descending": Descending}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.results.Sort(*args, **arguments)


class Row:

    def __init__(self, row=None):
        self.row = row

    @property
    def Application(self):
        return Application(self.row.Application)

    @property
    def Class(self):
        return OlObjectClass(self.row.Class)

    @property
    def Parent(self):
        return Row(self.row.Parent)

    @property
    def Session(self):
        return NameSpace(self.row.Session)

    def BinaryToString(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.row.BinaryToString(*args, **arguments)

    def GetValues(self):
        return self.row.GetValues()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.row.Item(*args, **arguments)

    def LocalTimeToUTC(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.row.LocalTimeToUTC(*args, **arguments)

    def UTCToLocalTime(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.row.UTCToLocalTime(*args, **arguments)


class Rule:

    def __init__(self, rule=None):
        self.rule = rule

    @property
    def Actions(self):
        return RuleActions(self.rule.Actions)

    @property
    def Application(self):
        return Application(self.rule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.rule.Class)

    @property
    def Conditions(self):
        return RuleConditions(self.rule.Conditions)

    @property
    def Enabled(self):
        return self.rule.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.rule.Enabled = value

    @property
    def Exceptions(self):
        return RuleConditions(self.rule.Exceptions)

    @property
    def ExecutionOrder(self):
        return Rules(self.rule.ExecutionOrder)

    @ExecutionOrder.setter
    def ExecutionOrder(self, value):
        self.rule.ExecutionOrder = value

    @property
    def IsLocalRule(self):
        return self.rule.IsLocalRule

    @property
    def Name(self):
        return self.rule.Name

    @Name.setter
    def Name(self, value):
        self.rule.Name = value

    @property
    def Parent(self):
        return self.rule.Parent

    @property
    def RuleType(self):
        return OlRuleType(self.rule.RuleType)

    @property
    def Session(self):
        return NameSpace(self.rule.Session)

    def Execute(self, *args, ShowProgress=None, Folder=None, IncludeSubfolders=None, RuleExecuteOption=None):
        arguments = {"ShowProgress": ShowProgress, "Folder": Folder, "IncludeSubfolders": IncludeSubfolders, "RuleExecuteOption": RuleExecuteOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.rule.Execute(*args, **arguments)


class RuleAction:

    def __init__(self, ruleaction=None):
        self.ruleaction = ruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.ruleaction.ActionType)

    @property
    def Application(self):
        return Application(self.ruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.ruleaction.Class)

    @property
    def Enabled(self):
        return RuleAction(self.ruleaction.Enabled)

    @Enabled.setter
    def Enabled(self, value):
        self.ruleaction.Enabled = value

    @property
    def Parent(self):
        return self.ruleaction.Parent

    @property
    def Session(self):
        return NameSpace(self.ruleaction.Session)


class RuleActions:

    def __init__(self, ruleactions=None):
        self.ruleactions = ruleactions

    @property
    def Application(self):
        return Application(self.ruleactions.Application)

    @property
    def AssignToCategory(self):
        return AssignToCategoryRuleAction(self.ruleactions.AssignToCategory)

    @property
    def CC(self):
        return SendRuleAction(self.ruleactions.CC)

    @property
    def Class(self):
        return OlObjectClass(self.ruleactions.Class)

    @property
    def ClearCategories(self):
        return RuleAction(self.ruleactions.ClearCategories)

    @property
    def CopyToFolder(self):
        return MoveOrCopyRuleAction(self.ruleactions.CopyToFolder)

    @property
    def Count(self):
        return self.ruleactions.Count

    @property
    def Delete(self):
        return RuleAction(self.ruleactions.Delete)

    @property
    def DeletePermanently(self):
        return RuleAction(self.ruleactions.DeletePermanently)

    @property
    def DesktopAlert(self):
        return RuleAction(self.ruleactions.DesktopAlert)

    @property
    def Forward(self):
        return SendRuleAction(self.ruleactions.Forward)

    @property
    def ForwardAsAttachment(self):
        return SendRuleAction(self.ruleactions.ForwardAsAttachment)

    @property
    def MarkAsTask(self):
        return MarkAsTaskRuleAction(self.ruleactions.MarkAsTask)

    @property
    def MoveToFolder(self):
        return MoveOrCopyRuleAction(self.ruleactions.MoveToFolder)

    @property
    def NewItemAlert(self):
        return NewItemAlertRuleAction(self.ruleactions.NewItemAlert)

    @property
    def NotifyDelivery(self):
        return RuleAction(self.ruleactions.NotifyDelivery)

    @property
    def NotifyRead(self):
        return RuleAction(self.ruleactions.NotifyRead)

    @property
    def Parent(self):
        return self.ruleactions.Parent

    @property
    def PlaySound(self):
        return PlaySoundRuleAction(self.ruleactions.PlaySound)

    @property
    def Redirect(self):
        return SendRuleAction(self.ruleactions.Redirect)

    @property
    def Session(self):
        return NameSpace(self.ruleactions.Session)

    @property
    def Stop(self):
        return RuleAction(self.ruleactions.Stop)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.ruleactions.Item(*args, **arguments)


class RuleCondition:

    def __init__(self, rulecondition=None):
        self.rulecondition = rulecondition

    @property
    def Application(self):
        return Application(self.rulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.rulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.rulecondition.ConditionType)

    @property
    def Enabled(self):
        return RuleCondition(self.rulecondition.Enabled)

    @Enabled.setter
    def Enabled(self, value):
        self.rulecondition.Enabled = value

    @property
    def Parent(self):
        return self.rulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.rulecondition.Session)


class RuleConditions:

    def __init__(self, ruleconditions=None):
        self.ruleconditions = ruleconditions

    @property
    def Account(self):
        return AccountRuleCondition(self.ruleconditions.Account)

    @property
    def AnyCategory(self):
        return RuleCondition(self.ruleconditions.AnyCategory)

    @property
    def Application(self):
        return Application(self.ruleconditions.Application)

    @property
    def Body(self):
        return TextRuleCondition(self.ruleconditions.Body)

    @property
    def BodyOrSubject(self):
        return TextRuleCondition(self.ruleconditions.BodyOrSubject)

    @property
    def Category(self):
        return CategoryRuleCondition(self.ruleconditions.Category)

    @property
    def CC(self):
        return RuleCondition(self.ruleconditions.CC)

    @property
    def Class(self):
        return OlObjectClass(self.ruleconditions.Class)

    @property
    def Count(self):
        return self.ruleconditions.Count

    @property
    def FormName(self):
        return FormNameRuleCondition(self.ruleconditions.FormName)

    @property
    def From(self):
        return ToOrFromRuleCondition(self.ruleconditions.From)

    @property
    def FromAnyRSSFeed(self):
        return RuleCondition(self.ruleconditions.FromAnyRSSFeed)

    @property
    def FromRssFeed(self):
        return FromRssFeedRuleCondition(self.ruleconditions.FromRssFeed)

    @property
    def HasAttachment(self):
        return RuleCondition(self.ruleconditions.HasAttachment)

    @property
    def Importance(self):
        return ImportanceRuleCondition(self.ruleconditions.Importance)

    @property
    def MeetingInviteOrUpdate(self):
        return RuleCondition(self.ruleconditions.MeetingInviteOrUpdate)

    @property
    def MessageHeader(self):
        return TextRuleCondition(self.ruleconditions.MessageHeader)

    @property
    def NotTo(self):
        return RuleCondition(self.ruleconditions.NotTo)

    @property
    def OnLocalMachine(self):
        return RuleCondition(self.ruleconditions.OnLocalMachine)

    @property
    def OnlyToMe(self):
        return RuleCondition(self.ruleconditions.OnlyToMe)

    @property
    def OnOtherMachine(self):
        return RuleCondition(self.ruleconditions.OnOtherMachine)

    @property
    def Parent(self):
        return self.ruleconditions.Parent

    @property
    def RecipientAddress(self):
        return AddressRuleCondition(self.ruleconditions.RecipientAddress)

    @property
    def SenderAddress(self):
        return AddressRuleCondition(self.ruleconditions.SenderAddress)

    @property
    def SenderInAddressList(self):
        return SenderInAddressListRuleCondition(self.ruleconditions.SenderInAddressList)

    @property
    def SentTo(self):
        return ToOrFromRuleCondition(self.ruleconditions.SentTo)

    @property
    def Session(self):
        return NameSpace(self.ruleconditions.Session)

    @property
    def Subject(self):
        return TextRuleCondition(self.ruleconditions.Subject)

    @property
    def ToMe(self):
        return RuleCondition(self.ruleconditions.ToMe)

    @property
    def ToOrCc(self):
        return RuleCondition(self.ruleconditions.ToOrCc)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.ruleconditions.Item(*args, **arguments)


class Rules:

    def __init__(self, rules=None):
        self.rules = rules

    @property
    def Application(self):
        return Application(self.rules.Application)

    @property
    def Class(self):
        return OlObjectClass(self.rules.Class)

    @property
    def Count(self):
        return self.rules.Count

    @property
    def IsRssRulesProcessingEnabled(self):
        return self.rules.IsRssRulesProcessingEnabled

    @IsRssRulesProcessingEnabled.setter
    def IsRssRulesProcessingEnabled(self, value):
        self.rules.IsRssRulesProcessingEnabled = value

    @property
    def Parent(self):
        return self.rules.Parent

    @property
    def Session(self):
        return NameSpace(self.rules.Session)

    def Create(self, *args, Name=None, RuleType=None):
        arguments = {"Name": Name, "RuleType": RuleType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.rules.Create(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.rules.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.rules.Remove(*args, **arguments)

    def Save(self, *args, ShowProgress=None):
        arguments = {"ShowProgress": ShowProgress}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.rules.Save(*args, **arguments)


class Search:

    def __init__(self, search=None):
        self.search = search

    @property
    def Application(self):
        return Application(self.search.Application)

    @property
    def Class(self):
        return OlObjectClass(self.search.Class)

    @property
    def Filter(self):
        return self.search.Filter

    @property
    def IsSynchronous(self):
        return self.search.IsSynchronous

    @property
    def Parent(self):
        return self.search.Parent

    @property
    def Results(self):
        return Results(self.search.Results)

    @property
    def Scope(self):
        return self.search.Scope

    @property
    def SearchSubFolders(self):
        return self.search.SearchSubFolders

    @property
    def Session(self):
        return NameSpace(self.search.Session)

    @property
    def Tag(self):
        return self.search.Tag

    def GetTable(self):
        return self.search.GetTable()

    def Save(self, *args, SchFldrName=None):
        arguments = {"SchFldrName": SchFldrName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.search.Save(*args, **arguments)

    def Stop(self):
        self.search.Stop()


class Selection:

    def __init__(self, selection=None):
        self.selection = selection

    @property
    def Application(self):
        return Application(self.selection.Application)

    @property
    def Class(self):
        return OlObjectClass(self.selection.Class)

    @property
    def Count(self):
        return self.selection.Count

    @property
    def Location(self):
        return OlSelectionLocation(self.selection.Location)

    @property
    def Parent(self):
        return self.selection.Parent

    @property
    def Session(self):
        return NameSpace(self.selection.Session)

    def GetSelection(self, *args, SelectionContents=None):
        arguments = {"SelectionContents": SelectionContents}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.GetSelection(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.Item(*args, **arguments)


class SelectNamesDialog:

    def __init__(self, selectnamesdialog=None):
        self.selectnamesdialog = selectnamesdialog

    @property
    def AllowMultipleSelection(self):
        return self.selectnamesdialog.AllowMultipleSelection

    @AllowMultipleSelection.setter
    def AllowMultipleSelection(self, value):
        self.selectnamesdialog.AllowMultipleSelection = value

    @property
    def Application(self):
        return Application(self.selectnamesdialog.Application)

    @property
    def BccLabel(self):
        return self.selectnamesdialog.BccLabel

    @BccLabel.setter
    def BccLabel(self, value):
        self.selectnamesdialog.BccLabel = value

    @property
    def Caption(self):
        return self.selectnamesdialog.Caption

    @Caption.setter
    def Caption(self, value):
        self.selectnamesdialog.Caption = value

    @property
    def CcLabel(self):
        return self.selectnamesdialog.CcLabel

    @CcLabel.setter
    def CcLabel(self, value):
        self.selectnamesdialog.CcLabel = value

    @property
    def Class(self):
        return OlObjectClass(self.selectnamesdialog.Class)

    @property
    def ForceResolution(self):
        return SelectNamesDialog.Recipients(self.selectnamesdialog.ForceResolution)

    @ForceResolution.setter
    def ForceResolution(self, value):
        self.selectnamesdialog.ForceResolution = value

    @property
    def InitialAddressList(self):
        return AddressList(self.selectnamesdialog.InitialAddressList)

    @InitialAddressList.setter
    def InitialAddressList(self, value):
        self.selectnamesdialog.InitialAddressList = value

    @property
    def NumberOfRecipientSelectors(self):
        return OlRecipientSelectors(self.selectnamesdialog.NumberOfRecipientSelectors)

    @NumberOfRecipientSelectors.setter
    def NumberOfRecipientSelectors(self, value):
        self.selectnamesdialog.NumberOfRecipientSelectors = value

    @property
    def Parent(self):
        return SelectNamesDialog(self.selectnamesdialog.Parent)

    @property
    def Recipients(self):
        return Recipients(self.selectnamesdialog.Recipients)

    @Recipients.setter
    def Recipients(self, value):
        self.selectnamesdialog.Recipients = value

    @property
    def Session(self):
        return NameSpace(self.selectnamesdialog.Session)

    @property
    def ShowOnlyInitialAddressList(self):
        return AddressList(self.selectnamesdialog.ShowOnlyInitialAddressList)

    @ShowOnlyInitialAddressList.setter
    def ShowOnlyInitialAddressList(self, value):
        self.selectnamesdialog.ShowOnlyInitialAddressList = value

    @property
    def ToLabel(self):
        return self.selectnamesdialog.ToLabel

    @ToLabel.setter
    def ToLabel(self, value):
        self.selectnamesdialog.ToLabel = value

    def Display(self):
        return self.selectnamesdialog.Display()

    def SetDefaultDisplayMode(self, *args, defaultMode=None):
        arguments = {"defaultMode": defaultMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selectnamesdialog.SetDefaultDisplayMode(*args, **arguments)


class SenderInAddressListRuleCondition:

    def __init__(self, senderinaddresslistrulecondition=None):
        self.senderinaddresslistrulecondition = senderinaddresslistrulecondition

    @property
    def AddressList(self):
        return AddressList(self.senderinaddresslistrulecondition.AddressList)

    @AddressList.setter
    def AddressList(self, value):
        self.senderinaddresslistrulecondition.AddressList = value

    @property
    def Application(self):
        return Application(self.senderinaddresslistrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.senderinaddresslistrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.senderinaddresslistrulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.senderinaddresslistrulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.senderinaddresslistrulecondition.Enabled = value

    @property
    def Parent(self):
        return self.senderinaddresslistrulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.senderinaddresslistrulecondition.Session)


class SendRuleAction:

    def __init__(self, sendruleaction=None):
        self.sendruleaction = sendruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.sendruleaction.ActionType)

    @property
    def Application(self):
        return Application(self.sendruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.sendruleaction.Class)

    @property
    def Enabled(self):
        return self.sendruleaction.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.sendruleaction.Enabled = value

    @property
    def Parent(self):
        return self.sendruleaction.Parent

    @property
    def Recipients(self):
        return Recipients(self.sendruleaction.Recipients)

    @property
    def Session(self):
        return NameSpace(self.sendruleaction.Session)


class SharingItem:

    def __init__(self, sharingitem=None):
        self.sharingitem = sharingitem

    @property
    def Actions(self):
        return Actions(self.sharingitem.Actions)

    @property
    def AllowWriteAccess(self):
        return self.sharingitem.AllowWriteAccess

    @AllowWriteAccess.setter
    def AllowWriteAccess(self, value):
        self.sharingitem.AllowWriteAccess = value

    @property
    def AlternateRecipientAllowed(self):
        return self.sharingitem.AlternateRecipientAllowed

    @AlternateRecipientAllowed.setter
    def AlternateRecipientAllowed(self, value):
        self.sharingitem.AlternateRecipientAllowed = value

    @property
    def Application(self):
        return Application(self.sharingitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.sharingitem.Attachments)

    @property
    def AutoForwarded(self):
        return self.sharingitem.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.sharingitem.AutoForwarded = value

    @property
    def BCC(self):
        return SharingItem(self.sharingitem.BCC)

    @BCC.setter
    def BCC(self, value):
        self.sharingitem.BCC = value

    @property
    def BillingInformation(self):
        return SharingItem(self.sharingitem.BillingInformation)

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.sharingitem.BillingInformation = value

    @property
    def Body(self):
        return SharingItem(self.sharingitem.Body)

    @Body.setter
    def Body(self, value):
        self.sharingitem.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.sharingitem.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.sharingitem.BodyFormat = value

    @property
    def Categories(self):
        return SharingItem(self.sharingitem.Categories)

    @Categories.setter
    def Categories(self, value):
        self.sharingitem.Categories = value

    @property
    def CC(self):
        return SharingItem(self.sharingitem.CC)

    @CC.setter
    def CC(self, value):
        self.sharingitem.CC = value

    @property
    def Class(self):
        return OlObjectClass(self.sharingitem.Class)

    @property
    def Companies(self):
        return SharingItem(self.sharingitem.Companies)

    @Companies.setter
    def Companies(self, value):
        self.sharingitem.Companies = value

    @property
    def Conflicts(self):
        return self.sharingitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.sharingitem.ConversationID)

    @property
    def ConversationIndex(self):
        return SharingItem(self.sharingitem.ConversationIndex)

    @property
    def ConversationTopic(self):
        return SharingItem(self.sharingitem.ConversationTopic)

    @property
    def CreationTime(self):
        return SharingItem(self.sharingitem.CreationTime)

    @property
    def DeferredDeliveryTime(self):
        return SharingItem(self.sharingitem.DeferredDeliveryTime)

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.sharingitem.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.sharingitem.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.sharingitem.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.sharingitem.DownloadState)

    @property
    def EntryID(self):
        return SharingItem(self.sharingitem.EntryID)

    @property
    def ExpiryTime(self):
        return SharingItem(self.sharingitem.ExpiryTime)

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.sharingitem.ExpiryTime = value

    @property
    def FlagRequest(self):
        return SharingItem(self.sharingitem.FlagRequest)

    @FlagRequest.setter
    def FlagRequest(self, value):
        self.sharingitem.FlagRequest = value

    @property
    def FormDescription(self):
        return FormDescription(self.sharingitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.sharingitem.GetInspector)

    @property
    def HTMLBody(self):
        return SharingItem(self.sharingitem.HTMLBody)

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.sharingitem.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.sharingitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.sharingitem.Importance = value

    @property
    def InternetCodepage(self):
        return self.sharingitem.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.sharingitem.InternetCodepage = value

    @property
    def IsConflict(self):
        return SharingItem(self.sharingitem.IsConflict)

    @property
    def IsMarkedAsTask(self):
        return SharingItem(self.sharingitem.IsMarkedAsTask)

    @property
    def ItemProperties(self):
        return ItemProperties(self.sharingitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return SharingItem(self.sharingitem.LastModificationTime)

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.sharingitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.sharingitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return SharingItem(self.sharingitem.MessageClass)

    @MessageClass.setter
    def MessageClass(self, value):
        self.sharingitem.MessageClass = value

    @property
    def Mileage(self):
        return SharingItem(self.sharingitem.Mileage)

    @Mileage.setter
    def Mileage(self, value):
        self.sharingitem.Mileage = value

    @property
    def NoAging(self):
        return SharingItem(self.sharingitem.NoAging)

    @NoAging.setter
    def NoAging(self, value):
        self.sharingitem.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return SharingItem(self.sharingitem.OriginatorDeliveryReportRequested)

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.sharingitem.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return SharingItem(self.sharingitem.OutlookInternalVersion)

    @property
    def OutlookVersion(self):
        return SharingItem(self.sharingitem.OutlookVersion)

    @property
    def Parent(self):
        return SharingItem(self.sharingitem.Parent)

    @property
    def Permission(self):
        return self.sharingitem.Permission

    @Permission.setter
    def Permission(self, value):
        self.sharingitem.Permission = value

    @property
    def PermissionService(self):
        return self.sharingitem.PermissionService

    @PermissionService.setter
    def PermissionService(self, value):
        self.sharingitem.PermissionService = value

    @property
    def PermissionTemplateGuid(self):
        return SharingItem(self.sharingitem.PermissionTemplateGuid)

    @PermissionTemplateGuid.setter
    def PermissionTemplateGuid(self, value):
        self.sharingitem.PermissionTemplateGuid = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.sharingitem.PropertyAccessor)

    @property
    def ReadReceiptRequested(self):
        return self.sharingitem.ReadReceiptRequested

    @property
    def ReceivedByEntryID(self):
        return self.sharingitem.ReceivedByEntryID

    @property
    def ReceivedByName(self):
        return SharingItem(self.sharingitem.ReceivedByName)

    @property
    def ReceivedOnBehalfOfEntryID(self):
        return self.sharingitem.ReceivedOnBehalfOfEntryID

    @property
    def ReceivedOnBehalfOfName(self):
        return SharingItem(self.sharingitem.ReceivedOnBehalfOfName)

    @property
    def ReceivedTime(self):
        return SharingItem(self.sharingitem.ReceivedTime)

    @property
    def RecipientReassignmentProhibited(self):
        return SharingItem(self.sharingitem.RecipientReassignmentProhibited)

    @RecipientReassignmentProhibited.setter
    def RecipientReassignmentProhibited(self, value):
        self.sharingitem.RecipientReassignmentProhibited = value

    @property
    def Recipients(self):
        return Recipients(self.sharingitem.Recipients)

    @property
    def ReminderOverrideDefault(self):
        return SharingItem(self.sharingitem.ReminderOverrideDefault)

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.sharingitem.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return SharingItem(self.sharingitem.ReminderPlaySound)

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.sharingitem.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return SharingItem(self.sharingitem.ReminderSet)

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.sharingitem.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.sharingitem.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.sharingitem.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return SharingItem(self.sharingitem.ReminderTime)

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.sharingitem.ReminderTime = value

    @property
    def RemoteID(self):
        return SharingItem(self.sharingitem.RemoteID)

    @property
    def RemoteName(self):
        return SharingItem(self.sharingitem.RemoteName)

    @property
    def RemotePath(self):
        return SharingItem(self.sharingitem.RemotePath)

    @property
    def RemoteStatus(self):
        return OlRemoteStatus(self.sharingitem.RemoteStatus)

    @RemoteStatus.setter
    def RemoteStatus(self, value):
        self.sharingitem.RemoteStatus = value

    @property
    def ReplyRecipientNames(self):
        return SharingItem(self.sharingitem.ReplyRecipientNames)

    @property
    def ReplyRecipients(self):
        return Recipients(self.sharingitem.ReplyRecipients)

    @property
    def RequestedFolder(self):
        return OlDefaultFolders(self.sharingitem.RequestedFolder)

    @property
    def RetentionExpirationDate(self):
        return SharingItem(self.sharingitem.RetentionExpirationDate)

    @property
    def RetentionPolicyName(self):
        return self.sharingitem.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.sharingitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.sharingitem.RTFBody = value

    @property
    def Saved(self):
        return SharingItem(self.sharingitem.Saved)

    @property
    def SaveSentMessageFolder(self):
        return Folder(self.sharingitem.SaveSentMessageFolder)

    @SaveSentMessageFolder.setter
    def SaveSentMessageFolder(self, value):
        self.sharingitem.SaveSentMessageFolder = value

    @property
    def SenderEmailAddress(self):
        return SharingItem(self.sharingitem.SenderEmailAddress)

    @property
    def SenderEmailType(self):
        return SharingItem(self.sharingitem.SenderEmailType)

    @property
    def SenderName(self):
        return SharingItem(self.sharingitem.SenderName)

    @property
    def SendUsingAccount(self):
        return Account(self.sharingitem.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.sharingitem.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.sharingitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.sharingitem.Sensitivity = value

    @property
    def Sent(self):
        return SharingItem(self.sharingitem.Sent)

    @property
    def SentOn(self):
        return SharingItem(self.sharingitem.SentOn)

    @property
    def SentOnBehalfOfName(self):
        return SharingItem(self.sharingitem.SentOnBehalfOfName)

    @SentOnBehalfOfName.setter
    def SentOnBehalfOfName(self, value):
        self.sharingitem.SentOnBehalfOfName = value

    @property
    def Session(self):
        return NameSpace(self.sharingitem.Session)

    @property
    def SharingProvider(self):
        return OlSharingProvider(self.sharingitem.SharingProvider)

    @property
    def SharingProviderGuid(self):
        return SharingItem(self.sharingitem.SharingProviderGuid)

    @property
    def Size(self):
        return SharingItem(self.sharingitem.Size)

    @property
    def Subject(self):
        return SharingItem(self.sharingitem.Subject)

    @Subject.setter
    def Subject(self, value):
        self.sharingitem.Subject = value

    @property
    def Submitted(self):
        return SharingItem(self.sharingitem.Submitted)

    @property
    def TaskCompletedDate(self):
        return SharingItem(self.sharingitem.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.sharingitem.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return SharingItem(self.sharingitem.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.sharingitem.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return SharingItem(self.sharingitem.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.sharingitem.TaskStartDate = value

    @property
    def TaskSubject(self):
        return SharingItem(self.sharingitem.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.sharingitem.TaskSubject = value

    @property
    def To(self):
        return SharingItem(self.sharingitem.To)

    @To.setter
    def To(self, value):
        self.sharingitem.To = value

    @property
    def ToDoTaskOrdinal(self):
        return SharingItem(self.sharingitem.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.sharingitem.ToDoTaskOrdinal = value

    @property
    def Type(self):
        return OlSharingMsgType(self.sharingitem.Type)

    @Type.setter
    def Type(self, value):
        self.sharingitem.Type = value

    @property
    def UnRead(self):
        return SharingItem(self.sharingitem.UnRead)

    @UnRead.setter
    def UnRead(self, value):
        self.sharingitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.sharingitem.UserProperties)

    def AddBusinessCard(self, *args, contact=None):
        arguments = {"contact": contact}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sharingitem.AddBusinessCard(*args, **arguments)

    def Allow(self):
        self.sharingitem.Allow()

    def ClearConversationIndex(self):
        self.sharingitem.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.sharingitem.ClearTaskFlag()

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sharingitem.Close(*args, **arguments)

    def Copy(self):
        self.sharingitem.Copy()

    def Delete(self):
        self.sharingitem.Delete()

    def Deny(self):
        return self.sharingitem.Deny()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sharingitem.Display(*args, **arguments)

    def Forward(self):
        return self.sharingitem.Forward()

    def GetConversation(self):
        return self.sharingitem.GetConversation()

    def MarkAsTask(self, *args, MarkInterval=None):
        arguments = {"MarkInterval": MarkInterval}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sharingitem.MarkAsTask(*args, **arguments)

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sharingitem.Move(*args, **arguments)

    def OpenSharedFolder(self):
        return Folder(self.sharingitem.OpenSharedFolder())

    def PrintOut(self):
        self.sharingitem.PrintOut()

    def Reply(self):
        return MailItem(self.sharingitem.Reply())

    def ReplyAll(self):
        return MailItem(self.sharingitem.ReplyAll())

    def Save(self):
        self.sharingitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sharingitem.SaveAs(*args, **arguments)

    def Send(self):
        self.sharingitem.Send()

    def ShowCategoriesDialog(self):
        self.sharingitem.ShowCategoriesDialog()


class SimpleItems:

    def __init__(self, simpleitems=None):
        self.simpleitems = simpleitems

    @property
    def Application(self):
        return Application(self.simpleitems.Application)

    @property
    def Class(self):
        return OlObjectClass(self.simpleitems.Class)

    @property
    def Count(self):
        return SimpleItems(self.simpleitems.Count)

    @property
    def Parent(self):
        return SimpleItems(self.simpleitems.Parent)

    @property
    def Session(self):
        return NameSpace(self.simpleitems.Session)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.simpleitems.Item(*args, **arguments)


class SolutionsModule:

    def __init__(self, solutionsmodule=None):
        self.solutionsmodule = solutionsmodule

    @property
    def Application(self):
        return Application(self.solutionsmodule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.solutionsmodule.Class)

    @property
    def Name(self):
        return SolutionsModule(self.solutionsmodule.Name)

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.solutionsmodule.NavigationModuleType)

    @property
    def Parent(self):
        return SolutionsModule(self.solutionsmodule.Parent)

    @property
    def Position(self):
        return SolutionsModule(self.solutionsmodule.Position)

    @Position.setter
    def Position(self, value):
        self.solutionsmodule.Position = value

    @property
    def Session(self):
        return NameSpace(self.solutionsmodule.Session)

    @property
    def Visible(self):
        return self.solutionsmodule.Visible

    @Visible.setter
    def Visible(self, value):
        self.solutionsmodule.Visible = value

    def AddSolution(self, *args, Solution=None, Scope=None):
        arguments = {"Solution": Solution, "Scope": Scope}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.solutionsmodule.AddSolution(*args, **arguments)


class StorageItem:

    def __init__(self, storageitem=None):
        self.storageitem = storageitem

    @property
    def Application(self):
        return Application(self.storageitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.storageitem.Attachments)

    @property
    def Body(self):
        return self.storageitem.Body

    @Body.setter
    def Body(self, value):
        self.storageitem.Body = value

    @property
    def Class(self):
        return OlObjectClass(self.storageitem.Class)

    @property
    def CreationTime(self):
        return StorageItem(self.storageitem.CreationTime)

    @property
    def Creator(self):
        return StorageItem(self.storageitem.Creator)

    @Creator.setter
    def Creator(self, value):
        self.storageitem.Creator = value

    @property
    def EntryID(self):
        return self.storageitem.EntryID

    @property
    def LastModificationTime(self):
        return self.storageitem.LastModificationTime

    @property
    def Parent(self):
        return self.storageitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.storageitem.PropertyAccessor)

    @property
    def Session(self):
        return NameSpace(self.storageitem.Session)

    @property
    def Size(self):
        return StorageItem(self.storageitem.Size)

    @property
    def Subject(self):
        return self.storageitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.storageitem.Subject = value

    @property
    def UserProperties(self):
        return UserProperties(self.storageitem.UserProperties)

    def Delete(self):
        self.storageitem.Delete()

    def Save(self):
        self.storageitem.Save()


class Store:

    def __init__(self, store=None):
        self.store = store

    @property
    def Application(self):
        return Application(self.store.Application)

    @property
    def Categories(self):
        return Categories(self.store.Categories)

    @property
    def Class(self):
        return OlObjectClass(self.store.Class)

    @property
    def DisplayName(self):
        return Store(self.store.DisplayName)

    @property
    def ExchangeStoreType(self):
        return OlExchangeStoreType(self.store.ExchangeStoreType)

    @property
    def FilePath(self):
        return self.store.FilePath

    @property
    def IsCachedExchange(self):
        return Store(self.store.IsCachedExchange)

    @property
    def IsConversationEnabled(self):
        return self.store.IsConversationEnabled

    @property
    def IsDataFileStore(self):
        return Store(self.store.IsDataFileStore)

    @property
    def IsInstantSearchEnabled(self):
        return self.store.IsInstantSearchEnabled

    @property
    def IsOpen(self):
        return Store(self.store.IsOpen)

    @property
    def Parent(self):
        return self.store.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.store.PropertyAccessor)

    @property
    def Session(self):
        return NameSpace(self.store.Session)

    @property
    def StoreID(self):
        return Store(self.store.StoreID)

    def GetDefaultFolder(self, *args, FolderType=None):
        arguments = {"FolderType": FolderType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.store.GetDefaultFolder(*args, **arguments)

    def GetRootFolder(self):
        return self.store.GetRootFolder()

    def GetRules(self):
        return self.store.GetRules()

    def GetSearchFolders(self):
        return self.store.GetSearchFolders()

    def GetSpecialFolder(self, *args, FolderType=None):
        arguments = {"FolderType": FolderType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.store.GetSpecialFolder(*args, **arguments)

    def RefreshQuotaDisplay(self):
        self.store.RefreshQuotaDisplay()


class Stores:

    def __init__(self, stores=None):
        self.stores = stores

    @property
    def Application(self):
        return Application(self.stores.Application)

    @property
    def Class(self):
        return OlObjectClass(self.stores.Class)

    @property
    def Count(self):
        return self.stores.Count

    @property
    def Parent(self):
        return self.stores.Parent

    @property
    def Session(self):
        return NameSpace(self.stores.Session)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Stores(self.stores.Item(*args, **arguments))


class SyncObject:

    def __init__(self, syncobject=None):
        self.syncobject = syncobject

    @property
    def Application(self):
        return Application(self.syncobject.Application)

    @property
    def Class(self):
        return OlObjectClass(self.syncobject.Class)

    @property
    def Name(self):
        return self.syncobject.Name

    @property
    def Parent(self):
        return self.syncobject.Parent

    @property
    def Session(self):
        return NameSpace(self.syncobject.Session)

    def Start(self):
        self.syncobject.Start()

    def Stop(self):
        self.syncobject.Stop()


class SyncObjects:

    def __init__(self, syncobjects=None):
        self.syncobjects = syncobjects

    @property
    def AppFolders(self):
        return self.syncobjects.AppFolders

    @property
    def Application(self):
        return Application(self.syncobjects.Application)

    @property
    def Class(self):
        return OlObjectClass(self.syncobjects.Class)

    @property
    def Count(self):
        return self.syncobjects.Count

    @property
    def Parent(self):
        return self.syncobjects.Parent

    @property
    def Session(self):
        return NameSpace(self.syncobjects.Session)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.syncobjects.Item(*args, **arguments)


class Table:

    def __init__(self, table=None):
        self.table = table

    @property
    def Application(self):
        return Application(self.table.Application)

    @property
    def Class(self):
        return OlObjectClass(self.table.Class)

    @property
    def Columns(self):
        return Columns(self.table.Columns)

    @property
    def EndOfTable(self):
        return Table(self.table.EndOfTable)

    @property
    def Parent(self):
        return Table(self.table.Parent)

    @property
    def Session(self):
        return NameSpace(self.table.Session)

    def FindNextRow(self):
        return Row(self.table.FindNextRow())

    def FindRow(self, *args, Filter=None):
        arguments = {"Filter": Filter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Row(self.table.FindRow(*args, **arguments))

    def GetArray(self, *args, MaxRows=None):
        arguments = {"MaxRows": MaxRows}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.table.GetArray(*args, **arguments)

    def GetNextRow(self):
        return Row(self.table.GetNextRow())

    def GetRowCount(self):
        return self.table.GetRowCount()

    def MoveToStart(self):
        self.table.MoveToStart()

    def Restrict(self, *args, Filter=None):
        arguments = {"Filter": Filter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.table.Restrict(*args, **arguments)

    def Sort(self, *args, SortProperty=None, Descending=None):
        arguments = {"SortProperty": SortProperty, "Descending": Descending}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.table.Sort(*args, **arguments)


class TableView:

    def __init__(self, tableview=None):
        self.tableview = tableview

    @property
    def AllowInCellEditing(self):
        return TableView(self.tableview.AllowInCellEditing)

    @AllowInCellEditing.setter
    def AllowInCellEditing(self, value):
        self.tableview.AllowInCellEditing = value

    @property
    def AlwaysExpandConversation(self):
        return self.tableview.AlwaysExpandConversation

    @AlwaysExpandConversation.setter
    def AlwaysExpandConversation(self, value):
        self.tableview.AlwaysExpandConversation = value

    @property
    def Application(self):
        return Application(self.tableview.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.tableview.AutoFormatRules)

    @property
    def AutomaticColumnSizing(self):
        return TableView(self.tableview.AutomaticColumnSizing)

    @AutomaticColumnSizing.setter
    def AutomaticColumnSizing(self, value):
        self.tableview.AutomaticColumnSizing = value

    @property
    def AutomaticGrouping(self):
        return TableView(self.tableview.AutomaticGrouping)

    @AutomaticGrouping.setter
    def AutomaticGrouping(self, value):
        self.tableview.AutomaticGrouping = value

    @property
    def AutoPreview(self):
        return OlAutoPreview(self.tableview.AutoPreview)

    @AutoPreview.setter
    def AutoPreview(self, value):
        self.tableview.AutoPreview = value

    @property
    def AutoPreviewFont(self):
        return ViewFont(self.tableview.AutoPreviewFont)

    @property
    def Class(self):
        return OlObjectClass(self.tableview.Class)

    @property
    def ColumnFont(self):
        return ViewFont(self.tableview.ColumnFont)

    @property
    def DefaultExpandCollapseSetting(self):
        return OlDefaultExpandCollapseSetting(self.tableview.DefaultExpandCollapseSetting)

    @DefaultExpandCollapseSetting.setter
    def DefaultExpandCollapseSetting(self, value):
        self.tableview.DefaultExpandCollapseSetting = value

    @property
    def Filter(self):
        return self.tableview.Filter

    @Filter.setter
    def Filter(self, value):
        self.tableview.Filter = value

    @property
    def GridLineStyle(self):
        return OlGridLineStyle(self.tableview.GridLineStyle)

    @GridLineStyle.setter
    def GridLineStyle(self, value):
        self.tableview.GridLineStyle = value

    @property
    def GroupByFields(self):
        return OrderFields(self.tableview.GroupByFields)

    @property
    def HideReadingPaneHeaderInfo(self):
        return TableView(self.tableview.HideReadingPaneHeaderInfo)

    @HideReadingPaneHeaderInfo.setter
    def HideReadingPaneHeaderInfo(self, value):
        self.tableview.HideReadingPaneHeaderInfo = value

    @property
    def Language(self):
        return self.tableview.Language

    @Language.setter
    def Language(self, value):
        self.tableview.Language = value

    @property
    def LockUserChanges(self):
        return self.tableview.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.tableview.LockUserChanges = value

    @property
    def MaxLinesInMultiLineView(self):
        return TableView(self.tableview.MaxLinesInMultiLineView)

    @MaxLinesInMultiLineView.setter
    def MaxLinesInMultiLineView(self, value):
        self.tableview.MaxLinesInMultiLineView = value

    @property
    def Multiline(self):
        return OlMultiLine(self.tableview.Multiline)

    @Multiline.setter
    def Multiline(self, value):
        self.tableview.Multiline = value

    @property
    def MultiLineWidth(self):
        return TableView(self.tableview.MultiLineWidth)

    @MultiLineWidth.setter
    def MultiLineWidth(self, value):
        self.tableview.MultiLineWidth = value

    @property
    def Name(self):
        return self.tableview.Name

    @Name.setter
    def Name(self, value):
        self.tableview.Name = value

    @property
    def Parent(self):
        return self.tableview.Parent

    @property
    def RowFont(self):
        return ViewFont(self.tableview.RowFont)

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.tableview.SaveOption)

    @property
    def Session(self):
        return NameSpace(self.tableview.Session)

    @property
    def ShowConversationByDate(self):
        return self.tableview.ShowConversationByDate

    @ShowConversationByDate.setter
    def ShowConversationByDate(self, value):
        self.tableview.ShowConversationByDate = value

    @property
    def ShowConversationSendersAboveSubject(self):
        return self.tableview.ShowConversationSendersAboveSubject

    @ShowConversationSendersAboveSubject.setter
    def ShowConversationSendersAboveSubject(self, value):
        self.tableview.ShowConversationSendersAboveSubject = value

    @property
    def ShowFullConversations(self):
        return self.tableview.ShowFullConversations

    @ShowFullConversations.setter
    def ShowFullConversations(self, value):
        self.tableview.ShowFullConversations = value

    @property
    def ShowItemsInGroups(self):
        return TableView(self.tableview.ShowItemsInGroups)

    @ShowItemsInGroups.setter
    def ShowItemsInGroups(self, value):
        self.tableview.ShowItemsInGroups = value

    @property
    def ShowNewItemRow(self):
        return TableView(self.tableview.ShowNewItemRow)

    @ShowNewItemRow.setter
    def ShowNewItemRow(self, value):
        self.tableview.ShowNewItemRow = value

    @property
    def ShowReadingPane(self):
        return TableView(self.tableview.ShowReadingPane)

    @ShowReadingPane.setter
    def ShowReadingPane(self, value):
        self.tableview.ShowReadingPane = value

    @property
    def SortFields(self):
        return OrderFields(self.tableview.SortFields)

    @property
    def Standard(self):
        return TableView(self.tableview.Standard)

    @property
    def ViewFields(self):
        return ViewFields(self.tableview.ViewFields)

    @property
    def ViewType(self):
        return OlViewType(self.tableview.ViewType)

    @property
    def XML(self):
        return self.tableview.XML

    @XML.setter
    def XML(self, value):
        self.tableview.XML = value

    def Apply(self):
        self.tableview.Apply()

    def Copy(self, *args, Name=None, SaveOption=None):
        arguments = {"Name": Name, "SaveOption": SaveOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.tableview.Copy(*args, **arguments)

    def Delete(self):
        self.tableview.Delete()

    def GetTable(self):
        return self.tableview.GetTable()

    def GoToDate(self, *args, Date=None):
        arguments = {"Date": Date}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.tableview.GoToDate(*args, **arguments)

    def Reset(self):
        self.tableview.Reset()

    def Save(self):
        self.tableview.Save()


class TaskItem:

    def __init__(self, taskitem=None):
        self.taskitem = taskitem

    @property
    def Actions(self):
        return Actions(self.taskitem.Actions)

    @property
    def ActualWork(self):
        return self.taskitem.ActualWork

    @ActualWork.setter
    def ActualWork(self, value):
        self.taskitem.ActualWork = value

    @property
    def Application(self):
        return Application(self.taskitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.taskitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskitem.BillingInformation = value

    @property
    def Body(self):
        return self.taskitem.Body

    @Body.setter
    def Body(self, value):
        self.taskitem.Body = value

    @property
    def CardData(self):
        return self.taskitem.CardData

    @CardData.setter
    def CardData(self, value):
        self.taskitem.CardData = value

    @property
    def Categories(self):
        return self.taskitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskitem.Class)

    @property
    def Companies(self):
        return self.taskitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskitem.Companies = value

    @property
    def Complete(self):
        return self.taskitem.Complete

    @Complete.setter
    def Complete(self, value):
        self.taskitem.Complete = value

    @property
    def Conflicts(self):
        return self.taskitem.Conflicts

    @property
    def ContactNames(self):
        return self.taskitem.ContactNames

    @ContactNames.setter
    def ContactNames(self, value):
        self.taskitem.ContactNames = value

    @property
    def ConversationID(self):
        return Conversation(self.taskitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.taskitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskitem.CreationTime

    @property
    def DateCompleted(self):
        return self.taskitem.DateCompleted

    @DateCompleted.setter
    def DateCompleted(self, value):
        self.taskitem.DateCompleted = value

    @property
    def DelegationState(self):
        return OlTaskDelegationState(self.taskitem.DelegationState)

    @property
    def Delegator(self):
        return self.taskitem.Delegator

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskitem.DownloadState)

    @property
    def DueDate(self):
        return self.taskitem.DueDate

    @DueDate.setter
    def DueDate(self, value):
        self.taskitem.DueDate = value

    @property
    def EntryID(self):
        return self.taskitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.taskitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.taskitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.taskitem.Importance = value

    @property
    def InternetCodepage(self):
        return self.taskitem.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.taskitem.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.taskitem.IsConflict

    @property
    def IsRecurring(self):
        return self.taskitem.IsRecurring

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.taskitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskitem.MessageClass = value

    @property
    def Mileage(self):
        return self.taskitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskitem.Mileage = value

    @property
    def NoAging(self):
        return self.taskitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskitem.NoAging = value

    @property
    def Ordinal(self):
        return self.taskitem.Ordinal

    @Ordinal.setter
    def Ordinal(self, value):
        self.taskitem.Ordinal = value

    @property
    def OutlookInternalVersion(self):
        return self.taskitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskitem.OutlookVersion

    @property
    def Owner(self):
        return self.taskitem.Owner

    @Owner.setter
    def Owner(self, value):
        self.taskitem.Owner = value

    @property
    def Ownership(self):
        return OlTaskOwnership(self.taskitem.Ownership)

    @property
    def Parent(self):
        return self.taskitem.Parent

    @property
    def PercentComplete(self):
        return self.taskitem.PercentComplete

    @PercentComplete.setter
    def PercentComplete(self, value):
        self.taskitem.PercentComplete = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskitem.PropertyAccessor)

    @property
    def Recipients(self):
        return Recipients(self.taskitem.Recipients)

    @property
    def ReminderOverrideDefault(self):
        return self.taskitem.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.taskitem.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.taskitem.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.taskitem.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.taskitem.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.taskitem.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.taskitem.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.taskitem.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.taskitem.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.taskitem.ReminderTime = value

    @property
    def ResponseState(self):
        return OlTaskResponse(self.taskitem.ResponseState)

    @property
    def Role(self):
        return self.taskitem.Role

    @Role.setter
    def Role(self, value):
        self.taskitem.Role = value

    @property
    def RTFBody(self):
        return self.taskitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskitem.RTFBody = value

    @property
    def Saved(self):
        return self.taskitem.Saved

    @property
    def SchedulePlusPriority(self):
        return self.taskitem.SchedulePlusPriority

    @SchedulePlusPriority.setter
    def SchedulePlusPriority(self, value):
        self.taskitem.SchedulePlusPriority = value

    @property
    def SendUsingAccount(self):
        return Account(self.taskitem.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.taskitem.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskitem.Session)

    @property
    def Size(self):
        return self.taskitem.Size

    @property
    def StartDate(self):
        return self.taskitem.StartDate

    @StartDate.setter
    def StartDate(self, value):
        self.taskitem.StartDate = value

    @property
    def Status(self):
        return OlTaskStatus(self.taskitem.Status)

    @Status.setter
    def Status(self, value):
        self.taskitem.Status = value

    @property
    def StatusOnCompletionRecipients(self):
        return self.taskitem.StatusOnCompletionRecipients

    @StatusOnCompletionRecipients.setter
    def StatusOnCompletionRecipients(self, value):
        self.taskitem.StatusOnCompletionRecipients = value

    @property
    def StatusUpdateRecipients(self):
        return self.taskitem.StatusUpdateRecipients

    @StatusUpdateRecipients.setter
    def StatusUpdateRecipients(self, value):
        self.taskitem.StatusUpdateRecipients = value

    @property
    def Subject(self):
        return self.taskitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskitem.Subject = value

    @property
    def TeamTask(self):
        return self.taskitem.TeamTask

    @TeamTask.setter
    def TeamTask(self, value):
        self.taskitem.TeamTask = value

    @property
    def ToDoTaskOrdinal(self):
        return TaskItem(self.taskitem.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.taskitem.ToDoTaskOrdinal = value

    @property
    def TotalWork(self):
        return self.taskitem.TotalWork

    @TotalWork.setter
    def TotalWork(self, value):
        self.taskitem.TotalWork = value

    @property
    def UnRead(self):
        return self.taskitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskitem.UserProperties)

    def Assign(self):
        return self.taskitem.Assign()

    def CancelResponseState(self):
        self.taskitem.CancelResponseState()

    def ClearRecurrencePattern(self):
        self.taskitem.ClearRecurrencePattern()

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskitem.Close(*args, **arguments)

    def Copy(self):
        self.taskitem.Copy()

    def Delete(self):
        self.taskitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskitem.Display(*args, **arguments)

    def GetConversation(self):
        return self.taskitem.GetConversation()

    def GetRecurrencePattern(self):
        return self.taskitem.GetRecurrencePattern()

    def MarkComplete(self):
        self.taskitem.MarkComplete()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskitem.Move(*args, **arguments)

    def PrintOut(self):
        self.taskitem.PrintOut()

    def Respond(self, *args, Response=None, fNoUI=None, fAdditionalTextDialog=None):
        arguments = {"Response": Response, "fNoUI": fNoUI, "fAdditionalTextDialog": fAdditionalTextDialog}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return TaskItem(self.taskitem.Respond(*args, **arguments))

    def Save(self):
        self.taskitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskitem.SaveAs(*args, **arguments)

    def Send(self):
        self.taskitem.Send()

    def ShowCategoriesDialog(self):
        self.taskitem.ShowCategoriesDialog()

    def SkipRecurrence(self):
        return self.taskitem.SkipRecurrence()

    def StatusReport(self):
        return self.taskitem.StatusReport()


class TaskRequestAcceptItem:

    def __init__(self, taskrequestacceptitem=None):
        self.taskrequestacceptitem = taskrequestacceptitem

    @property
    def Actions(self):
        return Actions(self.taskrequestacceptitem.Actions)

    @property
    def Application(self):
        return Application(self.taskrequestacceptitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskrequestacceptitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.taskrequestacceptitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskrequestacceptitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskrequestacceptitem.BillingInformation = value

    @property
    def Body(self):
        return self.taskrequestacceptitem.Body

    @Body.setter
    def Body(self, value):
        self.taskrequestacceptitem.Body = value

    @property
    def Categories(self):
        return self.taskrequestacceptitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskrequestacceptitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskrequestacceptitem.Class)

    @property
    def Companies(self):
        return self.taskrequestacceptitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskrequestacceptitem.Companies = value

    @property
    def Conflicts(self):
        return self.taskrequestacceptitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.taskrequestacceptitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.taskrequestacceptitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskrequestacceptitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskrequestacceptitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskrequestacceptitem.DownloadState)

    @property
    def EntryID(self):
        return self.taskrequestacceptitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskrequestacceptitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.taskrequestacceptitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.taskrequestacceptitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.taskrequestacceptitem.Importance = value

    @property
    def IsConflict(self):
        return self.taskrequestacceptitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskrequestacceptitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.taskrequestacceptitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskrequestacceptitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskrequestacceptitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskrequestacceptitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskrequestacceptitem.MessageClass = value

    @property
    def Mileage(self):
        return self.taskrequestacceptitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskrequestacceptitem.Mileage = value

    @property
    def NoAging(self):
        return self.taskrequestacceptitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskrequestacceptitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.taskrequestacceptitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskrequestacceptitem.OutlookVersion

    @property
    def Parent(self):
        return self.taskrequestacceptitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskrequestacceptitem.PropertyAccessor)

    @property
    def RTFBody(self):
        return self.taskrequestacceptitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskrequestacceptitem.RTFBody = value

    @property
    def Saved(self):
        return self.taskrequestacceptitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskrequestacceptitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskrequestacceptitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskrequestacceptitem.Session)

    @property
    def Size(self):
        return self.taskrequestacceptitem.Size

    @property
    def Subject(self):
        return self.taskrequestacceptitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskrequestacceptitem.Subject = value

    @property
    def UnRead(self):
        return self.taskrequestacceptitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskrequestacceptitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskrequestacceptitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestacceptitem.Close(*args, **arguments)

    def Copy(self):
        self.taskrequestacceptitem.Copy()

    def Delete(self):
        self.taskrequestacceptitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestacceptitem.Display(*args, **arguments)

    def GetAssociatedTask(self, *args, AddToTaskList=None):
        arguments = {"AddToTaskList": AddToTaskList}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskrequestacceptitem.GetAssociatedTask(*args, **arguments)

    def GetConversation(self):
        return self.taskrequestacceptitem.GetConversation()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskrequestacceptitem.Move(*args, **arguments)

    def PrintOut(self):
        self.taskrequestacceptitem.PrintOut()

    def Save(self):
        self.taskrequestacceptitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestacceptitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.taskrequestacceptitem.ShowCategoriesDialog()


class TaskRequestDeclineItem:

    def __init__(self, taskrequestdeclineitem=None):
        self.taskrequestdeclineitem = taskrequestdeclineitem

    @property
    def Actions(self):
        return Actions(self.taskrequestdeclineitem.Actions)

    @property
    def Application(self):
        return Application(self.taskrequestdeclineitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskrequestdeclineitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.taskrequestdeclineitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskrequestdeclineitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskrequestdeclineitem.BillingInformation = value

    @property
    def Body(self):
        return self.taskrequestdeclineitem.Body

    @Body.setter
    def Body(self, value):
        self.taskrequestdeclineitem.Body = value

    @property
    def Categories(self):
        return self.taskrequestdeclineitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskrequestdeclineitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskrequestdeclineitem.Class)

    @property
    def Companies(self):
        return self.taskrequestdeclineitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskrequestdeclineitem.Companies = value

    @property
    def Conflicts(self):
        return self.taskrequestdeclineitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.taskrequestdeclineitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.taskrequestdeclineitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskrequestdeclineitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskrequestdeclineitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskrequestdeclineitem.DownloadState)

    @property
    def EntryID(self):
        return self.taskrequestdeclineitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskrequestdeclineitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.taskrequestdeclineitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.taskrequestdeclineitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.taskrequestdeclineitem.Importance = value

    @property
    def IsConflict(self):
        return self.taskrequestdeclineitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskrequestdeclineitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.taskrequestdeclineitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskrequestdeclineitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskrequestdeclineitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskrequestdeclineitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskrequestdeclineitem.MessageClass = value

    @property
    def Mileage(self):
        return self.taskrequestdeclineitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskrequestdeclineitem.Mileage = value

    @property
    def NoAging(self):
        return self.taskrequestdeclineitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskrequestdeclineitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.taskrequestdeclineitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskrequestdeclineitem.OutlookVersion

    @property
    def Parent(self):
        return self.taskrequestdeclineitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskrequestdeclineitem.PropertyAccessor)

    @property
    def RTFBody(self):
        return self.taskrequestdeclineitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskrequestdeclineitem.RTFBody = value

    @property
    def Saved(self):
        return self.taskrequestdeclineitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskrequestdeclineitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskrequestdeclineitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskrequestdeclineitem.Session)

    @property
    def Size(self):
        return self.taskrequestdeclineitem.Size

    @property
    def Subject(self):
        return self.taskrequestdeclineitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskrequestdeclineitem.Subject = value

    @property
    def UnRead(self):
        return self.taskrequestdeclineitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskrequestdeclineitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskrequestdeclineitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestdeclineitem.Close(*args, **arguments)

    def Copy(self):
        self.taskrequestdeclineitem.Copy()

    def Delete(self):
        self.taskrequestdeclineitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestdeclineitem.Display(*args, **arguments)

    def GetAssociatedTask(self, *args, AddToTaskList=None):
        arguments = {"AddToTaskList": AddToTaskList}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskrequestdeclineitem.GetAssociatedTask(*args, **arguments)

    def GetConversation(self):
        return self.taskrequestdeclineitem.GetConversation()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskrequestdeclineitem.Move(*args, **arguments)

    def PrintOut(self):
        self.taskrequestdeclineitem.PrintOut()

    def Save(self):
        self.taskrequestdeclineitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestdeclineitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.taskrequestdeclineitem.ShowCategoriesDialog()


class TaskRequestItem:

    def __init__(self, taskrequestitem=None):
        self.taskrequestitem = taskrequestitem

    @property
    def Actions(self):
        return Actions(self.taskrequestitem.Actions)

    @property
    def Application(self):
        return Application(self.taskrequestitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskrequestitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.taskrequestitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskrequestitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskrequestitem.BillingInformation = value

    @property
    def Body(self):
        return self.taskrequestitem.Body

    @Body.setter
    def Body(self, value):
        self.taskrequestitem.Body = value

    @property
    def Categories(self):
        return self.taskrequestitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskrequestitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskrequestitem.Class)

    @property
    def Companies(self):
        return self.taskrequestitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskrequestitem.Companies = value

    @property
    def Conflicts(self):
        return self.taskrequestitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.taskrequestitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.taskrequestitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskrequestitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskrequestitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskrequestitem.DownloadState)

    @property
    def EntryID(self):
        return self.taskrequestitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskrequestitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.taskrequestitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.taskrequestitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.taskrequestitem.Importance = value

    @property
    def IsConflict(self):
        return self.taskrequestitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskrequestitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.taskrequestitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskrequestitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskrequestitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskrequestitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskrequestitem.MessageClass = value

    @property
    def Mileage(self):
        return self.taskrequestitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskrequestitem.Mileage = value

    @property
    def NoAging(self):
        return self.taskrequestitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskrequestitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.taskrequestitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskrequestitem.OutlookVersion

    @property
    def Parent(self):
        return self.taskrequestitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskrequestitem.PropertyAccessor)

    @property
    def RTFBody(self):
        return self.taskrequestitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskrequestitem.RTFBody = value

    @property
    def Saved(self):
        return self.taskrequestitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskrequestitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskrequestitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskrequestitem.Session)

    @property
    def Size(self):
        return self.taskrequestitem.Size

    @property
    def Subject(self):
        return self.taskrequestitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskrequestitem.Subject = value

    @property
    def UnRead(self):
        return self.taskrequestitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskrequestitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskrequestitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestitem.Close(*args, **arguments)

    def Copy(self):
        self.taskrequestitem.Copy()

    def Delete(self):
        self.taskrequestitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestitem.Display(*args, **arguments)

    def GetAssociatedTask(self, *args, AddToTaskList=None):
        arguments = {"AddToTaskList": AddToTaskList}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskrequestitem.GetAssociatedTask(*args, **arguments)

    def GetConversation(self):
        return self.taskrequestitem.GetConversation()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskrequestitem.Move(*args, **arguments)

    def PrintOut(self):
        self.taskrequestitem.PrintOut()

    def Save(self):
        self.taskrequestitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.taskrequestitem.ShowCategoriesDialog()


class TaskRequestUpdateItem:

    def __init__(self, taskrequestupdateitem=None):
        self.taskrequestupdateitem = taskrequestupdateitem

    @property
    def Actions(self):
        return Actions(self.taskrequestupdateitem.Actions)

    @property
    def Application(self):
        return Application(self.taskrequestupdateitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskrequestupdateitem.Attachments)

    @property
    def AutoResolvedWinner(self):
        return self.taskrequestupdateitem.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskrequestupdateitem.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskrequestupdateitem.BillingInformation = value

    @property
    def Body(self):
        return self.taskrequestupdateitem.Body

    @Body.setter
    def Body(self, value):
        self.taskrequestupdateitem.Body = value

    @property
    def Categories(self):
        return self.taskrequestupdateitem.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskrequestupdateitem.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskrequestupdateitem.Class)

    @property
    def Companies(self):
        return self.taskrequestupdateitem.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskrequestupdateitem.Companies = value

    @property
    def Conflicts(self):
        return self.taskrequestupdateitem.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.taskrequestupdateitem.ConversationID)

    @property
    def ConversationIndex(self):
        return self.taskrequestupdateitem.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskrequestupdateitem.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskrequestupdateitem.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskrequestupdateitem.DownloadState)

    @property
    def EntryID(self):
        return self.taskrequestupdateitem.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskrequestupdateitem.FormDescription)

    @property
    def GetInspector(self):
        return Inspector(self.taskrequestupdateitem.GetInspector)

    @property
    def Importance(self):
        return OlImportance(self.taskrequestupdateitem.Importance)

    @Importance.setter
    def Importance(self, value):
        self.taskrequestupdateitem.Importance = value

    @property
    def IsConflict(self):
        return self.taskrequestupdateitem.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskrequestupdateitem.ItemProperties)

    @property
    def LastModificationTime(self):
        return self.taskrequestupdateitem.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskrequestupdateitem.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskrequestupdateitem.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskrequestupdateitem.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskrequestupdateitem.MessageClass = value

    @property
    def Mileage(self):
        return self.taskrequestupdateitem.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskrequestupdateitem.Mileage = value

    @property
    def NoAging(self):
        return self.taskrequestupdateitem.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskrequestupdateitem.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.taskrequestupdateitem.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskrequestupdateitem.OutlookVersion

    @property
    def Parent(self):
        return self.taskrequestupdateitem.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskrequestupdateitem.PropertyAccessor)

    @property
    def RTFBody(self):
        return self.taskrequestupdateitem.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskrequestupdateitem.RTFBody = value

    @property
    def Saved(self):
        return self.taskrequestupdateitem.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskrequestupdateitem.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskrequestupdateitem.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskrequestupdateitem.Session)

    @property
    def Size(self):
        return self.taskrequestupdateitem.Size

    @property
    def Subject(self):
        return self.taskrequestupdateitem.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskrequestupdateitem.Subject = value

    @property
    def UnRead(self):
        return self.taskrequestupdateitem.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskrequestupdateitem.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskrequestupdateitem.UserProperties)

    def Close(self, *args, SaveMode=None):
        arguments = {"SaveMode": SaveMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestupdateitem.Close(*args, **arguments)

    def Copy(self):
        self.taskrequestupdateitem.Copy()

    def Delete(self):
        self.taskrequestupdateitem.Delete()

    def Display(self, *args, Modal=None):
        arguments = {"Modal": Modal}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestupdateitem.Display(*args, **arguments)

    def GetAssociatedTask(self, *args, AddToTaskList=None):
        arguments = {"AddToTaskList": AddToTaskList}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskrequestupdateitem.GetAssociatedTask(*args, **arguments)

    def GetConversation(self):
        return self.taskrequestupdateitem.GetConversation()

    def Move(self, *args, DestFldr=None):
        arguments = {"DestFldr": DestFldr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskrequestupdateitem.Move(*args, **arguments)

    def PrintOut(self):
        self.taskrequestupdateitem.PrintOut()

    def Save(self):
        self.taskrequestupdateitem.Save()

    def SaveAs(self, *args, Path=None, Type=None):
        arguments = {"Path": Path, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.taskrequestupdateitem.SaveAs(*args, **arguments)

    def ShowCategoriesDialog(self):
        self.taskrequestupdateitem.ShowCategoriesDialog()


class TasksModule:

    def __init__(self, tasksmodule=None):
        self.tasksmodule = tasksmodule

    @property
    def Application(self):
        return Application(self.tasksmodule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.tasksmodule.Class)

    @property
    def Name(self):
        return TasksModule(self.tasksmodule.Name)

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.tasksmodule.NavigationGroups)

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.tasksmodule.NavigationModuleType)

    @property
    def Parent(self):
        return self.tasksmodule.Parent

    @property
    def Position(self):
        return TasksModule(self.tasksmodule.Position)

    @Position.setter
    def Position(self, value):
        self.tasksmodule.Position = value

    @property
    def Session(self):
        return NameSpace(self.tasksmodule.Session)

    @property
    def Visible(self):
        return TasksModule(self.tasksmodule.Visible)

    @Visible.setter
    def Visible(self, value):
        self.tasksmodule.Visible = value


class TextRuleCondition:

    def __init__(self, textrulecondition=None):
        self.textrulecondition = textrulecondition

    @property
    def Application(self):
        return Application(self.textrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.textrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.textrulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.textrulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.textrulecondition.Enabled = value

    @property
    def Parent(self):
        return self.textrulecondition.Parent

    @property
    def Session(self):
        return NameSpace(self.textrulecondition.Session)

    @property
    def Text(self):
        return self.textrulecondition.Text

    @Text.setter
    def Text(self, value):
        self.textrulecondition.Text = value


class TimelineView:

    def __init__(self, timelineview=None):
        self.timelineview = timelineview

    @property
    def Application(self):
        return Application(self.timelineview.Application)

    @property
    def Class(self):
        return OlObjectClass(self.timelineview.Class)

    @property
    def DefaultExpandCollapseSetting(self):
        return OlDefaultExpandCollapseSetting(self.timelineview.DefaultExpandCollapseSetting)

    @DefaultExpandCollapseSetting.setter
    def DefaultExpandCollapseSetting(self, value):
        self.timelineview.DefaultExpandCollapseSetting = value

    @property
    def EndField(self):
        return TimelineView(self.timelineview.EndField)

    @EndField.setter
    def EndField(self, value):
        self.timelineview.EndField = value

    @property
    def Filter(self):
        return self.timelineview.Filter

    @Filter.setter
    def Filter(self, value):
        self.timelineview.Filter = value

    @property
    def GroupByFields(self):
        return OrderFields(self.timelineview.GroupByFields)

    @property
    def ItemFont(self):
        return ViewFont(self.timelineview.ItemFont)

    @property
    def Language(self):
        return self.timelineview.Language

    @Language.setter
    def Language(self, value):
        self.timelineview.Language = value

    @property
    def LockUserChanges(self):
        return self.timelineview.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.timelineview.LockUserChanges = value

    @property
    def LowerScaleFont(self):
        return ViewFont(self.timelineview.LowerScaleFont)

    @property
    def MaxLabelWidth(self):
        return TimelineView(self.timelineview.MaxLabelWidth)

    @MaxLabelWidth.setter
    def MaxLabelWidth(self, value):
        self.timelineview.MaxLabelWidth = value

    @property
    def Name(self):
        return self.timelineview.Name

    @Name.setter
    def Name(self, value):
        self.timelineview.Name = value

    @property
    def Parent(self):
        return self.timelineview.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.timelineview.SaveOption)

    @property
    def Session(self):
        return NameSpace(self.timelineview.Session)

    @property
    def ShowLabelWhenViewingByMonth(self):
        return TimelineView(self.timelineview.ShowLabelWhenViewingByMonth)

    @ShowLabelWhenViewingByMonth.setter
    def ShowLabelWhenViewingByMonth(self, value):
        self.timelineview.ShowLabelWhenViewingByMonth = value

    @property
    def ShowWeekNumbers(self):
        return TimelineView(self.timelineview.ShowWeekNumbers)

    @ShowWeekNumbers.setter
    def ShowWeekNumbers(self, value):
        self.timelineview.ShowWeekNumbers = value

    @property
    def Standard(self):
        return TimelineView(self.timelineview.Standard)

    @property
    def StartField(self):
        return TimelineView(self.timelineview.StartField)

    @StartField.setter
    def StartField(self, value):
        self.timelineview.StartField = value

    @property
    def TimelineViewMode(self):
        return OlTimelineViewMode(self.timelineview.TimelineViewMode)

    @TimelineViewMode.setter
    def TimelineViewMode(self, value):
        self.timelineview.TimelineViewMode = value

    @property
    def UpperScaleFont(self):
        return ViewFont(self.timelineview.UpperScaleFont)

    @property
    def ViewType(self):
        return OlViewType(self.timelineview.ViewType)

    @property
    def XML(self):
        return self.timelineview.XML

    @XML.setter
    def XML(self, value):
        self.timelineview.XML = value

    def Apply(self):
        self.timelineview.Apply()

    def Copy(self, *args, Name=None, SaveOption=None):
        arguments = {"Name": Name, "SaveOption": SaveOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.timelineview.Copy(*args, **arguments)

    def Delete(self):
        self.timelineview.Delete()

    def GoToDate(self, *args, Date=None):
        arguments = {"Date": Date}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.timelineview.GoToDate(*args, **arguments)

    def Reset(self):
        self.timelineview.Reset()

    def Save(self):
        self.timelineview.Save()


class TimeZone:

    def __init__(self, timezone=None):
        self.timezone = timezone

    @property
    def Application(self):
        return Application(self.timezone.Application)

    @property
    def Bias(self):
        return self.timezone.Bias

    @property
    def Class(self):
        return OlObjectClass(self.timezone.Class)

    @property
    def DaylightBias(self):
        return self.timezone.DaylightBias

    @property
    def DaylightDate(self):
        return self.timezone.DaylightDate

    @property
    def DaylightDesignation(self):
        return self.timezone.DaylightDesignation

    @property
    def ID(self):
        return self.timezone.ID

    @property
    def Name(self):
        return self.timezone.Name

    @property
    def Parent(self):
        return self.timezone.Parent

    @property
    def Session(self):
        return NameSpace(self.timezone.Session)

    @property
    def StandardBias(self):
        return self.timezone.StandardBias

    @property
    def StandardDate(self):
        return self.timezone.StandardDate

    @property
    def StandardDesignation(self):
        return self.timezone.StandardDesignation


class TimeZones:

    def __init__(self, timezones=None):
        self.timezones = timezones

    def __call__(self, item):
        return TimeZone(self.timezones(item))

    @property
    def Application(self):
        return Application(self.timezones.Application)

    @property
    def Class(self):
        return OlObjectClass(self.timezones.Class)

    @property
    def Count(self):
        return self.timezones.Count

    @property
    def CurrentTimeZone(self):
        return TimeZone(self.timezones.CurrentTimeZone)

    @property
    def Parent(self):
        return self.timezones.Parent

    @property
    def Session(self):
        return NameSpace(self.timezones.Session)

    def ConvertTime(self, *args, SourceDateTime=None, SourceTimeZone=None, DestinationTimeZone=None):
        arguments = {"SourceDateTime": SourceDateTime, "SourceTimeZone": SourceTimeZone, "DestinationTimeZone": DestinationTimeZone}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.timezones.ConvertTime(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.timezones.Item(*args, **arguments)


class ToOrFromRuleCondition:

    def __init__(self, toorfromrulecondition=None):
        self.toorfromrulecondition = toorfromrulecondition

    @property
    def Application(self):
        return Application(self.toorfromrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.toorfromrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.toorfromrulecondition.ConditionType)

    @property
    def Enabled(self):
        return self.toorfromrulecondition.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.toorfromrulecondition.Enabled = value

    @property
    def Parent(self):
        return self.toorfromrulecondition.Parent

    @property
    def Recipients(self):
        return Recipients(self.toorfromrulecondition.Recipients)

    @property
    def Session(self):
        return NameSpace(self.toorfromrulecondition.Session)


class UserDefinedProperties:

    def __init__(self, userdefinedproperties=None):
        self.userdefinedproperties = userdefinedproperties

    @property
    def Application(self):
        return Application(self.userdefinedproperties.Application)

    @property
    def Class(self):
        return OlObjectClass(self.userdefinedproperties.Class)

    @property
    def Count(self):
        return self.userdefinedproperties.Count

    @property
    def Parent(self):
        return self.userdefinedproperties.Parent

    @property
    def Session(self):
        return NameSpace(self.userdefinedproperties.Session)

    def Add(self, *args, Name=None, Type=None, DisplayFormat=None, Formula=None):
        arguments = {"Name": Name, "Type": Type, "DisplayFormat": DisplayFormat, "Formula": Formula}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.userdefinedproperties.Add(*args, **arguments)

    def Find(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.userdefinedproperties.Find(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return UserDefinedProperty(self.userdefinedproperties.Item(*args, **arguments))

    def Refresh(self):
        self.userdefinedproperties.Refresh()

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.userdefinedproperties.Remove(*args, **arguments)


class UserDefinedProperty:

    def __init__(self, userdefinedproperty=None):
        self.userdefinedproperty = userdefinedproperty

    @property
    def Application(self):
        return Application(self.userdefinedproperty.Application)

    @property
    def Class(self):
        return OlObjectClass(self.userdefinedproperty.Class)

    @property
    def DisplayFormat(self):
        return UserDefinedProperty(self.userdefinedproperty.DisplayFormat)

    @property
    def Formula(self):
        return self.userdefinedproperty.Formula

    @property
    def Name(self):
        return self.userdefinedproperty.Name

    @property
    def Parent(self):
        return self.userdefinedproperty.Parent

    @property
    def Session(self):
        return NameSpace(self.userdefinedproperty.Session)

    @property
    def Type(self):
        return OlUserPropertyType(self.userdefinedproperty.Type)

    def Delete(self):
        self.userdefinedproperty.Delete()


class UserProperties:

    def __init__(self, userproperties=None):
        self.userproperties = userproperties

    @property
    def Application(self):
        return Application(self.userproperties.Application)

    @property
    def Class(self):
        return OlObjectClass(self.userproperties.Class)

    @property
    def Count(self):
        return self.userproperties.Count

    @property
    def Parent(self):
        return self.userproperties.Parent

    @property
    def Session(self):
        return NameSpace(self.userproperties.Session)

    def Add(self, *args, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = {"Name": Name, "Type": Type, "AddToFolderFields": AddToFolderFields, "DisplayFormat": DisplayFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return UserProperty(self.userproperties.Add(*args, **arguments))

    def Find(self, *args, Name=None, Custom=None):
        arguments = {"Name": Name, "Custom": Custom}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.userproperties.Find(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.userproperties.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.userproperties.Remove(*args, **arguments)


class UserProperty:

    def __init__(self, userproperty=None):
        self.userproperty = userproperty

    @property
    def Application(self):
        return Application(self.userproperty.Application)

    @property
    def Class(self):
        return OlObjectClass(self.userproperty.Class)

    @property
    def Formula(self):
        return self.userproperty.Formula

    @Formula.setter
    def Formula(self, value):
        self.userproperty.Formula = value

    @property
    def Name(self):
        return self.userproperty.Name

    @property
    def Parent(self):
        return self.userproperty.Parent

    @property
    def Session(self):
        return NameSpace(self.userproperty.Session)

    @property
    def Type(self):
        return OlUserPropertyType(self.userproperty.Type)

    @property
    def ValidationFormula(self):
        return self.userproperty.ValidationFormula

    @ValidationFormula.setter
    def ValidationFormula(self, value):
        self.userproperty.ValidationFormula = value

    @property
    def ValidationText(self):
        return self.userproperty.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.userproperty.ValidationText = value

    @property
    def Value(self):
        return self.userproperty.Value

    @Value.setter
    def Value(self, value):
        self.userproperty.Value = value

    def Delete(self):
        self.userproperty.Delete()


class View:

    def __init__(self, view=None):
        self.view = view

    @property
    def Application(self):
        return Application(self.view.Application)

    @property
    def Class(self):
        return OlObjectClass(self.view.Class)

    @property
    def Filter(self):
        return self.view.Filter

    @Filter.setter
    def Filter(self, value):
        self.view.Filter = value

    @property
    def Language(self):
        return self.view.Language

    @Language.setter
    def Language(self, value):
        self.view.Language = value

    @property
    def LockUserChanges(self):
        return self.view.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.view.LockUserChanges = value

    @property
    def Name(self):
        return self.view.Name

    @Name.setter
    def Name(self, value):
        self.view.Name = value

    @property
    def Parent(self):
        return self.view.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.view.SaveOption)

    @property
    def Session(self):
        return NameSpace(self.view.Session)

    @property
    def Standard(self):
        return self.view.Standard

    @property
    def ViewType(self):
        return OlViewType(self.view.ViewType)

    @property
    def XML(self):
        return self.view.XML

    @XML.setter
    def XML(self, value):
        self.view.XML = value

    def Apply(self):
        self.view.Apply()

    def Copy(self, *args, Name=None, SaveOption=None):
        arguments = {"Name": Name, "SaveOption": SaveOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.view.Copy(*args, **arguments)

    def Delete(self):
        self.view.Delete()

    def GoToDate(self, *args, Date=None):
        arguments = {"Date": Date}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.view.GoToDate(*args, **arguments)

    def Reset(self):
        self.view.Reset()

    def Save(self):
        self.view.Save()


class ViewField:

    def __init__(self, viewfield=None):
        self.viewfield = viewfield

    @property
    def Application(self):
        return Application(self.viewfield.Application)

    @property
    def Class(self):
        return OlObjectClass(self.viewfield.Class)

    @property
    def ColumnFormat(self):
        return ColumnFormat(self.viewfield.ColumnFormat)

    @property
    def Parent(self):
        return self.viewfield.Parent

    @property
    def Session(self):
        return NameSpace(self.viewfield.Session)

    @property
    def ViewXMLSchemaName(self):
        return ViewField(self.viewfield.ViewXMLSchemaName)


class ViewFields:

    def __init__(self, viewfields=None):
        self.viewfields = viewfields

    @property
    def Application(self):
        return Application(self.viewfields.Application)

    @property
    def Class(self):
        return OlObjectClass(self.viewfields.Class)

    @property
    def Count(self):
        return ViewField(self.viewfields.Count)

    @property
    def Parent(self):
        return self.viewfields.Parent

    @property
    def Session(self):
        return NameSpace(self.viewfields.Session)

    def Add(self, *args, PropertyName=None):
        arguments = {"PropertyName": PropertyName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.viewfields.Add(*args, **arguments)

    def Insert(self, *args, PropertyName=None, Index=None):
        arguments = {"PropertyName": PropertyName, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.viewfields.Insert(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.viewfields.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.viewfields.Remove(*args, **arguments)


class ViewFont:

    def __init__(self, viewfont=None):
        self.viewfont = viewfont

    @property
    def Application(self):
        return Application(self.viewfont.Application)

    @property
    def Bold(self):
        return ViewFont(self.viewfont.Bold)

    @Bold.setter
    def Bold(self, value):
        self.viewfont.Bold = value

    @property
    def Class(self):
        return OlObjectClass(self.viewfont.Class)

    @property
    def Color(self):
        return OlColor(self.viewfont.Color)

    @Color.setter
    def Color(self, value):
        self.viewfont.Color = value

    @property
    def ExtendedColor(self):
        return OlCategoryColor(self.viewfont.ExtendedColor)

    @ExtendedColor.setter
    def ExtendedColor(self, value):
        self.viewfont.ExtendedColor = value

    @property
    def Italic(self):
        return ViewFont(self.viewfont.Italic)

    @Italic.setter
    def Italic(self, value):
        self.viewfont.Italic = value

    @property
    def Name(self):
        return self.viewfont.Name

    @Name.setter
    def Name(self, value):
        self.viewfont.Name = value

    @property
    def Parent(self):
        return self.viewfont.Parent

    @property
    def Session(self):
        return NameSpace(self.viewfont.Session)

    @property
    def Size(self):
        return self.viewfont.Size

    @Size.setter
    def Size(self, value):
        self.viewfont.Size = value

    @property
    def Strikethrough(self):
        return ViewFont(self.viewfont.Strikethrough)

    @Strikethrough.setter
    def Strikethrough(self, value):
        self.viewfont.Strikethrough = value

    @property
    def Underline(self):
        return ViewFont(self.viewfont.Underline)

    @Underline.setter
    def Underline(self, value):
        self.viewfont.Underline = value


class Views:

    def __init__(self, views=None):
        self.views = views

    @property
    def Application(self):
        return Application(self.views.Application)

    @property
    def Class(self):
        return OlObjectClass(self.views.Class)

    @property
    def Count(self):
        return self.views.Count

    @property
    def Parent(self):
        return self.views.Parent

    @property
    def Session(self):
        return NameSpace(self.views.Session)

    def Add(self, *args, Name=None, ViewType=None, SaveOption=None):
        arguments = {"Name": Name, "ViewType": ViewType, "SaveOption": SaveOption}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return View(self.views.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.views.Item(*args, **arguments)

    def Remove(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.views.Remove(*args, **arguments)

