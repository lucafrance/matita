from . import com_arguments

import win32com.client
import pythoncom

class Account:

    def __init__(self, account=None):
        self.account = account

    @property
    def AccountType(self):
        return OlAccountType(self.account.AccountType)

    # Lower case alias for AccountType
    @property
    def accounttype(self):
        return self.AccountType

    @property
    def Application(self):
        return Application(self.account.Application)

    @property
    def AutoDiscoverConnectionMode(self):
        return OlAutoDiscoverConnectionMode(self.account.AutoDiscoverConnectionMode)

    # Lower case alias for AutoDiscoverConnectionMode
    @property
    def autodiscoverconnectionmode(self):
        return self.AutoDiscoverConnectionMode

    @property
    def AutoDiscoverXml(self):
        return self.account.AutoDiscoverXml

    # Lower case alias for AutoDiscoverXml
    @property
    def autodiscoverxml(self):
        return self.AutoDiscoverXml

    @property
    def Class(self):
        return OlObjectClass(self.account.Class)

    @property
    def CurrentUser(self):
        return Recipient(self.account.CurrentUser)

    # Lower case alias for CurrentUser
    @property
    def currentuser(self):
        return self.CurrentUser

    @property
    def DeliveryStore(self):
        return Store(self.account.DeliveryStore)

    # Lower case alias for DeliveryStore
    @property
    def deliverystore(self):
        return self.DeliveryStore

    @property
    def DisplayName(self):
        return Account(self.account.DisplayName)

    # Lower case alias for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def ExchangeConnectionMode(self):
        return OlExchangeConnectionMode(self.account.ExchangeConnectionMode)

    # Lower case alias for ExchangeConnectionMode
    @property
    def exchangeconnectionmode(self):
        return self.ExchangeConnectionMode

    @property
    def ExchangeMailboxServerName(self):
        return self.account.ExchangeMailboxServerName

    # Lower case alias for ExchangeMailboxServerName
    @property
    def exchangemailboxservername(self):
        return self.ExchangeMailboxServerName

    @property
    def ExchangeMailboxServerVersion(self):
        return self.account.ExchangeMailboxServerVersion

    # Lower case alias for ExchangeMailboxServerVersion
    @property
    def exchangemailboxserverversion(self):
        return self.ExchangeMailboxServerVersion

    @property
    def Parent(self):
        return self.account.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.account.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def SmtpAddress(self):
        return Account(self.account.SmtpAddress)

    # Lower case alias for SmtpAddress
    @property
    def smtpaddress(self):
        return self.SmtpAddress

    @property
    def UserName(self):
        return Account(self.account.UserName)

    # Lower case alias for UserName
    @property
    def username(self):
        return self.UserName

    def GetAddressEntryFromID(self, ID=None):
        arguments = com_arguments([ID])
        return ID(self.account.GetAddressEntryFromID(*arguments))

    def GetRecipientFromID(self, EntryID=None):
        arguments = com_arguments([EntryID])
        return self.account.GetRecipientFromID(*arguments)


class AccountRuleCondition:

    def __init__(self, accountrulecondition=None):
        self.accountrulecondition = accountrulecondition

    @property
    def Account(self):
        return Account(self.accountrulecondition.Account)

    # Lower case alias for Account
    @property
    def account(self):
        return self.Account

    @Account.setter
    def Account(self, value):
        self.accountrulecondition.Account = value

    # Lower case alias for Account setter
    @account.setter
    def account(self, value):
        self.Account = value

    @property
    def Application(self):
        return Application(self.accountrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.accountrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.accountrulecondition.ConditionType)

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.accountrulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.accountrulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.accountrulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.accountrulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.accounts.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.accounts.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.accounts.Item(*arguments)


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

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SelectedAccount(self):
        return Account(self.accountselector.SelectedAccount)

    # Lower case alias for SelectedAccount
    @property
    def selectedaccount(self):
        return self.SelectedAccount

    @property
    def Session(self):
        return NameSpace(self.accountselector.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for CopyLike
    @property
    def copylike(self):
        return self.CopyLike

    @CopyLike.setter
    def CopyLike(self, value):
        self.action.CopyLike = value

    # Lower case alias for CopyLike setter
    @copylike.setter
    def copylike(self, value):
        self.CopyLike = value

    @property
    def Enabled(self):
        return self.action.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.action.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def MessageClass(self):
        return Action(self.action.MessageClass)

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.action.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Name(self):
        return self.action.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.action.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.action.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Prefix(self):
        return self.action.Prefix

    # Lower case alias for Prefix
    @property
    def prefix(self):
        return self.Prefix

    @Prefix.setter
    def Prefix(self, value):
        self.action.Prefix = value

    # Lower case alias for Prefix setter
    @prefix.setter
    def prefix(self, value):
        self.Prefix = value

    @property
    def ReplyStyle(self):
        return OlActionReplyStyle(self.action.ReplyStyle)

    # Lower case alias for ReplyStyle
    @property
    def replystyle(self):
        return self.ReplyStyle

    @ReplyStyle.setter
    def ReplyStyle(self, value):
        self.action.ReplyStyle = value

    # Lower case alias for ReplyStyle setter
    @replystyle.setter
    def replystyle(self, value):
        self.ReplyStyle = value

    @property
    def ResponseStyle(self):
        return OlActionResponseStyle(self.action.ResponseStyle)

    # Lower case alias for ResponseStyle
    @property
    def responsestyle(self):
        return self.ResponseStyle

    @ResponseStyle.setter
    def ResponseStyle(self, value):
        self.action.ResponseStyle = value

    # Lower case alias for ResponseStyle setter
    @responsestyle.setter
    def responsestyle(self, value):
        self.ResponseStyle = value

    @property
    def Session(self):
        return NameSpace(self.action.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowOn(self):
        return OlActionShowOn(self.action.ShowOn)

    # Lower case alias for ShowOn
    @property
    def showon(self):
        return self.ShowOn

    @ShowOn.setter
    def ShowOn(self, value):
        self.action.ShowOn = value

    # Lower case alias for ShowOn setter
    @showon.setter
    def showon(self, value):
        self.ShowOn = value

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.actions.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.actions.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self):
        return Action(self.actions.Add())

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.actions.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.actions.Remove(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.addressentries.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.addressentries.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Type=None, Name=None, Address=None):
        arguments = com_arguments([Type, Name, Address])
        return AddressEntry(self.addressentries.Add(*arguments))

    def GetFirst(self):
        return AddressEntry(self.addressentries.GetFirst())

    def GetLast(self):
        return AddressEntry(self.addressentries.GetLast())

    def GetNext(self):
        return AddressEntry(self.addressentries.GetNext())

    def GetPrevious(self):
        return AddressEntry(self.addressentries.GetPrevious())

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.addressentries.Item(*arguments)

    def Sort(self, Property=None, Order=None):
        arguments = com_arguments([Property, Order])
        self.addressentries.Sort(*arguments)


class AddressEntry:

    def __init__(self, addressentry=None):
        self.addressentry = addressentry

    @property
    def Address(self):
        return AddressEntry(self.addressentry.Address)

    # Lower case alias for Address
    @property
    def address(self):
        return self.Address

    @Address.setter
    def Address(self, value):
        self.addressentry.Address = value

    # Lower case alias for Address setter
    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.addressentry.AddressEntryUserType)

    # Lower case alias for AddressEntryUserType
    @property
    def addressentryusertype(self):
        return self.AddressEntryUserType

    @property
    def Application(self):
        return Application(self.addressentry.Application)

    @property
    def Class(self):
        return OlObjectClass(self.addressentry.Class)

    @property
    def DisplayType(self):
        return OlDisplayType(self.addressentry.DisplayType)

    # Lower case alias for DisplayType
    @property
    def displaytype(self):
        return self.DisplayType

    @property
    def ID(self):
        return self.addressentry.ID

    # Lower case alias for ID
    @property
    def id(self):
        return self.ID

    @property
    def Name(self):
        return self.addressentry.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.addressentry.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.addressentry.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.addressentry.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.addressentry.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return self.addressentry.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.addressentry.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    def Delete(self):
        self.addressentry.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([HWnd])
        self.addressentry.Details(*arguments)

    def GetContact(self):
        return self.addressentry.GetContact()

    def GetExchangeDistributionList(self):
        return self.addressentry.GetExchangeDistributionList()

    def GetExchangeUser(self):
        return self.addressentry.GetExchangeUser()

    def GetFreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([Start, MinPerChar, CompleteFormat])
        return self.addressentry.GetFreeBusy(*arguments)

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([MakePermanent, Refresh])
        self.addressentry.Update(*arguments)


class AddressList:

    def __init__(self, addresslist=None):
        self.addresslist = addresslist

    @property
    def AddressEntries(self):
        return AddressEntries(self.addresslist.AddressEntries)

    # Lower case alias for AddressEntries
    @property
    def addressentries(self):
        return self.AddressEntries

    @property
    def AddressListType(self):
        return OlAddressListType(self.addresslist.AddressListType)

    # Lower case alias for AddressListType
    @property
    def addresslisttype(self):
        return self.AddressListType

    @property
    def Application(self):
        return Application(self.addresslist.Application)

    @property
    def Class(self):
        return OlObjectClass(self.addresslist.Class)

    @property
    def ID(self):
        return self.addresslist.ID

    # Lower case alias for ID
    @property
    def id(self):
        return self.ID

    @property
    def Index(self):
        return self.addresslist.Index

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def IsInitialAddressList(self):
        return AddressList(self.addresslist.IsInitialAddressList)

    # Lower case alias for IsInitialAddressList
    @property
    def isinitialaddresslist(self):
        return self.IsInitialAddressList

    @property
    def IsReadOnly(self):
        return AddressList(self.addresslist.IsReadOnly)

    # Lower case alias for IsReadOnly
    @property
    def isreadonly(self):
        return self.IsReadOnly

    @property
    def Name(self):
        return self.addresslist.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.addresslist.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.addresslist.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ResolutionOrder(self):
        return AddressList(self.addresslist.ResolutionOrder)

    # Lower case alias for ResolutionOrder
    @property
    def resolutionorder(self):
        return self.ResolutionOrder

    @property
    def Session(self):
        return NameSpace(self.addresslist.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.addresslists.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.addresslists.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.addresslists.Item(*arguments)


class AddressRuleCondition:

    def __init__(self, addressrulecondition=None):
        self.addressrulecondition = addressrulecondition

    @property
    def Address(self):
        return self.addressrulecondition.Address

    # Lower case alias for Address
    @property
    def address(self):
        return self.Address

    @Address.setter
    def Address(self, value):
        self.addressrulecondition.Address = value

    # Lower case alias for Address setter
    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def Application(self):
        return Application(self.addressrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.addressrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.addressrulecondition.ConditionType)

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.addressrulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.addressrulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.addressrulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.addressrulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Assistance
    @property
    def assistance(self):
        return self.Assistance

    @property
    def Class(self):
        return OlObjectClass(self.application.Class)

    @property
    def COMAddIns(self):
        return self.application.COMAddIns

    # Lower case alias for COMAddIns
    @property
    def comaddins(self):
        return self.COMAddIns

    @property
    def DefaultProfileName(self):
        return self.application.DefaultProfileName

    # Lower case alias for DefaultProfileName
    @property
    def defaultprofilename(self):
        return self.DefaultProfileName

    @property
    def Explorers(self):
        return Explorers(self.application.Explorers)

    # Lower case alias for Explorers
    @property
    def explorers(self):
        return self.Explorers

    @property
    def Inspectors(self):
        return Inspectors(self.application.Inspectors)

    # Lower case alias for Inspectors
    @property
    def inspectors(self):
        return self.Inspectors

    @property
    def IsTrusted(self):
        return self.application.IsTrusted

    # Lower case alias for IsTrusted
    @property
    def istrusted(self):
        return self.IsTrusted

    @property
    def LanguageSettings(self):
        return self.application.LanguageSettings

    # Lower case alias for LanguageSettings
    @property
    def languagesettings(self):
        return self.LanguageSettings

    @property
    def Name(self):
        return self.application.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.application.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PickerDialog(self):
        return self.application.PickerDialog

    # Lower case alias for PickerDialog
    @property
    def pickerdialog(self):
        return self.PickerDialog

    @property
    def ProductCode(self):
        return self.application.ProductCode

    # Lower case alias for ProductCode
    @property
    def productcode(self):
        return self.ProductCode

    @property
    def Reminders(self):
        return Reminders(self.application.Reminders)

    # Lower case alias for Reminders
    @property
    def reminders(self):
        return self.Reminders

    @property
    def Session(self):
        return NameSpace(self.application.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def TimeZones(self):
        return TimeZones(self.application.TimeZones)

    # Lower case alias for TimeZones
    @property
    def timezones(self):
        return self.TimeZones

    @property
    def Version(self):
        return self.application.Version

    # Lower case alias for Version
    @property
    def version(self):
        return self.Version

    @Version.setter
    def Version(self, value):
        self.application.Version = value

    # Lower case alias for Version setter
    @version.setter
    def version(self, value):
        self.Version = value

    def ActiveExplorer(self):
        return self.application.ActiveExplorer()

    def ActiveInspector(self):
        return self.application.ActiveInspector()

    def ActiveWindow(self):
        return self.application.ActiveWindow()

    def AdvancedSearch(self, Scope=None, Filter=None, SearchSubFolders=None, Tag=None):
        arguments = com_arguments([Scope, Filter, SearchSubFolders, Tag])
        return Search(self.application.AdvancedSearch(*arguments))

    def CopyFile(self, FilePath=None, DestFolderPath=None):
        arguments = com_arguments([FilePath, DestFolderPath])
        return self.application.CopyFile(*arguments)

    def CreateItem(self, ItemType=None):
        arguments = com_arguments([ItemType])
        return self.application.CreateItem(*arguments)

    def CreateItemFromTemplate(self, TemplatePath=None, InFolder=None):
        arguments = com_arguments([TemplatePath, InFolder])
        return self.application.CreateItemFromTemplate(*arguments)

    def CreateObject(self, ObjectName=None):
        arguments = com_arguments([ObjectName])
        return self.application.CreateObject(*arguments)

    def GetNamespace(self, Type=None):
        arguments = com_arguments([Type])
        return self.application.GetNamespace(*arguments)

    def GetObjectReference(self, Item=None, ReferenceType=None):
        arguments = com_arguments([Item, ReferenceType])
        return self.application.GetObjectReference(*arguments)

    def IsSearchSynchronous(self, LookInFolders=None):
        arguments = com_arguments([LookInFolders])
        return self.application.IsSearchSynchronous(*arguments)

    def Quit(self):
        self.application.Quit()

    def RefreshFormRegionDefinition(self, RegionName=None):
        arguments = com_arguments([RegionName])
        self.application.RefreshFormRegionDefinition(*arguments)


class AppointmentItem:

    def __init__(self, appointmentitem=None):
        self.appointmentitem = appointmentitem

    @property
    def Actions(self):
        return Actions(self.appointmentitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def AllDayEvent(self):
        return self.appointmentitem.AllDayEvent

    # Lower case alias for AllDayEvent
    @property
    def alldayevent(self):
        return self.AllDayEvent

    @AllDayEvent.setter
    def AllDayEvent(self, value):
        self.appointmentitem.AllDayEvent = value

    # Lower case alias for AllDayEvent setter
    @alldayevent.setter
    def alldayevent(self, value):
        self.AllDayEvent = value

    @property
    def Application(self):
        return Application(self.appointmentitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.appointmentitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.appointmentitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.appointmentitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.appointmentitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.appointmentitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.appointmentitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def BusyStatus(self):
        return OlBusyStatus(self.appointmentitem.BusyStatus)

    # Lower case alias for BusyStatus
    @property
    def busystatus(self):
        return self.BusyStatus

    @BusyStatus.setter
    def BusyStatus(self, value):
        self.appointmentitem.BusyStatus = value

    # Lower case alias for BusyStatus setter
    @busystatus.setter
    def busystatus(self, value):
        self.BusyStatus = value

    @property
    def Categories(self):
        return self.appointmentitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.appointmentitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.appointmentitem.Class)

    @property
    def Companies(self):
        return self.appointmentitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.appointmentitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.appointmentitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.appointmentitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.appointmentitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.appointmentitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.appointmentitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.appointmentitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def Duration(self):
        return AppointmentItem(self.appointmentitem.Duration)

    # Lower case alias for Duration
    @property
    def duration(self):
        return self.Duration

    @Duration.setter
    def Duration(self, value):
        self.appointmentitem.Duration = value

    # Lower case alias for Duration setter
    @duration.setter
    def duration(self, value):
        self.Duration = value

    @property
    def End(self):
        return AppointmentItem(self.appointmentitem.End)

    # Lower case alias for End
    @property
    def end(self):
        return self.End

    @End.setter
    def End(self, value):
        self.appointmentitem.End = value

    # Lower case alias for End setter
    @end.setter
    def end(self, value):
        self.End = value

    @property
    def EndInEndTimeZone(self):
        return AppointmentItem.EndTimeZone(self.appointmentitem.EndInEndTimeZone)

    # Lower case alias for EndInEndTimeZone
    @property
    def endinendtimezone(self):
        return self.EndInEndTimeZone

    @EndInEndTimeZone.setter
    def EndInEndTimeZone(self, value):
        self.appointmentitem.EndInEndTimeZone = value

    # Lower case alias for EndInEndTimeZone setter
    @endinendtimezone.setter
    def endinendtimezone(self, value):
        self.EndInEndTimeZone = value

    @property
    def EndTimeZone(self):
        return TimeZone(self.appointmentitem.EndTimeZone)

    # Lower case alias for EndTimeZone
    @property
    def endtimezone(self):
        return self.EndTimeZone

    @EndTimeZone.setter
    def EndTimeZone(self, value):
        self.appointmentitem.EndTimeZone = value

    # Lower case alias for EndTimeZone setter
    @endtimezone.setter
    def endtimezone(self, value):
        self.EndTimeZone = value

    @property
    def EndUTC(self):
        return self.appointmentitem.EndUTC

    # Lower case alias for EndUTC
    @property
    def endutc(self):
        return self.EndUTC

    @EndUTC.setter
    def EndUTC(self, value):
        self.appointmentitem.EndUTC = value

    # Lower case alias for EndUTC setter
    @endutc.setter
    def endutc(self, value):
        self.EndUTC = value

    @property
    def EntryID(self):
        return self.appointmentitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ForceUpdateToAllAttendees(self):
        return self.appointmentitem.ForceUpdateToAllAttendees

    # Lower case alias for ForceUpdateToAllAttendees
    @property
    def forceupdatetoallattendees(self):
        return self.ForceUpdateToAllAttendees

    @ForceUpdateToAllAttendees.setter
    def ForceUpdateToAllAttendees(self, value):
        self.appointmentitem.ForceUpdateToAllAttendees = value

    # Lower case alias for ForceUpdateToAllAttendees setter
    @forceupdatetoallattendees.setter
    def forceupdatetoallattendees(self, value):
        self.ForceUpdateToAllAttendees = value

    @property
    def FormDescription(self):
        return FormDescription(self.appointmentitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.appointmentitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def GlobalAppointmentID(self):
        return AppointmentItem(self.appointmentitem.GlobalAppointmentID)

    # Lower case alias for GlobalAppointmentID
    @property
    def globalappointmentid(self):
        return self.GlobalAppointmentID

    @property
    def Importance(self):
        return OlImportance(self.appointmentitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.appointmentitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.appointmentitem.InternetCodepage

    # Lower case alias for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.appointmentitem.InternetCodepage = value

    # Lower case alias for InternetCodepage setter
    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.appointmentitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsRecurring(self):
        return self.appointmentitem.IsRecurring

    # Lower case alias for IsRecurring
    @property
    def isrecurring(self):
        return self.IsRecurring

    @property
    def ItemProperties(self):
        return ItemProperties(self.appointmentitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.appointmentitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def Location(self):
        return self.appointmentitem.Location

    # Lower case alias for Location
    @property
    def location(self):
        return self.Location

    @Location.setter
    def Location(self, value):
        self.appointmentitem.Location = value

    # Lower case alias for Location setter
    @location.setter
    def location(self, value):
        self.Location = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.appointmentitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.appointmentitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MeetingStatus(self):
        return OlMeetingStatus(self.appointmentitem.MeetingStatus)

    # Lower case alias for MeetingStatus
    @property
    def meetingstatus(self):
        return self.MeetingStatus

    @MeetingStatus.setter
    def MeetingStatus(self, value):
        self.appointmentitem.MeetingStatus = value

    # Lower case alias for MeetingStatus setter
    @meetingstatus.setter
    def meetingstatus(self, value):
        self.MeetingStatus = value

    @property
    def MeetingWorkspaceURL(self):
        return self.appointmentitem.MeetingWorkspaceURL

    # Lower case alias for MeetingWorkspaceURL
    @property
    def meetingworkspaceurl(self):
        return self.MeetingWorkspaceURL

    @property
    def MessageClass(self):
        return self.appointmentitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.appointmentitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.appointmentitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.appointmentitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.appointmentitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.appointmentitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OptionalAttendees(self):
        return self.appointmentitem.OptionalAttendees

    # Lower case alias for OptionalAttendees
    @property
    def optionalattendees(self):
        return self.OptionalAttendees

    @OptionalAttendees.setter
    def OptionalAttendees(self, value):
        self.appointmentitem.OptionalAttendees = value

    # Lower case alias for OptionalAttendees setter
    @optionalattendees.setter
    def optionalattendees(self, value):
        self.OptionalAttendees = value

    @property
    def Organizer(self):
        return self.appointmentitem.Organizer

    # Lower case alias for Organizer
    @property
    def organizer(self):
        return self.Organizer

    @property
    def OutlookInternalVersion(self):
        return self.appointmentitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.appointmentitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.appointmentitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.appointmentitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.appointmentitem.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def RecurrenceState(self):
        return OlRecurrenceState(self.appointmentitem.RecurrenceState)

    # Lower case alias for RecurrenceState
    @property
    def recurrencestate(self):
        return self.RecurrenceState

    @property
    def ReminderMinutesBeforeStart(self):
        return self.appointmentitem.ReminderMinutesBeforeStart

    # Lower case alias for ReminderMinutesBeforeStart
    @property
    def reminderminutesbeforestart(self):
        return self.ReminderMinutesBeforeStart

    @ReminderMinutesBeforeStart.setter
    def ReminderMinutesBeforeStart(self, value):
        self.appointmentitem.ReminderMinutesBeforeStart = value

    # Lower case alias for ReminderMinutesBeforeStart setter
    @reminderminutesbeforestart.setter
    def reminderminutesbeforestart(self, value):
        self.ReminderMinutesBeforeStart = value

    @property
    def ReminderOverrideDefault(self):
        return self.appointmentitem.ReminderOverrideDefault

    # Lower case alias for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.appointmentitem.ReminderOverrideDefault = value

    # Lower case alias for ReminderOverrideDefault setter
    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.appointmentitem.ReminderPlaySound

    # Lower case alias for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.appointmentitem.ReminderPlaySound = value

    # Lower case alias for ReminderPlaySound setter
    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.appointmentitem.ReminderSet

    # Lower case alias for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.appointmentitem.ReminderSet = value

    # Lower case alias for ReminderSet setter
    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.appointmentitem.ReminderSoundFile

    # Lower case alias for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.appointmentitem.ReminderSoundFile = value

    # Lower case alias for ReminderSoundFile setter
    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReplyTime(self):
        return self.appointmentitem.ReplyTime

    # Lower case alias for ReplyTime
    @property
    def replytime(self):
        return self.ReplyTime

    @ReplyTime.setter
    def ReplyTime(self, value):
        self.appointmentitem.ReplyTime = value

    # Lower case alias for ReplyTime setter
    @replytime.setter
    def replytime(self, value):
        self.ReplyTime = value

    @property
    def RequiredAttendees(self):
        return self.appointmentitem.RequiredAttendees

    # Lower case alias for RequiredAttendees
    @property
    def requiredattendees(self):
        return self.RequiredAttendees

    @RequiredAttendees.setter
    def RequiredAttendees(self, value):
        self.appointmentitem.RequiredAttendees = value

    # Lower case alias for RequiredAttendees setter
    @requiredattendees.setter
    def requiredattendees(self, value):
        self.RequiredAttendees = value

    @property
    def Resources(self):
        return self.appointmentitem.Resources

    # Lower case alias for Resources
    @property
    def resources(self):
        return self.Resources

    @Resources.setter
    def Resources(self, value):
        self.appointmentitem.Resources = value

    # Lower case alias for Resources setter
    @resources.setter
    def resources(self, value):
        self.Resources = value

    @property
    def ResponseRequested(self):
        return self.appointmentitem.ResponseRequested

    # Lower case alias for ResponseRequested
    @property
    def responserequested(self):
        return self.ResponseRequested

    @ResponseRequested.setter
    def ResponseRequested(self, value):
        self.appointmentitem.ResponseRequested = value

    # Lower case alias for ResponseRequested setter
    @responserequested.setter
    def responserequested(self, value):
        self.ResponseRequested = value

    @property
    def ResponseStatus(self):
        return OlResponseStatus(self.appointmentitem.ResponseStatus)

    # Lower case alias for ResponseStatus
    @property
    def responsestatus(self):
        return self.ResponseStatus

    @property
    def RTFBody(self):
        return self.appointmentitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.appointmentitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.appointmentitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SendUsingAccount(self):
        return Account(self.appointmentitem.SendUsingAccount)

    # Lower case alias for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.appointmentitem.SendUsingAccount = value

    # Lower case alias for SendUsingAccount setter
    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.appointmentitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.appointmentitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.appointmentitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.appointmentitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Start(self):
        return self.appointmentitem.Start

    # Lower case alias for Start
    @property
    def start(self):
        return self.Start

    @Start.setter
    def Start(self, value):
        self.appointmentitem.Start = value

    # Lower case alias for Start setter
    @start.setter
    def start(self, value):
        self.Start = value

    @property
    def StartInStartTimeZone(self):
        return AppointmentItem.StartTimeZone(self.appointmentitem.StartInStartTimeZone)

    # Lower case alias for StartInStartTimeZone
    @property
    def startinstarttimezone(self):
        return self.StartInStartTimeZone

    @StartInStartTimeZone.setter
    def StartInStartTimeZone(self, value):
        self.appointmentitem.StartInStartTimeZone = value

    # Lower case alias for StartInStartTimeZone setter
    @startinstarttimezone.setter
    def startinstarttimezone(self, value):
        self.StartInStartTimeZone = value

    @property
    def StartTimeZone(self):
        return TimeZone(self.appointmentitem.StartTimeZone)

    # Lower case alias for StartTimeZone
    @property
    def starttimezone(self):
        return self.StartTimeZone

    @StartTimeZone.setter
    def StartTimeZone(self, value):
        self.appointmentitem.StartTimeZone = value

    # Lower case alias for StartTimeZone setter
    @starttimezone.setter
    def starttimezone(self, value):
        self.StartTimeZone = value

    @property
    def StartUTC(self):
        return self.appointmentitem.StartUTC

    # Lower case alias for StartUTC
    @property
    def startutc(self):
        return self.StartUTC

    @StartUTC.setter
    def StartUTC(self, value):
        self.appointmentitem.StartUTC = value

    # Lower case alias for StartUTC setter
    @startutc.setter
    def startutc(self, value):
        self.StartUTC = value

    @property
    def Subject(self):
        return self.appointmentitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.appointmentitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.appointmentitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.appointmentitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.appointmentitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def ClearRecurrencePattern(self):
        self.appointmentitem.ClearRecurrencePattern()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.appointmentitem.Close(*arguments)

    def Copy(self):
        self.appointmentitem.Copy()

    def CopyTo(self, DestinationFolder=None, CopyOptions=None):
        arguments = com_arguments([DestinationFolder, CopyOptions])
        return self.appointmentitem.CopyTo(*arguments)

    def Delete(self):
        self.appointmentitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.appointmentitem.Display(*arguments)

    def ForwardAsVcal(self):
        return MailItem(self.appointmentitem.ForwardAsVcal())

    def GetConversation(self):
        return self.appointmentitem.GetConversation()

    def GetOrganizer(self):
        return self.appointmentitem.GetOrganizer()

    def GetRecurrencePattern(self):
        return self.appointmentitem.GetRecurrencePattern()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.appointmentitem.Move(*arguments)

    def PrintOut(self):
        self.appointmentitem.PrintOut()

    def Respond(self, Response=None, fNoUI=None, fAdditionalTextDialog=None):
        arguments = com_arguments([Response, fNoUI, fAdditionalTextDialog])
        return MeetingItem(self.appointmentitem.Respond(*arguments))

    def Save(self):
        self.appointmentitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.appointmentitem.SaveAs(*arguments)

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

    # Lower case alias for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.assigntocategoryruleaction.Application)

    @property
    def Categories(self):
        return self.assigntocategoryruleaction.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.assigntocategoryruleaction.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.assigntocategoryruleaction.Class)

    @property
    def Enabled(self):
        return self.assigntocategoryruleaction.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.assigntocategoryruleaction.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.assigntocategoryruleaction.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.assigntocategoryruleaction.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class Attachment:

    def __init__(self, attachment=None):
        self.attachment = attachment

    @property
    def Application(self):
        return Application(self.attachment.Application)

    @property
    def BlockLevel(self):
        return OlAttachmentBlockLevel(self.attachment.BlockLevel)

    # Lower case alias for BlockLevel
    @property
    def blocklevel(self):
        return self.BlockLevel

    @property
    def Class(self):
        return OlObjectClass(self.attachment.Class)

    @property
    def DisplayName(self):
        return self.attachment.DisplayName

    # Lower case alias for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @DisplayName.setter
    def DisplayName(self, value):
        self.attachment.DisplayName = value

    # Lower case alias for DisplayName setter
    @displayname.setter
    def displayname(self, value):
        self.DisplayName = value

    @property
    def FileName(self):
        return self.attachment.FileName

    # Lower case alias for FileName
    @property
    def filename(self):
        return self.FileName

    @property
    def Index(self):
        return self.attachment.Index

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def Parent(self):
        return self.attachment.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PathName(self):
        return self.attachment.PathName

    # Lower case alias for PathName
    @property
    def pathname(self):
        return self.PathName

    @property
    def Position(self):
        return self.attachment.Position

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.attachment.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.attachment.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.attachment.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.attachment.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Type(self):
        return OlAttachmentType(self.attachment.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    def Delete(self):
        self.attachment.Delete()

    def GetTemporaryFilePath(self):
        return self.attachment.GetTemporaryFilePath()

    def SaveAsFile(self, Path=None):
        arguments = com_arguments([Path])
        self.attachment.SaveAsFile(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.attachments.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.attachments.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Source=None, Type=None, Position=None, DisplayName=None):
        arguments = com_arguments([Source, Type, Position, DisplayName])
        return Attachment(self.attachments.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.attachments.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.attachments.Remove(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Location(self):
        return OlSelectionLocation(self.attachmentselection.Location)

    # Lower case alias for Location
    @property
    def location(self):
        return self.Location

    @property
    def Parent(self):
        return self.attachmentselection.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.attachmentselection.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def GetSelection(self, SelectionContents=None):
        arguments = com_arguments([SelectionContents])
        return self.attachmentselection.GetSelection(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.attachmentselection.Item(*arguments)


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

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.autoformatrule.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Filter(self):
        return self.autoformatrule.Filter

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @Filter.setter
    def Filter(self, value):
        self.autoformatrule.Filter = value

    # Lower case alias for Filter setter
    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def Font(self):
        return ViewFont(self.autoformatrule.Font)

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def Name(self):
        return self.autoformatrule.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.autoformatrule.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.autoformatrule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.autoformatrule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Standard(self):
        return AutoFormatRule(self.autoformatrule.Standard)

    # Lower case alias for Standard
    @property
    def standard(self):
        return self.Standard


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.autoformatrules.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.autoformatrules.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([Name])
        return self.autoformatrules.Add(*arguments)

    def Insert(self, Name=None, Index=None):
        arguments = com_arguments([Name, Index])
        return self.autoformatrules.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.autoformatrules.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.autoformatrules.Remove(*arguments)

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

    # Lower case alias for CardSize
    @property
    def cardsize(self):
        return self.CardSize

    @CardSize.setter
    def CardSize(self, value):
        self.businesscardview.CardSize = value

    # Lower case alias for CardSize setter
    @cardsize.setter
    def cardsize(self, value):
        self.CardSize = value

    @property
    def Class(self):
        return OlObjectClass(self.businesscardview.Class)

    @property
    def Filter(self):
        return self.businesscardview.Filter

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @Filter.setter
    def Filter(self, value):
        self.businesscardview.Filter = value

    # Lower case alias for Filter setter
    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def HeadingsFont(self):
        return ViewFont(self.businesscardview.HeadingsFont)

    # Lower case alias for HeadingsFont
    @property
    def headingsfont(self):
        return self.HeadingsFont

    @property
    def Language(self):
        return self.businesscardview.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @Language.setter
    def Language(self, value):
        self.businesscardview.Language = value

    # Lower case alias for Language setter
    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.businesscardview.LockUserChanges

    # Lower case alias for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.businesscardview.LockUserChanges = value

    # Lower case alias for LockUserChanges setter
    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.businesscardview.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.businesscardview.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.businesscardview.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.businesscardview.SaveOption)

    # Lower case alias for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.businesscardview.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def SortFields(self):
        return OrderFields(self.businesscardview.SortFields)

    # Lower case alias for SortFields
    @property
    def sortfields(self):
        return self.SortFields

    @property
    def Standard(self):
        return BusinessCardView(self.businesscardview.Standard)

    # Lower case alias for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.businesscardview.ViewType)

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.businesscardview.XML

    # Lower case alias for XML
    @property
    def xml(self):
        return self.XML

    @XML.setter
    def XML(self, value):
        self.businesscardview.XML = value

    # Lower case alias for XML setter
    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.businesscardview.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return self.businesscardview.Copy(*arguments)

    def Delete(self):
        self.businesscardview.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.businesscardview.GoToDate(*arguments)

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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.calendarmodule.NavigationGroups)

    # Lower case alias for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.calendarmodule.NavigationModuleType)

    # Lower case alias for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.calendarmodule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return CalendarModule(self.calendarmodule.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.calendarmodule.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.calendarmodule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return CalendarModule(self.calendarmodule.Visible)

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.calendarmodule.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


class CalendarSharing:

    def __init__(self, calendarsharing=None):
        self.calendarsharing = calendarsharing

    @property
    def Application(self):
        return Application(self.calendarsharing.Application)

    @property
    def CalendarDetail(self):
        return OlCalendarDetail(self.calendarsharing.CalendarDetail)

    # Lower case alias for CalendarDetail
    @property
    def calendardetail(self):
        return self.CalendarDetail

    @CalendarDetail.setter
    def CalendarDetail(self, value):
        self.calendarsharing.CalendarDetail = value

    # Lower case alias for CalendarDetail setter
    @calendardetail.setter
    def calendardetail(self, value):
        self.CalendarDetail = value

    @property
    def Class(self):
        return OlObjectClass(self.calendarsharing.Class)

    @property
    def EndDate(self):
        return CalendarSharing(self.calendarsharing.EndDate)

    # Lower case alias for EndDate
    @property
    def enddate(self):
        return self.EndDate

    @EndDate.setter
    def EndDate(self, value):
        self.calendarsharing.EndDate = value

    # Lower case alias for EndDate setter
    @enddate.setter
    def enddate(self, value):
        self.EndDate = value

    @property
    def Folder(self):
        return Folder(self.calendarsharing.Folder)

    # Lower case alias for Folder
    @property
    def folder(self):
        return self.Folder

    @property
    def IncludeAttachments(self):
        return self.calendarsharing.IncludeAttachments

    # Lower case alias for IncludeAttachments
    @property
    def includeattachments(self):
        return self.IncludeAttachments

    @IncludeAttachments.setter
    def IncludeAttachments(self, value):
        self.calendarsharing.IncludeAttachments = value

    # Lower case alias for IncludeAttachments setter
    @includeattachments.setter
    def includeattachments(self, value):
        self.IncludeAttachments = value

    @property
    def IncludePrivateDetails(self):
        return self.calendarsharing.IncludePrivateDetails

    # Lower case alias for IncludePrivateDetails
    @property
    def includeprivatedetails(self):
        return self.IncludePrivateDetails

    @IncludePrivateDetails.setter
    def IncludePrivateDetails(self, value):
        self.calendarsharing.IncludePrivateDetails = value

    # Lower case alias for IncludePrivateDetails setter
    @includeprivatedetails.setter
    def includeprivatedetails(self, value):
        self.IncludePrivateDetails = value

    @property
    def IncludeWholeCalendar(self):
        return self.calendarsharing.IncludeWholeCalendar

    # Lower case alias for IncludeWholeCalendar
    @property
    def includewholecalendar(self):
        return self.IncludeWholeCalendar

    @IncludeWholeCalendar.setter
    def IncludeWholeCalendar(self, value):
        self.calendarsharing.IncludeWholeCalendar = value

    # Lower case alias for IncludeWholeCalendar setter
    @includewholecalendar.setter
    def includewholecalendar(self, value):
        self.IncludeWholeCalendar = value

    @property
    def Parent(self):
        return self.calendarsharing.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RestrictToWorkingHours(self):
        return self.calendarsharing.RestrictToWorkingHours

    # Lower case alias for RestrictToWorkingHours
    @property
    def restricttoworkinghours(self):
        return self.RestrictToWorkingHours

    @RestrictToWorkingHours.setter
    def RestrictToWorkingHours(self, value):
        self.calendarsharing.RestrictToWorkingHours = value

    # Lower case alias for RestrictToWorkingHours setter
    @restricttoworkinghours.setter
    def restricttoworkinghours(self, value):
        self.RestrictToWorkingHours = value

    @property
    def Session(self):
        return NameSpace(self.calendarsharing.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def StartDate(self):
        return CalendarSharing(self.calendarsharing.StartDate)

    # Lower case alias for StartDate
    @property
    def startdate(self):
        return self.StartDate

    @StartDate.setter
    def StartDate(self, value):
        self.calendarsharing.StartDate = value

    # Lower case alias for StartDate setter
    @startdate.setter
    def startdate(self, value):
        self.StartDate = value

    def ForwardAsICal(self, MailFormat=None):
        arguments = com_arguments([MailFormat])
        return self.calendarsharing.ForwardAsICal(*arguments)

    def SaveAsICal(self, Path=None):
        arguments = com_arguments([Path])
        self.calendarsharing.SaveAsICal(*arguments)


class CalendarView:

    def __init__(self, calendarview=None):
        self.calendarview = calendarview

    @property
    def Application(self):
        return Application(self.calendarview.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.calendarview.AutoFormatRules)

    # Lower case alias for AutoFormatRules
    @property
    def autoformatrules(self):
        return self.AutoFormatRules

    @property
    def BoldDatesWithItems(self):
        return CalendarView(self.calendarview.BoldDatesWithItems)

    # Lower case alias for BoldDatesWithItems
    @property
    def bolddateswithitems(self):
        return self.BoldDatesWithItems

    @BoldDatesWithItems.setter
    def BoldDatesWithItems(self, value):
        self.calendarview.BoldDatesWithItems = value

    # Lower case alias for BoldDatesWithItems setter
    @bolddateswithitems.setter
    def bolddateswithitems(self, value):
        self.BoldDatesWithItems = value

    @property
    def BoldSubjects(self):
        return CalendarView(self.calendarview.BoldSubjects)

    # Lower case alias for BoldSubjects
    @property
    def boldsubjects(self):
        return self.BoldSubjects

    @BoldSubjects.setter
    def BoldSubjects(self, value):
        self.calendarview.BoldSubjects = value

    # Lower case alias for BoldSubjects setter
    @boldsubjects.setter
    def boldsubjects(self, value):
        self.BoldSubjects = value

    @property
    def CalendarViewMode(self):
        return OlCalendarViewMode(self.calendarview.CalendarViewMode)

    # Lower case alias for CalendarViewMode
    @property
    def calendarviewmode(self):
        return self.CalendarViewMode

    @CalendarViewMode.setter
    def CalendarViewMode(self, value):
        self.calendarview.CalendarViewMode = value

    # Lower case alias for CalendarViewMode setter
    @calendarviewmode.setter
    def calendarviewmode(self, value):
        self.CalendarViewMode = value

    @property
    def Class(self):
        return OlObjectClass(self.calendarview.Class)

    @property
    def DaysInMultiDayMode(self):
        return CalendarView(self.calendarview.DaysInMultiDayMode)

    # Lower case alias for DaysInMultiDayMode
    @property
    def daysinmultidaymode(self):
        return self.DaysInMultiDayMode

    @DaysInMultiDayMode.setter
    def DaysInMultiDayMode(self, value):
        self.calendarview.DaysInMultiDayMode = value

    # Lower case alias for DaysInMultiDayMode setter
    @daysinmultidaymode.setter
    def daysinmultidaymode(self, value):
        self.DaysInMultiDayMode = value

    @property
    def DayWeekTimeScale(self):
        return OlDayWeekTimeScale(self.calendarview.DayWeekTimeScale)

    # Lower case alias for DayWeekTimeScale
    @property
    def dayweektimescale(self):
        return self.DayWeekTimeScale

    @DayWeekTimeScale.setter
    def DayWeekTimeScale(self, value):
        self.calendarview.DayWeekTimeScale = value

    # Lower case alias for DayWeekTimeScale setter
    @dayweektimescale.setter
    def dayweektimescale(self, value):
        self.DayWeekTimeScale = value

    @property
    def DisplayedDates(self):
        return CalendarView(self.calendarview.DisplayedDates)

    # Lower case alias for DisplayedDates
    @property
    def displayeddates(self):
        return self.DisplayedDates

    @property
    def EndField(self):
        return CalendarView(self.calendarview.EndField)

    # Lower case alias for EndField
    @property
    def endfield(self):
        return self.EndField

    @EndField.setter
    def EndField(self, value):
        self.calendarview.EndField = value

    # Lower case alias for EndField setter
    @endfield.setter
    def endfield(self, value):
        self.EndField = value

    @property
    def Filter(self):
        return self.calendarview.Filter

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @Filter.setter
    def Filter(self, value):
        self.calendarview.Filter = value

    # Lower case alias for Filter setter
    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def Language(self):
        return self.calendarview.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @Language.setter
    def Language(self, value):
        self.calendarview.Language = value

    # Lower case alias for Language setter
    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.calendarview.LockUserChanges

    # Lower case alias for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.calendarview.LockUserChanges = value

    # Lower case alias for LockUserChanges setter
    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def MonthShowEndTime(self):
        return CalendarView(self.calendarview.MonthShowEndTime)

    # Lower case alias for MonthShowEndTime
    @property
    def monthshowendtime(self):
        return self.MonthShowEndTime

    @MonthShowEndTime.setter
    def MonthShowEndTime(self, value):
        self.calendarview.MonthShowEndTime = value

    # Lower case alias for MonthShowEndTime setter
    @monthshowendtime.setter
    def monthshowendtime(self, value):
        self.MonthShowEndTime = value

    @property
    def Name(self):
        return self.calendarview.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.calendarview.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.calendarview.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.calendarview.SaveOption)

    # Lower case alias for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def SelectedEndTime(self):
        return CalendarView(self.calendarview.SelectedEndTime)

    # Lower case alias for SelectedEndTime
    @property
    def selectedendtime(self):
        return self.SelectedEndTime

    @property
    def SelectedStartTime(self):
        return CalendarView(self.calendarview.SelectedStartTime)

    # Lower case alias for SelectedStartTime
    @property
    def selectedstarttime(self):
        return self.SelectedStartTime

    @property
    def Session(self):
        return NameSpace(self.calendarview.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Standard(self):
        return CalendarView(self.calendarview.Standard)

    # Lower case alias for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def StartField(self):
        return CalendarView(self.calendarview.StartField)

    # Lower case alias for StartField
    @property
    def startfield(self):
        return self.StartField

    @StartField.setter
    def StartField(self, value):
        self.calendarview.StartField = value

    # Lower case alias for StartField setter
    @startfield.setter
    def startfield(self, value):
        self.StartField = value

    @property
    def ViewType(self):
        return OlViewType(self.calendarview.ViewType)

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.calendarview.XML

    # Lower case alias for XML
    @property
    def xml(self):
        return self.XML

    @XML.setter
    def XML(self, value):
        self.calendarview.XML = value

    # Lower case alias for XML setter
    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.calendarview.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return self.calendarview.Copy(*arguments)

    def Delete(self):
        self.calendarview.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.calendarview.GoToDate(*arguments)

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

    # Lower case alias for AllowInCellEditing
    @property
    def allowincellediting(self):
        return self.AllowInCellEditing

    @AllowInCellEditing.setter
    def AllowInCellEditing(self, value):
        self.cardview.AllowInCellEditing = value

    # Lower case alias for AllowInCellEditing setter
    @allowincellediting.setter
    def allowincellediting(self, value):
        self.AllowInCellEditing = value

    @property
    def Application(self):
        return Application(self.cardview.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.cardview.AutoFormatRules)

    # Lower case alias for AutoFormatRules
    @property
    def autoformatrules(self):
        return self.AutoFormatRules

    @property
    def BodyFont(self):
        return ViewFont(self.cardview.BodyFont)

    # Lower case alias for BodyFont
    @property
    def bodyfont(self):
        return self.BodyFont

    @property
    def Class(self):
        return OlObjectClass(self.cardview.Class)

    @property
    def Filter(self):
        return self.cardview.Filter

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @Filter.setter
    def Filter(self, value):
        self.cardview.Filter = value

    # Lower case alias for Filter setter
    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def HeadingsFont(self):
        return ViewFont(self.cardview.HeadingsFont)

    # Lower case alias for HeadingsFont
    @property
    def headingsfont(self):
        return self.HeadingsFont

    @property
    def Language(self):
        return self.cardview.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @Language.setter
    def Language(self, value):
        self.cardview.Language = value

    # Lower case alias for Language setter
    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.cardview.LockUserChanges

    # Lower case alias for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.cardview.LockUserChanges = value

    # Lower case alias for LockUserChanges setter
    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def MultiLineFieldHeight(self):
        return CardView(self.cardview.MultiLineFieldHeight)

    # Lower case alias for MultiLineFieldHeight
    @property
    def multilinefieldheight(self):
        return self.MultiLineFieldHeight

    @MultiLineFieldHeight.setter
    def MultiLineFieldHeight(self, value):
        self.cardview.MultiLineFieldHeight = value

    # Lower case alias for MultiLineFieldHeight setter
    @multilinefieldheight.setter
    def multilinefieldheight(self, value):
        self.MultiLineFieldHeight = value

    @property
    def Name(self):
        return self.cardview.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.cardview.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.cardview.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.cardview.SaveOption)

    # Lower case alias for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.cardview.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowEmptyFields(self):
        return CardView(self.cardview.ShowEmptyFields)

    # Lower case alias for ShowEmptyFields
    @property
    def showemptyfields(self):
        return self.ShowEmptyFields

    @ShowEmptyFields.setter
    def ShowEmptyFields(self, value):
        self.cardview.ShowEmptyFields = value

    # Lower case alias for ShowEmptyFields setter
    @showemptyfields.setter
    def showemptyfields(self, value):
        self.ShowEmptyFields = value

    @property
    def SortFields(self):
        return OrderFields(self.cardview.SortFields)

    # Lower case alias for SortFields
    @property
    def sortfields(self):
        return self.SortFields

    @property
    def Standard(self):
        return CardView(self.cardview.Standard)

    # Lower case alias for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewFields(self):
        return ViewFields(self.cardview.ViewFields)

    # Lower case alias for ViewFields
    @property
    def viewfields(self):
        return self.ViewFields

    @property
    def ViewType(self):
        return OlViewType(self.cardview.ViewType)

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def Width(self):
        return CardView(self.cardview.Width)

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.cardview.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def XML(self):
        return self.cardview.XML

    # Lower case alias for XML
    @property
    def xml(self):
        return self.XML

    @XML.setter
    def XML(self, value):
        self.cardview.XML = value

    # Lower case alias for XML setter
    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.cardview.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return self.cardview.Copy(*arguments)

    def Delete(self):
        self.cardview.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.cardview.GoToDate(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.categories.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.categories.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Color=None, ShortcutKey=None):
        arguments = com_arguments([Name, Color, ShortcutKey])
        return self.categories.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.categories.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.categories.Remove(*arguments)


class Category:

    def __init__(self, category=None):
        self.category = category

    @property
    def Application(self):
        return Application(self.category.Application)

    @property
    def CategoryBorderColor(self):
        return Category(self.category.CategoryBorderColor)

    # Lower case alias for CategoryBorderColor
    @property
    def categorybordercolor(self):
        return self.CategoryBorderColor

    @property
    def CategoryGradientBottomColor(self):
        return Category(self.category.CategoryGradientBottomColor)

    # Lower case alias for CategoryGradientBottomColor
    @property
    def categorygradientbottomcolor(self):
        return self.CategoryGradientBottomColor

    @property
    def CategoryGradientTopColor(self):
        return Category(self.category.CategoryGradientTopColor)

    # Lower case alias for CategoryGradientTopColor
    @property
    def categorygradienttopcolor(self):
        return self.CategoryGradientTopColor

    @property
    def CategoryID(self):
        return Category(self.category.CategoryID)

    # Lower case alias for CategoryID
    @property
    def categoryid(self):
        return self.CategoryID

    @property
    def Class(self):
        return OlObjectClass(self.category.Class)

    @property
    def Color(self):
        return OlCategoryColor(self.category.Color)

    # Lower case alias for Color
    @property
    def color(self):
        return self.Color

    @Color.setter
    def Color(self, value):
        self.category.Color = value

    # Lower case alias for Color setter
    @color.setter
    def color(self, value):
        self.Color = value

    @property
    def Name(self):
        return self.category.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.category.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.category.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.category.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShortcutKey(self):
        return OlCategoryShortcutKey(self.category.ShortcutKey)

    # Lower case alias for ShortcutKey
    @property
    def shortcutkey(self):
        return self.ShortcutKey

    @ShortcutKey.setter
    def ShortcutKey(self, value):
        self.category.ShortcutKey = value

    # Lower case alias for ShortcutKey setter
    @shortcutkey.setter
    def shortcutkey(self, value):
        self.ShortcutKey = value


class CategoryRuleCondition:

    def __init__(self, categoryrulecondition=None):
        self.categoryrulecondition = categoryrulecondition

    @property
    def Application(self):
        return Application(self.categoryrulecondition.Application)

    @property
    def Categories(self):
        return self.categoryrulecondition.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.categoryrulecondition.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.categoryrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.categoryrulecondition.ConditionType)

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.categoryrulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.categoryrulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.categoryrulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.categoryrulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return Column(self.column.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.column.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class ColumnFormat:

    def __init__(self, columnformat=None):
        self.columnformat = columnformat

    @property
    def Align(self):
        return OlAlign(self.columnformat.Align)

    # Lower case alias for Align
    @property
    def align(self):
        return self.Align

    @Align.setter
    def Align(self, value):
        self.columnformat.Align = value

    # Lower case alias for Align setter
    @align.setter
    def align(self, value):
        self.Align = value

    @property
    def Application(self):
        return Application(self.columnformat.Application)

    @property
    def Class(self):
        return OlObjectClass(self.columnformat.Class)

    @property
    def FieldFormat(self):
        return ColumnFormat(self.columnformat.FieldFormat)

    # Lower case alias for FieldFormat
    @property
    def fieldformat(self):
        return self.FieldFormat

    @FieldFormat.setter
    def FieldFormat(self, value):
        self.columnformat.FieldFormat = value

    # Lower case alias for FieldFormat setter
    @fieldformat.setter
    def fieldformat(self, value):
        self.FieldFormat = value

    @property
    def FieldType(self):
        return OlUserPropertyType(self.columnformat.FieldType)

    # Lower case alias for FieldType
    @property
    def fieldtype(self):
        return self.FieldType

    @property
    def Label(self):
        return ColumnFormat(self.columnformat.Label)

    # Lower case alias for Label
    @property
    def label(self):
        return self.Label

    @Label.setter
    def Label(self, value):
        self.columnformat.Label = value

    # Lower case alias for Label setter
    @label.setter
    def label(self, value):
        self.Label = value

    @property
    def Parent(self):
        return self.columnformat.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.columnformat.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Width(self):
        return self.columnformat.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.columnformat.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return Columns(self.columns.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.columns.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([Name])
        return self.columns.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Table(self.columns.Item(*arguments))

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.columns.Remove(*arguments)

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

    # Lower case alias for Item
    @property
    def item(self):
        return self.Item

    @property
    def Name(self):
        return self.conflict.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.conflict.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.conflict.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return OlObjectClass(self.conflict.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.conflicts.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.conflicts.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def GetFirst(self):
        return Conflict(self.conflicts.GetFirst())

    def GetLast(self):
        return Conflict(self.conflicts.GetLast())

    def GetNext(self):
        return Conflict(self.conflicts.GetNext())

    def GetPrevious(self):
        return Conflict(self.conflicts.GetPrevious())

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.conflicts.Item(*arguments)


class ContactItem:

    def __init__(self, contactitem=None):
        self.contactitem = contactitem

    @property
    def Account(self):
        return self.contactitem.Account

    # Lower case alias for Account
    @property
    def account(self):
        return self.Account

    @Account.setter
    def Account(self, value):
        self.contactitem.Account = value

    # Lower case alias for Account setter
    @account.setter
    def account(self, value):
        self.Account = value

    @property
    def Actions(self):
        return Actions(self.contactitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Anniversary(self):
        return self.contactitem.Anniversary

    # Lower case alias for Anniversary
    @property
    def anniversary(self):
        return self.Anniversary

    @Anniversary.setter
    def Anniversary(self, value):
        self.contactitem.Anniversary = value

    # Lower case alias for Anniversary setter
    @anniversary.setter
    def anniversary(self, value):
        self.Anniversary = value

    @property
    def Application(self):
        return Application(self.contactitem.Application)

    @property
    def AssistantName(self):
        return self.contactitem.AssistantName

    # Lower case alias for AssistantName
    @property
    def assistantname(self):
        return self.AssistantName

    @AssistantName.setter
    def AssistantName(self, value):
        self.contactitem.AssistantName = value

    # Lower case alias for AssistantName setter
    @assistantname.setter
    def assistantname(self, value):
        self.AssistantName = value

    @property
    def AssistantTelephoneNumber(self):
        return self.contactitem.AssistantTelephoneNumber

    # Lower case alias for AssistantTelephoneNumber
    @property
    def assistanttelephonenumber(self):
        return self.AssistantTelephoneNumber

    @AssistantTelephoneNumber.setter
    def AssistantTelephoneNumber(self, value):
        self.contactitem.AssistantTelephoneNumber = value

    # Lower case alias for AssistantTelephoneNumber setter
    @assistanttelephonenumber.setter
    def assistanttelephonenumber(self, value):
        self.AssistantTelephoneNumber = value

    @property
    def Attachments(self):
        return Attachments(self.contactitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.contactitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.contactitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.contactitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Birthday(self):
        return self.contactitem.Birthday

    # Lower case alias for Birthday
    @property
    def birthday(self):
        return self.Birthday

    @Birthday.setter
    def Birthday(self, value):
        self.contactitem.Birthday = value

    # Lower case alias for Birthday setter
    @birthday.setter
    def birthday(self, value):
        self.Birthday = value

    @property
    def Body(self):
        return self.contactitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.contactitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Business2TelephoneNumber(self):
        return self.contactitem.Business2TelephoneNumber

    # Lower case alias for Business2TelephoneNumber
    @property
    def business2telephonenumber(self):
        return self.Business2TelephoneNumber

    @Business2TelephoneNumber.setter
    def Business2TelephoneNumber(self, value):
        self.contactitem.Business2TelephoneNumber = value

    # Lower case alias for Business2TelephoneNumber setter
    @business2telephonenumber.setter
    def business2telephonenumber(self, value):
        self.Business2TelephoneNumber = value

    @property
    def BusinessAddress(self):
        return self.contactitem.BusinessAddress

    # Lower case alias for BusinessAddress
    @property
    def businessaddress(self):
        return self.BusinessAddress

    @BusinessAddress.setter
    def BusinessAddress(self, value):
        self.contactitem.BusinessAddress = value

    # Lower case alias for BusinessAddress setter
    @businessaddress.setter
    def businessaddress(self, value):
        self.BusinessAddress = value

    @property
    def BusinessAddressCity(self):
        return self.contactitem.BusinessAddressCity

    # Lower case alias for BusinessAddressCity
    @property
    def businessaddresscity(self):
        return self.BusinessAddressCity

    @BusinessAddressCity.setter
    def BusinessAddressCity(self, value):
        self.contactitem.BusinessAddressCity = value

    # Lower case alias for BusinessAddressCity setter
    @businessaddresscity.setter
    def businessaddresscity(self, value):
        self.BusinessAddressCity = value

    @property
    def BusinessAddressCountry(self):
        return self.contactitem.BusinessAddressCountry

    # Lower case alias for BusinessAddressCountry
    @property
    def businessaddresscountry(self):
        return self.BusinessAddressCountry

    @BusinessAddressCountry.setter
    def BusinessAddressCountry(self, value):
        self.contactitem.BusinessAddressCountry = value

    # Lower case alias for BusinessAddressCountry setter
    @businessaddresscountry.setter
    def businessaddresscountry(self, value):
        self.BusinessAddressCountry = value

    @property
    def BusinessAddressPostalCode(self):
        return self.contactitem.BusinessAddressPostalCode

    # Lower case alias for BusinessAddressPostalCode
    @property
    def businessaddresspostalcode(self):
        return self.BusinessAddressPostalCode

    @BusinessAddressPostalCode.setter
    def BusinessAddressPostalCode(self, value):
        self.contactitem.BusinessAddressPostalCode = value

    # Lower case alias for BusinessAddressPostalCode setter
    @businessaddresspostalcode.setter
    def businessaddresspostalcode(self, value):
        self.BusinessAddressPostalCode = value

    @property
    def BusinessAddressPostOfficeBox(self):
        return self.contactitem.BusinessAddressPostOfficeBox

    # Lower case alias for BusinessAddressPostOfficeBox
    @property
    def businessaddresspostofficebox(self):
        return self.BusinessAddressPostOfficeBox

    @BusinessAddressPostOfficeBox.setter
    def BusinessAddressPostOfficeBox(self, value):
        self.contactitem.BusinessAddressPostOfficeBox = value

    # Lower case alias for BusinessAddressPostOfficeBox setter
    @businessaddresspostofficebox.setter
    def businessaddresspostofficebox(self, value):
        self.BusinessAddressPostOfficeBox = value

    @property
    def BusinessAddressState(self):
        return self.contactitem.BusinessAddressState

    # Lower case alias for BusinessAddressState
    @property
    def businessaddressstate(self):
        return self.BusinessAddressState

    @BusinessAddressState.setter
    def BusinessAddressState(self, value):
        self.contactitem.BusinessAddressState = value

    # Lower case alias for BusinessAddressState setter
    @businessaddressstate.setter
    def businessaddressstate(self, value):
        self.BusinessAddressState = value

    @property
    def BusinessAddressStreet(self):
        return self.contactitem.BusinessAddressStreet

    # Lower case alias for BusinessAddressStreet
    @property
    def businessaddressstreet(self):
        return self.BusinessAddressStreet

    @BusinessAddressStreet.setter
    def BusinessAddressStreet(self, value):
        self.contactitem.BusinessAddressStreet = value

    # Lower case alias for BusinessAddressStreet setter
    @businessaddressstreet.setter
    def businessaddressstreet(self, value):
        self.BusinessAddressStreet = value

    @property
    def BusinessCardLayoutXml(self):
        return self.contactitem.BusinessCardLayoutXml

    # Lower case alias for BusinessCardLayoutXml
    @property
    def businesscardlayoutxml(self):
        return self.BusinessCardLayoutXml

    @BusinessCardLayoutXml.setter
    def BusinessCardLayoutXml(self, value):
        self.contactitem.BusinessCardLayoutXml = value

    # Lower case alias for BusinessCardLayoutXml setter
    @businesscardlayoutxml.setter
    def businesscardlayoutxml(self, value):
        self.BusinessCardLayoutXml = value

    @property
    def BusinessCardType(self):
        return OlBusinessCardType(self.contactitem.BusinessCardType)

    # Lower case alias for BusinessCardType
    @property
    def businesscardtype(self):
        return self.BusinessCardType

    @property
    def BusinessFaxNumber(self):
        return self.contactitem.BusinessFaxNumber

    # Lower case alias for BusinessFaxNumber
    @property
    def businessfaxnumber(self):
        return self.BusinessFaxNumber

    @BusinessFaxNumber.setter
    def BusinessFaxNumber(self, value):
        self.contactitem.BusinessFaxNumber = value

    # Lower case alias for BusinessFaxNumber setter
    @businessfaxnumber.setter
    def businessfaxnumber(self, value):
        self.BusinessFaxNumber = value

    @property
    def BusinessHomePage(self):
        return self.contactitem.BusinessHomePage

    # Lower case alias for BusinessHomePage
    @property
    def businesshomepage(self):
        return self.BusinessHomePage

    @BusinessHomePage.setter
    def BusinessHomePage(self, value):
        self.contactitem.BusinessHomePage = value

    # Lower case alias for BusinessHomePage setter
    @businesshomepage.setter
    def businesshomepage(self, value):
        self.BusinessHomePage = value

    @property
    def BusinessTelephoneNumber(self):
        return self.contactitem.BusinessTelephoneNumber

    # Lower case alias for BusinessTelephoneNumber
    @property
    def businesstelephonenumber(self):
        return self.BusinessTelephoneNumber

    @BusinessTelephoneNumber.setter
    def BusinessTelephoneNumber(self, value):
        self.contactitem.BusinessTelephoneNumber = value

    # Lower case alias for BusinessTelephoneNumber setter
    @businesstelephonenumber.setter
    def businesstelephonenumber(self, value):
        self.BusinessTelephoneNumber = value

    @property
    def CallbackTelephoneNumber(self):
        return self.contactitem.CallbackTelephoneNumber

    # Lower case alias for CallbackTelephoneNumber
    @property
    def callbacktelephonenumber(self):
        return self.CallbackTelephoneNumber

    @CallbackTelephoneNumber.setter
    def CallbackTelephoneNumber(self, value):
        self.contactitem.CallbackTelephoneNumber = value

    # Lower case alias for CallbackTelephoneNumber setter
    @callbacktelephonenumber.setter
    def callbacktelephonenumber(self, value):
        self.CallbackTelephoneNumber = value

    @property
    def CarTelephoneNumber(self):
        return self.contactitem.CarTelephoneNumber

    # Lower case alias for CarTelephoneNumber
    @property
    def cartelephonenumber(self):
        return self.CarTelephoneNumber

    @CarTelephoneNumber.setter
    def CarTelephoneNumber(self, value):
        self.contactitem.CarTelephoneNumber = value

    # Lower case alias for CarTelephoneNumber setter
    @cartelephonenumber.setter
    def cartelephonenumber(self, value):
        self.CarTelephoneNumber = value

    @property
    def Categories(self):
        return self.contactitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.contactitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Children(self):
        return self.contactitem.Children

    # Lower case alias for Children
    @property
    def children(self):
        return self.Children

    @Children.setter
    def Children(self, value):
        self.contactitem.Children = value

    # Lower case alias for Children setter
    @children.setter
    def children(self, value):
        self.Children = value

    @property
    def Class(self):
        return OlObjectClass(self.contactitem.Class)

    @property
    def Companies(self):
        return self.contactitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.contactitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def CompanyAndFullName(self):
        return self.contactitem.CompanyAndFullName

    # Lower case alias for CompanyAndFullName
    @property
    def companyandfullname(self):
        return self.CompanyAndFullName

    @property
    def CompanyLastFirstNoSpace(self):
        return self.contactitem.CompanyLastFirstNoSpace

    # Lower case alias for CompanyLastFirstNoSpace
    @property
    def companylastfirstnospace(self):
        return self.CompanyLastFirstNoSpace

    @property
    def CompanyLastFirstSpaceOnly(self):
        return self.contactitem.CompanyLastFirstSpaceOnly

    # Lower case alias for CompanyLastFirstSpaceOnly
    @property
    def companylastfirstspaceonly(self):
        return self.CompanyLastFirstSpaceOnly

    @property
    def CompanyMainTelephoneNumber(self):
        return self.contactitem.CompanyMainTelephoneNumber

    # Lower case alias for CompanyMainTelephoneNumber
    @property
    def companymaintelephonenumber(self):
        return self.CompanyMainTelephoneNumber

    @CompanyMainTelephoneNumber.setter
    def CompanyMainTelephoneNumber(self, value):
        self.contactitem.CompanyMainTelephoneNumber = value

    # Lower case alias for CompanyMainTelephoneNumber setter
    @companymaintelephonenumber.setter
    def companymaintelephonenumber(self, value):
        self.CompanyMainTelephoneNumber = value

    @property
    def CompanyName(self):
        return self.contactitem.CompanyName

    # Lower case alias for CompanyName
    @property
    def companyname(self):
        return self.CompanyName

    @CompanyName.setter
    def CompanyName(self, value):
        self.contactitem.CompanyName = value

    # Lower case alias for CompanyName setter
    @companyname.setter
    def companyname(self, value):
        self.CompanyName = value

    @property
    def ComputerNetworkName(self):
        return self.contactitem.ComputerNetworkName

    # Lower case alias for ComputerNetworkName
    @property
    def computernetworkname(self):
        return self.ComputerNetworkName

    @ComputerNetworkName.setter
    def ComputerNetworkName(self, value):
        self.contactitem.ComputerNetworkName = value

    # Lower case alias for ComputerNetworkName setter
    @computernetworkname.setter
    def computernetworkname(self, value):
        self.ComputerNetworkName = value

    @property
    def Conflicts(self):
        return self.contactitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.contactitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.contactitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.contactitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.contactitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def CustomerID(self):
        return self.contactitem.CustomerID

    # Lower case alias for CustomerID
    @property
    def customerid(self):
        return self.CustomerID

    @CustomerID.setter
    def CustomerID(self, value):
        self.contactitem.CustomerID = value

    # Lower case alias for CustomerID setter
    @customerid.setter
    def customerid(self, value):
        self.CustomerID = value

    @property
    def Department(self):
        return self.contactitem.Department

    # Lower case alias for Department
    @property
    def department(self):
        return self.Department

    @Department.setter
    def Department(self, value):
        self.contactitem.Department = value

    # Lower case alias for Department setter
    @department.setter
    def department(self, value):
        self.Department = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.contactitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def Email1Address(self):
        return self.contactitem.Email1Address

    # Lower case alias for Email1Address
    @property
    def email1address(self):
        return self.Email1Address

    @Email1Address.setter
    def Email1Address(self, value):
        self.contactitem.Email1Address = value

    # Lower case alias for Email1Address setter
    @email1address.setter
    def email1address(self, value):
        self.Email1Address = value

    @property
    def Email1AddressType(self):
        return self.contactitem.Email1AddressType

    # Lower case alias for Email1AddressType
    @property
    def email1addresstype(self):
        return self.Email1AddressType

    @Email1AddressType.setter
    def Email1AddressType(self, value):
        self.contactitem.Email1AddressType = value

    # Lower case alias for Email1AddressType setter
    @email1addresstype.setter
    def email1addresstype(self, value):
        self.Email1AddressType = value

    @property
    def Email1DisplayName(self):
        return self.contactitem.Email1DisplayName

    # Lower case alias for Email1DisplayName
    @property
    def email1displayname(self):
        return self.Email1DisplayName

    @Email1DisplayName.setter
    def Email1DisplayName(self, value):
        self.contactitem.Email1DisplayName = value

    # Lower case alias for Email1DisplayName setter
    @email1displayname.setter
    def email1displayname(self, value):
        self.Email1DisplayName = value

    @property
    def Email1EntryID(self):
        return self.contactitem.Email1EntryID

    # Lower case alias for Email1EntryID
    @property
    def email1entryid(self):
        return self.Email1EntryID

    @property
    def Email2Address(self):
        return self.contactitem.Email2Address

    # Lower case alias for Email2Address
    @property
    def email2address(self):
        return self.Email2Address

    @Email2Address.setter
    def Email2Address(self, value):
        self.contactitem.Email2Address = value

    # Lower case alias for Email2Address setter
    @email2address.setter
    def email2address(self, value):
        self.Email2Address = value

    @property
    def Email2AddressType(self):
        return self.contactitem.Email2AddressType

    # Lower case alias for Email2AddressType
    @property
    def email2addresstype(self):
        return self.Email2AddressType

    @Email2AddressType.setter
    def Email2AddressType(self, value):
        self.contactitem.Email2AddressType = value

    # Lower case alias for Email2AddressType setter
    @email2addresstype.setter
    def email2addresstype(self, value):
        self.Email2AddressType = value

    @property
    def Email2DisplayName(self):
        return self.contactitem.Email2DisplayName

    # Lower case alias for Email2DisplayName
    @property
    def email2displayname(self):
        return self.Email2DisplayName

    @Email2DisplayName.setter
    def Email2DisplayName(self, value):
        self.contactitem.Email2DisplayName = value

    # Lower case alias for Email2DisplayName setter
    @email2displayname.setter
    def email2displayname(self, value):
        self.Email2DisplayName = value

    @property
    def Email2EntryID(self):
        return self.contactitem.Email2EntryID

    # Lower case alias for Email2EntryID
    @property
    def email2entryid(self):
        return self.Email2EntryID

    @property
    def Email3Address(self):
        return self.contactitem.Email3Address

    # Lower case alias for Email3Address
    @property
    def email3address(self):
        return self.Email3Address

    @Email3Address.setter
    def Email3Address(self, value):
        self.contactitem.Email3Address = value

    # Lower case alias for Email3Address setter
    @email3address.setter
    def email3address(self, value):
        self.Email3Address = value

    @property
    def Email3AddressType(self):
        return self.contactitem.Email3AddressType

    # Lower case alias for Email3AddressType
    @property
    def email3addresstype(self):
        return self.Email3AddressType

    @Email3AddressType.setter
    def Email3AddressType(self, value):
        self.contactitem.Email3AddressType = value

    # Lower case alias for Email3AddressType setter
    @email3addresstype.setter
    def email3addresstype(self, value):
        self.Email3AddressType = value

    @property
    def Email3DisplayName(self):
        return self.contactitem.Email3DisplayName

    # Lower case alias for Email3DisplayName
    @property
    def email3displayname(self):
        return self.Email3DisplayName

    @Email3DisplayName.setter
    def Email3DisplayName(self, value):
        self.contactitem.Email3DisplayName = value

    # Lower case alias for Email3DisplayName setter
    @email3displayname.setter
    def email3displayname(self, value):
        self.Email3DisplayName = value

    @property
    def Email3EntryID(self):
        return self.contactitem.Email3EntryID

    # Lower case alias for Email3EntryID
    @property
    def email3entryid(self):
        return self.Email3EntryID

    @property
    def EntryID(self):
        return self.contactitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FileAs(self):
        return self.contactitem.FileAs

    # Lower case alias for FileAs
    @property
    def fileas(self):
        return self.FileAs

    @FileAs.setter
    def FileAs(self, value):
        self.contactitem.FileAs = value

    # Lower case alias for FileAs setter
    @fileas.setter
    def fileas(self, value):
        self.FileAs = value

    @property
    def FirstName(self):
        return self.contactitem.FirstName

    # Lower case alias for FirstName
    @property
    def firstname(self):
        return self.FirstName

    @FirstName.setter
    def FirstName(self, value):
        self.contactitem.FirstName = value

    # Lower case alias for FirstName setter
    @firstname.setter
    def firstname(self, value):
        self.FirstName = value

    @property
    def FormDescription(self):
        return FormDescription(self.contactitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def FTPSite(self):
        return self.contactitem.FTPSite

    # Lower case alias for FTPSite
    @property
    def ftpsite(self):
        return self.FTPSite

    @FTPSite.setter
    def FTPSite(self, value):
        self.contactitem.FTPSite = value

    # Lower case alias for FTPSite setter
    @ftpsite.setter
    def ftpsite(self, value):
        self.FTPSite = value

    @property
    def FullName(self):
        return self.contactitem.FullName

    # Lower case alias for FullName
    @property
    def fullname(self):
        return self.FullName

    @FullName.setter
    def FullName(self, value):
        self.contactitem.FullName = value

    # Lower case alias for FullName setter
    @fullname.setter
    def fullname(self, value):
        self.FullName = value

    @property
    def FullNameAndCompany(self):
        return self.contactitem.FullNameAndCompany

    # Lower case alias for FullNameAndCompany
    @property
    def fullnameandcompany(self):
        return self.FullNameAndCompany

    @property
    def Gender(self):
        return OlGender(self.contactitem.Gender)

    # Lower case alias for Gender
    @property
    def gender(self):
        return self.Gender

    @Gender.setter
    def Gender(self, value):
        self.contactitem.Gender = value

    # Lower case alias for Gender setter
    @gender.setter
    def gender(self, value):
        self.Gender = value

    @property
    def GetInspector(self):
        return Inspector(self.contactitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def GovernmentIDNumber(self):
        return self.contactitem.GovernmentIDNumber

    # Lower case alias for GovernmentIDNumber
    @property
    def governmentidnumber(self):
        return self.GovernmentIDNumber

    @GovernmentIDNumber.setter
    def GovernmentIDNumber(self, value):
        self.contactitem.GovernmentIDNumber = value

    # Lower case alias for GovernmentIDNumber setter
    @governmentidnumber.setter
    def governmentidnumber(self, value):
        self.GovernmentIDNumber = value

    @property
    def HasPicture(self):
        return self.contactitem.HasPicture

    # Lower case alias for HasPicture
    @property
    def haspicture(self):
        return self.HasPicture

    @property
    def Hobby(self):
        return self.contactitem.Hobby

    # Lower case alias for Hobby
    @property
    def hobby(self):
        return self.Hobby

    @Hobby.setter
    def Hobby(self, value):
        self.contactitem.Hobby = value

    # Lower case alias for Hobby setter
    @hobby.setter
    def hobby(self, value):
        self.Hobby = value

    @property
    def Home2TelephoneNumber(self):
        return self.contactitem.Home2TelephoneNumber

    # Lower case alias for Home2TelephoneNumber
    @property
    def home2telephonenumber(self):
        return self.Home2TelephoneNumber

    @Home2TelephoneNumber.setter
    def Home2TelephoneNumber(self, value):
        self.contactitem.Home2TelephoneNumber = value

    # Lower case alias for Home2TelephoneNumber setter
    @home2telephonenumber.setter
    def home2telephonenumber(self, value):
        self.Home2TelephoneNumber = value

    @property
    def HomeAddress(self):
        return self.contactitem.HomeAddress

    # Lower case alias for HomeAddress
    @property
    def homeaddress(self):
        return self.HomeAddress

    @HomeAddress.setter
    def HomeAddress(self, value):
        self.contactitem.HomeAddress = value

    # Lower case alias for HomeAddress setter
    @homeaddress.setter
    def homeaddress(self, value):
        self.HomeAddress = value

    @property
    def HomeAddressCity(self):
        return self.contactitem.HomeAddressCity

    # Lower case alias for HomeAddressCity
    @property
    def homeaddresscity(self):
        return self.HomeAddressCity

    @HomeAddressCity.setter
    def HomeAddressCity(self, value):
        self.contactitem.HomeAddressCity = value

    # Lower case alias for HomeAddressCity setter
    @homeaddresscity.setter
    def homeaddresscity(self, value):
        self.HomeAddressCity = value

    @property
    def HomeAddressCountry(self):
        return self.contactitem.HomeAddressCountry

    # Lower case alias for HomeAddressCountry
    @property
    def homeaddresscountry(self):
        return self.HomeAddressCountry

    @HomeAddressCountry.setter
    def HomeAddressCountry(self, value):
        self.contactitem.HomeAddressCountry = value

    # Lower case alias for HomeAddressCountry setter
    @homeaddresscountry.setter
    def homeaddresscountry(self, value):
        self.HomeAddressCountry = value

    @property
    def HomeAddressPostalCode(self):
        return self.contactitem.HomeAddressPostalCode

    # Lower case alias for HomeAddressPostalCode
    @property
    def homeaddresspostalcode(self):
        return self.HomeAddressPostalCode

    @HomeAddressPostalCode.setter
    def HomeAddressPostalCode(self, value):
        self.contactitem.HomeAddressPostalCode = value

    # Lower case alias for HomeAddressPostalCode setter
    @homeaddresspostalcode.setter
    def homeaddresspostalcode(self, value):
        self.HomeAddressPostalCode = value

    @property
    def HomeAddressPostOfficeBox(self):
        return self.contactitem.HomeAddressPostOfficeBox

    # Lower case alias for HomeAddressPostOfficeBox
    @property
    def homeaddresspostofficebox(self):
        return self.HomeAddressPostOfficeBox

    @HomeAddressPostOfficeBox.setter
    def HomeAddressPostOfficeBox(self, value):
        self.contactitem.HomeAddressPostOfficeBox = value

    # Lower case alias for HomeAddressPostOfficeBox setter
    @homeaddresspostofficebox.setter
    def homeaddresspostofficebox(self, value):
        self.HomeAddressPostOfficeBox = value

    @property
    def HomeAddressState(self):
        return self.contactitem.HomeAddressState

    # Lower case alias for HomeAddressState
    @property
    def homeaddressstate(self):
        return self.HomeAddressState

    @HomeAddressState.setter
    def HomeAddressState(self, value):
        self.contactitem.HomeAddressState = value

    # Lower case alias for HomeAddressState setter
    @homeaddressstate.setter
    def homeaddressstate(self, value):
        self.HomeAddressState = value

    @property
    def HomeAddressStreet(self):
        return self.contactitem.HomeAddressStreet

    # Lower case alias for HomeAddressStreet
    @property
    def homeaddressstreet(self):
        return self.HomeAddressStreet

    @HomeAddressStreet.setter
    def HomeAddressStreet(self, value):
        self.contactitem.HomeAddressStreet = value

    # Lower case alias for HomeAddressStreet setter
    @homeaddressstreet.setter
    def homeaddressstreet(self, value):
        self.HomeAddressStreet = value

    @property
    def HomeFaxNumber(self):
        return self.contactitem.HomeFaxNumber

    # Lower case alias for HomeFaxNumber
    @property
    def homefaxnumber(self):
        return self.HomeFaxNumber

    @HomeFaxNumber.setter
    def HomeFaxNumber(self, value):
        self.contactitem.HomeFaxNumber = value

    # Lower case alias for HomeFaxNumber setter
    @homefaxnumber.setter
    def homefaxnumber(self, value):
        self.HomeFaxNumber = value

    @property
    def HomeTelephoneNumber(self):
        return self.contactitem.HomeTelephoneNumber

    # Lower case alias for HomeTelephoneNumber
    @property
    def hometelephonenumber(self):
        return self.HomeTelephoneNumber

    @HomeTelephoneNumber.setter
    def HomeTelephoneNumber(self, value):
        self.contactitem.HomeTelephoneNumber = value

    # Lower case alias for HomeTelephoneNumber setter
    @hometelephonenumber.setter
    def hometelephonenumber(self, value):
        self.HomeTelephoneNumber = value

    @property
    def IMAddress(self):
        return self.contactitem.IMAddress

    # Lower case alias for IMAddress
    @property
    def imaddress(self):
        return self.IMAddress

    @IMAddress.setter
    def IMAddress(self, value):
        self.contactitem.IMAddress = value

    # Lower case alias for IMAddress setter
    @imaddress.setter
    def imaddress(self, value):
        self.IMAddress = value

    @property
    def Importance(self):
        return OlImportance(self.contactitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.contactitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def Initials(self):
        return self.contactitem.Initials

    # Lower case alias for Initials
    @property
    def initials(self):
        return self.Initials

    @Initials.setter
    def Initials(self, value):
        self.contactitem.Initials = value

    # Lower case alias for Initials setter
    @initials.setter
    def initials(self, value):
        self.Initials = value

    @property
    def InternetFreeBusyAddress(self):
        return self.contactitem.InternetFreeBusyAddress

    # Lower case alias for InternetFreeBusyAddress
    @property
    def internetfreebusyaddress(self):
        return self.InternetFreeBusyAddress

    @InternetFreeBusyAddress.setter
    def InternetFreeBusyAddress(self, value):
        self.contactitem.InternetFreeBusyAddress = value

    # Lower case alias for InternetFreeBusyAddress setter
    @internetfreebusyaddress.setter
    def internetfreebusyaddress(self, value):
        self.InternetFreeBusyAddress = value

    @property
    def IsConflict(self):
        return self.contactitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ISDNNumber(self):
        return self.contactitem.ISDNNumber

    # Lower case alias for ISDNNumber
    @property
    def isdnnumber(self):
        return self.ISDNNumber

    @ISDNNumber.setter
    def ISDNNumber(self, value):
        self.contactitem.ISDNNumber = value

    # Lower case alias for ISDNNumber setter
    @isdnnumber.setter
    def isdnnumber(self, value):
        self.ISDNNumber = value

    @property
    def IsMarkedAsTask(self):
        return ContactItem(self.contactitem.IsMarkedAsTask)

    # Lower case alias for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.contactitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def JobTitle(self):
        return self.contactitem.JobTitle

    # Lower case alias for JobTitle
    @property
    def jobtitle(self):
        return self.JobTitle

    @JobTitle.setter
    def JobTitle(self, value):
        self.contactitem.JobTitle = value

    # Lower case alias for JobTitle setter
    @jobtitle.setter
    def jobtitle(self, value):
        self.JobTitle = value

    @property
    def Journal(self):
        return self.contactitem.Journal

    # Lower case alias for Journal
    @property
    def journal(self):
        return self.Journal

    @Journal.setter
    def Journal(self, value):
        self.contactitem.Journal = value

    # Lower case alias for Journal setter
    @journal.setter
    def journal(self, value):
        self.Journal = value

    @property
    def Language(self):
        return self.contactitem.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @Language.setter
    def Language(self, value):
        self.contactitem.Language = value

    # Lower case alias for Language setter
    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LastFirstAndSuffix(self):
        return self.contactitem.LastFirstAndSuffix

    # Lower case alias for LastFirstAndSuffix
    @property
    def lastfirstandsuffix(self):
        return self.LastFirstAndSuffix

    @property
    def LastFirstNoSpace(self):
        return self.contactitem.LastFirstNoSpace

    # Lower case alias for LastFirstNoSpace
    @property
    def lastfirstnospace(self):
        return self.LastFirstNoSpace

    @property
    def LastFirstNoSpaceAndSuffix(self):
        return self.contactitem.LastFirstNoSpaceAndSuffix

    # Lower case alias for LastFirstNoSpaceAndSuffix
    @property
    def lastfirstnospaceandsuffix(self):
        return self.LastFirstNoSpaceAndSuffix

    @property
    def LastFirstNoSpaceCompany(self):
        return self.contactitem.LastFirstNoSpaceCompany

    # Lower case alias for LastFirstNoSpaceCompany
    @property
    def lastfirstnospacecompany(self):
        return self.LastFirstNoSpaceCompany

    @property
    def LastFirstSpaceOnly(self):
        return self.contactitem.LastFirstSpaceOnly

    # Lower case alias for LastFirstSpaceOnly
    @property
    def lastfirstspaceonly(self):
        return self.LastFirstSpaceOnly

    @property
    def LastFirstSpaceOnlyCompany(self):
        return self.contactitem.LastFirstSpaceOnlyCompany

    # Lower case alias for LastFirstSpaceOnlyCompany
    @property
    def lastfirstspaceonlycompany(self):
        return self.LastFirstSpaceOnlyCompany

    @property
    def LastModificationTime(self):
        return self.contactitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def LastName(self):
        return self.contactitem.LastName

    # Lower case alias for LastName
    @property
    def lastname(self):
        return self.LastName

    @LastName.setter
    def LastName(self, value):
        self.contactitem.LastName = value

    # Lower case alias for LastName setter
    @lastname.setter
    def lastname(self, value):
        self.LastName = value

    @property
    def LastNameAndFirstName(self):
        return self.contactitem.LastNameAndFirstName

    # Lower case alias for LastNameAndFirstName
    @property
    def lastnameandfirstname(self):
        return self.LastNameAndFirstName

    @property
    def MailingAddress(self):
        return self.contactitem.MailingAddress

    # Lower case alias for MailingAddress
    @property
    def mailingaddress(self):
        return self.MailingAddress

    @MailingAddress.setter
    def MailingAddress(self, value):
        self.contactitem.MailingAddress = value

    # Lower case alias for MailingAddress setter
    @mailingaddress.setter
    def mailingaddress(self, value):
        self.MailingAddress = value

    @property
    def MailingAddressCity(self):
        return self.contactitem.MailingAddressCity

    # Lower case alias for MailingAddressCity
    @property
    def mailingaddresscity(self):
        return self.MailingAddressCity

    @MailingAddressCity.setter
    def MailingAddressCity(self, value):
        self.contactitem.MailingAddressCity = value

    # Lower case alias for MailingAddressCity setter
    @mailingaddresscity.setter
    def mailingaddresscity(self, value):
        self.MailingAddressCity = value

    @property
    def MailingAddressCountry(self):
        return self.contactitem.MailingAddressCountry

    # Lower case alias for MailingAddressCountry
    @property
    def mailingaddresscountry(self):
        return self.MailingAddressCountry

    @MailingAddressCountry.setter
    def MailingAddressCountry(self, value):
        self.contactitem.MailingAddressCountry = value

    # Lower case alias for MailingAddressCountry setter
    @mailingaddresscountry.setter
    def mailingaddresscountry(self, value):
        self.MailingAddressCountry = value

    @property
    def MailingAddressPostalCode(self):
        return self.contactitem.MailingAddressPostalCode

    # Lower case alias for MailingAddressPostalCode
    @property
    def mailingaddresspostalcode(self):
        return self.MailingAddressPostalCode

    @MailingAddressPostalCode.setter
    def MailingAddressPostalCode(self, value):
        self.contactitem.MailingAddressPostalCode = value

    # Lower case alias for MailingAddressPostalCode setter
    @mailingaddresspostalcode.setter
    def mailingaddresspostalcode(self, value):
        self.MailingAddressPostalCode = value

    @property
    def MailingAddressPostOfficeBox(self):
        return self.contactitem.MailingAddressPostOfficeBox

    # Lower case alias for MailingAddressPostOfficeBox
    @property
    def mailingaddresspostofficebox(self):
        return self.MailingAddressPostOfficeBox

    @MailingAddressPostOfficeBox.setter
    def MailingAddressPostOfficeBox(self, value):
        self.contactitem.MailingAddressPostOfficeBox = value

    # Lower case alias for MailingAddressPostOfficeBox setter
    @mailingaddresspostofficebox.setter
    def mailingaddresspostofficebox(self, value):
        self.MailingAddressPostOfficeBox = value

    @property
    def MailingAddressState(self):
        return self.contactitem.MailingAddressState

    # Lower case alias for MailingAddressState
    @property
    def mailingaddressstate(self):
        return self.MailingAddressState

    @MailingAddressState.setter
    def MailingAddressState(self, value):
        self.contactitem.MailingAddressState = value

    # Lower case alias for MailingAddressState setter
    @mailingaddressstate.setter
    def mailingaddressstate(self, value):
        self.MailingAddressState = value

    @property
    def MailingAddressStreet(self):
        return self.contactitem.MailingAddressStreet

    # Lower case alias for MailingAddressStreet
    @property
    def mailingaddressstreet(self):
        return self.MailingAddressStreet

    @MailingAddressStreet.setter
    def MailingAddressStreet(self, value):
        self.contactitem.MailingAddressStreet = value

    # Lower case alias for MailingAddressStreet setter
    @mailingaddressstreet.setter
    def mailingaddressstreet(self, value):
        self.MailingAddressStreet = value

    @property
    def ManagerName(self):
        return self.contactitem.ManagerName

    # Lower case alias for ManagerName
    @property
    def managername(self):
        return self.ManagerName

    @ManagerName.setter
    def ManagerName(self, value):
        self.contactitem.ManagerName = value

    # Lower case alias for ManagerName setter
    @managername.setter
    def managername(self, value):
        self.ManagerName = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.contactitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.contactitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.contactitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.contactitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def MiddleName(self):
        return self.contactitem.MiddleName

    # Lower case alias for MiddleName
    @property
    def middlename(self):
        return self.MiddleName

    @MiddleName.setter
    def MiddleName(self, value):
        self.contactitem.MiddleName = value

    # Lower case alias for MiddleName setter
    @middlename.setter
    def middlename(self, value):
        self.MiddleName = value

    @property
    def Mileage(self):
        return self.contactitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.contactitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def MobileTelephoneNumber(self):
        return self.contactitem.MobileTelephoneNumber

    # Lower case alias for MobileTelephoneNumber
    @property
    def mobiletelephonenumber(self):
        return self.MobileTelephoneNumber

    @MobileTelephoneNumber.setter
    def MobileTelephoneNumber(self, value):
        self.contactitem.MobileTelephoneNumber = value

    # Lower case alias for MobileTelephoneNumber setter
    @mobiletelephonenumber.setter
    def mobiletelephonenumber(self, value):
        self.MobileTelephoneNumber = value

    @property
    def NetMeetingAlias(self):
        return self.contactitem.NetMeetingAlias

    # Lower case alias for NetMeetingAlias
    @property
    def netmeetingalias(self):
        return self.NetMeetingAlias

    @NetMeetingAlias.setter
    def NetMeetingAlias(self, value):
        self.contactitem.NetMeetingAlias = value

    # Lower case alias for NetMeetingAlias setter
    @netmeetingalias.setter
    def netmeetingalias(self, value):
        self.NetMeetingAlias = value

    @property
    def NetMeetingServer(self):
        return self.contactitem.NetMeetingServer

    # Lower case alias for NetMeetingServer
    @property
    def netmeetingserver(self):
        return self.NetMeetingServer

    @NetMeetingServer.setter
    def NetMeetingServer(self, value):
        self.contactitem.NetMeetingServer = value

    # Lower case alias for NetMeetingServer setter
    @netmeetingserver.setter
    def netmeetingserver(self, value):
        self.NetMeetingServer = value

    @property
    def NickName(self):
        return self.contactitem.NickName

    # Lower case alias for NickName
    @property
    def nickname(self):
        return self.NickName

    @NickName.setter
    def NickName(self, value):
        self.contactitem.NickName = value

    # Lower case alias for NickName setter
    @nickname.setter
    def nickname(self, value):
        self.NickName = value

    @property
    def NoAging(self):
        return self.contactitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.contactitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OfficeLocation(self):
        return self.contactitem.OfficeLocation

    # Lower case alias for OfficeLocation
    @property
    def officelocation(self):
        return self.OfficeLocation

    @OfficeLocation.setter
    def OfficeLocation(self, value):
        self.contactitem.OfficeLocation = value

    # Lower case alias for OfficeLocation setter
    @officelocation.setter
    def officelocation(self, value):
        self.OfficeLocation = value

    @property
    def OrganizationalIDNumber(self):
        return self.contactitem.OrganizationalIDNumber

    # Lower case alias for OrganizationalIDNumber
    @property
    def organizationalidnumber(self):
        return self.OrganizationalIDNumber

    @OrganizationalIDNumber.setter
    def OrganizationalIDNumber(self, value):
        self.contactitem.OrganizationalIDNumber = value

    # Lower case alias for OrganizationalIDNumber setter
    @organizationalidnumber.setter
    def organizationalidnumber(self, value):
        self.OrganizationalIDNumber = value

    @property
    def OtherAddress(self):
        return self.contactitem.OtherAddress

    # Lower case alias for OtherAddress
    @property
    def otheraddress(self):
        return self.OtherAddress

    @OtherAddress.setter
    def OtherAddress(self, value):
        self.contactitem.OtherAddress = value

    # Lower case alias for OtherAddress setter
    @otheraddress.setter
    def otheraddress(self, value):
        self.OtherAddress = value

    @property
    def OtherAddressCity(self):
        return self.contactitem.OtherAddressCity

    # Lower case alias for OtherAddressCity
    @property
    def otheraddresscity(self):
        return self.OtherAddressCity

    @OtherAddressCity.setter
    def OtherAddressCity(self, value):
        self.contactitem.OtherAddressCity = value

    # Lower case alias for OtherAddressCity setter
    @otheraddresscity.setter
    def otheraddresscity(self, value):
        self.OtherAddressCity = value

    @property
    def OtherAddressCountry(self):
        return self.contactitem.OtherAddressCountry

    # Lower case alias for OtherAddressCountry
    @property
    def otheraddresscountry(self):
        return self.OtherAddressCountry

    @OtherAddressCountry.setter
    def OtherAddressCountry(self, value):
        self.contactitem.OtherAddressCountry = value

    # Lower case alias for OtherAddressCountry setter
    @otheraddresscountry.setter
    def otheraddresscountry(self, value):
        self.OtherAddressCountry = value

    @property
    def OtherAddressPostalCode(self):
        return self.contactitem.OtherAddressPostalCode

    # Lower case alias for OtherAddressPostalCode
    @property
    def otheraddresspostalcode(self):
        return self.OtherAddressPostalCode

    @OtherAddressPostalCode.setter
    def OtherAddressPostalCode(self, value):
        self.contactitem.OtherAddressPostalCode = value

    # Lower case alias for OtherAddressPostalCode setter
    @otheraddresspostalcode.setter
    def otheraddresspostalcode(self, value):
        self.OtherAddressPostalCode = value

    @property
    def OtherAddressPostOfficeBox(self):
        return self.contactitem.OtherAddressPostOfficeBox

    # Lower case alias for OtherAddressPostOfficeBox
    @property
    def otheraddresspostofficebox(self):
        return self.OtherAddressPostOfficeBox

    @OtherAddressPostOfficeBox.setter
    def OtherAddressPostOfficeBox(self, value):
        self.contactitem.OtherAddressPostOfficeBox = value

    # Lower case alias for OtherAddressPostOfficeBox setter
    @otheraddresspostofficebox.setter
    def otheraddresspostofficebox(self, value):
        self.OtherAddressPostOfficeBox = value

    @property
    def OtherAddressState(self):
        return self.contactitem.OtherAddressState

    # Lower case alias for OtherAddressState
    @property
    def otheraddressstate(self):
        return self.OtherAddressState

    @OtherAddressState.setter
    def OtherAddressState(self, value):
        self.contactitem.OtherAddressState = value

    # Lower case alias for OtherAddressState setter
    @otheraddressstate.setter
    def otheraddressstate(self, value):
        self.OtherAddressState = value

    @property
    def OtherAddressStreet(self):
        return self.contactitem.OtherAddressStreet

    # Lower case alias for OtherAddressStreet
    @property
    def otheraddressstreet(self):
        return self.OtherAddressStreet

    @OtherAddressStreet.setter
    def OtherAddressStreet(self, value):
        self.contactitem.OtherAddressStreet = value

    # Lower case alias for OtherAddressStreet setter
    @otheraddressstreet.setter
    def otheraddressstreet(self, value):
        self.OtherAddressStreet = value

    @property
    def OtherFaxNumber(self):
        return self.contactitem.OtherFaxNumber

    # Lower case alias for OtherFaxNumber
    @property
    def otherfaxnumber(self):
        return self.OtherFaxNumber

    @OtherFaxNumber.setter
    def OtherFaxNumber(self, value):
        self.contactitem.OtherFaxNumber = value

    # Lower case alias for OtherFaxNumber setter
    @otherfaxnumber.setter
    def otherfaxnumber(self, value):
        self.OtherFaxNumber = value

    @property
    def OtherTelephoneNumber(self):
        return self.contactitem.OtherTelephoneNumber

    # Lower case alias for OtherTelephoneNumber
    @property
    def othertelephonenumber(self):
        return self.OtherTelephoneNumber

    @OtherTelephoneNumber.setter
    def OtherTelephoneNumber(self, value):
        self.contactitem.OtherTelephoneNumber = value

    # Lower case alias for OtherTelephoneNumber setter
    @othertelephonenumber.setter
    def othertelephonenumber(self, value):
        self.OtherTelephoneNumber = value

    @property
    def OutlookInternalVersion(self):
        return self.contactitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.contactitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def PagerNumber(self):
        return self.contactitem.PagerNumber

    # Lower case alias for PagerNumber
    @property
    def pagernumber(self):
        return self.PagerNumber

    @PagerNumber.setter
    def PagerNumber(self, value):
        self.contactitem.PagerNumber = value

    # Lower case alias for PagerNumber setter
    @pagernumber.setter
    def pagernumber(self, value):
        self.PagerNumber = value

    @property
    def Parent(self):
        return self.contactitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PersonalHomePage(self):
        return self.contactitem.PersonalHomePage

    # Lower case alias for PersonalHomePage
    @property
    def personalhomepage(self):
        return self.PersonalHomePage

    @PersonalHomePage.setter
    def PersonalHomePage(self, value):
        self.contactitem.PersonalHomePage = value

    # Lower case alias for PersonalHomePage setter
    @personalhomepage.setter
    def personalhomepage(self, value):
        self.PersonalHomePage = value

    @property
    def PrimaryTelephoneNumber(self):
        return self.contactitem.PrimaryTelephoneNumber

    # Lower case alias for PrimaryTelephoneNumber
    @property
    def primarytelephonenumber(self):
        return self.PrimaryTelephoneNumber

    @PrimaryTelephoneNumber.setter
    def PrimaryTelephoneNumber(self, value):
        self.contactitem.PrimaryTelephoneNumber = value

    # Lower case alias for PrimaryTelephoneNumber setter
    @primarytelephonenumber.setter
    def primarytelephonenumber(self, value):
        self.PrimaryTelephoneNumber = value

    @property
    def Profession(self):
        return self.contactitem.Profession

    # Lower case alias for Profession
    @property
    def profession(self):
        return self.Profession

    @Profession.setter
    def Profession(self, value):
        self.contactitem.Profession = value

    # Lower case alias for Profession setter
    @profession.setter
    def profession(self, value):
        self.Profession = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.contactitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RadioTelephoneNumber(self):
        return self.contactitem.RadioTelephoneNumber

    # Lower case alias for RadioTelephoneNumber
    @property
    def radiotelephonenumber(self):
        return self.RadioTelephoneNumber

    @RadioTelephoneNumber.setter
    def RadioTelephoneNumber(self, value):
        self.contactitem.RadioTelephoneNumber = value

    # Lower case alias for RadioTelephoneNumber setter
    @radiotelephonenumber.setter
    def radiotelephonenumber(self, value):
        self.RadioTelephoneNumber = value

    @property
    def ReferredBy(self):
        return self.contactitem.ReferredBy

    # Lower case alias for ReferredBy
    @property
    def referredby(self):
        return self.ReferredBy

    @ReferredBy.setter
    def ReferredBy(self, value):
        self.contactitem.ReferredBy = value

    # Lower case alias for ReferredBy setter
    @referredby.setter
    def referredby(self, value):
        self.ReferredBy = value

    @property
    def ReminderOverrideDefault(self):
        return self.contactitem.ReminderOverrideDefault

    # Lower case alias for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.contactitem.ReminderOverrideDefault = value

    # Lower case alias for ReminderOverrideDefault setter
    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.contactitem.ReminderPlaySound

    # Lower case alias for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.contactitem.ReminderPlaySound = value

    # Lower case alias for ReminderPlaySound setter
    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.contactitem.ReminderSet

    # Lower case alias for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.contactitem.ReminderSet = value

    # Lower case alias for ReminderSet setter
    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.contactitem.ReminderSoundFile

    # Lower case alias for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.contactitem.ReminderSoundFile = value

    # Lower case alias for ReminderSoundFile setter
    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.contactitem.ReminderTime

    # Lower case alias for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.contactitem.ReminderTime = value

    # Lower case alias for ReminderTime setter
    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.contactitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.contactitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.contactitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SelectedMailingAddress(self):
        return OlMailingAddress(self.contactitem.SelectedMailingAddress)

    # Lower case alias for SelectedMailingAddress
    @property
    def selectedmailingaddress(self):
        return self.SelectedMailingAddress

    @SelectedMailingAddress.setter
    def SelectedMailingAddress(self, value):
        self.contactitem.SelectedMailingAddress = value

    # Lower case alias for SelectedMailingAddress setter
    @selectedmailingaddress.setter
    def selectedmailingaddress(self, value):
        self.SelectedMailingAddress = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.contactitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.contactitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.contactitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.contactitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Spouse(self):
        return self.contactitem.Spouse

    # Lower case alias for Spouse
    @property
    def spouse(self):
        return self.Spouse

    @Spouse.setter
    def Spouse(self, value):
        self.contactitem.Spouse = value

    # Lower case alias for Spouse setter
    @spouse.setter
    def spouse(self, value):
        self.Spouse = value

    @property
    def Subject(self):
        return self.contactitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.contactitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Suffix(self):
        return self.contactitem.Suffix

    # Lower case alias for Suffix
    @property
    def suffix(self):
        return self.Suffix

    @Suffix.setter
    def Suffix(self, value):
        self.contactitem.Suffix = value

    # Lower case alias for Suffix setter
    @suffix.setter
    def suffix(self, value):
        self.Suffix = value

    @property
    def TaskCompletedDate(self):
        return ContactItem(self.contactitem.TaskCompletedDate)

    # Lower case alias for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.contactitem.TaskCompletedDate = value

    # Lower case alias for TaskCompletedDate setter
    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return ContactItem(self.contactitem.TaskDueDate)

    # Lower case alias for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.contactitem.TaskDueDate = value

    # Lower case alias for TaskDueDate setter
    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return ContactItem(self.contactitem.TaskStartDate)

    # Lower case alias for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.contactitem.TaskStartDate = value

    # Lower case alias for TaskStartDate setter
    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return ContactItem(self.contactitem.TaskSubject)

    # Lower case alias for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.contactitem.TaskSubject = value

    # Lower case alias for TaskSubject setter
    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def TelexNumber(self):
        return self.contactitem.TelexNumber

    # Lower case alias for TelexNumber
    @property
    def telexnumber(self):
        return self.TelexNumber

    @TelexNumber.setter
    def TelexNumber(self, value):
        self.contactitem.TelexNumber = value

    # Lower case alias for TelexNumber setter
    @telexnumber.setter
    def telexnumber(self, value):
        self.TelexNumber = value

    @property
    def Title(self):
        return self.contactitem.Title

    # Lower case alias for Title
    @property
    def title(self):
        return self.Title

    @Title.setter
    def Title(self, value):
        self.contactitem.Title = value

    # Lower case alias for Title setter
    @title.setter
    def title(self, value):
        self.Title = value

    @property
    def ToDoTaskOrdinal(self):
        return ContactItem(self.contactitem.ToDoTaskOrdinal)

    # Lower case alias for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.contactitem.ToDoTaskOrdinal = value

    # Lower case alias for ToDoTaskOrdinal setter
    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def TTYTDDTelephoneNumber(self):
        return self.contactitem.TTYTDDTelephoneNumber

    # Lower case alias for TTYTDDTelephoneNumber
    @property
    def ttytddtelephonenumber(self):
        return self.TTYTDDTelephoneNumber

    @TTYTDDTelephoneNumber.setter
    def TTYTDDTelephoneNumber(self, value):
        self.contactitem.TTYTDDTelephoneNumber = value

    # Lower case alias for TTYTDDTelephoneNumber setter
    @ttytddtelephonenumber.setter
    def ttytddtelephonenumber(self, value):
        self.TTYTDDTelephoneNumber = value

    @property
    def UnRead(self):
        return self.contactitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.contactitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def User1(self):
        return self.contactitem.User1

    # Lower case alias for User1
    @property
    def user1(self):
        return self.User1

    @User1.setter
    def User1(self, value):
        self.contactitem.User1 = value

    # Lower case alias for User1 setter
    @user1.setter
    def user1(self, value):
        self.User1 = value

    @property
    def User2(self):
        return self.contactitem.User2

    # Lower case alias for User2
    @property
    def user2(self):
        return self.User2

    @User2.setter
    def User2(self, value):
        self.contactitem.User2 = value

    # Lower case alias for User2 setter
    @user2.setter
    def user2(self, value):
        self.User2 = value

    @property
    def User3(self):
        return self.contactitem.User3

    # Lower case alias for User3
    @property
    def user3(self):
        return self.User3

    @User3.setter
    def User3(self, value):
        self.contactitem.User3 = value

    # Lower case alias for User3 setter
    @user3.setter
    def user3(self, value):
        self.User3 = value

    @property
    def User4(self):
        return self.contactitem.User4

    # Lower case alias for User4
    @property
    def user4(self):
        return self.User4

    @User4.setter
    def User4(self, value):
        self.contactitem.User4 = value

    # Lower case alias for User4 setter
    @user4.setter
    def user4(self, value):
        self.User4 = value

    @property
    def UserProperties(self):
        return UserProperties(self.contactitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    @property
    def WebPage(self):
        return self.contactitem.WebPage

    # Lower case alias for WebPage
    @property
    def webpage(self):
        return self.WebPage

    @WebPage.setter
    def WebPage(self, value):
        self.contactitem.WebPage = value

    # Lower case alias for WebPage setter
    @webpage.setter
    def webpage(self, value):
        self.WebPage = value

    @property
    def YomiCompanyName(self):
        return self.contactitem.YomiCompanyName

    # Lower case alias for YomiCompanyName
    @property
    def yomicompanyname(self):
        return self.YomiCompanyName

    @YomiCompanyName.setter
    def YomiCompanyName(self, value):
        self.contactitem.YomiCompanyName = value

    # Lower case alias for YomiCompanyName setter
    @yomicompanyname.setter
    def yomicompanyname(self, value):
        self.YomiCompanyName = value

    @property
    def YomiFirstName(self):
        return self.contactitem.YomiFirstName

    # Lower case alias for YomiFirstName
    @property
    def yomifirstname(self):
        return self.YomiFirstName

    @YomiFirstName.setter
    def YomiFirstName(self, value):
        self.contactitem.YomiFirstName = value

    # Lower case alias for YomiFirstName setter
    @yomifirstname.setter
    def yomifirstname(self, value):
        self.YomiFirstName = value

    @property
    def YomiLastName(self):
        return self.contactitem.YomiLastName

    # Lower case alias for YomiLastName
    @property
    def yomilastname(self):
        return self.YomiLastName

    @YomiLastName.setter
    def YomiLastName(self, value):
        self.contactitem.YomiLastName = value

    # Lower case alias for YomiLastName setter
    @yomilastname.setter
    def yomilastname(self, value):
        self.YomiLastName = value

    def AddBusinessCardLogoPicture(self, Path=None):
        arguments = com_arguments([Path])
        self.contactitem.AddBusinessCardLogoPicture(*arguments)

    def AddPicture(self, Path=None):
        arguments = com_arguments([Path])
        self.contactitem.AddPicture(*arguments)

    def ClearTaskFlag(self):
        self.contactitem.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.contactitem.Close(*arguments)

    def Copy(self):
        self.contactitem.Copy()

    def Delete(self):
        self.contactitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.contactitem.Display(*arguments)

    def ForwardAsBusinessCard(self):
        return self.contactitem.ForwardAsBusinessCard()

    def ForwardAsVcard(self):
        return self.contactitem.ForwardAsVcard()

    def GetConversation(self):
        return self.contactitem.GetConversation()

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.contactitem.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.contactitem.Move(*arguments)

    def PrintOut(self):
        self.contactitem.PrintOut()

    def RemovePicture(self):
        self.contactitem.RemovePicture()

    def ResetBusinessCard(self):
        self.contactitem.ResetBusinessCard()

    def Save(self):
        self.contactitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.contactitem.SaveAs(*arguments)

    def SaveBusinessCardImage(self, Path=None):
        arguments = com_arguments([Path])
        self.contactitem.SaveBusinessCardImage(*arguments)

    def ShowBusinessCardEditor(self):
        self.contactitem.ShowBusinessCardEditor()

    def ShowCategoriesDialog(self):
        self.contactitem.ShowCategoriesDialog()

    def ShowCheckPhoneDialog(self, PhoneNumber=None):
        arguments = com_arguments([PhoneNumber])
        self.contactitem.ShowCheckPhoneDialog(*arguments)


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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.contactsmodule.NavigationGroups)

    # Lower case alias for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.contactsmodule.NavigationModuleType)

    # Lower case alias for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.contactsmodule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return ContactsModule(self.contactsmodule.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.contactsmodule.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.contactsmodule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return ContactsModule(self.contactsmodule.Visible)

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.contactsmodule.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


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

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def Parent(self):
        return Conversation(self.conversation.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.conversation.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def ClearAlwaysAssignCategories(self, Store=None):
        arguments = com_arguments([Store])
        self.conversation.ClearAlwaysAssignCategories(*arguments)

    def GetAlwaysAssignCategories(self, Store=None):
        arguments = com_arguments([Store])
        return self.conversation.GetAlwaysAssignCategories(*arguments)

    def GetAlwaysDelete(self, Store=None):
        arguments = com_arguments([Store])
        return self.conversation.GetAlwaysDelete(*arguments)

    def GetAlwaysMoveToFolder(self, Store=None):
        arguments = com_arguments([Store])
        return self.conversation.GetAlwaysMoveToFolder(*arguments)

    def GetChildren(self, Item=None):
        arguments = com_arguments([Item])
        return self.conversation.GetChildren(*arguments)

    def GetParent(self, Item=None):
        arguments = com_arguments([Item])
        return self.conversation.GetParent(*arguments)

    def GetRootItems(self):
        return self.conversation.GetRootItems()

    def GetTable(self):
        return self.conversation.GetTable()

    def MarkAsRead(self):
        self.conversation.MarkAsRead()

    def MarkAsUnread(self):
        self.conversation.MarkAsUnread()

    def SetAlwaysAssignCategories(self, Categories=None, Store=None):
        arguments = com_arguments([Categories, Store])
        self.conversation.SetAlwaysAssignCategories(*arguments)

    def SetAlwaysDelete(self, AlwaysDelete=None, Store=None):
        arguments = com_arguments([AlwaysDelete, Store])
        self.conversation.SetAlwaysDelete(*arguments)

    def SetAlwaysMoveToFolder(self, MoveToFolder=None, Store=None):
        arguments = com_arguments([MoveToFolder, Store])
        self.conversation.SetAlwaysMoveToFolder(*arguments)

    def StopAlwaysDelete(self, Store=None):
        arguments = com_arguments([Store])
        self.conversation.StopAlwaysDelete(*arguments)

    def StopAlwaysMoveToFolder(self, Store=None):
        arguments = com_arguments([Store])
        self.conversation.StopAlwaysMoveToFolder(*arguments)


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

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationTopic(self):
        return self.conversationheader.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def Parent(self):
        return self.conversationheader.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.conversationheader.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

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

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.distlistitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.distlistitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.distlistitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.distlistitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.distlistitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.distlistitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.distlistitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.distlistitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.distlistitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.distlistitem.Class)

    @property
    def Companies(self):
        return self.distlistitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.distlistitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.distlistitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.distlistitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.distlistitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.distlistitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.distlistitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DLName(self):
        return self.distlistitem.DLName

    # Lower case alias for DLName
    @property
    def dlname(self):
        return self.DLName

    @DLName.setter
    def DLName(self, value):
        self.distlistitem.DLName = value

    # Lower case alias for DLName setter
    @dlname.setter
    def dlname(self, value):
        self.DLName = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.distlistitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.distlistitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.distlistitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.distlistitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.distlistitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.distlistitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.distlistitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return DistListItem(self.distlistitem.IsMarkedAsTask)

    # Lower case alias for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.distlistitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.distlistitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.distlistitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.distlistitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MemberCount(self):
        return self.distlistitem.MemberCount

    # Lower case alias for MemberCount
    @property
    def membercount(self):
        return self.MemberCount

    @property
    def MessageClass(self):
        return self.distlistitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.distlistitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.distlistitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.distlistitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.distlistitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.distlistitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.distlistitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.distlistitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.distlistitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.distlistitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReminderOverrideDefault(self):
        return self.distlistitem.ReminderOverrideDefault

    # Lower case alias for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.distlistitem.ReminderOverrideDefault = value

    # Lower case alias for ReminderOverrideDefault setter
    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.distlistitem.ReminderPlaySound

    # Lower case alias for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.distlistitem.ReminderPlaySound = value

    # Lower case alias for ReminderPlaySound setter
    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.distlistitem.ReminderSet

    # Lower case alias for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.distlistitem.ReminderSet = value

    # Lower case alias for ReminderSet setter
    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.distlistitem.ReminderSoundFile

    # Lower case alias for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.distlistitem.ReminderSoundFile = value

    # Lower case alias for ReminderSoundFile setter
    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.distlistitem.ReminderTime

    # Lower case alias for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.distlistitem.ReminderTime = value

    # Lower case alias for ReminderTime setter
    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.distlistitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.distlistitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.distlistitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.distlistitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.distlistitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.distlistitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.distlistitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.distlistitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.distlistitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def TaskCompletedDate(self):
        return DistListItem(self.distlistitem.TaskCompletedDate)

    # Lower case alias for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.distlistitem.TaskCompletedDate = value

    # Lower case alias for TaskCompletedDate setter
    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return DistListItem(self.distlistitem.TaskDueDate)

    # Lower case alias for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.distlistitem.TaskDueDate = value

    # Lower case alias for TaskDueDate setter
    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return DistListItem(self.distlistitem.TaskStartDate)

    # Lower case alias for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.distlistitem.TaskStartDate = value

    # Lower case alias for TaskStartDate setter
    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return DistListItem(self.distlistitem.TaskSubject)

    # Lower case alias for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.distlistitem.TaskSubject = value

    # Lower case alias for TaskSubject setter
    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def ToDoTaskOrdinal(self):
        return DistListItem(self.distlistitem.ToDoTaskOrdinal)

    # Lower case alias for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.distlistitem.ToDoTaskOrdinal = value

    # Lower case alias for ToDoTaskOrdinal setter
    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.distlistitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.distlistitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.distlistitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def AddMember(self, Recipient=None):
        arguments = com_arguments([Recipient])
        self.distlistitem.AddMember(*arguments)

    def AddMembers(self, Recipients=None):
        arguments = com_arguments([Recipients])
        self.distlistitem.AddMembers(*arguments)

    def ClearTaskFlag(self):
        self.distlistitem.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.distlistitem.Close(*arguments)

    def Copy(self):
        self.distlistitem.Copy()

    def Delete(self):
        self.distlistitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.distlistitem.Display(*arguments)

    def GetConversation(self):
        return self.distlistitem.GetConversation()

    def GetMember(self, Index=None):
        arguments = com_arguments([Index])
        return self.distlistitem.GetMember(*arguments)

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.distlistitem.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.distlistitem.Move(*arguments)

    def PrintOut(self):
        self.distlistitem.PrintOut()

    def RemoveMember(self, Recipient=None):
        arguments = com_arguments([Recipient])
        self.distlistitem.RemoveMember(*arguments)

    def RemoveMembers(self, Recipients=None):
        arguments = com_arguments([Recipients])
        self.distlistitem.RemoveMembers(*arguments)

    def Save(self):
        self.distlistitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.distlistitem.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.distlistitem.ShowCategoriesDialog()


class DocumentItem:

    def __init__(self, documentitem=None):
        self.documentitem = documentitem

    @property
    def Actions(self):
        return Actions(self.documentitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.documentitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.documentitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.documentitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.documentitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.documentitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.documentitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @property
    def Categories(self):
        return self.documentitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.documentitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.documentitem.Class)

    @property
    def Companies(self):
        return self.documentitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.documentitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.documentitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationIndex(self):
        return self.documentitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.documentitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.documentitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.documentitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.documentitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.documentitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return self.documentitem.GetInspector

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.documentitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.documentitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.documentitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.documentitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.documentitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return self.documentitem.MarkForDownload

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @property
    def MessageClass(self):
        return self.documentitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.documentitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.documentitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.documentitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.documentitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.documentitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.documentitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.documentitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.documentitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.documentitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Saved(self):
        return self.documentitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.documentitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.documentitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.documentitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.documentitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.documentitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.documentitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.documentitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.documentitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.documentitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.documentitem.Close(*arguments)

    def Copy(self):
        self.documentitem.Copy()

    def Delete(self):
        self.documentitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.documentitem.Display(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.documentitem.Move(*arguments)

    def PrintOut(self):
        self.documentitem.PrintOut()

    def Save(self):
        self.documentitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.documentitem.SaveAs(*arguments)

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

    # Lower case alias for AppointmentItem
    @property
    def appointmentitem(self):
        return self.AppointmentItem

    @property
    def Class(self):
        return OlObjectClass(self.exception.Class)

    @property
    def Deleted(self):
        return AppointmentItem(self.exception.Deleted)

    # Lower case alias for Deleted
    @property
    def deleted(self):
        return self.Deleted

    @property
    def OriginalDate(self):
        return AppointmentItem(self.exception.OriginalDate)

    # Lower case alias for OriginalDate
    @property
    def originaldate(self):
        return self.OriginalDate

    @property
    def Parent(self):
        return self.exception.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.exception.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.exceptions.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.exceptions.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.exceptions.Item(*arguments)


class ExchangeDistributionList:

    def __init__(self, exchangedistributionlist=None):
        self.exchangedistributionlist = exchangedistributionlist

    @property
    def Address(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Address)

    # Lower case alias for Address
    @property
    def address(self):
        return self.Address

    @Address.setter
    def Address(self, value):
        self.exchangedistributionlist.Address = value

    # Lower case alias for Address setter
    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.exchangedistributionlist.AddressEntryUserType)

    # Lower case alias for AddressEntryUserType
    @property
    def addressentryusertype(self):
        return self.AddressEntryUserType

    @property
    def Alias(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Alias)

    # Lower case alias for Alias
    @property
    def alias(self):
        return self.Alias

    @property
    def Application(self):
        return Application(self.exchangedistributionlist.Application)

    @property
    def Class(self):
        return OlObjectClass(self.exchangedistributionlist.Class)

    @property
    def Comments(self):
        return self.exchangedistributionlist.Comments

    # Lower case alias for Comments
    @property
    def comments(self):
        return self.Comments

    @Comments.setter
    def Comments(self, value):
        self.exchangedistributionlist.Comments = value

    # Lower case alias for Comments setter
    @comments.setter
    def comments(self, value):
        self.Comments = value

    @property
    def DisplayType(self):
        return OlDisplayType(self.exchangedistributionlist.DisplayType)

    # Lower case alias for DisplayType
    @property
    def displaytype(self):
        return self.DisplayType

    @property
    def ID(self):
        return ExchangeDistributionList(self.exchangedistributionlist.ID)

    # Lower case alias for ID
    @property
    def id(self):
        return self.ID

    @property
    def Name(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Name)

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.exchangedistributionlist.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrimarySmtpAddress(self):
        return ExchangeDistributionList(self.exchangedistributionlist.PrimarySmtpAddress)

    # Lower case alias for PrimarySmtpAddress
    @property
    def primarysmtpaddress(self):
        return self.PrimarySmtpAddress

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.exchangedistributionlist.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.exchangedistributionlist.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return ExchangeDistributionList(self.exchangedistributionlist.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.exchangedistributionlist.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    def Delete(self):
        self.exchangedistributionlist.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([HWnd])
        self.exchangedistributionlist.Details(*arguments)

    def GetContact(self):
        return self.exchangedistributionlist.GetContact()

    def GetExchangeDistributionList(self):
        return self.exchangedistributionlist.GetExchangeDistributionList()

    def GetExchangeDistributionListMembers(self):
        return AddressEntry(self.exchangedistributionlist.GetExchangeDistributionListMembers())

    def GetExchangeUser(self):
        return self.exchangedistributionlist.GetExchangeUser()

    def GetFreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([Start, MinPerChar, CompleteFormat])
        self.exchangedistributionlist.GetFreeBusy(*arguments)

    def GetMemberOfList(self):
        return self.exchangedistributionlist.GetMemberOfList()

    def GetOwners(self):
        return AddressEntry(self.exchangedistributionlist.GetOwners())

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([MakePermanent, Refresh])
        self.exchangedistributionlist.Update(*arguments)


class ExchangeUser:

    def __init__(self, exchangeuser=None):
        self.exchangeuser = exchangeuser

    @property
    def Address(self):
        return ExchangeUser(self.exchangeuser.Address)

    # Lower case alias for Address
    @property
    def address(self):
        return self.Address

    @Address.setter
    def Address(self, value):
        self.exchangeuser.Address = value

    # Lower case alias for Address setter
    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.exchangeuser.AddressEntryUserType)

    # Lower case alias for AddressEntryUserType
    @property
    def addressentryusertype(self):
        return self.AddressEntryUserType

    @property
    def Alias(self):
        return ExchangeUser(self.exchangeuser.Alias)

    # Lower case alias for Alias
    @property
    def alias(self):
        return self.Alias

    @property
    def Application(self):
        return Application(self.exchangeuser.Application)

    @property
    def AssistantName(self):
        return ExchangeUser(self.exchangeuser.AssistantName)

    # Lower case alias for AssistantName
    @property
    def assistantname(self):
        return self.AssistantName

    @AssistantName.setter
    def AssistantName(self, value):
        self.exchangeuser.AssistantName = value

    # Lower case alias for AssistantName setter
    @assistantname.setter
    def assistantname(self, value):
        self.AssistantName = value

    @property
    def BusinessTelephoneNumber(self):
        return ExchangeUser(self.exchangeuser.BusinessTelephoneNumber)

    # Lower case alias for BusinessTelephoneNumber
    @property
    def businesstelephonenumber(self):
        return self.BusinessTelephoneNumber

    @BusinessTelephoneNumber.setter
    def BusinessTelephoneNumber(self, value):
        self.exchangeuser.BusinessTelephoneNumber = value

    # Lower case alias for BusinessTelephoneNumber setter
    @businesstelephonenumber.setter
    def businesstelephonenumber(self, value):
        self.BusinessTelephoneNumber = value

    @property
    def City(self):
        return ExchangeUser(self.exchangeuser.City)

    # Lower case alias for City
    @property
    def city(self):
        return self.City

    @City.setter
    def City(self, value):
        self.exchangeuser.City = value

    # Lower case alias for City setter
    @city.setter
    def city(self, value):
        self.City = value

    @property
    def Class(self):
        return OlObjectClass(self.exchangeuser.Class)

    @property
    def Comments(self):
        return self.exchangeuser.Comments

    # Lower case alias for Comments
    @property
    def comments(self):
        return self.Comments

    @Comments.setter
    def Comments(self, value):
        self.exchangeuser.Comments = value

    # Lower case alias for Comments setter
    @comments.setter
    def comments(self, value):
        self.Comments = value

    @property
    def CompanyName(self):
        return ExchangeUser(self.exchangeuser.CompanyName)

    # Lower case alias for CompanyName
    @property
    def companyname(self):
        return self.CompanyName

    @CompanyName.setter
    def CompanyName(self, value):
        self.exchangeuser.CompanyName = value

    # Lower case alias for CompanyName setter
    @companyname.setter
    def companyname(self, value):
        self.CompanyName = value

    @property
    def Department(self):
        return ExchangeUser(self.exchangeuser.Department)

    # Lower case alias for Department
    @property
    def department(self):
        return self.Department

    @Department.setter
    def Department(self, value):
        self.exchangeuser.Department = value

    # Lower case alias for Department setter
    @department.setter
    def department(self, value):
        self.Department = value

    @property
    def DisplayType(self):
        return OlDisplayType(self.exchangeuser.DisplayType)

    # Lower case alias for DisplayType
    @property
    def displaytype(self):
        return self.DisplayType

    @property
    def FirstName(self):
        return ExchangeUser(self.exchangeuser.FirstName)

    # Lower case alias for FirstName
    @property
    def firstname(self):
        return self.FirstName

    @FirstName.setter
    def FirstName(self, value):
        self.exchangeuser.FirstName = value

    # Lower case alias for FirstName setter
    @firstname.setter
    def firstname(self, value):
        self.FirstName = value

    @property
    def ID(self):
        return ExchangeUser(self.exchangeuser.ID)

    # Lower case alias for ID
    @property
    def id(self):
        return self.ID

    @property
    def JobTitle(self):
        return ExchangeUser(self.exchangeuser.JobTitle)

    # Lower case alias for JobTitle
    @property
    def jobtitle(self):
        return self.JobTitle

    @JobTitle.setter
    def JobTitle(self, value):
        self.exchangeuser.JobTitle = value

    # Lower case alias for JobTitle setter
    @jobtitle.setter
    def jobtitle(self, value):
        self.JobTitle = value

    @property
    def LastName(self):
        return ExchangeUser(self.exchangeuser.LastName)

    # Lower case alias for LastName
    @property
    def lastname(self):
        return self.LastName

    @LastName.setter
    def LastName(self, value):
        self.exchangeuser.LastName = value

    # Lower case alias for LastName setter
    @lastname.setter
    def lastname(self, value):
        self.LastName = value

    @property
    def MobileTelephoneNumber(self):
        return ExchangeUser(self.exchangeuser.MobileTelephoneNumber)

    # Lower case alias for MobileTelephoneNumber
    @property
    def mobiletelephonenumber(self):
        return self.MobileTelephoneNumber

    @MobileTelephoneNumber.setter
    def MobileTelephoneNumber(self, value):
        self.exchangeuser.MobileTelephoneNumber = value

    # Lower case alias for MobileTelephoneNumber setter
    @mobiletelephonenumber.setter
    def mobiletelephonenumber(self, value):
        self.MobileTelephoneNumber = value

    @property
    def Name(self):
        return ExchangeUser(self.exchangeuser.Name)

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.exchangeuser.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def OfficeLocation(self):
        return ExchangeUser(self.exchangeuser.OfficeLocation)

    # Lower case alias for OfficeLocation
    @property
    def officelocation(self):
        return self.OfficeLocation

    @OfficeLocation.setter
    def OfficeLocation(self, value):
        self.exchangeuser.OfficeLocation = value

    # Lower case alias for OfficeLocation setter
    @officelocation.setter
    def officelocation(self, value):
        self.OfficeLocation = value

    @property
    def Parent(self):
        return ExchangeUser(self.exchangeuser.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PostalCode(self):
        return ExchangeUser(self.exchangeuser.PostalCode)

    # Lower case alias for PostalCode
    @property
    def postalcode(self):
        return self.PostalCode

    @PostalCode.setter
    def PostalCode(self, value):
        self.exchangeuser.PostalCode = value

    # Lower case alias for PostalCode setter
    @postalcode.setter
    def postalcode(self, value):
        self.PostalCode = value

    @property
    def PrimarySmtpAddress(self):
        return ExchangeUser(self.exchangeuser.PrimarySmtpAddress)

    # Lower case alias for PrimarySmtpAddress
    @property
    def primarysmtpaddress(self):
        return self.PrimarySmtpAddress

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.exchangeuser.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.exchangeuser.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def StateOrProvince(self):
        return ExchangeUser(self.exchangeuser.StateOrProvince)

    # Lower case alias for StateOrProvince
    @property
    def stateorprovince(self):
        return self.StateOrProvince

    @StateOrProvince.setter
    def StateOrProvince(self, value):
        self.exchangeuser.StateOrProvince = value

    # Lower case alias for StateOrProvince setter
    @stateorprovince.setter
    def stateorprovince(self, value):
        self.StateOrProvince = value

    @property
    def StreetAddress(self):
        return ExchangeUser(self.exchangeuser.StreetAddress)

    # Lower case alias for StreetAddress
    @property
    def streetaddress(self):
        return self.StreetAddress

    @StreetAddress.setter
    def StreetAddress(self, value):
        self.exchangeuser.StreetAddress = value

    # Lower case alias for StreetAddress setter
    @streetaddress.setter
    def streetaddress(self, value):
        self.StreetAddress = value

    @property
    def Type(self):
        return ExchangeUser(self.exchangeuser.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.exchangeuser.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def YomiCompanyName(self):
        return ExchangeUser(self.exchangeuser.YomiCompanyName)

    # Lower case alias for YomiCompanyName
    @property
    def yomicompanyname(self):
        return self.YomiCompanyName

    @YomiCompanyName.setter
    def YomiCompanyName(self, value):
        self.exchangeuser.YomiCompanyName = value

    # Lower case alias for YomiCompanyName setter
    @yomicompanyname.setter
    def yomicompanyname(self, value):
        self.YomiCompanyName = value

    @property
    def YomiDepartment(self):
        return ExchangeUser(self.exchangeuser.YomiDepartment)

    # Lower case alias for YomiDepartment
    @property
    def yomidepartment(self):
        return self.YomiDepartment

    @YomiDepartment.setter
    def YomiDepartment(self, value):
        self.exchangeuser.YomiDepartment = value

    # Lower case alias for YomiDepartment setter
    @yomidepartment.setter
    def yomidepartment(self, value):
        self.YomiDepartment = value

    @property
    def YomiDisplayName(self):
        return ExchangeUser(self.exchangeuser.YomiDisplayName)

    # Lower case alias for YomiDisplayName
    @property
    def yomidisplayname(self):
        return self.YomiDisplayName

    @YomiDisplayName.setter
    def YomiDisplayName(self, value):
        self.exchangeuser.YomiDisplayName = value

    # Lower case alias for YomiDisplayName setter
    @yomidisplayname.setter
    def yomidisplayname(self, value):
        self.YomiDisplayName = value

    @property
    def YomiFirstName(self):
        return ExchangeUser(self.exchangeuser.YomiFirstName)

    # Lower case alias for YomiFirstName
    @property
    def yomifirstname(self):
        return self.YomiFirstName

    @YomiFirstName.setter
    def YomiFirstName(self, value):
        self.exchangeuser.YomiFirstName = value

    # Lower case alias for YomiFirstName setter
    @yomifirstname.setter
    def yomifirstname(self, value):
        self.YomiFirstName = value

    @property
    def YomiLastName(self):
        return ExchangeUser(self.exchangeuser.YomiLastName)

    # Lower case alias for YomiLastName
    @property
    def yomilastname(self):
        return self.YomiLastName

    @YomiLastName.setter
    def YomiLastName(self, value):
        self.exchangeuser.YomiLastName = value

    # Lower case alias for YomiLastName setter
    @yomilastname.setter
    def yomilastname(self, value):
        self.YomiLastName = value

    def Delete(self):
        self.exchangeuser.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([HWnd])
        self.exchangeuser.Details(*arguments)

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

    def GetFreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([Start, MinPerChar, CompleteFormat])
        return self.exchangeuser.GetFreeBusy(*arguments)

    def GetMemberOfList(self):
        return ExchangeUser(self.exchangeuser.GetMemberOfList())

    def GetPicture(self):
        return self.exchangeuser.GetPicture()

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([MakePermanent, Refresh])
        self.exchangeuser.Update(*arguments)


class Explorer:

    def __init__(self, explorer=None):
        self.explorer = explorer

    @property
    def AccountSelector(self):
        return AccountSelector(self.explorer.AccountSelector)

    # Lower case alias for AccountSelector
    @property
    def accountselector(self):
        return self.AccountSelector

    @property
    def Application(self):
        return Application(self.explorer.Application)

    @property
    def AttachmentSelection(self):
        return AttachmentSelection(self.explorer.AttachmentSelection)

    # Lower case alias for AttachmentSelection
    @property
    def attachmentselection(self):
        return self.AttachmentSelection

    @property
    def Caption(self):
        return self.explorer.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Class(self):
        return OlObjectClass(self.explorer.Class)

    @property
    def CurrentFolder(self):
        return Folder(self.explorer.CurrentFolder)

    # Lower case alias for CurrentFolder
    @property
    def currentfolder(self):
        return self.CurrentFolder

    @CurrentFolder.setter
    def CurrentFolder(self, value):
        self.explorer.CurrentFolder = value

    # Lower case alias for CurrentFolder setter
    @currentfolder.setter
    def currentfolder(self, value):
        self.CurrentFolder = value

    @property
    def CurrentView(self):
        return self.explorer.CurrentView

    # Lower case alias for CurrentView
    @property
    def currentview(self):
        return self.CurrentView

    @CurrentView.setter
    def CurrentView(self, value):
        self.explorer.CurrentView = value

    # Lower case alias for CurrentView setter
    @currentview.setter
    def currentview(self, value):
        self.CurrentView = value

    @property
    def Height(self):
        return self.explorer.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.explorer.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def HTMLDocument(self):
        return self.explorer.HTMLDocument

    # Lower case alias for HTMLDocument
    @property
    def htmldocument(self):
        return self.HTMLDocument

    @property
    def Left(self):
        return self.explorer.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.explorer.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def NavigationPane(self):
        return NavigationPane(self.explorer.NavigationPane)

    # Lower case alias for NavigationPane
    @property
    def navigationpane(self):
        return self.NavigationPane

    @property
    def Panes(self):
        return Panes(self.explorer.Panes)

    # Lower case alias for Panes
    @property
    def panes(self):
        return self.Panes

    @property
    def Parent(self):
        return self.explorer.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Selection(self):
        return Selection(self.explorer.Selection)

    # Lower case alias for Selection
    @property
    def selection(self):
        return self.Selection

    @property
    def Session(self):
        return NameSpace(self.explorer.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Top(self):
        return self.explorer.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.explorer.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.explorer.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.explorer.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def WindowState(self):
        return OlWindowState(self.explorer.WindowState)

    # Lower case alias for WindowState
    @property
    def windowstate(self):
        return self.WindowState

    @WindowState.setter
    def WindowState(self, value):
        self.explorer.WindowState = value

    # Lower case alias for WindowState setter
    @windowstate.setter
    def windowstate(self, value):
        self.WindowState = value

    def Activate(self):
        self.explorer.Activate()

    def AddToSelection(self, Item=None):
        arguments = com_arguments([Item])
        self.explorer.AddToSelection(*arguments)

    def ClearSearch(self):
        self.explorer.ClearSearch()

    def ClearSelection(self):
        self.explorer.ClearSelection()

    def Close(self):
        self.explorer.Close()

    def Display(self):
        self.explorer.Display()

    def IsItemSelectableInView(self, Item=None):
        arguments = com_arguments([Item])
        return self.explorer.IsItemSelectableInView(*arguments)

    def IsPaneVisible(self, Pane=None):
        arguments = com_arguments([Pane])
        return self.explorer.IsPaneVisible(*arguments)

    def RemoveFromSelection(self, Item=None):
        arguments = com_arguments([Item])
        self.explorer.RemoveFromSelection(*arguments)

    def Search(self, Query=None, SearchScope=None):
        arguments = com_arguments([Query, SearchScope])
        self.explorer.Search(*arguments)

    def SelectAllItems(self):
        self.explorer.SelectAllItems()

    def ShowPane(self, Pane=None, Visible=None):
        arguments = com_arguments([Pane, Visible])
        self.explorer.ShowPane(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.explorers.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.explorers.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Folder=None, DisplayMode=None):
        arguments = com_arguments([Folder, DisplayMode])
        return Explorer(self.explorers.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.explorers.Item(*arguments)


class Folder:

    def __init__(self, folder=None):
        self.folder = folder

    @property
    def AddressBookName(self):
        return Folder(self.folder.AddressBookName)

    # Lower case alias for AddressBookName
    @property
    def addressbookname(self):
        return self.AddressBookName

    @AddressBookName.setter
    def AddressBookName(self, value):
        self.folder.AddressBookName = value

    # Lower case alias for AddressBookName setter
    @addressbookname.setter
    def addressbookname(self, value):
        self.AddressBookName = value

    @property
    def Application(self):
        return Application(self.folder.Application)

    @property
    def Class(self):
        return OlObjectClass(self.folder.Class)

    @property
    def CurrentView(self):
        return View(self.folder.CurrentView)

    # Lower case alias for CurrentView
    @property
    def currentview(self):
        return self.CurrentView

    @property
    def CustomViewsOnly(self):
        return self.folder.CustomViewsOnly

    # Lower case alias for CustomViewsOnly
    @property
    def customviewsonly(self):
        return self.CustomViewsOnly

    @CustomViewsOnly.setter
    def CustomViewsOnly(self, value):
        self.folder.CustomViewsOnly = value

    # Lower case alias for CustomViewsOnly setter
    @customviewsonly.setter
    def customviewsonly(self, value):
        self.CustomViewsOnly = value

    @property
    def DefaultItemType(self):
        return OlItemType(self.folder.DefaultItemType)

    # Lower case alias for DefaultItemType
    @property
    def defaultitemtype(self):
        return self.DefaultItemType

    @property
    def DefaultMessageClass(self):
        return self.folder.DefaultMessageClass

    # Lower case alias for DefaultMessageClass
    @property
    def defaultmessageclass(self):
        return self.DefaultMessageClass

    @property
    def Description(self):
        return self.folder.Description

    # Lower case alias for Description
    @property
    def description(self):
        return self.Description

    @Description.setter
    def Description(self, value):
        self.folder.Description = value

    # Lower case alias for Description setter
    @description.setter
    def description(self, value):
        self.Description = value

    @property
    def EntryID(self):
        return self.folder.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FolderPath(self):
        return self.folder.FolderPath

    # Lower case alias for FolderPath
    @property
    def folderpath(self):
        return self.FolderPath

    @property
    def Folders(self):
        return Folders(self.folder.Folders)

    # Lower case alias for Folders
    @property
    def folders(self):
        return self.Folders

    @property
    def InAppFolderSyncObject(self):
        return self.folder.InAppFolderSyncObject

    # Lower case alias for InAppFolderSyncObject
    @property
    def inappfoldersyncobject(self):
        return self.InAppFolderSyncObject

    @InAppFolderSyncObject.setter
    def InAppFolderSyncObject(self, value):
        self.folder.InAppFolderSyncObject = value

    # Lower case alias for InAppFolderSyncObject setter
    @inappfoldersyncobject.setter
    def inappfoldersyncobject(self, value):
        self.InAppFolderSyncObject = value

    @property
    def IsSharePointFolder(self):
        return self.folder.IsSharePointFolder

    # Lower case alias for IsSharePointFolder
    @property
    def issharepointfolder(self):
        return self.IsSharePointFolder

    @property
    def Items(self):
        return Items(self.folder.Items)

    # Lower case alias for Items
    @property
    def items(self):
        return self.Items

    @property
    def Name(self):
        return self.folder.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.folder.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.folder.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.folder.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.folder.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowAsOutlookAB(self):
        return self.folder.ShowAsOutlookAB

    # Lower case alias for ShowAsOutlookAB
    @property
    def showasoutlookab(self):
        return self.ShowAsOutlookAB

    @ShowAsOutlookAB.setter
    def ShowAsOutlookAB(self, value):
        self.folder.ShowAsOutlookAB = value

    # Lower case alias for ShowAsOutlookAB setter
    @showasoutlookab.setter
    def showasoutlookab(self, value):
        self.ShowAsOutlookAB = value

    @property
    def ShowItemCount(self):
        return self.folder.ShowItemCount

    # Lower case alias for ShowItemCount
    @property
    def showitemcount(self):
        return self.ShowItemCount

    @ShowItemCount.setter
    def ShowItemCount(self, value):
        self.folder.ShowItemCount = value

    # Lower case alias for ShowItemCount setter
    @showitemcount.setter
    def showitemcount(self, value):
        self.ShowItemCount = value

    @property
    def Store(self):
        return Store(self.folder.Store)

    # Lower case alias for Store
    @property
    def store(self):
        return self.Store

    @property
    def StoreID(self):
        return self.folder.StoreID

    # Lower case alias for StoreID
    @property
    def storeid(self):
        return self.StoreID

    @property
    def UnReadItemCount(self):
        return self.folder.UnReadItemCount

    # Lower case alias for UnReadItemCount
    @property
    def unreaditemcount(self):
        return self.UnReadItemCount

    @property
    def UserDefinedProperties(self):
        return UserDefinedProperties(self.folder.UserDefinedProperties)

    # Lower case alias for UserDefinedProperties
    @property
    def userdefinedproperties(self):
        return self.UserDefinedProperties

    @property
    def Views(self):
        return Views(self.folder.Views)

    # Lower case alias for Views
    @property
    def views(self):
        return self.Views

    @property
    def WebViewOn(self):
        return self.folder.WebViewOn

    # Lower case alias for WebViewOn
    @property
    def webviewon(self):
        return self.WebViewOn

    @WebViewOn.setter
    def WebViewOn(self, value):
        self.folder.WebViewOn = value

    # Lower case alias for WebViewOn setter
    @webviewon.setter
    def webviewon(self, value):
        self.WebViewOn = value

    @property
    def WebViewURL(self):
        return self.folder.WebViewURL

    # Lower case alias for WebViewURL
    @property
    def webviewurl(self):
        return self.WebViewURL

    @WebViewURL.setter
    def WebViewURL(self, value):
        self.folder.WebViewURL = value

    # Lower case alias for WebViewURL setter
    @webviewurl.setter
    def webviewurl(self, value):
        self.WebViewURL = value

    def AddToPFFavorites(self):
        self.folder.AddToPFFavorites()

    def CopyTo(self, DestinationFolder=None):
        arguments = com_arguments([DestinationFolder])
        return self.folder.CopyTo(*arguments)

    def Delete(self):
        self.folder.Delete()

    def Display(self):
        self.folder.Display()

    def GetCalendarExporter(self):
        return self.folder.GetCalendarExporter()

    def GetCustomIcon(self):
        return self.folder.GetCustomIcon()

    def GetExplorer(self, DisplayMode=None):
        arguments = com_arguments([DisplayMode])
        return self.folder.GetExplorer(*arguments)

    def GetOrganizer(self):
        return self.folder.GetOrganizer()

    def GetStorage(self, StorageIdentifier=None, StorageIdentifierType=None):
        arguments = com_arguments([StorageIdentifier, StorageIdentifierType])
        return self.folder.GetStorage(*arguments)

    def GetTable(self, Filter=None, TableContents=None):
        arguments = com_arguments([Filter, TableContents])
        return Folder(self.folder.GetTable(*arguments))

    def MoveTo(self, DestinationFolder=None):
        arguments = com_arguments([DestinationFolder])
        self.folder.MoveTo(*arguments)

    def SetCustomIcon(self, Picture=None):
        arguments = com_arguments([Picture])
        self.folder.SetCustomIcon(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.folders.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.folders.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Type=None):
        arguments = com_arguments([Name, Type])
        return Folder(self.folders.Add(*arguments))

    def GetFirst(self):
        return Folder(self.folders.GetFirst())

    def GetLast(self):
        return Folder(self.folders.GetLast())

    def GetNext(self):
        return Folder(self.folders.GetNext())

    def GetPrevious(self):
        return Folder(self.folders.GetPrevious())

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.folders.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.folders.Remove(*arguments)


class FormDescription:

    def __init__(self, formdescription=None):
        self.formdescription = formdescription

    @property
    def Application(self):
        return Application(self.formdescription.Application)

    @property
    def Category(self):
        return self.formdescription.Category

    # Lower case alias for Category
    @property
    def category(self):
        return self.Category

    @Category.setter
    def Category(self, value):
        self.formdescription.Category = value

    # Lower case alias for Category setter
    @category.setter
    def category(self, value):
        self.Category = value

    @property
    def CategorySub(self):
        return self.formdescription.CategorySub

    # Lower case alias for CategorySub
    @property
    def categorysub(self):
        return self.CategorySub

    @CategorySub.setter
    def CategorySub(self, value):
        self.formdescription.CategorySub = value

    # Lower case alias for CategorySub setter
    @categorysub.setter
    def categorysub(self, value):
        self.CategorySub = value

    @property
    def Class(self):
        return OlObjectClass(self.formdescription.Class)

    @property
    def Comment(self):
        return self.formdescription.Comment

    # Lower case alias for Comment
    @property
    def comment(self):
        return self.Comment

    @Comment.setter
    def Comment(self, value):
        self.formdescription.Comment = value

    # Lower case alias for Comment setter
    @comment.setter
    def comment(self, value):
        self.Comment = value

    @property
    def ContactName(self):
        return FormDescription(self.formdescription.ContactName)

    # Lower case alias for ContactName
    @property
    def contactname(self):
        return self.ContactName

    @ContactName.setter
    def ContactName(self, value):
        self.formdescription.ContactName = value

    # Lower case alias for ContactName setter
    @contactname.setter
    def contactname(self, value):
        self.ContactName = value

    @property
    def DisplayName(self):
        return self.formdescription.DisplayName

    # Lower case alias for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @DisplayName.setter
    def DisplayName(self, value):
        self.formdescription.DisplayName = value

    # Lower case alias for DisplayName setter
    @displayname.setter
    def displayname(self, value):
        self.DisplayName = value

    @property
    def Hidden(self):
        return self.formdescription.Hidden

    # Lower case alias for Hidden
    @property
    def hidden(self):
        return self.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.formdescription.Hidden = value

    # Lower case alias for Hidden setter
    @hidden.setter
    def hidden(self, value):
        self.Hidden = value

    @property
    def Icon(self):
        return self.formdescription.Icon

    # Lower case alias for Icon
    @property
    def icon(self):
        return self.Icon

    @Icon.setter
    def Icon(self, value):
        self.formdescription.Icon = value

    # Lower case alias for Icon setter
    @icon.setter
    def icon(self, value):
        self.Icon = value

    @property
    def Locked(self):
        return self.formdescription.Locked

    # Lower case alias for Locked
    @property
    def locked(self):
        return self.Locked

    @Locked.setter
    def Locked(self, value):
        self.formdescription.Locked = value

    # Lower case alias for Locked setter
    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MessageClass(self):
        return FormDescription(self.formdescription.MessageClass)

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @property
    def MiniIcon(self):
        return self.formdescription.MiniIcon

    # Lower case alias for MiniIcon
    @property
    def miniicon(self):
        return self.MiniIcon

    @MiniIcon.setter
    def MiniIcon(self, value):
        self.formdescription.MiniIcon = value

    # Lower case alias for MiniIcon setter
    @miniicon.setter
    def miniicon(self, value):
        self.MiniIcon = value

    @property
    def Name(self):
        return self.formdescription.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.formdescription.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Number(self):
        return self.formdescription.Number

    # Lower case alias for Number
    @property
    def number(self):
        return self.Number

    @Number.setter
    def Number(self, value):
        self.formdescription.Number = value

    # Lower case alias for Number setter
    @number.setter
    def number(self, value):
        self.Number = value

    @property
    def OneOff(self):
        return self.formdescription.OneOff

    # Lower case alias for OneOff
    @property
    def oneoff(self):
        return self.OneOff

    @OneOff.setter
    def OneOff(self, value):
        self.formdescription.OneOff = value

    # Lower case alias for OneOff setter
    @oneoff.setter
    def oneoff(self, value):
        self.OneOff = value

    @property
    def Parent(self):
        return self.formdescription.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ScriptText(self):
        return self.formdescription.ScriptText

    # Lower case alias for ScriptText
    @property
    def scripttext(self):
        return self.ScriptText

    @property
    def Session(self):
        return NameSpace(self.formdescription.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Template(self):
        return self.formdescription.Template

    # Lower case alias for Template
    @property
    def template(self):
        return self.Template

    @Template.setter
    def Template(self, value):
        self.formdescription.Template = value

    # Lower case alias for Template setter
    @template.setter
    def template(self, value):
        self.Template = value

    @property
    def UseWordMail(self):
        return self.formdescription.UseWordMail

    # Lower case alias for UseWordMail
    @property
    def usewordmail(self):
        return self.UseWordMail

    @UseWordMail.setter
    def UseWordMail(self, value):
        self.formdescription.UseWordMail = value

    # Lower case alias for UseWordMail setter
    @usewordmail.setter
    def usewordmail(self, value):
        self.UseWordMail = value

    @property
    def Version(self):
        return self.formdescription.Version

    # Lower case alias for Version
    @property
    def version(self):
        return self.Version

    @Version.setter
    def Version(self, value):
        self.formdescription.Version = value

    # Lower case alias for Version setter
    @version.setter
    def version(self, value):
        self.Version = value

    def PublishForm(self, Registry=None, Folder=None):
        arguments = com_arguments([Registry, Folder])
        self.formdescription.PublishForm(*arguments)


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

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.formnamerulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.formnamerulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def FormName(self):
        return self.formnamerulecondition.FormName

    # Lower case alias for FormName
    @property
    def formname(self):
        return self.FormName

    @FormName.setter
    def FormName(self, value):
        self.formnamerulecondition.FormName = value

    # Lower case alias for FormName setter
    @formname.setter
    def formname(self, value):
        self.FormName = value

    @property
    def Parent(self):
        return self.formnamerulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.formnamerulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Detail
    @property
    def detail(self):
        return self.Detail

    @Detail.setter
    def Detail(self, value):
        self.formregion.Detail = value

    # Lower case alias for Detail setter
    @detail.setter
    def detail(self, value):
        self.Detail = value

    @property
    def DisplayName(self):
        return self.formregion.DisplayName

    # Lower case alias for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def EnableAutoLayout(self):
        return self.formregion.EnableAutoLayout

    # Lower case alias for EnableAutoLayout
    @property
    def enableautolayout(self):
        return self.EnableAutoLayout

    @EnableAutoLayout.setter
    def EnableAutoLayout(self, value):
        self.formregion.EnableAutoLayout = value

    # Lower case alias for EnableAutoLayout setter
    @enableautolayout.setter
    def enableautolayout(self, value):
        self.EnableAutoLayout = value

    @property
    def Form(self):
        return self.formregion.Form

    # Lower case alias for Form
    @property
    def form(self):
        return self.Form

    @property
    def FormRegionMode(self):
        return OlFormRegionMode(self.formregion.FormRegionMode)

    # Lower case alias for FormRegionMode
    @property
    def formregionmode(self):
        return self.FormRegionMode

    @property
    def Inspector(self):
        return Inspector(self.formregion.Inspector)

    # Lower case alias for Inspector
    @property
    def inspector(self):
        return self.Inspector

    @property
    def InternalName(self):
        return self.formregion.InternalName

    # Lower case alias for InternalName
    @property
    def internalname(self):
        return self.InternalName

    @property
    def IsExpanded(self):
        return self.formregion.IsExpanded

    # Lower case alias for IsExpanded
    @property
    def isexpanded(self):
        return self.IsExpanded

    @property
    def Item(self):
        return self.formregion.Item

    # Lower case alias for Item
    @property
    def item(self):
        return self.Item

    @property
    def Language(self):
        return self.formregion.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @property
    def Parent(self):
        return self.formregion.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.formregion.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def SuppressControlReplacement(self):
        return self.formregion.SuppressControlReplacement

    # Lower case alias for SuppressControlReplacement
    @property
    def suppresscontrolreplacement(self):
        return self.SuppressControlReplacement

    @SuppressControlReplacement.setter
    def SuppressControlReplacement(self, value):
        self.formregion.SuppressControlReplacement = value

    # Lower case alias for SuppressControlReplacement setter
    @suppresscontrolreplacement.setter
    def suppresscontrolreplacement(self, value):
        self.SuppressControlReplacement = value

    @property
    def Visible(self):
        return self.formregion.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.formregion.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

    def Reflow(self):
        self.formregion.Reflow()

    def Select(self):
        self.formregion.Select()

    def SetControlItemProperty(self, Control=None, PropertyName=None):
        arguments = com_arguments([Control, PropertyName])
        self.formregion.SetControlItemProperty(*arguments)


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

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.fromrssfeedrulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.fromrssfeedrulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def FromRssFeed(self):
        return self.fromrssfeedrulecondition.FromRssFeed

    # Lower case alias for FromRssFeed
    @property
    def fromrssfeed(self):
        return self.FromRssFeed

    @FromRssFeed.setter
    def FromRssFeed(self, value):
        self.fromrssfeedrulecondition.FromRssFeed = value

    # Lower case alias for FromRssFeed setter
    @fromrssfeed.setter
    def fromrssfeed(self, value):
        self.FromRssFeed = value

    @property
    def Parent(self):
        return self.fromrssfeedrulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.fromrssfeedrulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @Filter.setter
    def Filter(self, value):
        self.iconview.Filter = value

    # Lower case alias for Filter setter
    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def IconPlacement(self):
        return OlIconViewPlacement(self.iconview.IconPlacement)

    # Lower case alias for IconPlacement
    @property
    def iconplacement(self):
        return self.IconPlacement

    @IconPlacement.setter
    def IconPlacement(self, value):
        self.iconview.IconPlacement = value

    # Lower case alias for IconPlacement setter
    @iconplacement.setter
    def iconplacement(self, value):
        self.IconPlacement = value

    @property
    def IconViewType(self):
        return OlIconViewType(self.iconview.IconViewType)

    # Lower case alias for IconViewType
    @property
    def iconviewtype(self):
        return self.IconViewType

    @IconViewType.setter
    def IconViewType(self, value):
        self.iconview.IconViewType = value

    # Lower case alias for IconViewType setter
    @iconviewtype.setter
    def iconviewtype(self, value):
        self.IconViewType = value

    @property
    def Language(self):
        return self.iconview.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @Language.setter
    def Language(self, value):
        self.iconview.Language = value

    # Lower case alias for Language setter
    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.iconview.LockUserChanges

    # Lower case alias for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.iconview.LockUserChanges = value

    # Lower case alias for LockUserChanges setter
    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.iconview.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.iconview.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.iconview.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.iconview.SaveOption)

    # Lower case alias for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.iconview.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def SortFields(self):
        return OrderFields(self.iconview.SortFields)

    # Lower case alias for SortFields
    @property
    def sortfields(self):
        return self.SortFields

    @property
    def Standard(self):
        return IconView(self.iconview.Standard)

    # Lower case alias for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.iconview.ViewType)

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.iconview.XML

    # Lower case alias for XML
    @property
    def xml(self):
        return self.XML

    @XML.setter
    def XML(self, value):
        self.iconview.XML = value

    # Lower case alias for XML setter
    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.iconview.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return self.iconview.Copy(*arguments)

    def Delete(self):
        self.iconview.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.iconview.GoToDate(*arguments)

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

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.importancerulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.importancerulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Importance(self):
        return OlImportance(self.importancerulecondition.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.importancerulecondition.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def Parent(self):
        return self.importancerulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.importancerulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class Inspector:

    def __init__(self, inspector=None):
        self.inspector = inspector

    @property
    def Application(self):
        return Application(self.inspector.Application)

    @property
    def AttachmentSelection(self):
        return AttachmentSelection(self.inspector.AttachmentSelection)

    # Lower case alias for AttachmentSelection
    @property
    def attachmentselection(self):
        return self.AttachmentSelection

    @property
    def Caption(self):
        return self.inspector.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Class(self):
        return OlObjectClass(self.inspector.Class)

    @property
    def CurrentItem(self):
        return self.inspector.CurrentItem

    # Lower case alias for CurrentItem
    @property
    def currentitem(self):
        return self.CurrentItem

    @property
    def EditorType(self):
        return OlEditorType(self.inspector.EditorType)

    # Lower case alias for EditorType
    @property
    def editortype(self):
        return self.EditorType

    @property
    def Height(self):
        return self.inspector.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.inspector.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def Left(self):
        return self.inspector.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.inspector.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def ModifiedFormPages(self):
        return Pages(self.inspector.ModifiedFormPages)

    # Lower case alias for ModifiedFormPages
    @property
    def modifiedformpages(self):
        return self.ModifiedFormPages

    @property
    def Parent(self):
        return self.inspector.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.inspector.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Top(self):
        return self.inspector.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.inspector.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.inspector.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.inspector.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def WindowState(self):
        return OlWindowState(self.inspector.WindowState)

    # Lower case alias for WindowState
    @property
    def windowstate(self):
        return self.WindowState

    @WindowState.setter
    def WindowState(self, value):
        self.inspector.WindowState = value

    # Lower case alias for WindowState setter
    @windowstate.setter
    def windowstate(self, value):
        self.WindowState = value

    @property
    def WordEditor(self):
        return self.inspector.WordEditor

    # Lower case alias for WordEditor
    @property
    def wordeditor(self):
        return self.WordEditor

    def Activate(self):
        self.inspector.Activate()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.inspector.Close(*arguments)

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.inspector.Display(*arguments)

    def HideFormPage(self, PageName=None):
        arguments = com_arguments([PageName])
        self.inspector.HideFormPage(*arguments)

    def IsWordMail(self):
        return self.inspector.IsWordMail()

    def NewFormRegion(self):
        return self.inspector.NewFormRegion()

    def OpenFormRegion(self, Path=None):
        arguments = com_arguments([Path])
        return self.inspector.OpenFormRegion(*arguments)

    def SaveFormRegion(self, Page=None, FileName=None):
        arguments = com_arguments([Page, FileName])
        self.inspector.SaveFormRegion(*arguments)

    def SetControlItemProperty(self, Control=None, PropertyName=None):
        arguments = com_arguments([Control, PropertyName])
        self.inspector.SetControlItemProperty(*arguments)

    def SetCurrentFormPage(self, PageName=None):
        arguments = com_arguments([PageName])
        self.inspector.SetCurrentFormPage(*arguments)

    def SetSchedulingStartTime(self, Start=None):
        arguments = com_arguments([Start])
        self.inspector.SetSchedulingStartTime(*arguments)

    def ShowFormPage(self, PageName=None):
        arguments = com_arguments([PageName])
        self.inspector.ShowFormPage(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.inspectors.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.inspectors.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self):
        return Inspector(self.inspectors.Add())

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.inspectors.Item(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.itemproperties.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.itemproperties.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = com_arguments([Name, Type, AddToFolderFields, DisplayFormat])
        self.itemproperties.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.itemproperties.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.itemproperties.Remove(*arguments)


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

    # Lower case alias for IsUserProperty
    @property
    def isuserproperty(self):
        return self.IsUserProperty

    @property
    def Name(self):
        return self.itemproperty.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.itemproperty.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.itemproperty.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.itemproperty.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.itemproperty.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @property
    def Value(self):
        return self.itemproperty.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.itemproperty.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def IncludeRecurrences(self):
        return Items(self.items.IncludeRecurrences)

    # Lower case alias for IncludeRecurrences
    @property
    def includerecurrences(self):
        return self.IncludeRecurrences

    @IncludeRecurrences.setter
    def IncludeRecurrences(self, value):
        self.items.IncludeRecurrences = value

    # Lower case alias for IncludeRecurrences setter
    @includerecurrences.setter
    def includerecurrences(self, value):
        self.IncludeRecurrences = value

    @property
    def Parent(self):
        return self.items.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.items.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self):
        return self.items.Add()

    def Find(self, Filter=None):
        arguments = com_arguments([Filter])
        return self.items.Find(*arguments)

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.items.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.items.Remove(*arguments)

    def ResetColumns(self):
        self.items.ResetColumns()

    def Restrict(self, Filter=None):
        arguments = com_arguments([Filter])
        return self.items.Restrict(*arguments)

    def SetColumns(self, Columns=None):
        arguments = com_arguments([Columns])
        self.items.SetColumns(*arguments)

    def Sort(self, Property=None, Descending=None):
        arguments = com_arguments([Property, Descending])
        self.items.Sort(*arguments)


class JournalItem:

    def __init__(self, journalitem=None):
        self.journalitem = journalitem

    @property
    def Actions(self):
        return Actions(self.journalitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.journalitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.journalitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.journalitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.journalitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.journalitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.journalitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.journalitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.journalitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.journalitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.journalitem.Class)

    @property
    def Companies(self):
        return self.journalitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.journalitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return Conflicts(self.journalitem.Conflicts)

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ContactNames(self):
        return self.journalitem.ContactNames

    # Lower case alias for ContactNames
    @property
    def contactnames(self):
        return self.ContactNames

    @ContactNames.setter
    def ContactNames(self, value):
        self.journalitem.ContactNames = value

    # Lower case alias for ContactNames setter
    @contactnames.setter
    def contactnames(self, value):
        self.ContactNames = value

    @property
    def ConversationID(self):
        return Conversation(self.journalitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.journalitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.journalitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.journalitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DocPosted(self):
        return self.journalitem.DocPosted

    # Lower case alias for DocPosted
    @property
    def docposted(self):
        return self.DocPosted

    @DocPosted.setter
    def DocPosted(self, value):
        self.journalitem.DocPosted = value

    # Lower case alias for DocPosted setter
    @docposted.setter
    def docposted(self, value):
        self.DocPosted = value

    @property
    def DocPrinted(self):
        return self.journalitem.DocPrinted

    # Lower case alias for DocPrinted
    @property
    def docprinted(self):
        return self.DocPrinted

    @DocPrinted.setter
    def DocPrinted(self, value):
        self.journalitem.DocPrinted = value

    # Lower case alias for DocPrinted setter
    @docprinted.setter
    def docprinted(self, value):
        self.DocPrinted = value

    @property
    def DocRouted(self):
        return self.journalitem.DocRouted

    # Lower case alias for DocRouted
    @property
    def docrouted(self):
        return self.DocRouted

    @DocRouted.setter
    def DocRouted(self, value):
        self.journalitem.DocRouted = value

    # Lower case alias for DocRouted setter
    @docrouted.setter
    def docrouted(self, value):
        self.DocRouted = value

    @property
    def DocSaved(self):
        return self.journalitem.DocSaved

    # Lower case alias for DocSaved
    @property
    def docsaved(self):
        return self.DocSaved

    @DocSaved.setter
    def DocSaved(self, value):
        self.journalitem.DocSaved = value

    # Lower case alias for DocSaved setter
    @docsaved.setter
    def docsaved(self, value):
        self.DocSaved = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.journalitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def Duration(self):
        return JournalItem(self.journalitem.Duration)

    # Lower case alias for Duration
    @property
    def duration(self):
        return self.Duration

    @Duration.setter
    def Duration(self, value):
        self.journalitem.Duration = value

    # Lower case alias for Duration setter
    @duration.setter
    def duration(self, value):
        self.Duration = value

    @property
    def End(self):
        return self.journalitem.End

    # Lower case alias for End
    @property
    def end(self):
        return self.End

    @End.setter
    def End(self, value):
        self.journalitem.End = value

    # Lower case alias for End setter
    @end.setter
    def end(self, value):
        self.End = value

    @property
    def EntryID(self):
        return self.journalitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.journalitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.journalitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.journalitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.journalitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.journalitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.journalitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.journalitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.journalitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.journalitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.journalitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.journalitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.journalitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.journalitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.journalitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.journalitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.journalitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.journalitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.journalitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.journalitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.journalitem.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def Saved(self):
        return self.journalitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.journalitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.journalitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.journalitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.journalitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Start(self):
        return self.journalitem.Start

    # Lower case alias for Start
    @property
    def start(self):
        return self.Start

    @Start.setter
    def Start(self, value):
        self.journalitem.Start = value

    # Lower case alias for Start setter
    @start.setter
    def start(self, value):
        self.Start = value

    @property
    def Subject(self):
        return self.journalitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.journalitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Type(self):
        return self.journalitem.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.journalitem.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def UnRead(self):
        return self.journalitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.journalitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.journalitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.journalitem.Close(*arguments)

    def Copy(self):
        self.journalitem.Copy()

    def Delete(self):
        self.journalitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.journalitem.Display(*arguments)

    def Forward(self):
        return self.journalitem.Forward()

    def GetConversation(self):
        return self.journalitem.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.journalitem.Move(*arguments)

    def PrintOut(self):
        self.journalitem.PrintOut()

    def Reply(self):
        return MailItem(self.journalitem.Reply())

    def ReplyAll(self):
        return MailItem(self.journalitem.ReplyAll())

    def Save(self):
        self.journalitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.journalitem.SaveAs(*arguments)

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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.journalmodule.NavigationGroups)

    # Lower case alias for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.journalmodule.NavigationModuleType)

    # Lower case alias for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.journalmodule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return JournalModule(self.journalmodule.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.journalmodule.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.journalmodule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return JournalModule(self.journalmodule.Visible)

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.journalmodule.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


class MailItem:

    def __init__(self, mailitem=None):
        self.mailitem = mailitem

    @property
    def Actions(self):
        return Actions(self.mailitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def AlternateRecipientAllowed(self):
        return self.mailitem.AlternateRecipientAllowed

    # Lower case alias for AlternateRecipientAllowed
    @property
    def alternaterecipientallowed(self):
        return self.AlternateRecipientAllowed

    @AlternateRecipientAllowed.setter
    def AlternateRecipientAllowed(self, value):
        self.mailitem.AlternateRecipientAllowed = value

    # Lower case alias for AlternateRecipientAllowed setter
    @alternaterecipientallowed.setter
    def alternaterecipientallowed(self, value):
        self.AlternateRecipientAllowed = value

    @property
    def Application(self):
        return Application(self.mailitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.mailitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.mailitem.AutoForwarded

    # Lower case alias for AutoForwarded
    @property
    def autoforwarded(self):
        return self.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.mailitem.AutoForwarded = value

    # Lower case alias for AutoForwarded setter
    @autoforwarded.setter
    def autoforwarded(self, value):
        self.AutoForwarded = value

    @property
    def AutoResolvedWinner(self):
        return self.mailitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BCC(self):
        return MailItem(self.mailitem.BCC)

    # Lower case alias for BCC
    @property
    def bcc(self):
        return self.BCC

    @BCC.setter
    def BCC(self, value):
        self.mailitem.BCC = value

    # Lower case alias for BCC setter
    @bcc.setter
    def bcc(self, value):
        self.BCC = value

    @property
    def BillingInformation(self):
        return self.mailitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.mailitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.mailitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.mailitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.mailitem.BodyFormat)

    # Lower case alias for BodyFormat
    @property
    def bodyformat(self):
        return self.BodyFormat

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.mailitem.BodyFormat = value

    # Lower case alias for BodyFormat setter
    @bodyformat.setter
    def bodyformat(self, value):
        self.BodyFormat = value

    @property
    def Categories(self):
        return self.mailitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.mailitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def CC(self):
        return MailItem(self.mailitem.CC)

    # Lower case alias for CC
    @property
    def cc(self):
        return self.CC

    @CC.setter
    def CC(self, value):
        self.mailitem.CC = value

    # Lower case alias for CC setter
    @cc.setter
    def cc(self, value):
        self.CC = value

    @property
    def Class(self):
        return OlObjectClass(self.mailitem.Class)

    @property
    def Companies(self):
        return self.mailitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.mailitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.mailitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.mailitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.mailitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.mailitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.mailitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return self.mailitem.DeferredDeliveryTime

    # Lower case alias for DeferredDeliveryTime
    @property
    def deferreddeliverytime(self):
        return self.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.mailitem.DeferredDeliveryTime = value

    # Lower case alias for DeferredDeliveryTime setter
    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.mailitem.DeleteAfterSubmit

    # Lower case alias for DeleteAfterSubmit
    @property
    def deleteaftersubmit(self):
        return self.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.mailitem.DeleteAfterSubmit = value

    # Lower case alias for DeleteAfterSubmit setter
    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.mailitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.mailitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.mailitem.ExpiryTime

    # Lower case alias for ExpiryTime
    @property
    def expirytime(self):
        return self.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.mailitem.ExpiryTime = value

    # Lower case alias for ExpiryTime setter
    @expirytime.setter
    def expirytime(self, value):
        self.ExpiryTime = value

    @property
    def FlagRequest(self):
        return self.mailitem.FlagRequest

    # Lower case alias for FlagRequest
    @property
    def flagrequest(self):
        return self.FlagRequest

    @FlagRequest.setter
    def FlagRequest(self, value):
        self.mailitem.FlagRequest = value

    # Lower case alias for FlagRequest setter
    @flagrequest.setter
    def flagrequest(self, value):
        self.FlagRequest = value

    @property
    def FormDescription(self):
        return FormDescription(self.mailitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.mailitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def HTMLBody(self):
        return self.mailitem.HTMLBody

    # Lower case alias for HTMLBody
    @property
    def htmlbody(self):
        return self.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.mailitem.HTMLBody = value

    # Lower case alias for HTMLBody setter
    @htmlbody.setter
    def htmlbody(self, value):
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.mailitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.mailitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.mailitem.InternetCodepage

    # Lower case alias for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.mailitem.InternetCodepage = value

    # Lower case alias for InternetCodepage setter
    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.mailitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return MailItem(self.mailitem.IsMarkedAsTask)

    # Lower case alias for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.mailitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.mailitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.mailitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.mailitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.mailitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.mailitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.mailitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.mailitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.mailitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.mailitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return self.mailitem.OriginatorDeliveryReportRequested

    # Lower case alias for OriginatorDeliveryReportRequested
    @property
    def originatordeliveryreportrequested(self):
        return self.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.mailitem.OriginatorDeliveryReportRequested = value

    # Lower case alias for OriginatorDeliveryReportRequested setter
    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return self.mailitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.mailitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.mailitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Permission(self):
        return self.mailitem.Permission

    # Lower case alias for Permission
    @property
    def permission(self):
        return self.Permission

    @Permission.setter
    def Permission(self, value):
        self.mailitem.Permission = value

    # Lower case alias for Permission setter
    @permission.setter
    def permission(self, value):
        self.Permission = value

    @property
    def PermissionService(self):
        return self.mailitem.PermissionService

    # Lower case alias for PermissionService
    @property
    def permissionservice(self):
        return self.PermissionService

    @PermissionService.setter
    def PermissionService(self, value):
        self.mailitem.PermissionService = value

    # Lower case alias for PermissionService setter
    @permissionservice.setter
    def permissionservice(self, value):
        self.PermissionService = value

    @property
    def PermissionTemplateGuid(self):
        return MailItem(self.mailitem.PermissionTemplateGuid)

    # Lower case alias for PermissionTemplateGuid
    @property
    def permissiontemplateguid(self):
        return self.PermissionTemplateGuid

    @PermissionTemplateGuid.setter
    def PermissionTemplateGuid(self, value):
        self.mailitem.PermissionTemplateGuid = value

    # Lower case alias for PermissionTemplateGuid setter
    @permissiontemplateguid.setter
    def permissiontemplateguid(self, value):
        self.PermissionTemplateGuid = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.mailitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReadReceiptRequested(self):
        return self.mailitem.ReadReceiptRequested

    # Lower case alias for ReadReceiptRequested
    @property
    def readreceiptrequested(self):
        return self.ReadReceiptRequested

    @property
    def ReceivedByEntryID(self):
        return self.mailitem.ReceivedByEntryID

    # Lower case alias for ReceivedByEntryID
    @property
    def receivedbyentryid(self):
        return self.ReceivedByEntryID

    @property
    def ReceivedByName(self):
        return self.mailitem.ReceivedByName

    # Lower case alias for ReceivedByName
    @property
    def receivedbyname(self):
        return self.ReceivedByName

    @property
    def ReceivedOnBehalfOfEntryID(self):
        return self.mailitem.ReceivedOnBehalfOfEntryID

    # Lower case alias for ReceivedOnBehalfOfEntryID
    @property
    def receivedonbehalfofentryid(self):
        return self.ReceivedOnBehalfOfEntryID

    @property
    def ReceivedOnBehalfOfName(self):
        return self.mailitem.ReceivedOnBehalfOfName

    # Lower case alias for ReceivedOnBehalfOfName
    @property
    def receivedonbehalfofname(self):
        return self.ReceivedOnBehalfOfName

    @property
    def ReceivedTime(self):
        return self.mailitem.ReceivedTime

    # Lower case alias for ReceivedTime
    @property
    def receivedtime(self):
        return self.ReceivedTime

    @property
    def RecipientReassignmentProhibited(self):
        return self.mailitem.RecipientReassignmentProhibited

    # Lower case alias for RecipientReassignmentProhibited
    @property
    def recipientreassignmentprohibited(self):
        return self.RecipientReassignmentProhibited

    @RecipientReassignmentProhibited.setter
    def RecipientReassignmentProhibited(self, value):
        self.mailitem.RecipientReassignmentProhibited = value

    # Lower case alias for RecipientReassignmentProhibited setter
    @recipientreassignmentprohibited.setter
    def recipientreassignmentprohibited(self, value):
        self.RecipientReassignmentProhibited = value

    @property
    def Recipients(self):
        return Recipients(self.mailitem.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return self.mailitem.ReminderOverrideDefault

    # Lower case alias for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.mailitem.ReminderOverrideDefault = value

    # Lower case alias for ReminderOverrideDefault setter
    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.mailitem.ReminderPlaySound

    # Lower case alias for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.mailitem.ReminderPlaySound = value

    # Lower case alias for ReminderPlaySound setter
    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.mailitem.ReminderSet

    # Lower case alias for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.mailitem.ReminderSet = value

    # Lower case alias for ReminderSet setter
    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.mailitem.ReminderSoundFile

    # Lower case alias for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.mailitem.ReminderSoundFile = value

    # Lower case alias for ReminderSoundFile setter
    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.mailitem.ReminderTime

    # Lower case alias for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.mailitem.ReminderTime = value

    # Lower case alias for ReminderTime setter
    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RemoteStatus(self):
        return OlRemoteStatus(self.mailitem.RemoteStatus)

    # Lower case alias for RemoteStatus
    @property
    def remotestatus(self):
        return self.RemoteStatus

    @RemoteStatus.setter
    def RemoteStatus(self, value):
        self.mailitem.RemoteStatus = value

    # Lower case alias for RemoteStatus setter
    @remotestatus.setter
    def remotestatus(self, value):
        self.RemoteStatus = value

    @property
    def ReplyRecipientNames(self):
        return self.mailitem.ReplyRecipientNames

    # Lower case alias for ReplyRecipientNames
    @property
    def replyrecipientnames(self):
        return self.ReplyRecipientNames

    @property
    def ReplyRecipients(self):
        return Recipients(self.mailitem.ReplyRecipients)

    # Lower case alias for ReplyRecipients
    @property
    def replyrecipients(self):
        return self.ReplyRecipients

    @property
    def RetentionExpirationDate(self):
        return MailItem(self.mailitem.RetentionExpirationDate)

    # Lower case alias for RetentionExpirationDate
    @property
    def retentionexpirationdate(self):
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.mailitem.RetentionPolicyName

    # Lower case alias for RetentionPolicyName
    @property
    def retentionpolicyname(self):
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.mailitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.mailitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.mailitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return Folder(self.mailitem.SaveSentMessageFolder)

    # Lower case alias for SaveSentMessageFolder
    @property
    def savesentmessagefolder(self):
        return self.SaveSentMessageFolder

    @SaveSentMessageFolder.setter
    def SaveSentMessageFolder(self, value):
        self.mailitem.SaveSentMessageFolder = value

    # Lower case alias for SaveSentMessageFolder setter
    @savesentmessagefolder.setter
    def savesentmessagefolder(self, value):
        self.SaveSentMessageFolder = value

    @property
    def Sender(self):
        return self.mailitem.Sender

    # Lower case alias for Sender
    @property
    def sender(self):
        return self.Sender

    @Sender.setter
    def Sender(self, value):
        self.mailitem.Sender = value

    # Lower case alias for Sender setter
    @sender.setter
    def sender(self, value):
        self.Sender = value

    @property
    def SenderEmailAddress(self):
        return self.mailitem.SenderEmailAddress

    # Lower case alias for SenderEmailAddress
    @property
    def senderemailaddress(self):
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.mailitem.SenderEmailType

    # Lower case alias for SenderEmailType
    @property
    def senderemailtype(self):
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.mailitem.SenderName

    # Lower case alias for SenderName
    @property
    def sendername(self):
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.mailitem.SendUsingAccount)

    # Lower case alias for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.mailitem.SendUsingAccount = value

    # Lower case alias for SendUsingAccount setter
    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.mailitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.mailitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Sent(self):
        return self.mailitem.Sent

    # Lower case alias for Sent
    @property
    def sent(self):
        return self.Sent

    @property
    def SentOn(self):
        return self.mailitem.SentOn

    # Lower case alias for SentOn
    @property
    def senton(self):
        return self.SentOn

    @property
    def SentOnBehalfOfName(self):
        return self.mailitem.SentOnBehalfOfName

    # Lower case alias for SentOnBehalfOfName
    @property
    def sentonbehalfofname(self):
        return self.SentOnBehalfOfName

    @SentOnBehalfOfName.setter
    def SentOnBehalfOfName(self, value):
        self.mailitem.SentOnBehalfOfName = value

    # Lower case alias for SentOnBehalfOfName setter
    @sentonbehalfofname.setter
    def sentonbehalfofname(self, value):
        self.SentOnBehalfOfName = value

    @property
    def Session(self):
        return NameSpace(self.mailitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.mailitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.mailitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.mailitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Submitted(self):
        return self.mailitem.Submitted

    # Lower case alias for Submitted
    @property
    def submitted(self):
        return self.Submitted

    @property
    def TaskCompletedDate(self):
        return MailItem(self.mailitem.TaskCompletedDate)

    # Lower case alias for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.mailitem.TaskCompletedDate = value

    # Lower case alias for TaskCompletedDate setter
    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return MailItem(self.mailitem.TaskDueDate)

    # Lower case alias for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.mailitem.TaskDueDate = value

    # Lower case alias for TaskDueDate setter
    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return MailItem(self.mailitem.TaskStartDate)

    # Lower case alias for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.mailitem.TaskStartDate = value

    # Lower case alias for TaskStartDate setter
    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return MailItem(self.mailitem.TaskSubject)

    # Lower case alias for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.mailitem.TaskSubject = value

    # Lower case alias for TaskSubject setter
    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def To(self):
        return self.mailitem.To

    # Lower case alias for To
    @property
    def to(self):
        return self.To

    @To.setter
    def To(self, value):
        self.mailitem.To = value

    # Lower case alias for To setter
    @to.setter
    def to(self, value):
        self.To = value

    @property
    def ToDoTaskOrdinal(self):
        return MailItem(self.mailitem.ToDoTaskOrdinal)

    # Lower case alias for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.mailitem.ToDoTaskOrdinal = value

    # Lower case alias for ToDoTaskOrdinal setter
    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.mailitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.mailitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.mailitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    @property
    def VotingOptions(self):
        return self.mailitem.VotingOptions

    # Lower case alias for VotingOptions
    @property
    def votingoptions(self):
        return self.VotingOptions

    @VotingOptions.setter
    def VotingOptions(self, value):
        self.mailitem.VotingOptions = value

    # Lower case alias for VotingOptions setter
    @votingoptions.setter
    def votingoptions(self, value):
        self.VotingOptions = value

    @property
    def VotingResponse(self):
        return self.mailitem.VotingResponse

    # Lower case alias for VotingResponse
    @property
    def votingresponse(self):
        return self.VotingResponse

    @VotingResponse.setter
    def VotingResponse(self, value):
        self.mailitem.VotingResponse = value

    # Lower case alias for VotingResponse setter
    @votingresponse.setter
    def votingresponse(self, value):
        self.VotingResponse = value

    def AddBusinessCard(self, contact=None):
        arguments = com_arguments([contact])
        self.mailitem.AddBusinessCard(*arguments)

    def ClearConversationIndex(self):
        self.mailitem.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.mailitem.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.mailitem.Close(*arguments)

    def Copy(self):
        self.mailitem.Copy()

    def Delete(self):
        self.mailitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.mailitem.Display(*arguments)

    def Forward(self):
        return self.mailitem.Forward()

    def GetConversation(self):
        return self.mailitem.GetConversation()

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.mailitem.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.mailitem.Move(*arguments)

    def PrintOut(self):
        self.mailitem.PrintOut()

    def Reply(self):
        return MailItem(self.mailitem.Reply())

    def ReplyAll(self):
        return MailItem(self.mailitem.ReplyAll())

    def Save(self):
        self.mailitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.mailitem.SaveAs(*arguments)

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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.mailmodule.NavigationGroups)

    # Lower case alias for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.mailmodule.NavigationModuleType)

    # Lower case alias for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.mailmodule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return MailModule(self.mailmodule.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.mailmodule.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.mailmodule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return MailModule(self.mailmodule.Visible)

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.mailmodule.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


class MarkAsTaskRuleAction:

    def __init__(self, markastaskruleaction=None):
        self.markastaskruleaction = markastaskruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.markastaskruleaction.ActionType)

    # Lower case alias for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.markastaskruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.markastaskruleaction.Class)

    @property
    def Enabled(self):
        return self.markastaskruleaction.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.markastaskruleaction.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def FlagTo(self):
        return self.markastaskruleaction.FlagTo

    # Lower case alias for FlagTo
    @property
    def flagto(self):
        return self.FlagTo

    @FlagTo.setter
    def FlagTo(self, value):
        self.markastaskruleaction.FlagTo = value

    # Lower case alias for FlagTo setter
    @flagto.setter
    def flagto(self, value):
        self.FlagTo = value

    @property
    def MarkInterval(self):
        return OlMarkInterval(self.markastaskruleaction.MarkInterval)

    # Lower case alias for MarkInterval
    @property
    def markinterval(self):
        return self.MarkInterval

    @MarkInterval.setter
    def MarkInterval(self, value):
        self.markastaskruleaction.MarkInterval = value

    # Lower case alias for MarkInterval setter
    @markinterval.setter
    def markinterval(self, value):
        self.MarkInterval = value

    @property
    def Parent(self):
        return self.markastaskruleaction.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.markastaskruleaction.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class MeetingItem:

    def __init__(self, meetingitem=None):
        self.meetingitem = meetingitem

    @property
    def Actions(self):
        return Actions(self.meetingitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.meetingitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.meetingitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.meetingitem.AutoForwarded

    # Lower case alias for AutoForwarded
    @property
    def autoforwarded(self):
        return self.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.meetingitem.AutoForwarded = value

    # Lower case alias for AutoForwarded setter
    @autoforwarded.setter
    def autoforwarded(self, value):
        self.AutoForwarded = value

    @property
    def AutoResolvedWinner(self):
        return self.meetingitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.meetingitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.meetingitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.meetingitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.meetingitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.meetingitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.meetingitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.meetingitem.Class)

    @property
    def Companies(self):
        return self.meetingitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.meetingitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.meetingitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.meetingitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.meetingitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.meetingitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.meetingitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return self.meetingitem.DeferredDeliveryTime

    # Lower case alias for DeferredDeliveryTime
    @property
    def deferreddeliverytime(self):
        return self.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.meetingitem.DeferredDeliveryTime = value

    # Lower case alias for DeferredDeliveryTime setter
    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.meetingitem.DeleteAfterSubmit

    # Lower case alias for DeleteAfterSubmit
    @property
    def deleteaftersubmit(self):
        return self.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.meetingitem.DeleteAfterSubmit = value

    # Lower case alias for DeleteAfterSubmit setter
    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.meetingitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.meetingitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.meetingitem.ExpiryTime

    # Lower case alias for ExpiryTime
    @property
    def expirytime(self):
        return self.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.meetingitem.ExpiryTime = value

    # Lower case alias for ExpiryTime setter
    @expirytime.setter
    def expirytime(self, value):
        self.ExpiryTime = value

    @property
    def FormDescription(self):
        return FormDescription(self.meetingitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.meetingitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.meetingitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.meetingitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.meetingitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsLatestVersion(self):
        return MeetingItem(self.meetingitem.IsLatestVersion)

    # Lower case alias for IsLatestVersion
    @property
    def islatestversion(self):
        return self.IsLatestVersion

    @property
    def ItemProperties(self):
        return ItemProperties(self.meetingitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.meetingitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.meetingitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.meetingitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MeetingWorkspaceURL(self):
        return self.meetingitem.MeetingWorkspaceURL

    # Lower case alias for MeetingWorkspaceURL
    @property
    def meetingworkspaceurl(self):
        return self.MeetingWorkspaceURL

    @property
    def MessageClass(self):
        return self.meetingitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.meetingitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.meetingitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.meetingitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.meetingitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.meetingitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return self.meetingitem.OriginatorDeliveryReportRequested

    # Lower case alias for OriginatorDeliveryReportRequested
    @property
    def originatordeliveryreportrequested(self):
        return self.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.meetingitem.OriginatorDeliveryReportRequested = value

    # Lower case alias for OriginatorDeliveryReportRequested setter
    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return self.meetingitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.meetingitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.meetingitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.meetingitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReceivedTime(self):
        return self.meetingitem.ReceivedTime

    # Lower case alias for ReceivedTime
    @property
    def receivedtime(self):
        return self.ReceivedTime

    @ReceivedTime.setter
    def ReceivedTime(self, value):
        self.meetingitem.ReceivedTime = value

    # Lower case alias for ReceivedTime setter
    @receivedtime.setter
    def receivedtime(self, value):
        self.ReceivedTime = value

    @property
    def Recipients(self):
        return Recipients(self.meetingitem.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def ReminderSet(self):
        return self.meetingitem.ReminderSet

    # Lower case alias for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.meetingitem.ReminderSet = value

    # Lower case alias for ReminderSet setter
    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderTime(self):
        return self.meetingitem.ReminderTime

    # Lower case alias for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.meetingitem.ReminderTime = value

    # Lower case alias for ReminderTime setter
    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def ReplyRecipients(self):
        return Recipients(self.meetingitem.ReplyRecipients)

    # Lower case alias for ReplyRecipients
    @property
    def replyrecipients(self):
        return self.ReplyRecipients

    @property
    def RetentionExpirationDate(self):
        return MeetingItem(self.meetingitem.RetentionExpirationDate)

    # Lower case alias for RetentionExpirationDate
    @property
    def retentionexpirationdate(self):
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.meetingitem.RetentionPolicyName

    # Lower case alias for RetentionPolicyName
    @property
    def retentionpolicyname(self):
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.meetingitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.meetingitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.meetingitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return self.meetingitem.SaveSentMessageFolder

    # Lower case alias for SaveSentMessageFolder
    @property
    def savesentmessagefolder(self):
        return self.SaveSentMessageFolder

    @property
    def SenderEmailAddress(self):
        return self.meetingitem.SenderEmailAddress

    # Lower case alias for SenderEmailAddress
    @property
    def senderemailaddress(self):
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.meetingitem.SenderEmailType

    # Lower case alias for SenderEmailType
    @property
    def senderemailtype(self):
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.meetingitem.SenderName

    # Lower case alias for SenderName
    @property
    def sendername(self):
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.meetingitem.SendUsingAccount)

    # Lower case alias for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.meetingitem.SendUsingAccount = value

    # Lower case alias for SendUsingAccount setter
    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.meetingitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.meetingitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Sent(self):
        return self.meetingitem.Sent

    # Lower case alias for Sent
    @property
    def sent(self):
        return self.Sent

    @property
    def SentOn(self):
        return self.meetingitem.SentOn

    # Lower case alias for SentOn
    @property
    def senton(self):
        return self.SentOn

    @property
    def Session(self):
        return NameSpace(self.meetingitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.meetingitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.meetingitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.meetingitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Submitted(self):
        return self.meetingitem.Submitted

    # Lower case alias for Submitted
    @property
    def submitted(self):
        return self.Submitted

    @property
    def UnRead(self):
        return self.meetingitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.meetingitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.meetingitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.meetingitem.Close(*arguments)

    def Copy(self):
        self.meetingitem.Copy()

    def Delete(self):
        self.meetingitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.meetingitem.Display(*arguments)

    def Forward(self):
        return self.meetingitem.Forward()

    def GetAssociatedAppointment(self, AddToCalendar=None):
        arguments = com_arguments([AddToCalendar])
        return self.meetingitem.GetAssociatedAppointment(*arguments)

    def GetConversation(self):
        return self.meetingitem.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.meetingitem.Move(*arguments)

    def PrintOut(self):
        self.meetingitem.PrintOut()

    def Reply(self):
        return MailItem(self.meetingitem.Reply())

    def ReplyAll(self):
        return MailItem(self.meetingitem.ReplyAll())

    def Save(self):
        self.meetingitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.meetingitem.SaveAs(*arguments)

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

    # Lower case alias for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.moveorcopyruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.moveorcopyruleaction.Class)

    @property
    def Enabled(self):
        return self.moveorcopyruleaction.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.moveorcopyruleaction.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Folder(self):
        return Folder(self.moveorcopyruleaction.Folder)

    # Lower case alias for Folder
    @property
    def folder(self):
        return self.Folder

    @Folder.setter
    def Folder(self, value):
        self.moveorcopyruleaction.Folder = value

    # Lower case alias for Folder setter
    @folder.setter
    def folder(self, value):
        self.Folder = value

    @property
    def Parent(self):
        return self.moveorcopyruleaction.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.moveorcopyruleaction.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class NameSpace:

    def __init__(self, namespace=None):
        self.namespace = namespace

    @property
    def Accounts(self):
        return Accounts(self.namespace.Accounts)

    # Lower case alias for Accounts
    @property
    def accounts(self):
        return self.Accounts

    @property
    def AddressLists(self):
        return AddressLists(self.namespace.AddressLists)

    # Lower case alias for AddressLists
    @property
    def addresslists(self):
        return self.AddressLists

    @property
    def Application(self):
        return Application(self.namespace.Application)

    @property
    def AutoDiscoverConnectionMode(self):
        return OlAutoDiscoverConnectionMode(self.namespace.AutoDiscoverConnectionMode)

    # Lower case alias for AutoDiscoverConnectionMode
    @property
    def autodiscoverconnectionmode(self):
        return self.AutoDiscoverConnectionMode

    @property
    def AutoDiscoverXml(self):
        return self.namespace.AutoDiscoverXml

    # Lower case alias for AutoDiscoverXml
    @property
    def autodiscoverxml(self):
        return self.AutoDiscoverXml

    @property
    def Categories(self):
        return Categories(self.namespace.Categories)

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.namespace.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.namespace.Class)

    @property
    def CurrentProfileName(self):
        return self.namespace.CurrentProfileName

    # Lower case alias for CurrentProfileName
    @property
    def currentprofilename(self):
        return self.CurrentProfileName

    @property
    def CurrentUser(self):
        return Recipient(self.namespace.CurrentUser)

    # Lower case alias for CurrentUser
    @property
    def currentuser(self):
        return self.CurrentUser

    @property
    def DefaultStore(self):
        return Store(self.namespace.DefaultStore)

    # Lower case alias for DefaultStore
    @property
    def defaultstore(self):
        return self.DefaultStore

    @property
    def ExchangeConnectionMode(self):
        return OlExchangeConnectionMode(self.namespace.ExchangeConnectionMode)

    # Lower case alias for ExchangeConnectionMode
    @property
    def exchangeconnectionmode(self):
        return self.ExchangeConnectionMode

    @property
    def ExchangeMailboxServerName(self):
        return self.namespace.ExchangeMailboxServerName

    # Lower case alias for ExchangeMailboxServerName
    @property
    def exchangemailboxservername(self):
        return self.ExchangeMailboxServerName

    @property
    def ExchangeMailboxServerVersion(self):
        return self.namespace.ExchangeMailboxServerVersion

    # Lower case alias for ExchangeMailboxServerVersion
    @property
    def exchangemailboxserverversion(self):
        return self.ExchangeMailboxServerVersion

    @property
    def Folders(self):
        return Folders(self.namespace.Folders)

    # Lower case alias for Folders
    @property
    def folders(self):
        return self.Folders

    @property
    def Offline(self):
        return self.namespace.Offline

    # Lower case alias for Offline
    @property
    def offline(self):
        return self.Offline

    @property
    def Parent(self):
        return self.namespace.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.namespace.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Stores(self):
        return Stores(self.namespace.Stores)

    # Lower case alias for Stores
    @property
    def stores(self):
        return self.Stores

    @property
    def SyncObjects(self):
        return SyncObjects(self.namespace.SyncObjects)

    # Lower case alias for SyncObjects
    @property
    def syncobjects(self):
        return self.SyncObjects

    @property
    def Type(self):
        return self.namespace.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    def AddStore(self, Store=None):
        arguments = com_arguments([Store])
        self.namespace.AddStore(*arguments)

    def AddStoreEx(self, Store=None, Type=None):
        arguments = com_arguments([Store, Type])
        self.namespace.AddStoreEx(*arguments)

    def CompareEntryIDs(self, FirstEntryID=None, SecondEntryID=None):
        arguments = com_arguments([FirstEntryID, SecondEntryID])
        return self.namespace.CompareEntryIDs(*arguments)

    def CreateContactCard(self, Address=None):
        arguments = com_arguments([Address])
        return self.namespace.CreateContactCard(*arguments)

    def CreateRecipient(self, RecipientName=None):
        arguments = com_arguments([RecipientName])
        return self.namespace.CreateRecipient(*arguments)

    def CreateSharingItem(self, Context=None, Provider=None):
        arguments = com_arguments([Context, Provider])
        return self.namespace.CreateSharingItem(*arguments)

    def Dial(self, ContactItem=None):
        arguments = com_arguments([ContactItem])
        self.namespace.Dial(*arguments)

    def GetAddressEntryFromID(self, ID=None):
        arguments = com_arguments([ID])
        return ID(self.namespace.GetAddressEntryFromID(*arguments))

    def GetDefaultFolder(self, FolderType=None):
        arguments = com_arguments([FolderType])
        return self.namespace.GetDefaultFolder(*arguments)

    def GetFolderFromID(self, EntryIDFolder=None, EntryIDStore=None):
        arguments = com_arguments([EntryIDFolder, EntryIDStore])
        return self.namespace.GetFolderFromID(*arguments)

    def GetGlobalAddressList(self):
        return self.namespace.GetGlobalAddressList()

    def GetItemFromID(self, EntryIDItem=None, EntryIDStore=None):
        arguments = com_arguments([EntryIDItem, EntryIDStore])
        return self.namespace.GetItemFromID(*arguments)

    def GetRecipientFromID(self, EntryID=None):
        arguments = com_arguments([EntryID])
        return self.namespace.GetRecipientFromID(*arguments)

    def GetSelectNamesDialog(self):
        return self.namespace.GetSelectNamesDialog()

    def GetSharedDefaultFolder(self, Recipient=None, FolderType=None):
        arguments = com_arguments([Recipient, FolderType])
        return self.namespace.GetSharedDefaultFolder(*arguments)

    def GetStoreFromID(self, ID=None):
        arguments = com_arguments([ID])
        return StoreID(self.namespace.GetStoreFromID(*arguments))

    def Logoff(self):
        self.namespace.Logoff()

    def Logon(self, Profile=None, Password=None, ShowDialog=None, NewSession=None):
        arguments = com_arguments([Profile, Password, ShowDialog, NewSession])
        self.namespace.Logon(*arguments)

    def OpenSharedFolder(self, Path=None, Name=None, DownloadAttachments=None, UseTTL=None):
        arguments = com_arguments([Path, Name, DownloadAttachments, UseTTL])
        return Folder(self.namespace.OpenSharedFolder(*arguments))

    def OpenSharedItem(self, Path=None):
        arguments = com_arguments([Path])
        return self.namespace.OpenSharedItem(*arguments)

    def PickFolder(self):
        return Folder(self.namespace.PickFolder())

    def RemoveStore(self, Folder=None):
        arguments = com_arguments([Folder])
        self.namespace.RemoveStore(*arguments)

    def SendAndReceive(self, showProgressDialog=None):
        arguments = com_arguments([showProgressDialog])
        self.namespace.SendAndReceive(*arguments)


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

    # Lower case alias for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def Folder(self):
        return Folder(self.navigationfolder.Folder)

    # Lower case alias for Folder
    @property
    def folder(self):
        return self.Folder

    @property
    def IsRemovable(self):
        return NavigationFolder(self.navigationfolder.IsRemovable)

    # Lower case alias for IsRemovable
    @property
    def isremovable(self):
        return self.IsRemovable

    @property
    def IsSelected(self):
        return NavigationFolder(self.navigationfolder.IsSelected)

    # Lower case alias for IsSelected
    @property
    def isselected(self):
        return self.IsSelected

    @IsSelected.setter
    def IsSelected(self, value):
        self.navigationfolder.IsSelected = value

    # Lower case alias for IsSelected setter
    @isselected.setter
    def isselected(self, value):
        self.IsSelected = value

    @property
    def IsSideBySide(self):
        return NavigationFolder(self.navigationfolder.IsSideBySide)

    # Lower case alias for IsSideBySide
    @property
    def issidebyside(self):
        return self.IsSideBySide

    @IsSideBySide.setter
    def IsSideBySide(self, value):
        self.navigationfolder.IsSideBySide = value

    # Lower case alias for IsSideBySide setter
    @issidebyside.setter
    def issidebyside(self, value):
        self.IsSideBySide = value

    @property
    def Parent(self):
        return self.navigationfolder.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return NavigationFolder(self.navigationfolder.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.navigationfolder.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.navigationfolder.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.navigationfolders.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.navigationfolders.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Folder=None):
        arguments = com_arguments([Folder])
        return self.navigationfolders.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.navigationfolders.Item(*arguments)

    def Remove(self, RemovableFolder=None):
        arguments = com_arguments([RemovableFolder])
        self.navigationfolders.Remove(*arguments)


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

    # Lower case alias for GroupType
    @property
    def grouptype(self):
        return self.GroupType

    @property
    def Name(self):
        return NavigationGroup(self.navigationgroup.Name)

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.navigationgroup.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def NavigationFolders(self):
        return NavigationFolders(self.navigationgroup.NavigationFolders)

    # Lower case alias for NavigationFolders
    @property
    def navigationfolders(self):
        return self.NavigationFolders

    @property
    def Parent(self):
        return self.navigationgroup.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return NavigationGroup(self.navigationgroup.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.navigationgroup.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.navigationgroup.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.navigationgroups.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.navigationgroups.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Create(self, GroupDisplayName=None):
        arguments = com_arguments([GroupDisplayName])
        return self.navigationgroups.Create(*arguments)

    def Delete(self, Group=None):
        arguments = com_arguments([Group])
        self.navigationgroups.Delete(*arguments)

    def GetDefaultNavigationGroup(self, DefaultFolderGroup=None):
        arguments = com_arguments([DefaultFolderGroup])
        return self.navigationgroups.GetDefaultNavigationGroup(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.navigationgroups.Item(*arguments)


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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.navigationmodule.NavigationModuleType)

    # Lower case alias for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.navigationmodule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return NavigationModule(self.navigationmodule.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.navigationmodule.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.navigationmodule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return NavigationModule(self.navigationmodule.Visible)

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.navigationmodule.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.navigationmodules.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.navigationmodules.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def GetNavigationModule(self, ModuleType=None):
        arguments = com_arguments([ModuleType])
        return self.navigationmodules.GetNavigationModule(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.navigationmodules.Item(*arguments)


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

    # Lower case alias for CurrentModule
    @property
    def currentmodule(self):
        return self.CurrentModule

    @CurrentModule.setter
    def CurrentModule(self, value):
        self.navigationpane.CurrentModule = value

    # Lower case alias for CurrentModule setter
    @currentmodule.setter
    def currentmodule(self, value):
        self.CurrentModule = value

    @property
    def DisplayedModuleCount(self):
        return NavigationModule(self.navigationpane.DisplayedModuleCount)

    # Lower case alias for DisplayedModuleCount
    @property
    def displayedmodulecount(self):
        return self.DisplayedModuleCount

    @DisplayedModuleCount.setter
    def DisplayedModuleCount(self, value):
        self.navigationpane.DisplayedModuleCount = value

    # Lower case alias for DisplayedModuleCount setter
    @displayedmodulecount.setter
    def displayedmodulecount(self, value):
        self.DisplayedModuleCount = value

    @property
    def IsCollapsed(self):
        return self.navigationpane.IsCollapsed

    # Lower case alias for IsCollapsed
    @property
    def iscollapsed(self):
        return self.IsCollapsed

    @IsCollapsed.setter
    def IsCollapsed(self, value):
        self.navigationpane.IsCollapsed = value

    # Lower case alias for IsCollapsed setter
    @iscollapsed.setter
    def iscollapsed(self, value):
        self.IsCollapsed = value

    @property
    def Modules(self):
        return NavigationModules(self.navigationpane.Modules)

    # Lower case alias for Modules
    @property
    def modules(self):
        return self.Modules

    @property
    def Parent(self):
        return self.navigationpane.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.navigationpane.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class NewItemAlertRuleAction:

    def __init__(self, newitemalertruleaction=None):
        self.newitemalertruleaction = newitemalertruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.newitemalertruleaction.ActionType)

    # Lower case alias for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.newitemalertruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.newitemalertruleaction.Class)

    @property
    def Enabled(self):
        return self.newitemalertruleaction.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.newitemalertruleaction.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.newitemalertruleaction.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.newitemalertruleaction.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Text(self):
        return self.newitemalertruleaction.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.newitemalertruleaction.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value


class NoteItem:

    def __init__(self, noteitem=None):
        self.noteitem = noteitem

    @property
    def Application(self):
        return Application(self.noteitem.Application)

    @property
    def AutoResolvedWinner(self):
        return self.noteitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def Body(self):
        return self.noteitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.noteitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.noteitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.noteitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.noteitem.Class)

    @property
    def Conflicts(self):
        return self.noteitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def CreationTime(self):
        return self.noteitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.noteitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.noteitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def GetInspector(self):
        return Inspector(self.noteitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Height(self):
        return self.noteitem.Height

    # Lower case alias for Height
    @property
    def height(self):
        return self.Height

    @Height.setter
    def Height(self, value):
        self.noteitem.Height = value

    # Lower case alias for Height setter
    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def IsConflict(self):
        return self.noteitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.noteitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.noteitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def Left(self):
        return self.noteitem.Left

    # Lower case alias for Left
    @property
    def left(self):
        return self.Left

    @Left.setter
    def Left(self, value):
        self.noteitem.Left = value

    # Lower case alias for Left setter
    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.noteitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.noteitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.noteitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.noteitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Parent(self):
        return self.noteitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.noteitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Saved(self):
        return self.noteitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Session(self):
        return NameSpace(self.noteitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.noteitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.noteitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.noteitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Top(self):
        return self.noteitem.Top

    # Lower case alias for Top
    @property
    def top(self):
        return self.Top

    @Top.setter
    def Top(self, value):
        self.noteitem.Top = value

    # Lower case alias for Top setter
    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.noteitem.Width

    # Lower case alias for Width
    @property
    def width(self):
        return self.Width

    @Width.setter
    def Width(self, value):
        self.noteitem.Width = value

    # Lower case alias for Width setter
    @width.setter
    def width(self, value):
        self.Width = value

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.noteitem.Close(*arguments)

    def Copy(self):
        return NoteItem(self.noteitem.Copy())

    def Delete(self):
        self.noteitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.noteitem.Display(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.noteitem.Move(*arguments)

    def PrintOut(self):
        self.noteitem.PrintOut()

    def Save(self):
        self.noteitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.noteitem.SaveAs(*arguments)


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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.notesmodule.NavigationGroups)

    # Lower case alias for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.notesmodule.NavigationModuleType)

    # Lower case alias for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.notesmodule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return NotesModule(self.notesmodule.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.notesmodule.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.notesmodule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return NotesModule(self.notesmodule.Visible)

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.notesmodule.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


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

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkbusinesscardcontrol.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkbusinesscardcontrol.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkbusinesscardcontrol.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkCategory:

    def __init__(self, olkcategory=None):
        self.olkcategory = olkcategory

    @property
    def AutoSize(self):
        return self.olkcategory.AutoSize

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olkcategory.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def BackColor(self):
        return self.olkcategory.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkcategory.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olkcategory.BackStyle)

    # Lower case alias for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.olkcategory.BackStyle = value

    # Lower case alias for BackStyle setter
    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Enabled(self):
        return self.olkcategory.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcategory.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def ForeColor(self):
        return self.olkcategory.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkcategory.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.olkcategory.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcategory.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcategory.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcategory.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkCheckBox:

    def __init__(self, olkcheckbox=None):
        self.olkcheckbox = olkcheckbox

    @property
    def Accelerator(self):
        return self.olkcheckbox.Accelerator

    # Lower case alias for Accelerator
    @property
    def accelerator(self):
        return self.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.olkcheckbox.Accelerator = value

    # Lower case alias for Accelerator setter
    @accelerator.setter
    def accelerator(self, value):
        self.Accelerator = value

    @property
    def Alignment(self):
        return olAlignment(self.olkcheckbox.Alignment)

    # Lower case alias for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.olkcheckbox.Alignment = value

    # Lower case alias for Alignment setter
    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def BackColor(self):
        return self.olkcheckbox.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkcheckbox.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olkcheckbox.BackStyle)

    # Lower case alias for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.olkcheckbox.BackStyle = value

    # Lower case alias for BackStyle setter
    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Caption(self):
        return self.olkcheckbox.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.olkcheckbox.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def Enabled(self):
        return self.olkcheckbox.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcheckbox.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.olkcheckbox.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olkcheckbox.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkcheckbox.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.olkcheckbox.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcheckbox.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcheckbox.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcheckbox.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def TripleState(self):
        return self.olkcheckbox.TripleState

    # Lower case alias for TripleState
    @property
    def triplestate(self):
        return self.TripleState

    @TripleState.setter
    def TripleState(self, value):
        self.olkcheckbox.TripleState = value

    # Lower case alias for TripleState setter
    @triplestate.setter
    def triplestate(self, value):
        self.TripleState = value

    @property
    def Value(self):
        return self.olkcheckbox.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.olkcheckbox.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

    @property
    def WordWrap(self):
        return self.olkcheckbox.WordWrap

    # Lower case alias for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olkcheckbox.WordWrap = value

    # Lower case alias for WordWrap setter
    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class OlkComboBox:

    def __init__(self, olkcombobox=None):
        self.olkcombobox = olkcombobox

    @property
    def AutoSize(self):
        return self.olkcombobox.AutoSize

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olkcombobox.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def AutoTab(self):
        return self.olkcombobox.AutoTab

    # Lower case alias for AutoTab
    @property
    def autotab(self):
        return self.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.olkcombobox.AutoTab = value

    # Lower case alias for AutoTab setter
    @autotab.setter
    def autotab(self, value):
        self.AutoTab = value

    @property
    def AutoWordSelect(self):
        return self.olkcombobox.AutoWordSelect

    # Lower case alias for AutoWordSelect
    @property
    def autowordselect(self):
        return self.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.olkcombobox.AutoWordSelect = value

    # Lower case alias for AutoWordSelect setter
    @autowordselect.setter
    def autowordselect(self, value):
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.olkcombobox.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkcombobox.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olkcombobox.BorderStyle)

    # Lower case alias for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olkcombobox.BorderStyle = value

    # Lower case alias for BorderStyle setter
    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def DragBehavior(self):
        return self.olkcombobox.DragBehavior

    # Lower case alias for DragBehavior
    @property
    def dragbehavior(self):
        return self.DragBehavior

    @DragBehavior.setter
    def DragBehavior(self, value):
        self.olkcombobox.DragBehavior = value

    # Lower case alias for DragBehavior setter
    @dragbehavior.setter
    def dragbehavior(self, value):
        self.DragBehavior = value

    @property
    def Enabled(self):
        return self.olkcombobox.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcombobox.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.olkcombobox.EnterFieldBehavior)

    # Lower case alias for EnterFieldBehavior
    @property
    def enterfieldbehavior(self):
        return self.EnterFieldBehavior

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.olkcombobox.EnterFieldBehavior = value

    # Lower case alias for EnterFieldBehavior setter
    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.olkcombobox.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olkcombobox.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkcombobox.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.olkcombobox.HideSelection

    # Lower case alias for HideSelection
    @property
    def hideselection(self):
        return self.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.olkcombobox.HideSelection = value

    # Lower case alias for HideSelection setter
    @hideselection.setter
    def hideselection(self, value):
        self.HideSelection = value

    @property
    def ListCount(self):
        return self.olkcombobox.ListCount

    # Lower case alias for ListCount
    @property
    def listcount(self):
        return self.ListCount

    @property
    def ListIndex(self):
        return self.olkcombobox.ListIndex

    # Lower case alias for ListIndex
    @property
    def listindex(self):
        return self.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.olkcombobox.ListIndex = value

    # Lower case alias for ListIndex setter
    @listindex.setter
    def listindex(self, value):
        self.ListIndex = value

    @property
    def Locked(self):
        return self.olkcombobox.Locked

    # Lower case alias for Locked
    @property
    def locked(self):
        return self.Locked

    @Locked.setter
    def Locked(self, value):
        self.olkcombobox.Locked = value

    # Lower case alias for Locked setter
    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MaxLength(self):
        return self.olkcombobox.MaxLength

    # Lower case alias for MaxLength
    @property
    def maxlength(self):
        return self.MaxLength

    @MaxLength.setter
    def MaxLength(self, value):
        self.olkcombobox.MaxLength = value

    # Lower case alias for MaxLength setter
    @maxlength.setter
    def maxlength(self, value):
        self.MaxLength = value

    @property
    def MouseIcon(self):
        return self.olkcombobox.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcombobox.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcombobox.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcombobox.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def SelectionMargin(self):
        return self.olkcombobox.SelectionMargin

    # Lower case alias for SelectionMargin
    @property
    def selectionmargin(self):
        return self.SelectionMargin

    @SelectionMargin.setter
    def SelectionMargin(self, value):
        self.olkcombobox.SelectionMargin = value

    # Lower case alias for SelectionMargin setter
    @selectionmargin.setter
    def selectionmargin(self, value):
        self.SelectionMargin = value

    @property
    def SelLength(self):
        return self.olkcombobox.SelLength

    # Lower case alias for SelLength
    @property
    def sellength(self):
        return self.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.olkcombobox.SelLength = value

    # Lower case alias for SelLength setter
    @sellength.setter
    def sellength(self, value):
        self.SelLength = value

    @property
    def SelStart(self):
        return self.olkcombobox.SelStart

    # Lower case alias for SelStart
    @property
    def selstart(self):
        return self.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.olkcombobox.SelStart = value

    # Lower case alias for SelStart setter
    @selstart.setter
    def selstart(self, value):
        self.SelStart = value

    @property
    def SelText(self):
        return self.olkcombobox.SelText

    # Lower case alias for SelText
    @property
    def seltext(self):
        return self.SelText

    @property
    def Style(self):
        return OlComboBoxStyle(self.olkcombobox.Style)

    # Lower case alias for Style
    @property
    def style(self):
        return self.Style

    @Style.setter
    def Style(self, value):
        self.olkcombobox.Style = value

    # Lower case alias for Style setter
    @style.setter
    def style(self, value):
        self.Style = value

    @property
    def Text(self):
        return self.olkcombobox.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.olkcombobox.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olkcombobox.TextAlign)

    # Lower case alias for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.olkcombobox.TextAlign = value

    # Lower case alias for TextAlign setter
    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def TopIndex(self):
        return self.olkcombobox.TopIndex

    # Lower case alias for TopIndex
    @property
    def topindex(self):
        return self.TopIndex

    @TopIndex.setter
    def TopIndex(self, value):
        self.olkcombobox.TopIndex = value

    # Lower case alias for TopIndex setter
    @topindex.setter
    def topindex(self, value):
        self.TopIndex = value

    @property
    def Value(self):
        return self.olkcombobox.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.olkcombobox.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

    def AddItem(self, ItemText=None, Index=None):
        arguments = com_arguments([ItemText, Index])
        self.olkcombobox.AddItem(*arguments)

    def Clear(self):
        self.olkcombobox.Clear()

    def Copy(self):
        self.olkcombobox.Copy()

    def Cut(self):
        self.olkcombobox.Cut()

    def DropDown(self):
        self.olkcombobox.DropDown()

    def GetItem(self, Index=None):
        arguments = com_arguments([Index])
        return self.olkcombobox.GetItem(*arguments)

    def Paste(self):
        self.olkcombobox.Paste()

    def RemoveItem(self, Index=None):
        arguments = com_arguments([Index])
        self.olkcombobox.RemoveItem(*arguments)

    def SetItem(self, Index=None, Item=None):
        arguments = com_arguments([Index, Item])
        self.olkcombobox.SetItem(*arguments)


class OlkCommandButton:

    def __init__(self, olkcommandbutton=None):
        self.olkcommandbutton = olkcommandbutton

    @property
    def Accelerator(self):
        return self.olkcommandbutton.Accelerator

    # Lower case alias for Accelerator
    @property
    def accelerator(self):
        return self.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.olkcommandbutton.Accelerator = value

    # Lower case alias for Accelerator setter
    @accelerator.setter
    def accelerator(self, value):
        self.Accelerator = value

    @property
    def AutoSize(self):
        return self.olkcommandbutton.AutoSize

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olkcommandbutton.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def Caption(self):
        return self.olkcommandbutton.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.olkcommandbutton.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def DisplayDropArrow(self):
        return self.olkcommandbutton.DisplayDropArrow

    # Lower case alias for DisplayDropArrow
    @property
    def displaydroparrow(self):
        return self.DisplayDropArrow

    @DisplayDropArrow.setter
    def DisplayDropArrow(self, value):
        self.olkcommandbutton.DisplayDropArrow = value

    # Lower case alias for DisplayDropArrow setter
    @displaydroparrow.setter
    def displaydroparrow(self, value):
        self.DisplayDropArrow = value

    @property
    def Enabled(self):
        return self.olkcommandbutton.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcommandbutton.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.olkcommandbutton.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def MouseIcon(self):
        return self.olkcommandbutton.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcommandbutton.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcommandbutton.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcommandbutton.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def Picture(self):
        return self.olkcommandbutton.Picture

    # Lower case alias for Picture
    @property
    def picture(self):
        return self.Picture

    @Picture.setter
    def Picture(self, value):
        self.olkcommandbutton.Picture = value

    # Lower case alias for Picture setter
    @picture.setter
    def picture(self, value):
        self.Picture = value

    @property
    def PictureAlignment(self):
        return OlPictureAlignment(self.olkcommandbutton.PictureAlignment)

    # Lower case alias for PictureAlignment
    @property
    def picturealignment(self):
        return self.PictureAlignment

    @PictureAlignment.setter
    def PictureAlignment(self, value):
        self.olkcommandbutton.PictureAlignment = value

    # Lower case alias for PictureAlignment setter
    @picturealignment.setter
    def picturealignment(self, value):
        self.PictureAlignment = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olkcommandbutton.TextAlign)

    # Lower case alias for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.olkcommandbutton.TextAlign = value

    # Lower case alias for TextAlign setter
    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def WordWrap(self):
        return self.olkcommandbutton.WordWrap

    # Lower case alias for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olkcommandbutton.WordWrap = value

    # Lower case alias for WordWrap setter
    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class OlkContactPhoto:

    def __init__(self, olkcontactphoto=None):
        self.olkcontactphoto = olkcontactphoto

    @property
    def Enabled(self):
        return self.olkcontactphoto.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkcontactphoto.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def MouseIcon(self):
        return self.olkcontactphoto.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkcontactphoto.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkcontactphoto.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkcontactphoto.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkDateControl:

    def __init__(self, olkdatecontrol=None):
        self.olkdatecontrol = olkdatecontrol

    @property
    def AutoSize(self):
        return self.olkdatecontrol.AutoSize

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olkdatecontrol.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def AutoWordSelect(self):
        return self.olkdatecontrol.AutoWordSelect

    # Lower case alias for AutoWordSelect
    @property
    def autowordselect(self):
        return self.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.olkdatecontrol.AutoWordSelect = value

    # Lower case alias for AutoWordSelect setter
    @autowordselect.setter
    def autowordselect(self, value):
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.olkdatecontrol.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkdatecontrol.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olkdatecontrol.BackStyle)

    # Lower case alias for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.olkdatecontrol.BackStyle = value

    # Lower case alias for BackStyle setter
    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Date(self):
        return self.olkdatecontrol.Date

    # Lower case alias for Date
    @property
    def date(self):
        return self.Date

    @Date.setter
    def Date(self, value):
        self.olkdatecontrol.Date = value

    # Lower case alias for Date setter
    @date.setter
    def date(self, value):
        self.Date = value

    @property
    def Enabled(self):
        return self.olkdatecontrol.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkdatecontrol.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.olkdatecontrol.EnterFieldBehavior)

    # Lower case alias for EnterFieldBehavior
    @property
    def enterfieldbehavior(self):
        return self.EnterFieldBehavior

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.olkdatecontrol.EnterFieldBehavior = value

    # Lower case alias for EnterFieldBehavior setter
    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.olkdatecontrol.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olkdatecontrol.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkdatecontrol.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.olkdatecontrol.HideSelection

    # Lower case alias for HideSelection
    @property
    def hideselection(self):
        return self.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.olkdatecontrol.HideSelection = value

    # Lower case alias for HideSelection setter
    @hideselection.setter
    def hideselection(self, value):
        self.HideSelection = value

    @property
    def Locked(self):
        return self.olkdatecontrol.Locked

    # Lower case alias for Locked
    @property
    def locked(self):
        return self.Locked

    @Locked.setter
    def Locked(self, value):
        self.olkdatecontrol.Locked = value

    # Lower case alias for Locked setter
    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.olkdatecontrol.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkdatecontrol.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkdatecontrol.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkdatecontrol.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def ShowNoneButton(self):
        return self.olkdatecontrol.ShowNoneButton

    # Lower case alias for ShowNoneButton
    @property
    def shownonebutton(self):
        return self.ShowNoneButton

    @ShowNoneButton.setter
    def ShowNoneButton(self, value):
        self.olkdatecontrol.ShowNoneButton = value

    # Lower case alias for ShowNoneButton setter
    @shownonebutton.setter
    def shownonebutton(self, value):
        self.ShowNoneButton = value

    @property
    def Text(self):
        return self.olkdatecontrol.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.olkdatecontrol.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olkdatecontrol.TextAlign)

    # Lower case alias for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.olkdatecontrol.TextAlign = value

    # Lower case alias for TextAlign setter
    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def Value(self):
        return self.olkdatecontrol.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.olkdatecontrol.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

    def DropDown(self):
        self.olkdatecontrol.DropDown()


class OlkFrameHeader:

    def __init__(self, olkframeheader=None):
        self.olkframeheader = olkframeheader

    @property
    def Alignment(self):
        return olAlignment(self.olkframeheader.Alignment)

    # Lower case alias for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.olkframeheader.Alignment = value

    # Lower case alias for Alignment setter
    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def Caption(self):
        return self.olkframeheader.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.olkframeheader.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def Enabled(self):
        return self.olkframeheader.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkframeheader.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.olkframeheader.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olkframeheader.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkframeheader.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.olkframeheader.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkframeheader.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkframeheader.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkframeheader.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkInfoBar:

    def __init__(self, olkinfobar=None):
        self.olkinfobar = olkinfobar

    @property
    def MouseIcon(self):
        return self.olkinfobar.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkinfobar.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkinfobar.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkinfobar.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkLabel:

    def __init__(self, olklabel=None):
        self.olklabel = olklabel

    @property
    def Accelerator(self):
        return self.olklabel.Accelerator

    # Lower case alias for Accelerator
    @property
    def accelerator(self):
        return self.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.olklabel.Accelerator = value

    # Lower case alias for Accelerator setter
    @accelerator.setter
    def accelerator(self, value):
        self.Accelerator = value

    @property
    def AutoSize(self):
        return self.olklabel.AutoSize

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olklabel.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def BackColor(self):
        return self.olklabel.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olklabel.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olklabel.BackStyle)

    # Lower case alias for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.olklabel.BackStyle = value

    # Lower case alias for BackStyle setter
    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olklabel.BorderStyle)

    # Lower case alias for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olklabel.BorderStyle = value

    # Lower case alias for BorderStyle setter
    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def Caption(self):
        return self.olklabel.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.olklabel.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def Enabled(self):
        return self.olklabel.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olklabel.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.olklabel.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olklabel.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olklabel.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.olklabel.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olklabel.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olklabel.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olklabel.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olklabel.TextAlign)

    # Lower case alias for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.olklabel.TextAlign = value

    # Lower case alias for TextAlign setter
    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def UseHeaderColor(self):
        return self.olklabel.UseHeaderColor

    # Lower case alias for UseHeaderColor
    @property
    def useheadercolor(self):
        return self.UseHeaderColor

    @UseHeaderColor.setter
    def UseHeaderColor(self, value):
        self.olklabel.UseHeaderColor = value

    # Lower case alias for UseHeaderColor setter
    @useheadercolor.setter
    def useheadercolor(self, value):
        self.UseHeaderColor = value

    @property
    def WordWrap(self):
        return self.olklabel.WordWrap

    # Lower case alias for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olklabel.WordWrap = value

    # Lower case alias for WordWrap setter
    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class OlkListBox:

    def __init__(self, olklistbox=None):
        self.olklistbox = olklistbox

    @property
    def BackColor(self):
        return self.olklistbox.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olklistbox.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olklistbox.BorderStyle)

    # Lower case alias for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olklistbox.BorderStyle = value

    # Lower case alias for BorderStyle setter
    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def Enabled(self):
        return self.olklistbox.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olklistbox.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.olklistbox.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olklistbox.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olklistbox.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def ListCount(self):
        return self.olklistbox.ListCount

    # Lower case alias for ListCount
    @property
    def listcount(self):
        return self.ListCount

    @property
    def ListIndex(self):
        return self.olklistbox.ListIndex

    # Lower case alias for ListIndex
    @property
    def listindex(self):
        return self.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.olklistbox.ListIndex = value

    # Lower case alias for ListIndex setter
    @listindex.setter
    def listindex(self, value):
        self.ListIndex = value

    @property
    def Locked(self):
        return self.olklistbox.Locked

    # Lower case alias for Locked
    @property
    def locked(self):
        return self.Locked

    @Locked.setter
    def Locked(self, value):
        self.olklistbox.Locked = value

    # Lower case alias for Locked setter
    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MatchEntry(self):
        return olMatchEntry(self.olklistbox.MatchEntry)

    # Lower case alias for MatchEntry
    @property
    def matchentry(self):
        return self.MatchEntry

    @MatchEntry.setter
    def MatchEntry(self, value):
        self.olklistbox.MatchEntry = value

    # Lower case alias for MatchEntry setter
    @matchentry.setter
    def matchentry(self, value):
        self.MatchEntry = value

    @property
    def MouseIcon(self):
        return self.olklistbox.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olklistbox.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olklistbox.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olklistbox.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def MultiSelect(self):
        return OlMultiSelect(self.olklistbox.MultiSelect)

    # Lower case alias for MultiSelect
    @property
    def multiselect(self):
        return self.MultiSelect

    @MultiSelect.setter
    def MultiSelect(self, value):
        self.olklistbox.MultiSelect = value

    # Lower case alias for MultiSelect setter
    @multiselect.setter
    def multiselect(self, value):
        self.MultiSelect = value

    @property
    def Text(self):
        return self.olklistbox.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.olklistbox.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olklistbox.TextAlign)

    # Lower case alias for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.olklistbox.TextAlign = value

    # Lower case alias for TextAlign setter
    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def TopIndex(self):
        return self.olklistbox.TopIndex

    # Lower case alias for TopIndex
    @property
    def topindex(self):
        return self.TopIndex

    @TopIndex.setter
    def TopIndex(self, value):
        self.olklistbox.TopIndex = value

    # Lower case alias for TopIndex setter
    @topindex.setter
    def topindex(self, value):
        self.TopIndex = value

    @property
    def Value(self):
        return self.olklistbox.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.olklistbox.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

    def AddItem(self, ItemText=None, Index=None):
        arguments = com_arguments([ItemText, Index])
        self.olklistbox.AddItem(*arguments)

    def Clear(self):
        self.olklistbox.Clear()

    def Copy(self):
        self.olklistbox.Copy()

    def GetItem(self, Index=None):
        arguments = com_arguments([Index])
        return self.olklistbox.GetItem(*arguments)

    def GetSelected(self, Index=None):
        arguments = com_arguments([Index])
        return self.olklistbox.GetSelected(*arguments)

    def RemoveItem(self, Index=None):
        arguments = com_arguments([Index])
        self.olklistbox.RemoveItem(*arguments)

    def SetItem(self, Index=None, Item=None):
        arguments = com_arguments([Index, Item])
        self.olklistbox.SetItem(*arguments)

    def SetSelected(self, Index=None, Selected=None):
        arguments = com_arguments([Index, Selected])
        self.olklistbox.SetSelected(*arguments)


class OlkOptionButton:

    def __init__(self, olkoptionbutton=None):
        self.olkoptionbutton = olkoptionbutton

    @property
    def Accelerator(self):
        return self.olkoptionbutton.Accelerator

    # Lower case alias for Accelerator
    @property
    def accelerator(self):
        return self.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.olkoptionbutton.Accelerator = value

    # Lower case alias for Accelerator setter
    @accelerator.setter
    def accelerator(self, value):
        self.Accelerator = value

    @property
    def Alignment(self):
        return olAlignment(self.olkoptionbutton.Alignment)

    # Lower case alias for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.olkoptionbutton.Alignment = value

    # Lower case alias for Alignment setter
    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def BackColor(self):
        return self.olkoptionbutton.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olkoptionbutton.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olkoptionbutton.BackStyle)

    # Lower case alias for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.olkoptionbutton.BackStyle = value

    # Lower case alias for BackStyle setter
    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Caption(self):
        return self.olkoptionbutton.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.olkoptionbutton.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def Enabled(self):
        return self.olkoptionbutton.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olkoptionbutton.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.olkoptionbutton.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olkoptionbutton.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olkoptionbutton.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def GroupName(self):
        return self.olkoptionbutton.GroupName

    # Lower case alias for GroupName
    @property
    def groupname(self):
        return self.GroupName

    @GroupName.setter
    def GroupName(self, value):
        self.olkoptionbutton.GroupName = value

    # Lower case alias for GroupName setter
    @groupname.setter
    def groupname(self, value):
        self.GroupName = value

    @property
    def MouseIcon(self):
        return self.olkoptionbutton.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olkoptionbutton.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olkoptionbutton.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olkoptionbutton.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def Value(self):
        return self.olkoptionbutton.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.olkoptionbutton.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

    @property
    def WordWrap(self):
        return self.olkoptionbutton.WordWrap

    # Lower case alias for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olkoptionbutton.WordWrap = value

    # Lower case alias for WordWrap setter
    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class OlkPageControl:

    def __init__(self, olkpagecontrol=None):
        self.olkpagecontrol = olkpagecontrol

    @property
    def Page(self):
        return OlPageType(self.olkpagecontrol.Page)

    # Lower case alias for Page
    @property
    def page(self):
        return self.Page

    @Page.setter
    def Page(self, value):
        self.olkpagecontrol.Page = value

    # Lower case alias for Page setter
    @page.setter
    def page(self, value):
        self.Page = value


class OlkSenderPhoto:

    def __init__(self, olksenderphoto=None):
        self.olksenderphoto = olksenderphoto

    @property
    def Enabled(self):
        return self.olksenderphoto.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olksenderphoto.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def MouseIcon(self):
        return self.olksenderphoto.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olksenderphoto.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olksenderphoto.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olksenderphoto.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def PreferredHeight(self):
        return self.olksenderphoto.PreferredHeight

    # Lower case alias for PreferredHeight
    @property
    def preferredheight(self):
        return self.PreferredHeight

    @property
    def PreferredWidth(self):
        return self.olksenderphoto.PreferredWidth

    # Lower case alias for PreferredWidth
    @property
    def preferredwidth(self):
        return self.PreferredWidth


class OlkTextBox:

    def __init__(self, olktextbox=None):
        self.olktextbox = olktextbox

    @property
    def AutoSize(self):
        return self.olktextbox.AutoSize

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olktextbox.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def AutoTab(self):
        return self.olktextbox.AutoTab

    # Lower case alias for AutoTab
    @property
    def autotab(self):
        return self.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.olktextbox.AutoTab = value

    # Lower case alias for AutoTab setter
    @autotab.setter
    def autotab(self, value):
        self.AutoTab = value

    @property
    def AutoWordSelect(self):
        return self.olktextbox.AutoWordSelect

    # Lower case alias for AutoWordSelect
    @property
    def autowordselect(self):
        return self.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.olktextbox.AutoWordSelect = value

    # Lower case alias for AutoWordSelect setter
    @autowordselect.setter
    def autowordselect(self, value):
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.olktextbox.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olktextbox.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olktextbox.BorderStyle)

    # Lower case alias for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olktextbox.BorderStyle = value

    # Lower case alias for BorderStyle setter
    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def DragBehavior(self):
        return self.olktextbox.DragBehavior

    # Lower case alias for DragBehavior
    @property
    def dragbehavior(self):
        return self.DragBehavior

    @DragBehavior.setter
    def DragBehavior(self, value):
        self.olktextbox.DragBehavior = value

    # Lower case alias for DragBehavior setter
    @dragbehavior.setter
    def dragbehavior(self, value):
        self.DragBehavior = value

    @property
    def Enabled(self):
        return self.olktextbox.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olktextbox.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.olktextbox.EnterFieldBehavior)

    # Lower case alias for EnterFieldBehavior
    @property
    def enterfieldbehavior(self):
        return self.EnterFieldBehavior

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.olktextbox.EnterFieldBehavior = value

    # Lower case alias for EnterFieldBehavior setter
    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        self.EnterFieldBehavior = value

    @property
    def EnterKeyBehavior(self):
        return self.olktextbox.EnterKeyBehavior

    # Lower case alias for EnterKeyBehavior
    @property
    def enterkeybehavior(self):
        return self.EnterKeyBehavior

    @EnterKeyBehavior.setter
    def EnterKeyBehavior(self, value):
        self.olktextbox.EnterKeyBehavior = value

    # Lower case alias for EnterKeyBehavior setter
    @enterkeybehavior.setter
    def enterkeybehavior(self, value):
        self.EnterKeyBehavior = value

    @property
    def Font(self):
        return self.olktextbox.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olktextbox.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olktextbox.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.olktextbox.HideSelection

    # Lower case alias for HideSelection
    @property
    def hideselection(self):
        return self.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.olktextbox.HideSelection = value

    # Lower case alias for HideSelection setter
    @hideselection.setter
    def hideselection(self, value):
        self.HideSelection = value

    @property
    def IntegralHeight(self):
        return self.olktextbox.IntegralHeight

    # Lower case alias for IntegralHeight
    @property
    def integralheight(self):
        return self.IntegralHeight

    @IntegralHeight.setter
    def IntegralHeight(self, value):
        self.olktextbox.IntegralHeight = value

    # Lower case alias for IntegralHeight setter
    @integralheight.setter
    def integralheight(self, value):
        self.IntegralHeight = value

    @property
    def Locked(self):
        return self.olktextbox.Locked

    # Lower case alias for Locked
    @property
    def locked(self):
        return self.Locked

    @Locked.setter
    def Locked(self, value):
        self.olktextbox.Locked = value

    # Lower case alias for Locked setter
    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MaxLength(self):
        return self.olktextbox.MaxLength

    # Lower case alias for MaxLength
    @property
    def maxlength(self):
        return self.MaxLength

    @MaxLength.setter
    def MaxLength(self, value):
        self.olktextbox.MaxLength = value

    # Lower case alias for MaxLength setter
    @maxlength.setter
    def maxlength(self, value):
        self.MaxLength = value

    @property
    def MouseIcon(self):
        return self.olktextbox.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olktextbox.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olktextbox.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olktextbox.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def Multiline(self):
        return self.olktextbox.Multiline

    # Lower case alias for Multiline
    @property
    def multiline(self):
        return self.Multiline

    @Multiline.setter
    def Multiline(self, value):
        self.olktextbox.Multiline = value

    # Lower case alias for Multiline setter
    @multiline.setter
    def multiline(self, value):
        self.Multiline = value

    @property
    def PasswordChar(self):
        return self.olktextbox.PasswordChar

    # Lower case alias for PasswordChar
    @property
    def passwordchar(self):
        return self.PasswordChar

    @PasswordChar.setter
    def PasswordChar(self, value):
        self.olktextbox.PasswordChar = value

    # Lower case alias for PasswordChar setter
    @passwordchar.setter
    def passwordchar(self, value):
        self.PasswordChar = value

    @property
    def Scrollbars(self):
        return olScrollBars(self.olktextbox.Scrollbars)

    # Lower case alias for Scrollbars
    @property
    def scrollbars(self):
        return self.Scrollbars

    @Scrollbars.setter
    def Scrollbars(self, value):
        self.olktextbox.Scrollbars = value

    # Lower case alias for Scrollbars setter
    @scrollbars.setter
    def scrollbars(self, value):
        self.Scrollbars = value

    @property
    def SelectionMargin(self):
        return self.olktextbox.SelectionMargin

    # Lower case alias for SelectionMargin
    @property
    def selectionmargin(self):
        return self.SelectionMargin

    @SelectionMargin.setter
    def SelectionMargin(self, value):
        self.olktextbox.SelectionMargin = value

    # Lower case alias for SelectionMargin setter
    @selectionmargin.setter
    def selectionmargin(self, value):
        self.SelectionMargin = value

    @property
    def SelLength(self):
        return self.olktextbox.SelLength

    # Lower case alias for SelLength
    @property
    def sellength(self):
        return self.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.olktextbox.SelLength = value

    # Lower case alias for SelLength setter
    @sellength.setter
    def sellength(self, value):
        self.SelLength = value

    @property
    def SelStart(self):
        return self.olktextbox.SelStart

    # Lower case alias for SelStart
    @property
    def selstart(self):
        return self.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.olktextbox.SelStart = value

    # Lower case alias for SelStart setter
    @selstart.setter
    def selstart(self, value):
        self.SelStart = value

    @property
    def SelText(self):
        return self.olktextbox.SelText

    # Lower case alias for SelText
    @property
    def seltext(self):
        return self.SelText

    @property
    def TabKeyBehavior(self):
        return self.olktextbox.TabKeyBehavior

    # Lower case alias for TabKeyBehavior
    @property
    def tabkeybehavior(self):
        return self.TabKeyBehavior

    @TabKeyBehavior.setter
    def TabKeyBehavior(self, value):
        self.olktextbox.TabKeyBehavior = value

    # Lower case alias for TabKeyBehavior setter
    @tabkeybehavior.setter
    def tabkeybehavior(self, value):
        self.TabKeyBehavior = value

    @property
    def Text(self):
        return self.olktextbox.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.olktextbox.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olktextbox.TextAlign)

    # Lower case alias for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.olktextbox.TextAlign = value

    # Lower case alias for TextAlign setter
    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def Value(self):
        return self.olktextbox.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.olktextbox.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

    @property
    def WordWrap(self):
        return self.olktextbox.WordWrap

    # Lower case alias for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.olktextbox.WordWrap = value

    # Lower case alias for WordWrap setter
    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value

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

    # Lower case alias for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.olktimecontrol.AutoSize = value

    # Lower case alias for AutoSize setter
    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def AutoWordSelect(self):
        return self.olktimecontrol.AutoWordSelect

    # Lower case alias for AutoWordSelect
    @property
    def autowordselect(self):
        return self.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.olktimecontrol.AutoWordSelect = value

    # Lower case alias for AutoWordSelect setter
    @autowordselect.setter
    def autowordselect(self, value):
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.olktimecontrol.BackColor

    # Lower case alias for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.olktimecontrol.BackColor = value

    # Lower case alias for BackColor setter
    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.olktimecontrol.BackStyle)

    # Lower case alias for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @BackStyle.setter
    def BackStyle(self, value):
        self.olktimecontrol.BackStyle = value

    # Lower case alias for BackStyle setter
    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Enabled(self):
        return self.olktimecontrol.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olktimecontrol.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.olktimecontrol.EnterFieldBehavior)

    # Lower case alias for EnterFieldBehavior
    @property
    def enterfieldbehavior(self):
        return self.EnterFieldBehavior

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.olktimecontrol.EnterFieldBehavior = value

    # Lower case alias for EnterFieldBehavior setter
    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.olktimecontrol.Font

    # Lower case alias for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.olktimecontrol.ForeColor

    # Lower case alias for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.olktimecontrol.ForeColor = value

    # Lower case alias for ForeColor setter
    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.olktimecontrol.HideSelection

    # Lower case alias for HideSelection
    @property
    def hideselection(self):
        return self.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.olktimecontrol.HideSelection = value

    # Lower case alias for HideSelection setter
    @hideselection.setter
    def hideselection(self, value):
        self.HideSelection = value

    @property
    def IntervalTime(self):
        return self.olktimecontrol.IntervalTime

    # Lower case alias for IntervalTime
    @property
    def intervaltime(self):
        return self.IntervalTime

    @IntervalTime.setter
    def IntervalTime(self, value):
        self.olktimecontrol.IntervalTime = value

    # Lower case alias for IntervalTime setter
    @intervaltime.setter
    def intervaltime(self, value):
        self.IntervalTime = value

    @property
    def Locked(self):
        return self.olktimecontrol.Locked

    # Lower case alias for Locked
    @property
    def locked(self):
        return self.Locked

    @Locked.setter
    def Locked(self, value):
        self.olktimecontrol.Locked = value

    # Lower case alias for Locked setter
    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.olktimecontrol.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olktimecontrol.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olktimecontrol.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olktimecontrol.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def ReferenceTime(self):
        return self.olktimecontrol.ReferenceTime

    # Lower case alias for ReferenceTime
    @property
    def referencetime(self):
        return self.ReferenceTime

    @ReferenceTime.setter
    def ReferenceTime(self, value):
        self.olktimecontrol.ReferenceTime = value

    # Lower case alias for ReferenceTime setter
    @referencetime.setter
    def referencetime(self, value):
        self.ReferenceTime = value

    @property
    def Style(self):
        return OlTimeStyle(self.olktimecontrol.Style)

    # Lower case alias for Style
    @property
    def style(self):
        return self.Style

    @Style.setter
    def Style(self, value):
        self.olktimecontrol.Style = value

    # Lower case alias for Style setter
    @style.setter
    def style(self, value):
        self.Style = value

    @property
    def Text(self):
        return self.olktimecontrol.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.olktimecontrol.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.olktimecontrol.TextAlign)

    # Lower case alias for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @TextAlign.setter
    def TextAlign(self, value):
        self.olktimecontrol.TextAlign = value

    # Lower case alias for TextAlign setter
    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def Time(self):
        return self.olktimecontrol.Time

    # Lower case alias for Time
    @property
    def time(self):
        return self.Time

    @Time.setter
    def Time(self, value):
        self.olktimecontrol.Time = value

    # Lower case alias for Time setter
    @time.setter
    def time(self, value):
        self.Time = value

    @property
    def Value(self):
        return self.olktimecontrol.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.olktimecontrol.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

    def DropDown(self):
        self.olktimecontrol.DropDown()


class OlkTimeZoneControl:

    def __init__(self, olktimezonecontrol=None):
        self.olktimezonecontrol = olktimezonecontrol

    @property
    def AppointmentTimeField(self):
        return OlAppointmentTimeField(self.olktimezonecontrol.AppointmentTimeField)

    # Lower case alias for AppointmentTimeField
    @property
    def appointmenttimefield(self):
        return self.AppointmentTimeField

    @AppointmentTimeField.setter
    def AppointmentTimeField(self, value):
        self.olktimezonecontrol.AppointmentTimeField = value

    # Lower case alias for AppointmentTimeField setter
    @appointmenttimefield.setter
    def appointmenttimefield(self, value):
        self.AppointmentTimeField = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.olktimezonecontrol.BorderStyle)

    # Lower case alias for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.olktimezonecontrol.BorderStyle = value

    # Lower case alias for BorderStyle setter
    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def Enabled(self):
        return self.olktimezonecontrol.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.olktimezonecontrol.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Locked(self):
        return self.olktimezonecontrol.Locked

    # Lower case alias for Locked
    @property
    def locked(self):
        return self.Locked

    @Locked.setter
    def Locked(self, value):
        self.olktimezonecontrol.Locked = value

    # Lower case alias for Locked setter
    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.olktimezonecontrol.MouseIcon

    # Lower case alias for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.olktimezonecontrol.MouseIcon = value

    # Lower case alias for MouseIcon setter
    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.olktimezonecontrol.MousePointer)

    # Lower case alias for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @MousePointer.setter
    def MousePointer(self, value):
        self.olktimezonecontrol.MousePointer = value

    # Lower case alias for MousePointer setter
    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def SelectedTimeZoneIndex(self):
        return Application.TimeZones(self.olktimezonecontrol.SelectedTimeZoneIndex)

    # Lower case alias for SelectedTimeZoneIndex
    @property
    def selectedtimezoneindex(self):
        return self.SelectedTimeZoneIndex

    @SelectedTimeZoneIndex.setter
    def SelectedTimeZoneIndex(self, value):
        self.olktimezonecontrol.SelectedTimeZoneIndex = value

    # Lower case alias for SelectedTimeZoneIndex setter
    @selectedtimezoneindex.setter
    def selectedtimezoneindex(self, value):
        self.SelectedTimeZoneIndex = value

    @property
    def Value(self):
        return self.olktimezonecontrol.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.olktimezonecontrol.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

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

    # Lower case alias for IsDescending
    @property
    def isdescending(self):
        return self.IsDescending

    @IsDescending.setter
    def IsDescending(self, value):
        self.orderfield.IsDescending = value

    # Lower case alias for IsDescending setter
    @isdescending.setter
    def isdescending(self, value):
        self.IsDescending = value

    @property
    def Parent(self):
        return self.orderfield.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.orderfield.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ViewXMLSchemaName(self):
        return OrderField(self.orderfield.ViewXMLSchemaName)

    # Lower case alias for ViewXMLSchemaName
    @property
    def viewxmlschemaname(self):
        return self.ViewXMLSchemaName


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.orderfields.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.orderfields.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, PropertyName=None, IsDescending=None):
        arguments = com_arguments([PropertyName, IsDescending])
        return self.orderfields.Add(*arguments)

    def Insert(self, PropertyName=None, Index=None, IsDescending=None):
        arguments = com_arguments([PropertyName, Index, IsDescending])
        return self.orderfields.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.orderfields.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.orderfields.Remove(*arguments)

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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.outlookbargroup.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.outlookbargroup.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbargroup.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Shortcuts(self):
        return OutlookBarShortcuts(self.outlookbargroup.Shortcuts)

    # Lower case alias for Shortcuts
    @property
    def shortcuts(self):
        return self.Shortcuts

    @property
    def ViewType(self):
        return OlOutlookBarViewType(self.outlookbargroup.ViewType)

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @ViewType.setter
    def ViewType(self, value):
        self.outlookbargroup.ViewType = value

    # Lower case alias for ViewType setter
    @viewtype.setter
    def viewtype(self, value):
        self.ViewType = value


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.outlookbargroups.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbargroups.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Index=None):
        arguments = com_arguments([Name, Index])
        return OutlookBarGroup(self.outlookbargroups.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.outlookbargroups.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.outlookbargroups.Remove(*arguments)


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

    # Lower case alias for Contents
    @property
    def contents(self):
        return self.Contents

    @property
    def Name(self):
        return self.outlookbarpane.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.outlookbarpane.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbarpane.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return self.outlookbarpane.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.outlookbarpane.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.outlookbarshortcut.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.outlookbarshortcut.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbarshortcut.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Target(self):
        return self.outlookbarshortcut.Target

    # Lower case alias for Target
    @property
    def target(self):
        return self.Target

    def SetIcon(self, Icon=None):
        arguments = com_arguments([Icon])
        self.outlookbarshortcut.SetIcon(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.outlookbarshortcuts.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbarshortcuts.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Target=None, Name=None, Index=None):
        arguments = com_arguments([Target, Name, Index])
        return OutlookBarShortcut(self.outlookbarshortcuts.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.outlookbarshortcuts.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.outlookbarshortcuts.Remove(*arguments)


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

    # Lower case alias for Groups
    @property
    def groups(self):
        return self.Groups

    @property
    def Parent(self):
        return self.outlookbarstorage.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.outlookbarstorage.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.pages.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.pages.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self):
        return Page(self.pages.Add())

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.pages.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.pages.Remove(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.panes.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.panes.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.panes.Item(*arguments)


class PlaySoundRuleAction:

    def __init__(self, playsoundruleaction=None):
        self.playsoundruleaction = playsoundruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.playsoundruleaction.ActionType)

    # Lower case alias for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.playsoundruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.playsoundruleaction.Class)

    @property
    def Enabled(self):
        return self.playsoundruleaction.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.playsoundruleaction.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def FilePath(self):
        return self.playsoundruleaction.FilePath

    # Lower case alias for FilePath
    @property
    def filepath(self):
        return self.FilePath

    @FilePath.setter
    def FilePath(self, value):
        self.playsoundruleaction.FilePath = value

    # Lower case alias for FilePath setter
    @filepath.setter
    def filepath(self, value):
        self.FilePath = value

    @property
    def Parent(self):
        return self.playsoundruleaction.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.playsoundruleaction.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class PostItem:

    def __init__(self, postitem=None):
        self.postitem = postitem

    @property
    def Actions(self):
        return Actions(self.postitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.postitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.postitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.postitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.postitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.postitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.postitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.postitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.postitem.BodyFormat)

    # Lower case alias for BodyFormat
    @property
    def bodyformat(self):
        return self.BodyFormat

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.postitem.BodyFormat = value

    # Lower case alias for BodyFormat setter
    @bodyformat.setter
    def bodyformat(self, value):
        self.BodyFormat = value

    @property
    def Categories(self):
        return self.postitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.postitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.postitem.Class)

    @property
    def Companies(self):
        return self.postitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.postitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.postitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.postitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.postitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.postitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.postitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.postitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.postitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.postitem.ExpiryTime

    # Lower case alias for ExpiryTime
    @property
    def expirytime(self):
        return self.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.postitem.ExpiryTime = value

    # Lower case alias for ExpiryTime setter
    @expirytime.setter
    def expirytime(self, value):
        self.ExpiryTime = value

    @property
    def FormDescription(self):
        return FormDescription(self.postitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.postitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def HTMLBody(self):
        return self.postitem.HTMLBody

    # Lower case alias for HTMLBody
    @property
    def htmlbody(self):
        return self.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.postitem.HTMLBody = value

    # Lower case alias for HTMLBody setter
    @htmlbody.setter
    def htmlbody(self, value):
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.postitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.postitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.postitem.InternetCodepage

    # Lower case alias for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.postitem.InternetCodepage = value

    # Lower case alias for InternetCodepage setter
    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.postitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return PostItem(self.postitem.IsMarkedAsTask)

    # Lower case alias for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.postitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.postitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.postitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.postitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.postitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.postitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.postitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.postitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.postitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.postitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.postitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.postitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.postitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.postitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReceivedTime(self):
        return self.postitem.ReceivedTime

    # Lower case alias for ReceivedTime
    @property
    def receivedtime(self):
        return self.ReceivedTime

    @property
    def ReminderOverrideDefault(self):
        return self.postitem.ReminderOverrideDefault

    # Lower case alias for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.postitem.ReminderOverrideDefault = value

    # Lower case alias for ReminderOverrideDefault setter
    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.postitem.ReminderPlaySound

    # Lower case alias for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.postitem.ReminderPlaySound = value

    # Lower case alias for ReminderPlaySound setter
    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.postitem.ReminderSet

    # Lower case alias for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.postitem.ReminderSet = value

    # Lower case alias for ReminderSet setter
    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.postitem.ReminderSoundFile

    # Lower case alias for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.postitem.ReminderSoundFile = value

    # Lower case alias for ReminderSoundFile setter
    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.postitem.ReminderTime

    # Lower case alias for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.postitem.ReminderTime = value

    # Lower case alias for ReminderTime setter
    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.postitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.postitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.postitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SenderEmailAddress(self):
        return self.postitem.SenderEmailAddress

    # Lower case alias for SenderEmailAddress
    @property
    def senderemailaddress(self):
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.postitem.SenderEmailType

    # Lower case alias for SenderEmailType
    @property
    def senderemailtype(self):
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.postitem.SenderName

    # Lower case alias for SenderName
    @property
    def sendername(self):
        return self.SenderName

    @property
    def Sensitivity(self):
        return OlSensitivity(self.postitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.postitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def SentOn(self):
        return self.postitem.SentOn

    # Lower case alias for SentOn
    @property
    def senton(self):
        return self.SentOn

    @property
    def Session(self):
        return NameSpace(self.postitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.postitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.postitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.postitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def TaskCompletedDate(self):
        return PostItem(self.postitem.TaskCompletedDate)

    # Lower case alias for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.postitem.TaskCompletedDate = value

    # Lower case alias for TaskCompletedDate setter
    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return PostItem(self.postitem.TaskDueDate)

    # Lower case alias for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.postitem.TaskDueDate = value

    # Lower case alias for TaskDueDate setter
    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return PostItem(self.postitem.TaskStartDate)

    # Lower case alias for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.postitem.TaskStartDate = value

    # Lower case alias for TaskStartDate setter
    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return PostItem(self.postitem.TaskSubject)

    # Lower case alias for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.postitem.TaskSubject = value

    # Lower case alias for TaskSubject setter
    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def ToDoTaskOrdinal(self):
        return PostItem(self.postitem.ToDoTaskOrdinal)

    # Lower case alias for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.postitem.ToDoTaskOrdinal = value

    # Lower case alias for ToDoTaskOrdinal setter
    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.postitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.postitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.postitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def ClearConversationIndex(self):
        self.postitem.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.postitem.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.postitem.Close(*arguments)

    def Copy(self):
        self.postitem.Copy()

    def Delete(self):
        self.postitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.postitem.Display(*arguments)

    def Forward(self):
        return self.postitem.Forward()

    def GetConversation(self):
        return self.postitem.GetConversation()

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.postitem.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.postitem.Move(*arguments)

    def Post(self):
        self.postitem.Post()

    def PrintOut(self):
        self.postitem.PrintOut()

    def Reply(self):
        return MailItem(self.postitem.Reply())

    def Save(self):
        self.postitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.postitem.SaveAs(*arguments)

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

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.propertyaccessor.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def BinaryToString(self, Value=None):
        arguments = com_arguments([Value])
        return self.propertyaccessor.BinaryToString(*arguments)

    def DeleteProperties(self, SchemaNames=None):
        arguments = com_arguments([SchemaNames])
        return Err(self.propertyaccessor.DeleteProperties(*arguments))

    def DeleteProperty(self, SchemaName=None):
        arguments = com_arguments([SchemaName])
        self.propertyaccessor.DeleteProperty(*arguments)

    def GetProperties(self, SchemaNames=None):
        arguments = com_arguments([SchemaNames])
        return self.propertyaccessor.GetProperties(*arguments)

    def GetProperty(self, SchemaName=None):
        arguments = com_arguments([SchemaName])
        return self.propertyaccessor.GetProperty(*arguments)

    def LocalTimeToUTC(self, Value=None):
        arguments = com_arguments([Value])
        return self.propertyaccessor.LocalTimeToUTC(*arguments)

    def SetProperties(self, SchemaNames=None, Values=None):
        arguments = com_arguments([SchemaNames, Values])
        return self.propertyaccessor.SetProperties(*arguments)

    def SetProperty(self, SchemaName=None, Value=None):
        arguments = com_arguments([SchemaName, Value])
        self.propertyaccessor.SetProperty(*arguments)

    def StringToBinary(self, Value=None):
        arguments = com_arguments([Value])
        return self.propertyaccessor.StringToBinary(*arguments)

    def UTCToLocalTime(self, Value=None):
        arguments = com_arguments([Value])
        return self.propertyaccessor.UTCToLocalTime(*arguments)


class PropertyPage:

    def __init__(self, propertypage=None):
        self.propertypage = propertypage

    def Dirty(self, Dirty=None):
        arguments = com_arguments([Dirty])
        if callable(self.propertypage.Dirty):
            return self.propertypage.Dirty(*arguments)
        else:
            return self.propertypage.GetDirty(*arguments)

    # Lower case alias for Dirty
    def dirty(self, Dirty=None):
        arguments = [Dirty]
        return self.Dirty(*arguments)

    def Apply(self):
        return self.propertypage.Apply()

    def GetPageInfo(self, HelpFile=None, HelpContext=None):
        arguments = com_arguments([HelpFile, HelpContext])
        return self.propertypage.GetPageInfo(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.propertypages.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.propertypages.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Page=None, Title=None):
        arguments = com_arguments([Page, Title])
        self.propertypages.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.propertypages.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.propertypages.Remove(*arguments)


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

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.propertypagesite.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def OnStatusChange(self):
        self.propertypagesite.OnStatusChange()


class Recipient:

    def __init__(self, recipient=None):
        self.recipient = recipient

    @property
    def Address(self):
        return self.recipient.Address

    # Lower case alias for Address
    @property
    def address(self):
        return self.Address

    @property
    def AddressEntry(self):
        return AddressEntry(self.recipient.AddressEntry)

    # Lower case alias for AddressEntry
    @property
    def addressentry(self):
        return self.AddressEntry

    @AddressEntry.setter
    def AddressEntry(self, value):
        self.recipient.AddressEntry = value

    # Lower case alias for AddressEntry setter
    @addressentry.setter
    def addressentry(self, value):
        self.AddressEntry = value

    @property
    def Application(self):
        return Application(self.recipient.Application)

    @property
    def AutoResponse(self):
        return Recipient(self.recipient.AutoResponse)

    # Lower case alias for AutoResponse
    @property
    def autoresponse(self):
        return self.AutoResponse

    @AutoResponse.setter
    def AutoResponse(self, value):
        self.recipient.AutoResponse = value

    # Lower case alias for AutoResponse setter
    @autoresponse.setter
    def autoresponse(self, value):
        self.AutoResponse = value

    @property
    def Class(self):
        return OlObjectClass(self.recipient.Class)

    @property
    def DisplayType(self):
        return OlDisplayType(self.recipient.DisplayType)

    # Lower case alias for DisplayType
    @property
    def displaytype(self):
        return self.DisplayType

    @property
    def EntryID(self):
        return self.recipient.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def Index(self):
        return self.recipient.Index

    # Lower case alias for Index
    @property
    def index(self):
        return self.Index

    @property
    def MeetingResponseStatus(self):
        return OlResponseStatus(self.recipient.MeetingResponseStatus)

    # Lower case alias for MeetingResponseStatus
    @property
    def meetingresponsestatus(self):
        return self.MeetingResponseStatus

    @property
    def Name(self):
        return self.recipient.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.recipient.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.recipient.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Resolved(self):
        return self.recipient.Resolved

    # Lower case alias for Resolved
    @property
    def resolved(self):
        return self.Resolved

    @property
    def Sendable(self):
        return Recipient(self.recipient.Sendable)

    # Lower case alias for Sendable
    @property
    def sendable(self):
        return self.Sendable

    @Sendable.setter
    def Sendable(self, value):
        self.recipient.Sendable = value

    # Lower case alias for Sendable setter
    @sendable.setter
    def sendable(self, value):
        self.Sendable = value

    @property
    def Session(self):
        return NameSpace(self.recipient.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def TrackingStatus(self):
        return OlTrackingStatus(self.recipient.TrackingStatus)

    # Lower case alias for TrackingStatus
    @property
    def trackingstatus(self):
        return self.TrackingStatus

    @TrackingStatus.setter
    def TrackingStatus(self, value):
        self.recipient.TrackingStatus = value

    # Lower case alias for TrackingStatus setter
    @trackingstatus.setter
    def trackingstatus(self, value):
        self.TrackingStatus = value

    @property
    def TrackingStatusTime(self):
        return self.recipient.TrackingStatusTime

    # Lower case alias for TrackingStatusTime
    @property
    def trackingstatustime(self):
        return self.TrackingStatusTime

    @TrackingStatusTime.setter
    def TrackingStatusTime(self, value):
        self.recipient.TrackingStatusTime = value

    # Lower case alias for TrackingStatusTime setter
    @trackingstatustime.setter
    def trackingstatustime(self, value):
        self.TrackingStatusTime = value

    @property
    def Type(self):
        return self.recipient.Type

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.recipient.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    def Delete(self):
        self.recipient.Delete()

    def FreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([Start, MinPerChar, CompleteFormat])
        return self.recipient.FreeBusy(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.recipients.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.recipients.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([Name])
        return Recipient(self.recipients.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.recipients.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.recipients.Remove(*arguments)

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

    # Lower case alias for DayOfMonth
    @property
    def dayofmonth(self):
        return self.DayOfMonth

    @DayOfMonth.setter
    def DayOfMonth(self, value):
        self.recurrencepattern.DayOfMonth = value

    # Lower case alias for DayOfMonth setter
    @dayofmonth.setter
    def dayofmonth(self, value):
        self.DayOfMonth = value

    @property
    def DayOfWeekMask(self):
        return OlDaysOfWeek(self.recurrencepattern.DayOfWeekMask)

    # Lower case alias for DayOfWeekMask
    @property
    def dayofweekmask(self):
        return self.DayOfWeekMask

    @DayOfWeekMask.setter
    def DayOfWeekMask(self, value):
        self.recurrencepattern.DayOfWeekMask = value

    # Lower case alias for DayOfWeekMask setter
    @dayofweekmask.setter
    def dayofweekmask(self, value):
        self.DayOfWeekMask = value

    @property
    def Duration(self):
        return RecurrencePattern(self.recurrencepattern.Duration)

    # Lower case alias for Duration
    @property
    def duration(self):
        return self.Duration

    @Duration.setter
    def Duration(self, value):
        self.recurrencepattern.Duration = value

    # Lower case alias for Duration setter
    @duration.setter
    def duration(self, value):
        self.Duration = value

    @property
    def EndTime(self):
        return self.recurrencepattern.EndTime

    # Lower case alias for EndTime
    @property
    def endtime(self):
        return self.EndTime

    @EndTime.setter
    def EndTime(self, value):
        self.recurrencepattern.EndTime = value

    # Lower case alias for EndTime setter
    @endtime.setter
    def endtime(self, value):
        self.EndTime = value

    @property
    def Exceptions(self):
        return Exceptions(self.recurrencepattern.Exceptions)

    # Lower case alias for Exceptions
    @property
    def exceptions(self):
        return self.Exceptions

    @property
    def Instance(self):
        return self.recurrencepattern.Instance

    # Lower case alias for Instance
    @property
    def instance(self):
        return self.Instance

    @Instance.setter
    def Instance(self, value):
        self.recurrencepattern.Instance = value

    # Lower case alias for Instance setter
    @instance.setter
    def instance(self, value):
        self.Instance = value

    @property
    def Interval(self):
        return self.recurrencepattern.Interval

    # Lower case alias for Interval
    @property
    def interval(self):
        return self.Interval

    @Interval.setter
    def Interval(self, value):
        self.recurrencepattern.Interval = value

    # Lower case alias for Interval setter
    @interval.setter
    def interval(self, value):
        self.Interval = value

    @property
    def MonthOfYear(self):
        return self.recurrencepattern.MonthOfYear

    # Lower case alias for MonthOfYear
    @property
    def monthofyear(self):
        return self.MonthOfYear

    @MonthOfYear.setter
    def MonthOfYear(self, value):
        self.recurrencepattern.MonthOfYear = value

    # Lower case alias for MonthOfYear setter
    @monthofyear.setter
    def monthofyear(self, value):
        self.MonthOfYear = value

    @property
    def NoEndDate(self):
        return self.recurrencepattern.NoEndDate

    # Lower case alias for NoEndDate
    @property
    def noenddate(self):
        return self.NoEndDate

    @NoEndDate.setter
    def NoEndDate(self, value):
        self.recurrencepattern.NoEndDate = value

    # Lower case alias for NoEndDate setter
    @noenddate.setter
    def noenddate(self, value):
        self.NoEndDate = value

    @property
    def Occurrences(self):
        return self.recurrencepattern.Occurrences

    # Lower case alias for Occurrences
    @property
    def occurrences(self):
        return self.Occurrences

    @Occurrences.setter
    def Occurrences(self, value):
        self.recurrencepattern.Occurrences = value

    # Lower case alias for Occurrences setter
    @occurrences.setter
    def occurrences(self, value):
        self.Occurrences = value

    @property
    def Parent(self):
        return self.recurrencepattern.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PatternEndDate(self):
        return self.recurrencepattern.PatternEndDate

    # Lower case alias for PatternEndDate
    @property
    def patternenddate(self):
        return self.PatternEndDate

    @PatternEndDate.setter
    def PatternEndDate(self, value):
        self.recurrencepattern.PatternEndDate = value

    # Lower case alias for PatternEndDate setter
    @patternenddate.setter
    def patternenddate(self, value):
        self.PatternEndDate = value

    @property
    def PatternStartDate(self):
        return self.recurrencepattern.PatternStartDate

    # Lower case alias for PatternStartDate
    @property
    def patternstartdate(self):
        return self.PatternStartDate

    @PatternStartDate.setter
    def PatternStartDate(self, value):
        self.recurrencepattern.PatternStartDate = value

    # Lower case alias for PatternStartDate setter
    @patternstartdate.setter
    def patternstartdate(self, value):
        self.PatternStartDate = value

    @property
    def RecurrenceType(self):
        return OlRecurrenceType(self.recurrencepattern.RecurrenceType)

    # Lower case alias for RecurrenceType
    @property
    def recurrencetype(self):
        return self.RecurrenceType

    @RecurrenceType.setter
    def RecurrenceType(self, value):
        self.recurrencepattern.RecurrenceType = value

    # Lower case alias for RecurrenceType setter
    @recurrencetype.setter
    def recurrencetype(self, value):
        self.RecurrenceType = value

    @property
    def Regenerate(self):
        return self.recurrencepattern.Regenerate

    # Lower case alias for Regenerate
    @property
    def regenerate(self):
        return self.Regenerate

    @Regenerate.setter
    def Regenerate(self, value):
        self.recurrencepattern.Regenerate = value

    # Lower case alias for Regenerate setter
    @regenerate.setter
    def regenerate(self, value):
        self.Regenerate = value

    @property
    def Session(self):
        return NameSpace(self.recurrencepattern.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def StartTime(self):
        return self.recurrencepattern.StartTime

    # Lower case alias for StartTime
    @property
    def starttime(self):
        return self.StartTime

    @StartTime.setter
    def StartTime(self, value):
        self.recurrencepattern.StartTime = value

    # Lower case alias for StartTime setter
    @starttime.setter
    def starttime(self, value):
        self.StartTime = value

    def GetOccurrence(self, StartDate=None):
        arguments = com_arguments([StartDate])
        return self.recurrencepattern.GetOccurrence(*arguments)


class Reminder:

    def __init__(self, reminder=None):
        self.reminder = reminder

    @property
    def Application(self):
        return Application(self.reminder.Application)

    @property
    def Caption(self):
        return self.reminder.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Class(self):
        return OlObjectClass(self.reminder.Class)

    @property
    def IsVisible(self):
        return self.reminder.IsVisible

    # Lower case alias for IsVisible
    @property
    def isvisible(self):
        return self.IsVisible

    @property
    def Item(self):
        return self.reminder.Item

    # Lower case alias for Item
    @property
    def item(self):
        return self.Item

    @property
    def NextReminderDate(self):
        return self.reminder.NextReminderDate

    # Lower case alias for NextReminderDate
    @property
    def nextreminderdate(self):
        return self.NextReminderDate

    @property
    def OriginalReminderDate(self):
        return self.reminder.OriginalReminderDate

    # Lower case alias for OriginalReminderDate
    @property
    def originalreminderdate(self):
        return self.OriginalReminderDate

    @property
    def Parent(self):
        return self.reminder.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.reminder.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Dismiss(self):
        self.reminder.Dismiss()

    def Snooze(self, SnoozeTime=None):
        arguments = com_arguments([SnoozeTime])
        self.reminder.Snooze(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.reminders.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.reminders.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.reminders.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.reminders.Remove(*arguments)


class RemoteItem:

    def __init__(self, remoteitem=None):
        self.remoteitem = remoteitem

    @property
    def Actions(self):
        return Actions(self.remoteitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.remoteitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.remoteitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.remoteitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.remoteitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.remoteitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.remoteitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.remoteitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.remoteitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.remoteitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.remoteitem.Class)

    @property
    def Companies(self):
        return self.remoteitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.remoteitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.remoteitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.remoteitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.remoteitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.remoteitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.remoteitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.remoteitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.remoteitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.remoteitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.remoteitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def HasAttachment(self):
        return self.remoteitem.HasAttachment

    # Lower case alias for HasAttachment
    @property
    def hasattachment(self):
        return self.HasAttachment

    @property
    def Importance(self):
        return OlImportance(self.remoteitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.remoteitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.remoteitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.remoteitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.remoteitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.remoteitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.remoteitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.remoteitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.remoteitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.remoteitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.remoteitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.remoteitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.remoteitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.remoteitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.remoteitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.remoteitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.remoteitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RemoteMessageClass(self):
        return self.remoteitem.RemoteMessageClass

    # Lower case alias for RemoteMessageClass
    @property
    def remotemessageclass(self):
        return self.RemoteMessageClass

    @property
    def Saved(self):
        return self.remoteitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.remoteitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.remoteitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.remoteitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.remoteitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.remoteitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.remoteitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def TransferSize(self):
        return self.remoteitem.TransferSize

    # Lower case alias for TransferSize
    @property
    def transfersize(self):
        return self.TransferSize

    @property
    def TransferTime(self):
        return self.remoteitem.TransferTime

    # Lower case alias for TransferTime
    @property
    def transfertime(self):
        return self.TransferTime

    @property
    def UnRead(self):
        return self.remoteitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.remoteitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.remoteitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.remoteitem.Close(*arguments)

    def Copy(self):
        self.remoteitem.Copy()

    def Delete(self):
        self.remoteitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.remoteitem.Display(*arguments)

    def GetConversation(self):
        return self.remoteitem.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.remoteitem.Move(*arguments)

    def PrintOut(self):
        self.remoteitem.PrintOut()

    def Save(self):
        self.remoteitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.remoteitem.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.remoteitem.ShowCategoriesDialog()


class ReportItem:

    def __init__(self, reportitem=None):
        self.reportitem = reportitem

    @property
    def Actions(self):
        return Actions(self.reportitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.reportitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.reportitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.reportitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.reportitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.reportitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.reportitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.reportitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.reportitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.reportitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.reportitem.Class)

    @property
    def Companies(self):
        return self.reportitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.reportitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.reportitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.reportitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.reportitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.reportitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.reportitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.reportitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.reportitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.reportitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.reportitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.reportitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.reportitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.reportitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.reportitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.reportitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.reportitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.reportitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.reportitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.reportitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.reportitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.reportitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.reportitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.reportitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.reportitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.reportitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.reportitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.reportitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RetentionExpirationDate(self):
        return ReportItem(self.reportitem.RetentionExpirationDate)

    # Lower case alias for RetentionExpirationDate
    @property
    def retentionexpirationdate(self):
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.reportitem.RetentionPolicyName

    # Lower case alias for RetentionPolicyName
    @property
    def retentionpolicyname(self):
        return self.RetentionPolicyName

    @property
    def Saved(self):
        return self.reportitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.reportitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.reportitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.reportitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.reportitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.reportitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.reportitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.reportitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.reportitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.reportitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.reportitem.Close(*arguments)

    def Copy(self):
        self.reportitem.Copy()

    def Delete(self):
        self.reportitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.reportitem.Display(*arguments)

    def GetConversation(self):
        return self.reportitem.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.reportitem.Move(*arguments)

    def PrintOut(self):
        self.reportitem.PrintOut()

    def Save(self):
        self.reportitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.reportitem.SaveAs(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def DefaultItemType(self):
        return OlItemType(self.results.DefaultItemType)

    # Lower case alias for DefaultItemType
    @property
    def defaultitemtype(self):
        return self.DefaultItemType

    @DefaultItemType.setter
    def DefaultItemType(self, value):
        self.results.DefaultItemType = value

    # Lower case alias for DefaultItemType setter
    @defaultitemtype.setter
    def defaultitemtype(self, value):
        self.DefaultItemType = value

    @property
    def Parent(self):
        return self.results.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.results.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def GetFirst(self):
        return self.results.GetFirst()

    def GetLast(self):
        return self.results.GetLast()

    def GetNext(self):
        return self.results.GetNext()

    def GetPrevious(self):
        return self.results.GetPrevious()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.results.Item(*arguments)

    def ResetColumns(self):
        self.results.ResetColumns()

    def SetColumns(self, Columns=None):
        arguments = com_arguments([Columns])
        self.results.SetColumns(*arguments)

    def Sort(self, Property=None, Descending=None):
        arguments = com_arguments([Property, Descending])
        self.results.Sort(*arguments)


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

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.row.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def BinaryToString(self, Index=None):
        arguments = com_arguments([Index])
        return self.row.BinaryToString(*arguments)

    def GetValues(self):
        return self.row.GetValues()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.row.Item(*arguments)

    def LocalTimeToUTC(self, Index=None):
        arguments = com_arguments([Index])
        return self.row.LocalTimeToUTC(*arguments)

    def UTCToLocalTime(self, Index=None):
        arguments = com_arguments([Index])
        return self.row.UTCToLocalTime(*arguments)


class Rule:

    def __init__(self, rule=None):
        self.rule = rule

    @property
    def Actions(self):
        return RuleActions(self.rule.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.rule.Application)

    @property
    def Class(self):
        return OlObjectClass(self.rule.Class)

    @property
    def Conditions(self):
        return RuleConditions(self.rule.Conditions)

    # Lower case alias for Conditions
    @property
    def conditions(self):
        return self.Conditions

    @property
    def Enabled(self):
        return self.rule.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.rule.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Exceptions(self):
        return RuleConditions(self.rule.Exceptions)

    # Lower case alias for Exceptions
    @property
    def exceptions(self):
        return self.Exceptions

    @property
    def ExecutionOrder(self):
        return Rules(self.rule.ExecutionOrder)

    # Lower case alias for ExecutionOrder
    @property
    def executionorder(self):
        return self.ExecutionOrder

    @ExecutionOrder.setter
    def ExecutionOrder(self, value):
        self.rule.ExecutionOrder = value

    # Lower case alias for ExecutionOrder setter
    @executionorder.setter
    def executionorder(self, value):
        self.ExecutionOrder = value

    @property
    def IsLocalRule(self):
        return self.rule.IsLocalRule

    # Lower case alias for IsLocalRule
    @property
    def islocalrule(self):
        return self.IsLocalRule

    @property
    def Name(self):
        return self.rule.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.rule.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.rule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RuleType(self):
        return OlRuleType(self.rule.RuleType)

    # Lower case alias for RuleType
    @property
    def ruletype(self):
        return self.RuleType

    @property
    def Session(self):
        return NameSpace(self.rule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Execute(self, ShowProgress=None, Folder=None, IncludeSubfolders=None, RuleExecuteOption=None):
        arguments = com_arguments([ShowProgress, Folder, IncludeSubfolders, RuleExecuteOption])
        self.rule.Execute(*arguments)


class RuleAction:

    def __init__(self, ruleaction=None):
        self.ruleaction = ruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.ruleaction.ActionType)

    # Lower case alias for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.ruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.ruleaction.Class)

    @property
    def Enabled(self):
        return RuleAction(self.ruleaction.Enabled)

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.ruleaction.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.ruleaction.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.ruleaction.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class RuleActions:

    def __init__(self, ruleactions=None):
        self.ruleactions = ruleactions

    @property
    def Application(self):
        return Application(self.ruleactions.Application)

    @property
    def AssignToCategory(self):
        return AssignToCategoryRuleAction(self.ruleactions.AssignToCategory)

    # Lower case alias for AssignToCategory
    @property
    def assigntocategory(self):
        return self.AssignToCategory

    @property
    def CC(self):
        return SendRuleAction(self.ruleactions.CC)

    # Lower case alias for CC
    @property
    def cc(self):
        return self.CC

    @property
    def Class(self):
        return OlObjectClass(self.ruleactions.Class)

    @property
    def ClearCategories(self):
        return RuleAction(self.ruleactions.ClearCategories)

    # Lower case alias for ClearCategories
    @property
    def clearcategories(self):
        return self.ClearCategories

    @property
    def CopyToFolder(self):
        return MoveOrCopyRuleAction(self.ruleactions.CopyToFolder)

    # Lower case alias for CopyToFolder
    @property
    def copytofolder(self):
        return self.CopyToFolder

    @property
    def Count(self):
        return self.ruleactions.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Delete(self):
        return RuleAction(self.ruleactions.Delete)

    # Lower case alias for Delete
    @property
    def delete(self):
        return self.Delete

    @property
    def DeletePermanently(self):
        return RuleAction(self.ruleactions.DeletePermanently)

    # Lower case alias for DeletePermanently
    @property
    def deletepermanently(self):
        return self.DeletePermanently

    @property
    def DesktopAlert(self):
        return RuleAction(self.ruleactions.DesktopAlert)

    # Lower case alias for DesktopAlert
    @property
    def desktopalert(self):
        return self.DesktopAlert

    @property
    def Forward(self):
        return SendRuleAction(self.ruleactions.Forward)

    # Lower case alias for Forward
    @property
    def forward(self):
        return self.Forward

    @property
    def ForwardAsAttachment(self):
        return SendRuleAction(self.ruleactions.ForwardAsAttachment)

    # Lower case alias for ForwardAsAttachment
    @property
    def forwardasattachment(self):
        return self.ForwardAsAttachment

    @property
    def MarkAsTask(self):
        return MarkAsTaskRuleAction(self.ruleactions.MarkAsTask)

    # Lower case alias for MarkAsTask
    @property
    def markastask(self):
        return self.MarkAsTask

    @property
    def MoveToFolder(self):
        return MoveOrCopyRuleAction(self.ruleactions.MoveToFolder)

    # Lower case alias for MoveToFolder
    @property
    def movetofolder(self):
        return self.MoveToFolder

    @property
    def NewItemAlert(self):
        return NewItemAlertRuleAction(self.ruleactions.NewItemAlert)

    # Lower case alias for NewItemAlert
    @property
    def newitemalert(self):
        return self.NewItemAlert

    @property
    def NotifyDelivery(self):
        return RuleAction(self.ruleactions.NotifyDelivery)

    # Lower case alias for NotifyDelivery
    @property
    def notifydelivery(self):
        return self.NotifyDelivery

    @property
    def NotifyRead(self):
        return RuleAction(self.ruleactions.NotifyRead)

    # Lower case alias for NotifyRead
    @property
    def notifyread(self):
        return self.NotifyRead

    @property
    def Parent(self):
        return self.ruleactions.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PlaySound(self):
        return PlaySoundRuleAction(self.ruleactions.PlaySound)

    # Lower case alias for PlaySound
    @property
    def playsound(self):
        return self.PlaySound

    @property
    def Redirect(self):
        return SendRuleAction(self.ruleactions.Redirect)

    # Lower case alias for Redirect
    @property
    def redirect(self):
        return self.Redirect

    @property
    def Session(self):
        return NameSpace(self.ruleactions.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Stop(self):
        return RuleAction(self.ruleactions.Stop)

    # Lower case alias for Stop
    @property
    def stop(self):
        return self.Stop

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.ruleactions.Item(*arguments)


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

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return RuleCondition(self.rulecondition.Enabled)

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.rulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.rulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.rulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class RuleConditions:

    def __init__(self, ruleconditions=None):
        self.ruleconditions = ruleconditions

    @property
    def Account(self):
        return AccountRuleCondition(self.ruleconditions.Account)

    # Lower case alias for Account
    @property
    def account(self):
        return self.Account

    @property
    def AnyCategory(self):
        return RuleCondition(self.ruleconditions.AnyCategory)

    # Lower case alias for AnyCategory
    @property
    def anycategory(self):
        return self.AnyCategory

    @property
    def Application(self):
        return Application(self.ruleconditions.Application)

    @property
    def Body(self):
        return TextRuleCondition(self.ruleconditions.Body)

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @property
    def BodyOrSubject(self):
        return TextRuleCondition(self.ruleconditions.BodyOrSubject)

    # Lower case alias for BodyOrSubject
    @property
    def bodyorsubject(self):
        return self.BodyOrSubject

    @property
    def Category(self):
        return CategoryRuleCondition(self.ruleconditions.Category)

    # Lower case alias for Category
    @property
    def category(self):
        return self.Category

    @property
    def CC(self):
        return RuleCondition(self.ruleconditions.CC)

    # Lower case alias for CC
    @property
    def cc(self):
        return self.CC

    @property
    def Class(self):
        return OlObjectClass(self.ruleconditions.Class)

    @property
    def Count(self):
        return self.ruleconditions.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def FormName(self):
        return FormNameRuleCondition(self.ruleconditions.FormName)

    # Lower case alias for FormName
    @property
    def formname(self):
        return self.FormName

    @property
    def From(self):
        return ToOrFromRuleCondition(self.ruleconditions.From)

    @property
    def FromAnyRSSFeed(self):
        return RuleCondition(self.ruleconditions.FromAnyRSSFeed)

    # Lower case alias for FromAnyRSSFeed
    @property
    def fromanyrssfeed(self):
        return self.FromAnyRSSFeed

    @property
    def FromRssFeed(self):
        return FromRssFeedRuleCondition(self.ruleconditions.FromRssFeed)

    # Lower case alias for FromRssFeed
    @property
    def fromrssfeed(self):
        return self.FromRssFeed

    @property
    def HasAttachment(self):
        return RuleCondition(self.ruleconditions.HasAttachment)

    # Lower case alias for HasAttachment
    @property
    def hasattachment(self):
        return self.HasAttachment

    @property
    def Importance(self):
        return ImportanceRuleCondition(self.ruleconditions.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @property
    def MeetingInviteOrUpdate(self):
        return RuleCondition(self.ruleconditions.MeetingInviteOrUpdate)

    # Lower case alias for MeetingInviteOrUpdate
    @property
    def meetinginviteorupdate(self):
        return self.MeetingInviteOrUpdate

    @property
    def MessageHeader(self):
        return TextRuleCondition(self.ruleconditions.MessageHeader)

    # Lower case alias for MessageHeader
    @property
    def messageheader(self):
        return self.MessageHeader

    @property
    def NotTo(self):
        return RuleCondition(self.ruleconditions.NotTo)

    # Lower case alias for NotTo
    @property
    def notto(self):
        return self.NotTo

    @property
    def OnLocalMachine(self):
        return RuleCondition(self.ruleconditions.OnLocalMachine)

    # Lower case alias for OnLocalMachine
    @property
    def onlocalmachine(self):
        return self.OnLocalMachine

    @property
    def OnlyToMe(self):
        return RuleCondition(self.ruleconditions.OnlyToMe)

    # Lower case alias for OnlyToMe
    @property
    def onlytome(self):
        return self.OnlyToMe

    @property
    def OnOtherMachine(self):
        return RuleCondition(self.ruleconditions.OnOtherMachine)

    # Lower case alias for OnOtherMachine
    @property
    def onothermachine(self):
        return self.OnOtherMachine

    @property
    def Parent(self):
        return self.ruleconditions.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RecipientAddress(self):
        return AddressRuleCondition(self.ruleconditions.RecipientAddress)

    # Lower case alias for RecipientAddress
    @property
    def recipientaddress(self):
        return self.RecipientAddress

    @property
    def SenderAddress(self):
        return AddressRuleCondition(self.ruleconditions.SenderAddress)

    # Lower case alias for SenderAddress
    @property
    def senderaddress(self):
        return self.SenderAddress

    @property
    def SenderInAddressList(self):
        return SenderInAddressListRuleCondition(self.ruleconditions.SenderInAddressList)

    # Lower case alias for SenderInAddressList
    @property
    def senderinaddresslist(self):
        return self.SenderInAddressList

    @property
    def SentTo(self):
        return ToOrFromRuleCondition(self.ruleconditions.SentTo)

    # Lower case alias for SentTo
    @property
    def sentto(self):
        return self.SentTo

    @property
    def Session(self):
        return NameSpace(self.ruleconditions.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Subject(self):
        return TextRuleCondition(self.ruleconditions.Subject)

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @property
    def ToMe(self):
        return RuleCondition(self.ruleconditions.ToMe)

    # Lower case alias for ToMe
    @property
    def tome(self):
        return self.ToMe

    @property
    def ToOrCc(self):
        return RuleCondition(self.ruleconditions.ToOrCc)

    # Lower case alias for ToOrCc
    @property
    def toorcc(self):
        return self.ToOrCc

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.ruleconditions.Item(*arguments)


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def IsRssRulesProcessingEnabled(self):
        return self.rules.IsRssRulesProcessingEnabled

    # Lower case alias for IsRssRulesProcessingEnabled
    @property
    def isrssrulesprocessingenabled(self):
        return self.IsRssRulesProcessingEnabled

    @IsRssRulesProcessingEnabled.setter
    def IsRssRulesProcessingEnabled(self, value):
        self.rules.IsRssRulesProcessingEnabled = value

    # Lower case alias for IsRssRulesProcessingEnabled setter
    @isrssrulesprocessingenabled.setter
    def isrssrulesprocessingenabled(self, value):
        self.IsRssRulesProcessingEnabled = value

    @property
    def Parent(self):
        return self.rules.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.rules.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Create(self, Name=None, RuleType=None):
        arguments = com_arguments([Name, RuleType])
        return self.rules.Create(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.rules.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.rules.Remove(*arguments)

    def Save(self, ShowProgress=None):
        arguments = com_arguments([ShowProgress])
        self.rules.Save(*arguments)


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

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @property
    def IsSynchronous(self):
        return self.search.IsSynchronous

    # Lower case alias for IsSynchronous
    @property
    def issynchronous(self):
        return self.IsSynchronous

    @property
    def Parent(self):
        return self.search.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Results(self):
        return Results(self.search.Results)

    # Lower case alias for Results
    @property
    def results(self):
        return self.Results

    @property
    def Scope(self):
        return self.search.Scope

    # Lower case alias for Scope
    @property
    def scope(self):
        return self.Scope

    @property
    def SearchSubFolders(self):
        return self.search.SearchSubFolders

    # Lower case alias for SearchSubFolders
    @property
    def searchsubfolders(self):
        return self.SearchSubFolders

    @property
    def Session(self):
        return NameSpace(self.search.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Tag(self):
        return self.search.Tag

    # Lower case alias for Tag
    @property
    def tag(self):
        return self.Tag

    def GetTable(self):
        return self.search.GetTable()

    def Save(self, SchFldrName=None):
        arguments = com_arguments([SchFldrName])
        self.search.Save(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Location(self):
        return OlSelectionLocation(self.selection.Location)

    # Lower case alias for Location
    @property
    def location(self):
        return self.Location

    @property
    def Parent(self):
        return self.selection.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.selection.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def GetSelection(self, SelectionContents=None):
        arguments = com_arguments([SelectionContents])
        return self.selection.GetSelection(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.selection.Item(*arguments)


class SelectNamesDialog:

    def __init__(self, selectnamesdialog=None):
        self.selectnamesdialog = selectnamesdialog

    @property
    def AllowMultipleSelection(self):
        return self.selectnamesdialog.AllowMultipleSelection

    # Lower case alias for AllowMultipleSelection
    @property
    def allowmultipleselection(self):
        return self.AllowMultipleSelection

    @AllowMultipleSelection.setter
    def AllowMultipleSelection(self, value):
        self.selectnamesdialog.AllowMultipleSelection = value

    # Lower case alias for AllowMultipleSelection setter
    @allowmultipleselection.setter
    def allowmultipleselection(self, value):
        self.AllowMultipleSelection = value

    @property
    def Application(self):
        return Application(self.selectnamesdialog.Application)

    @property
    def BccLabel(self):
        return self.selectnamesdialog.BccLabel

    # Lower case alias for BccLabel
    @property
    def bcclabel(self):
        return self.BccLabel

    @BccLabel.setter
    def BccLabel(self, value):
        self.selectnamesdialog.BccLabel = value

    # Lower case alias for BccLabel setter
    @bcclabel.setter
    def bcclabel(self, value):
        self.BccLabel = value

    @property
    def Caption(self):
        return self.selectnamesdialog.Caption

    # Lower case alias for Caption
    @property
    def caption(self):
        return self.Caption

    @Caption.setter
    def Caption(self, value):
        self.selectnamesdialog.Caption = value

    # Lower case alias for Caption setter
    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def CcLabel(self):
        return self.selectnamesdialog.CcLabel

    # Lower case alias for CcLabel
    @property
    def cclabel(self):
        return self.CcLabel

    @CcLabel.setter
    def CcLabel(self, value):
        self.selectnamesdialog.CcLabel = value

    # Lower case alias for CcLabel setter
    @cclabel.setter
    def cclabel(self, value):
        self.CcLabel = value

    @property
    def Class(self):
        return OlObjectClass(self.selectnamesdialog.Class)

    @property
    def ForceResolution(self):
        return SelectNamesDialog.Recipients(self.selectnamesdialog.ForceResolution)

    # Lower case alias for ForceResolution
    @property
    def forceresolution(self):
        return self.ForceResolution

    @ForceResolution.setter
    def ForceResolution(self, value):
        self.selectnamesdialog.ForceResolution = value

    # Lower case alias for ForceResolution setter
    @forceresolution.setter
    def forceresolution(self, value):
        self.ForceResolution = value

    @property
    def InitialAddressList(self):
        return AddressList(self.selectnamesdialog.InitialAddressList)

    # Lower case alias for InitialAddressList
    @property
    def initialaddresslist(self):
        return self.InitialAddressList

    @InitialAddressList.setter
    def InitialAddressList(self, value):
        self.selectnamesdialog.InitialAddressList = value

    # Lower case alias for InitialAddressList setter
    @initialaddresslist.setter
    def initialaddresslist(self, value):
        self.InitialAddressList = value

    @property
    def NumberOfRecipientSelectors(self):
        return OlRecipientSelectors(self.selectnamesdialog.NumberOfRecipientSelectors)

    # Lower case alias for NumberOfRecipientSelectors
    @property
    def numberofrecipientselectors(self):
        return self.NumberOfRecipientSelectors

    @NumberOfRecipientSelectors.setter
    def NumberOfRecipientSelectors(self, value):
        self.selectnamesdialog.NumberOfRecipientSelectors = value

    # Lower case alias for NumberOfRecipientSelectors setter
    @numberofrecipientselectors.setter
    def numberofrecipientselectors(self, value):
        self.NumberOfRecipientSelectors = value

    @property
    def Parent(self):
        return SelectNamesDialog(self.selectnamesdialog.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.selectnamesdialog.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @Recipients.setter
    def Recipients(self, value):
        self.selectnamesdialog.Recipients = value

    # Lower case alias for Recipients setter
    @recipients.setter
    def recipients(self, value):
        self.Recipients = value

    @property
    def Session(self):
        return NameSpace(self.selectnamesdialog.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowOnlyInitialAddressList(self):
        return AddressList(self.selectnamesdialog.ShowOnlyInitialAddressList)

    # Lower case alias for ShowOnlyInitialAddressList
    @property
    def showonlyinitialaddresslist(self):
        return self.ShowOnlyInitialAddressList

    @ShowOnlyInitialAddressList.setter
    def ShowOnlyInitialAddressList(self, value):
        self.selectnamesdialog.ShowOnlyInitialAddressList = value

    # Lower case alias for ShowOnlyInitialAddressList setter
    @showonlyinitialaddresslist.setter
    def showonlyinitialaddresslist(self, value):
        self.ShowOnlyInitialAddressList = value

    @property
    def ToLabel(self):
        return self.selectnamesdialog.ToLabel

    # Lower case alias for ToLabel
    @property
    def tolabel(self):
        return self.ToLabel

    @ToLabel.setter
    def ToLabel(self, value):
        self.selectnamesdialog.ToLabel = value

    # Lower case alias for ToLabel setter
    @tolabel.setter
    def tolabel(self, value):
        self.ToLabel = value

    def Display(self):
        return self.selectnamesdialog.Display()

    def SetDefaultDisplayMode(self, defaultMode=None):
        arguments = com_arguments([defaultMode])
        self.selectnamesdialog.SetDefaultDisplayMode(*arguments)


class SenderInAddressListRuleCondition:

    def __init__(self, senderinaddresslistrulecondition=None):
        self.senderinaddresslistrulecondition = senderinaddresslistrulecondition

    @property
    def AddressList(self):
        return AddressList(self.senderinaddresslistrulecondition.AddressList)

    # Lower case alias for AddressList
    @property
    def addresslist(self):
        return self.AddressList

    @AddressList.setter
    def AddressList(self, value):
        self.senderinaddresslistrulecondition.AddressList = value

    # Lower case alias for AddressList setter
    @addresslist.setter
    def addresslist(self, value):
        self.AddressList = value

    @property
    def Application(self):
        return Application(self.senderinaddresslistrulecondition.Application)

    @property
    def Class(self):
        return OlObjectClass(self.senderinaddresslistrulecondition.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.senderinaddresslistrulecondition.ConditionType)

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.senderinaddresslistrulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.senderinaddresslistrulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.senderinaddresslistrulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.senderinaddresslistrulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class SendRuleAction:

    def __init__(self, sendruleaction=None):
        self.sendruleaction = sendruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.sendruleaction.ActionType)

    # Lower case alias for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.sendruleaction.Application)

    @property
    def Class(self):
        return OlObjectClass(self.sendruleaction.Class)

    @property
    def Enabled(self):
        return self.sendruleaction.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.sendruleaction.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.sendruleaction.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.sendruleaction.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def Session(self):
        return NameSpace(self.sendruleaction.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


class SharingItem:

    def __init__(self, sharingitem=None):
        self.sharingitem = sharingitem

    @property
    def Actions(self):
        return Actions(self.sharingitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def AllowWriteAccess(self):
        return self.sharingitem.AllowWriteAccess

    # Lower case alias for AllowWriteAccess
    @property
    def allowwriteaccess(self):
        return self.AllowWriteAccess

    @AllowWriteAccess.setter
    def AllowWriteAccess(self, value):
        self.sharingitem.AllowWriteAccess = value

    # Lower case alias for AllowWriteAccess setter
    @allowwriteaccess.setter
    def allowwriteaccess(self, value):
        self.AllowWriteAccess = value

    @property
    def AlternateRecipientAllowed(self):
        return self.sharingitem.AlternateRecipientAllowed

    # Lower case alias for AlternateRecipientAllowed
    @property
    def alternaterecipientallowed(self):
        return self.AlternateRecipientAllowed

    @AlternateRecipientAllowed.setter
    def AlternateRecipientAllowed(self, value):
        self.sharingitem.AlternateRecipientAllowed = value

    # Lower case alias for AlternateRecipientAllowed setter
    @alternaterecipientallowed.setter
    def alternaterecipientallowed(self, value):
        self.AlternateRecipientAllowed = value

    @property
    def Application(self):
        return Application(self.sharingitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.sharingitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.sharingitem.AutoForwarded

    # Lower case alias for AutoForwarded
    @property
    def autoforwarded(self):
        return self.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.sharingitem.AutoForwarded = value

    # Lower case alias for AutoForwarded setter
    @autoforwarded.setter
    def autoforwarded(self, value):
        self.AutoForwarded = value

    @property
    def BCC(self):
        return SharingItem(self.sharingitem.BCC)

    # Lower case alias for BCC
    @property
    def bcc(self):
        return self.BCC

    @BCC.setter
    def BCC(self, value):
        self.sharingitem.BCC = value

    # Lower case alias for BCC setter
    @bcc.setter
    def bcc(self, value):
        self.BCC = value

    @property
    def BillingInformation(self):
        return SharingItem(self.sharingitem.BillingInformation)

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.sharingitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return SharingItem(self.sharingitem.Body)

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.sharingitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.sharingitem.BodyFormat)

    # Lower case alias for BodyFormat
    @property
    def bodyformat(self):
        return self.BodyFormat

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.sharingitem.BodyFormat = value

    # Lower case alias for BodyFormat setter
    @bodyformat.setter
    def bodyformat(self, value):
        self.BodyFormat = value

    @property
    def Categories(self):
        return SharingItem(self.sharingitem.Categories)

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.sharingitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def CC(self):
        return SharingItem(self.sharingitem.CC)

    # Lower case alias for CC
    @property
    def cc(self):
        return self.CC

    @CC.setter
    def CC(self, value):
        self.sharingitem.CC = value

    # Lower case alias for CC setter
    @cc.setter
    def cc(self, value):
        self.CC = value

    @property
    def Class(self):
        return OlObjectClass(self.sharingitem.Class)

    @property
    def Companies(self):
        return SharingItem(self.sharingitem.Companies)

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.sharingitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.sharingitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.sharingitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return SharingItem(self.sharingitem.ConversationIndex)

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return SharingItem(self.sharingitem.ConversationTopic)

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return SharingItem(self.sharingitem.CreationTime)

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return SharingItem(self.sharingitem.DeferredDeliveryTime)

    # Lower case alias for DeferredDeliveryTime
    @property
    def deferreddeliverytime(self):
        return self.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.sharingitem.DeferredDeliveryTime = value

    # Lower case alias for DeferredDeliveryTime setter
    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.sharingitem.DeleteAfterSubmit

    # Lower case alias for DeleteAfterSubmit
    @property
    def deleteaftersubmit(self):
        return self.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.sharingitem.DeleteAfterSubmit = value

    # Lower case alias for DeleteAfterSubmit setter
    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.sharingitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return SharingItem(self.sharingitem.EntryID)

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ExpiryTime(self):
        return SharingItem(self.sharingitem.ExpiryTime)

    # Lower case alias for ExpiryTime
    @property
    def expirytime(self):
        return self.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.sharingitem.ExpiryTime = value

    # Lower case alias for ExpiryTime setter
    @expirytime.setter
    def expirytime(self, value):
        self.ExpiryTime = value

    @property
    def FlagRequest(self):
        return SharingItem(self.sharingitem.FlagRequest)

    # Lower case alias for FlagRequest
    @property
    def flagrequest(self):
        return self.FlagRequest

    @FlagRequest.setter
    def FlagRequest(self, value):
        self.sharingitem.FlagRequest = value

    # Lower case alias for FlagRequest setter
    @flagrequest.setter
    def flagrequest(self, value):
        self.FlagRequest = value

    @property
    def FormDescription(self):
        return FormDescription(self.sharingitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.sharingitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def HTMLBody(self):
        return SharingItem(self.sharingitem.HTMLBody)

    # Lower case alias for HTMLBody
    @property
    def htmlbody(self):
        return self.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.sharingitem.HTMLBody = value

    # Lower case alias for HTMLBody setter
    @htmlbody.setter
    def htmlbody(self, value):
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.sharingitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.sharingitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.sharingitem.InternetCodepage

    # Lower case alias for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.sharingitem.InternetCodepage = value

    # Lower case alias for InternetCodepage setter
    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return SharingItem(self.sharingitem.IsConflict)

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return SharingItem(self.sharingitem.IsMarkedAsTask)

    # Lower case alias for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.sharingitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return SharingItem(self.sharingitem.LastModificationTime)

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.sharingitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.sharingitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return SharingItem(self.sharingitem.MessageClass)

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.sharingitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return SharingItem(self.sharingitem.Mileage)

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.sharingitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return SharingItem(self.sharingitem.NoAging)

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.sharingitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return SharingItem(self.sharingitem.OriginatorDeliveryReportRequested)

    # Lower case alias for OriginatorDeliveryReportRequested
    @property
    def originatordeliveryreportrequested(self):
        return self.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.sharingitem.OriginatorDeliveryReportRequested = value

    # Lower case alias for OriginatorDeliveryReportRequested setter
    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return SharingItem(self.sharingitem.OutlookInternalVersion)

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return SharingItem(self.sharingitem.OutlookVersion)

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return SharingItem(self.sharingitem.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Permission(self):
        return self.sharingitem.Permission

    # Lower case alias for Permission
    @property
    def permission(self):
        return self.Permission

    @Permission.setter
    def Permission(self, value):
        self.sharingitem.Permission = value

    # Lower case alias for Permission setter
    @permission.setter
    def permission(self, value):
        self.Permission = value

    @property
    def PermissionService(self):
        return self.sharingitem.PermissionService

    # Lower case alias for PermissionService
    @property
    def permissionservice(self):
        return self.PermissionService

    @PermissionService.setter
    def PermissionService(self, value):
        self.sharingitem.PermissionService = value

    # Lower case alias for PermissionService setter
    @permissionservice.setter
    def permissionservice(self, value):
        self.PermissionService = value

    @property
    def PermissionTemplateGuid(self):
        return SharingItem(self.sharingitem.PermissionTemplateGuid)

    # Lower case alias for PermissionTemplateGuid
    @property
    def permissiontemplateguid(self):
        return self.PermissionTemplateGuid

    @PermissionTemplateGuid.setter
    def PermissionTemplateGuid(self, value):
        self.sharingitem.PermissionTemplateGuid = value

    # Lower case alias for PermissionTemplateGuid setter
    @permissiontemplateguid.setter
    def permissiontemplateguid(self, value):
        self.PermissionTemplateGuid = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.sharingitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReadReceiptRequested(self):
        return self.sharingitem.ReadReceiptRequested

    # Lower case alias for ReadReceiptRequested
    @property
    def readreceiptrequested(self):
        return self.ReadReceiptRequested

    @property
    def ReceivedByEntryID(self):
        return self.sharingitem.ReceivedByEntryID

    # Lower case alias for ReceivedByEntryID
    @property
    def receivedbyentryid(self):
        return self.ReceivedByEntryID

    @property
    def ReceivedByName(self):
        return SharingItem(self.sharingitem.ReceivedByName)

    # Lower case alias for ReceivedByName
    @property
    def receivedbyname(self):
        return self.ReceivedByName

    @property
    def ReceivedOnBehalfOfEntryID(self):
        return self.sharingitem.ReceivedOnBehalfOfEntryID

    # Lower case alias for ReceivedOnBehalfOfEntryID
    @property
    def receivedonbehalfofentryid(self):
        return self.ReceivedOnBehalfOfEntryID

    @property
    def ReceivedOnBehalfOfName(self):
        return SharingItem(self.sharingitem.ReceivedOnBehalfOfName)

    # Lower case alias for ReceivedOnBehalfOfName
    @property
    def receivedonbehalfofname(self):
        return self.ReceivedOnBehalfOfName

    @property
    def ReceivedTime(self):
        return SharingItem(self.sharingitem.ReceivedTime)

    # Lower case alias for ReceivedTime
    @property
    def receivedtime(self):
        return self.ReceivedTime

    @property
    def RecipientReassignmentProhibited(self):
        return SharingItem(self.sharingitem.RecipientReassignmentProhibited)

    # Lower case alias for RecipientReassignmentProhibited
    @property
    def recipientreassignmentprohibited(self):
        return self.RecipientReassignmentProhibited

    @RecipientReassignmentProhibited.setter
    def RecipientReassignmentProhibited(self, value):
        self.sharingitem.RecipientReassignmentProhibited = value

    # Lower case alias for RecipientReassignmentProhibited setter
    @recipientreassignmentprohibited.setter
    def recipientreassignmentprohibited(self, value):
        self.RecipientReassignmentProhibited = value

    @property
    def Recipients(self):
        return Recipients(self.sharingitem.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return SharingItem(self.sharingitem.ReminderOverrideDefault)

    # Lower case alias for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.sharingitem.ReminderOverrideDefault = value

    # Lower case alias for ReminderOverrideDefault setter
    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return SharingItem(self.sharingitem.ReminderPlaySound)

    # Lower case alias for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.sharingitem.ReminderPlaySound = value

    # Lower case alias for ReminderPlaySound setter
    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return SharingItem(self.sharingitem.ReminderSet)

    # Lower case alias for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.sharingitem.ReminderSet = value

    # Lower case alias for ReminderSet setter
    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.sharingitem.ReminderSoundFile

    # Lower case alias for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.sharingitem.ReminderSoundFile = value

    # Lower case alias for ReminderSoundFile setter
    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return SharingItem(self.sharingitem.ReminderTime)

    # Lower case alias for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.sharingitem.ReminderTime = value

    # Lower case alias for ReminderTime setter
    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RemoteID(self):
        return SharingItem(self.sharingitem.RemoteID)

    # Lower case alias for RemoteID
    @property
    def remoteid(self):
        return self.RemoteID

    @property
    def RemoteName(self):
        return SharingItem(self.sharingitem.RemoteName)

    # Lower case alias for RemoteName
    @property
    def remotename(self):
        return self.RemoteName

    @property
    def RemotePath(self):
        return SharingItem(self.sharingitem.RemotePath)

    # Lower case alias for RemotePath
    @property
    def remotepath(self):
        return self.RemotePath

    @property
    def RemoteStatus(self):
        return OlRemoteStatus(self.sharingitem.RemoteStatus)

    # Lower case alias for RemoteStatus
    @property
    def remotestatus(self):
        return self.RemoteStatus

    @RemoteStatus.setter
    def RemoteStatus(self, value):
        self.sharingitem.RemoteStatus = value

    # Lower case alias for RemoteStatus setter
    @remotestatus.setter
    def remotestatus(self, value):
        self.RemoteStatus = value

    @property
    def ReplyRecipientNames(self):
        return SharingItem(self.sharingitem.ReplyRecipientNames)

    # Lower case alias for ReplyRecipientNames
    @property
    def replyrecipientnames(self):
        return self.ReplyRecipientNames

    @property
    def ReplyRecipients(self):
        return Recipients(self.sharingitem.ReplyRecipients)

    # Lower case alias for ReplyRecipients
    @property
    def replyrecipients(self):
        return self.ReplyRecipients

    @property
    def RequestedFolder(self):
        return OlDefaultFolders(self.sharingitem.RequestedFolder)

    # Lower case alias for RequestedFolder
    @property
    def requestedfolder(self):
        return self.RequestedFolder

    @property
    def RetentionExpirationDate(self):
        return SharingItem(self.sharingitem.RetentionExpirationDate)

    # Lower case alias for RetentionExpirationDate
    @property
    def retentionexpirationdate(self):
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.sharingitem.RetentionPolicyName

    # Lower case alias for RetentionPolicyName
    @property
    def retentionpolicyname(self):
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.sharingitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.sharingitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return SharingItem(self.sharingitem.Saved)

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return Folder(self.sharingitem.SaveSentMessageFolder)

    # Lower case alias for SaveSentMessageFolder
    @property
    def savesentmessagefolder(self):
        return self.SaveSentMessageFolder

    @SaveSentMessageFolder.setter
    def SaveSentMessageFolder(self, value):
        self.sharingitem.SaveSentMessageFolder = value

    # Lower case alias for SaveSentMessageFolder setter
    @savesentmessagefolder.setter
    def savesentmessagefolder(self, value):
        self.SaveSentMessageFolder = value

    @property
    def SenderEmailAddress(self):
        return SharingItem(self.sharingitem.SenderEmailAddress)

    # Lower case alias for SenderEmailAddress
    @property
    def senderemailaddress(self):
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return SharingItem(self.sharingitem.SenderEmailType)

    # Lower case alias for SenderEmailType
    @property
    def senderemailtype(self):
        return self.SenderEmailType

    @property
    def SenderName(self):
        return SharingItem(self.sharingitem.SenderName)

    # Lower case alias for SenderName
    @property
    def sendername(self):
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.sharingitem.SendUsingAccount)

    # Lower case alias for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.sharingitem.SendUsingAccount = value

    # Lower case alias for SendUsingAccount setter
    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.sharingitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.sharingitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Sent(self):
        return SharingItem(self.sharingitem.Sent)

    # Lower case alias for Sent
    @property
    def sent(self):
        return self.Sent

    @property
    def SentOn(self):
        return SharingItem(self.sharingitem.SentOn)

    # Lower case alias for SentOn
    @property
    def senton(self):
        return self.SentOn

    @property
    def SentOnBehalfOfName(self):
        return SharingItem(self.sharingitem.SentOnBehalfOfName)

    # Lower case alias for SentOnBehalfOfName
    @property
    def sentonbehalfofname(self):
        return self.SentOnBehalfOfName

    @SentOnBehalfOfName.setter
    def SentOnBehalfOfName(self, value):
        self.sharingitem.SentOnBehalfOfName = value

    # Lower case alias for SentOnBehalfOfName setter
    @sentonbehalfofname.setter
    def sentonbehalfofname(self, value):
        self.SentOnBehalfOfName = value

    @property
    def Session(self):
        return NameSpace(self.sharingitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def SharingProvider(self):
        return OlSharingProvider(self.sharingitem.SharingProvider)

    # Lower case alias for SharingProvider
    @property
    def sharingprovider(self):
        return self.SharingProvider

    @property
    def SharingProviderGuid(self):
        return SharingItem(self.sharingitem.SharingProviderGuid)

    # Lower case alias for SharingProviderGuid
    @property
    def sharingproviderguid(self):
        return self.SharingProviderGuid

    @property
    def Size(self):
        return SharingItem(self.sharingitem.Size)

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return SharingItem(self.sharingitem.Subject)

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.sharingitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Submitted(self):
        return SharingItem(self.sharingitem.Submitted)

    # Lower case alias for Submitted
    @property
    def submitted(self):
        return self.Submitted

    @property
    def TaskCompletedDate(self):
        return SharingItem(self.sharingitem.TaskCompletedDate)

    # Lower case alias for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.sharingitem.TaskCompletedDate = value

    # Lower case alias for TaskCompletedDate setter
    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return SharingItem(self.sharingitem.TaskDueDate)

    # Lower case alias for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.sharingitem.TaskDueDate = value

    # Lower case alias for TaskDueDate setter
    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return SharingItem(self.sharingitem.TaskStartDate)

    # Lower case alias for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.sharingitem.TaskStartDate = value

    # Lower case alias for TaskStartDate setter
    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return SharingItem(self.sharingitem.TaskSubject)

    # Lower case alias for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.sharingitem.TaskSubject = value

    # Lower case alias for TaskSubject setter
    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def To(self):
        return SharingItem(self.sharingitem.To)

    # Lower case alias for To
    @property
    def to(self):
        return self.To

    @To.setter
    def To(self, value):
        self.sharingitem.To = value

    # Lower case alias for To setter
    @to.setter
    def to(self, value):
        self.To = value

    @property
    def ToDoTaskOrdinal(self):
        return SharingItem(self.sharingitem.ToDoTaskOrdinal)

    # Lower case alias for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.sharingitem.ToDoTaskOrdinal = value

    # Lower case alias for ToDoTaskOrdinal setter
    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def Type(self):
        return OlSharingMsgType(self.sharingitem.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @Type.setter
    def Type(self, value):
        self.sharingitem.Type = value

    # Lower case alias for Type setter
    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def UnRead(self):
        return SharingItem(self.sharingitem.UnRead)

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.sharingitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.sharingitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def AddBusinessCard(self, contact=None):
        arguments = com_arguments([contact])
        self.sharingitem.AddBusinessCard(*arguments)

    def Allow(self):
        self.sharingitem.Allow()

    def ClearConversationIndex(self):
        self.sharingitem.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.sharingitem.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.sharingitem.Close(*arguments)

    def Copy(self):
        self.sharingitem.Copy()

    def Delete(self):
        self.sharingitem.Delete()

    def Deny(self):
        return self.sharingitem.Deny()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.sharingitem.Display(*arguments)

    def Forward(self):
        return self.sharingitem.Forward()

    def GetConversation(self):
        return self.sharingitem.GetConversation()

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.sharingitem.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.sharingitem.Move(*arguments)

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

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.sharingitem.SaveAs(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return SimpleItems(self.simpleitems.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.simpleitems.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.simpleitems.Item(*arguments)


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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.solutionsmodule.NavigationModuleType)

    # Lower case alias for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return SolutionsModule(self.solutionsmodule.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return SolutionsModule(self.solutionsmodule.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.solutionsmodule.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.solutionsmodule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return self.solutionsmodule.Visible

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.solutionsmodule.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value

    def AddSolution(self, Solution=None, Scope=None):
        arguments = com_arguments([Solution, Scope])
        self.solutionsmodule.AddSolution(*arguments)


class StorageItem:

    def __init__(self, storageitem=None):
        self.storageitem = storageitem

    @property
    def Application(self):
        return Application(self.storageitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.storageitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def Body(self):
        return self.storageitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.storageitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Class(self):
        return OlObjectClass(self.storageitem.Class)

    @property
    def CreationTime(self):
        return StorageItem(self.storageitem.CreationTime)

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def Creator(self):
        return StorageItem(self.storageitem.Creator)

    # Lower case alias for Creator
    @property
    def creator(self):
        return self.Creator

    @Creator.setter
    def Creator(self, value):
        self.storageitem.Creator = value

    # Lower case alias for Creator setter
    @creator.setter
    def creator(self, value):
        self.Creator = value

    @property
    def EntryID(self):
        return self.storageitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def LastModificationTime(self):
        return self.storageitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def Parent(self):
        return self.storageitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.storageitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.storageitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return StorageItem(self.storageitem.Size)

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.storageitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.storageitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UserProperties(self):
        return UserProperties(self.storageitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

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

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @property
    def Class(self):
        return OlObjectClass(self.store.Class)

    @property
    def DisplayName(self):
        return Store(self.store.DisplayName)

    # Lower case alias for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def ExchangeStoreType(self):
        return OlExchangeStoreType(self.store.ExchangeStoreType)

    # Lower case alias for ExchangeStoreType
    @property
    def exchangestoretype(self):
        return self.ExchangeStoreType

    @property
    def FilePath(self):
        return self.store.FilePath

    # Lower case alias for FilePath
    @property
    def filepath(self):
        return self.FilePath

    @property
    def IsCachedExchange(self):
        return Store(self.store.IsCachedExchange)

    # Lower case alias for IsCachedExchange
    @property
    def iscachedexchange(self):
        return self.IsCachedExchange

    @property
    def IsConversationEnabled(self):
        return self.store.IsConversationEnabled

    # Lower case alias for IsConversationEnabled
    @property
    def isconversationenabled(self):
        return self.IsConversationEnabled

    @property
    def IsDataFileStore(self):
        return Store(self.store.IsDataFileStore)

    # Lower case alias for IsDataFileStore
    @property
    def isdatafilestore(self):
        return self.IsDataFileStore

    @property
    def IsInstantSearchEnabled(self):
        return self.store.IsInstantSearchEnabled

    # Lower case alias for IsInstantSearchEnabled
    @property
    def isinstantsearchenabled(self):
        return self.IsInstantSearchEnabled

    @property
    def IsOpen(self):
        return Store(self.store.IsOpen)

    # Lower case alias for IsOpen
    @property
    def isopen(self):
        return self.IsOpen

    @property
    def Parent(self):
        return self.store.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.store.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.store.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def StoreID(self):
        return Store(self.store.StoreID)

    # Lower case alias for StoreID
    @property
    def storeid(self):
        return self.StoreID

    def GetDefaultFolder(self, FolderType=None):
        arguments = com_arguments([FolderType])
        return self.store.GetDefaultFolder(*arguments)

    def GetRootFolder(self):
        return self.store.GetRootFolder()

    def GetRules(self):
        return self.store.GetRules()

    def GetSearchFolders(self):
        return self.store.GetSearchFolders()

    def GetSpecialFolder(self, FolderType=None):
        arguments = com_arguments([FolderType])
        return self.store.GetSpecialFolder(*arguments)

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.stores.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.stores.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Stores(self.stores.Item(*arguments))


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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.syncobject.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.syncobject.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

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

    # Lower case alias for AppFolders
    @property
    def appfolders(self):
        return self.AppFolders

    @property
    def Application(self):
        return Application(self.syncobjects.Application)

    @property
    def Class(self):
        return OlObjectClass(self.syncobjects.Class)

    @property
    def Count(self):
        return self.syncobjects.Count

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.syncobjects.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.syncobjects.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.syncobjects.Item(*arguments)


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

    # Lower case alias for Columns
    @property
    def columns(self):
        return self.Columns

    @property
    def EndOfTable(self):
        return Table(self.table.EndOfTable)

    # Lower case alias for EndOfTable
    @property
    def endoftable(self):
        return self.EndOfTable

    @property
    def Parent(self):
        return Table(self.table.Parent)

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.table.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def FindNextRow(self):
        return Row(self.table.FindNextRow())

    def FindRow(self, Filter=None):
        arguments = com_arguments([Filter])
        return Row(self.table.FindRow(*arguments))

    def GetArray(self, MaxRows=None):
        arguments = com_arguments([MaxRows])
        return self.table.GetArray(*arguments)

    def GetNextRow(self):
        return Row(self.table.GetNextRow())

    def GetRowCount(self):
        return self.table.GetRowCount()

    def MoveToStart(self):
        self.table.MoveToStart()

    def Restrict(self, Filter=None):
        arguments = com_arguments([Filter])
        return self.table.Restrict(*arguments)

    def Sort(self, SortProperty=None, Descending=None):
        arguments = com_arguments([SortProperty, Descending])
        self.table.Sort(*arguments)


class TableView:

    def __init__(self, tableview=None):
        self.tableview = tableview

    @property
    def AllowInCellEditing(self):
        return TableView(self.tableview.AllowInCellEditing)

    # Lower case alias for AllowInCellEditing
    @property
    def allowincellediting(self):
        return self.AllowInCellEditing

    @AllowInCellEditing.setter
    def AllowInCellEditing(self, value):
        self.tableview.AllowInCellEditing = value

    # Lower case alias for AllowInCellEditing setter
    @allowincellediting.setter
    def allowincellediting(self, value):
        self.AllowInCellEditing = value

    @property
    def AlwaysExpandConversation(self):
        return self.tableview.AlwaysExpandConversation

    # Lower case alias for AlwaysExpandConversation
    @property
    def alwaysexpandconversation(self):
        return self.AlwaysExpandConversation

    @AlwaysExpandConversation.setter
    def AlwaysExpandConversation(self, value):
        self.tableview.AlwaysExpandConversation = value

    # Lower case alias for AlwaysExpandConversation setter
    @alwaysexpandconversation.setter
    def alwaysexpandconversation(self, value):
        self.AlwaysExpandConversation = value

    @property
    def Application(self):
        return Application(self.tableview.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.tableview.AutoFormatRules)

    # Lower case alias for AutoFormatRules
    @property
    def autoformatrules(self):
        return self.AutoFormatRules

    @property
    def AutomaticColumnSizing(self):
        return TableView(self.tableview.AutomaticColumnSizing)

    # Lower case alias for AutomaticColumnSizing
    @property
    def automaticcolumnsizing(self):
        return self.AutomaticColumnSizing

    @AutomaticColumnSizing.setter
    def AutomaticColumnSizing(self, value):
        self.tableview.AutomaticColumnSizing = value

    # Lower case alias for AutomaticColumnSizing setter
    @automaticcolumnsizing.setter
    def automaticcolumnsizing(self, value):
        self.AutomaticColumnSizing = value

    @property
    def AutomaticGrouping(self):
        return TableView(self.tableview.AutomaticGrouping)

    # Lower case alias for AutomaticGrouping
    @property
    def automaticgrouping(self):
        return self.AutomaticGrouping

    @AutomaticGrouping.setter
    def AutomaticGrouping(self, value):
        self.tableview.AutomaticGrouping = value

    # Lower case alias for AutomaticGrouping setter
    @automaticgrouping.setter
    def automaticgrouping(self, value):
        self.AutomaticGrouping = value

    @property
    def AutoPreview(self):
        return OlAutoPreview(self.tableview.AutoPreview)

    # Lower case alias for AutoPreview
    @property
    def autopreview(self):
        return self.AutoPreview

    @AutoPreview.setter
    def AutoPreview(self, value):
        self.tableview.AutoPreview = value

    # Lower case alias for AutoPreview setter
    @autopreview.setter
    def autopreview(self, value):
        self.AutoPreview = value

    @property
    def AutoPreviewFont(self):
        return ViewFont(self.tableview.AutoPreviewFont)

    # Lower case alias for AutoPreviewFont
    @property
    def autopreviewfont(self):
        return self.AutoPreviewFont

    @property
    def Class(self):
        return OlObjectClass(self.tableview.Class)

    @property
    def ColumnFont(self):
        return ViewFont(self.tableview.ColumnFont)

    # Lower case alias for ColumnFont
    @property
    def columnfont(self):
        return self.ColumnFont

    @property
    def DefaultExpandCollapseSetting(self):
        return OlDefaultExpandCollapseSetting(self.tableview.DefaultExpandCollapseSetting)

    # Lower case alias for DefaultExpandCollapseSetting
    @property
    def defaultexpandcollapsesetting(self):
        return self.DefaultExpandCollapseSetting

    @DefaultExpandCollapseSetting.setter
    def DefaultExpandCollapseSetting(self, value):
        self.tableview.DefaultExpandCollapseSetting = value

    # Lower case alias for DefaultExpandCollapseSetting setter
    @defaultexpandcollapsesetting.setter
    def defaultexpandcollapsesetting(self, value):
        self.DefaultExpandCollapseSetting = value

    @property
    def Filter(self):
        return self.tableview.Filter

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @Filter.setter
    def Filter(self, value):
        self.tableview.Filter = value

    # Lower case alias for Filter setter
    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def GridLineStyle(self):
        return OlGridLineStyle(self.tableview.GridLineStyle)

    # Lower case alias for GridLineStyle
    @property
    def gridlinestyle(self):
        return self.GridLineStyle

    @GridLineStyle.setter
    def GridLineStyle(self, value):
        self.tableview.GridLineStyle = value

    # Lower case alias for GridLineStyle setter
    @gridlinestyle.setter
    def gridlinestyle(self, value):
        self.GridLineStyle = value

    @property
    def GroupByFields(self):
        return OrderFields(self.tableview.GroupByFields)

    # Lower case alias for GroupByFields
    @property
    def groupbyfields(self):
        return self.GroupByFields

    @property
    def HideReadingPaneHeaderInfo(self):
        return TableView(self.tableview.HideReadingPaneHeaderInfo)

    # Lower case alias for HideReadingPaneHeaderInfo
    @property
    def hidereadingpaneheaderinfo(self):
        return self.HideReadingPaneHeaderInfo

    @HideReadingPaneHeaderInfo.setter
    def HideReadingPaneHeaderInfo(self, value):
        self.tableview.HideReadingPaneHeaderInfo = value

    # Lower case alias for HideReadingPaneHeaderInfo setter
    @hidereadingpaneheaderinfo.setter
    def hidereadingpaneheaderinfo(self, value):
        self.HideReadingPaneHeaderInfo = value

    @property
    def Language(self):
        return self.tableview.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @Language.setter
    def Language(self, value):
        self.tableview.Language = value

    # Lower case alias for Language setter
    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.tableview.LockUserChanges

    # Lower case alias for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.tableview.LockUserChanges = value

    # Lower case alias for LockUserChanges setter
    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def MaxLinesInMultiLineView(self):
        return TableView(self.tableview.MaxLinesInMultiLineView)

    # Lower case alias for MaxLinesInMultiLineView
    @property
    def maxlinesinmultilineview(self):
        return self.MaxLinesInMultiLineView

    @MaxLinesInMultiLineView.setter
    def MaxLinesInMultiLineView(self, value):
        self.tableview.MaxLinesInMultiLineView = value

    # Lower case alias for MaxLinesInMultiLineView setter
    @maxlinesinmultilineview.setter
    def maxlinesinmultilineview(self, value):
        self.MaxLinesInMultiLineView = value

    @property
    def Multiline(self):
        return OlMultiLine(self.tableview.Multiline)

    # Lower case alias for Multiline
    @property
    def multiline(self):
        return self.Multiline

    @Multiline.setter
    def Multiline(self, value):
        self.tableview.Multiline = value

    # Lower case alias for Multiline setter
    @multiline.setter
    def multiline(self, value):
        self.Multiline = value

    @property
    def MultiLineWidth(self):
        return TableView(self.tableview.MultiLineWidth)

    # Lower case alias for MultiLineWidth
    @property
    def multilinewidth(self):
        return self.MultiLineWidth

    @MultiLineWidth.setter
    def MultiLineWidth(self, value):
        self.tableview.MultiLineWidth = value

    # Lower case alias for MultiLineWidth setter
    @multilinewidth.setter
    def multilinewidth(self, value):
        self.MultiLineWidth = value

    @property
    def Name(self):
        return self.tableview.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.tableview.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.tableview.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RowFont(self):
        return ViewFont(self.tableview.RowFont)

    # Lower case alias for RowFont
    @property
    def rowfont(self):
        return self.RowFont

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.tableview.SaveOption)

    # Lower case alias for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.tableview.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowConversationByDate(self):
        return self.tableview.ShowConversationByDate

    # Lower case alias for ShowConversationByDate
    @property
    def showconversationbydate(self):
        return self.ShowConversationByDate

    @ShowConversationByDate.setter
    def ShowConversationByDate(self, value):
        self.tableview.ShowConversationByDate = value

    # Lower case alias for ShowConversationByDate setter
    @showconversationbydate.setter
    def showconversationbydate(self, value):
        self.ShowConversationByDate = value

    @property
    def ShowConversationSendersAboveSubject(self):
        return self.tableview.ShowConversationSendersAboveSubject

    # Lower case alias for ShowConversationSendersAboveSubject
    @property
    def showconversationsendersabovesubject(self):
        return self.ShowConversationSendersAboveSubject

    @ShowConversationSendersAboveSubject.setter
    def ShowConversationSendersAboveSubject(self, value):
        self.tableview.ShowConversationSendersAboveSubject = value

    # Lower case alias for ShowConversationSendersAboveSubject setter
    @showconversationsendersabovesubject.setter
    def showconversationsendersabovesubject(self, value):
        self.ShowConversationSendersAboveSubject = value

    @property
    def ShowFullConversations(self):
        return self.tableview.ShowFullConversations

    # Lower case alias for ShowFullConversations
    @property
    def showfullconversations(self):
        return self.ShowFullConversations

    @ShowFullConversations.setter
    def ShowFullConversations(self, value):
        self.tableview.ShowFullConversations = value

    # Lower case alias for ShowFullConversations setter
    @showfullconversations.setter
    def showfullconversations(self, value):
        self.ShowFullConversations = value

    @property
    def ShowItemsInGroups(self):
        return TableView(self.tableview.ShowItemsInGroups)

    # Lower case alias for ShowItemsInGroups
    @property
    def showitemsingroups(self):
        return self.ShowItemsInGroups

    @ShowItemsInGroups.setter
    def ShowItemsInGroups(self, value):
        self.tableview.ShowItemsInGroups = value

    # Lower case alias for ShowItemsInGroups setter
    @showitemsingroups.setter
    def showitemsingroups(self, value):
        self.ShowItemsInGroups = value

    @property
    def ShowNewItemRow(self):
        return TableView(self.tableview.ShowNewItemRow)

    # Lower case alias for ShowNewItemRow
    @property
    def shownewitemrow(self):
        return self.ShowNewItemRow

    @ShowNewItemRow.setter
    def ShowNewItemRow(self, value):
        self.tableview.ShowNewItemRow = value

    # Lower case alias for ShowNewItemRow setter
    @shownewitemrow.setter
    def shownewitemrow(self, value):
        self.ShowNewItemRow = value

    @property
    def ShowReadingPane(self):
        return TableView(self.tableview.ShowReadingPane)

    # Lower case alias for ShowReadingPane
    @property
    def showreadingpane(self):
        return self.ShowReadingPane

    @ShowReadingPane.setter
    def ShowReadingPane(self, value):
        self.tableview.ShowReadingPane = value

    # Lower case alias for ShowReadingPane setter
    @showreadingpane.setter
    def showreadingpane(self, value):
        self.ShowReadingPane = value

    @property
    def SortFields(self):
        return OrderFields(self.tableview.SortFields)

    # Lower case alias for SortFields
    @property
    def sortfields(self):
        return self.SortFields

    @property
    def Standard(self):
        return TableView(self.tableview.Standard)

    # Lower case alias for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewFields(self):
        return ViewFields(self.tableview.ViewFields)

    # Lower case alias for ViewFields
    @property
    def viewfields(self):
        return self.ViewFields

    @property
    def ViewType(self):
        return OlViewType(self.tableview.ViewType)

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.tableview.XML

    # Lower case alias for XML
    @property
    def xml(self):
        return self.XML

    @XML.setter
    def XML(self, value):
        self.tableview.XML = value

    # Lower case alias for XML setter
    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.tableview.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return self.tableview.Copy(*arguments)

    def Delete(self):
        self.tableview.Delete()

    def GetTable(self):
        return self.tableview.GetTable()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.tableview.GoToDate(*arguments)

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

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def ActualWork(self):
        return self.taskitem.ActualWork

    # Lower case alias for ActualWork
    @property
    def actualwork(self):
        return self.ActualWork

    @ActualWork.setter
    def ActualWork(self, value):
        self.taskitem.ActualWork = value

    # Lower case alias for ActualWork setter
    @actualwork.setter
    def actualwork(self, value):
        self.ActualWork = value

    @property
    def Application(self):
        return Application(self.taskitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.taskitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.taskitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.taskitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def CardData(self):
        return self.taskitem.CardData

    # Lower case alias for CardData
    @property
    def carddata(self):
        return self.CardData

    @CardData.setter
    def CardData(self, value):
        self.taskitem.CardData = value

    # Lower case alias for CardData setter
    @carddata.setter
    def carddata(self, value):
        self.CardData = value

    @property
    def Categories(self):
        return self.taskitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskitem.Class)

    @property
    def Companies(self):
        return self.taskitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Complete(self):
        return self.taskitem.Complete

    # Lower case alias for Complete
    @property
    def complete(self):
        return self.Complete

    @Complete.setter
    def Complete(self, value):
        self.taskitem.Complete = value

    # Lower case alias for Complete setter
    @complete.setter
    def complete(self, value):
        self.Complete = value

    @property
    def Conflicts(self):
        return self.taskitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ContactNames(self):
        return self.taskitem.ContactNames

    # Lower case alias for ContactNames
    @property
    def contactnames(self):
        return self.ContactNames

    @ContactNames.setter
    def ContactNames(self, value):
        self.taskitem.ContactNames = value

    # Lower case alias for ContactNames setter
    @contactnames.setter
    def contactnames(self, value):
        self.ContactNames = value

    @property
    def ConversationID(self):
        return Conversation(self.taskitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.taskitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DateCompleted(self):
        return self.taskitem.DateCompleted

    # Lower case alias for DateCompleted
    @property
    def datecompleted(self):
        return self.DateCompleted

    @DateCompleted.setter
    def DateCompleted(self, value):
        self.taskitem.DateCompleted = value

    # Lower case alias for DateCompleted setter
    @datecompleted.setter
    def datecompleted(self, value):
        self.DateCompleted = value

    @property
    def DelegationState(self):
        return OlTaskDelegationState(self.taskitem.DelegationState)

    # Lower case alias for DelegationState
    @property
    def delegationstate(self):
        return self.DelegationState

    @property
    def Delegator(self):
        return self.taskitem.Delegator

    # Lower case alias for Delegator
    @property
    def delegator(self):
        return self.Delegator

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def DueDate(self):
        return self.taskitem.DueDate

    # Lower case alias for DueDate
    @property
    def duedate(self):
        return self.DueDate

    @DueDate.setter
    def DueDate(self, value):
        self.taskitem.DueDate = value

    # Lower case alias for DueDate setter
    @duedate.setter
    def duedate(self, value):
        self.DueDate = value

    @property
    def EntryID(self):
        return self.taskitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.taskitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.taskitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.taskitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.taskitem.InternetCodepage

    # Lower case alias for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.taskitem.InternetCodepage = value

    # Lower case alias for InternetCodepage setter
    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.taskitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsRecurring(self):
        return self.taskitem.IsRecurring

    # Lower case alias for IsRecurring
    @property
    def isrecurring(self):
        return self.IsRecurring

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.taskitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.taskitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.taskitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def Ordinal(self):
        return self.taskitem.Ordinal

    # Lower case alias for Ordinal
    @property
    def ordinal(self):
        return self.Ordinal

    @Ordinal.setter
    def Ordinal(self, value):
        self.taskitem.Ordinal = value

    # Lower case alias for Ordinal setter
    @ordinal.setter
    def ordinal(self, value):
        self.Ordinal = value

    @property
    def OutlookInternalVersion(self):
        return self.taskitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Owner(self):
        return self.taskitem.Owner

    # Lower case alias for Owner
    @property
    def owner(self):
        return self.Owner

    @Owner.setter
    def Owner(self, value):
        self.taskitem.Owner = value

    # Lower case alias for Owner setter
    @owner.setter
    def owner(self, value):
        self.Owner = value

    @property
    def Ownership(self):
        return OlTaskOwnership(self.taskitem.Ownership)

    # Lower case alias for Ownership
    @property
    def ownership(self):
        return self.Ownership

    @property
    def Parent(self):
        return self.taskitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PercentComplete(self):
        return self.taskitem.PercentComplete

    # Lower case alias for PercentComplete
    @property
    def percentcomplete(self):
        return self.PercentComplete

    @PercentComplete.setter
    def PercentComplete(self, value):
        self.taskitem.PercentComplete = value

    # Lower case alias for PercentComplete setter
    @percentcomplete.setter
    def percentcomplete(self, value):
        self.PercentComplete = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.taskitem.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return self.taskitem.ReminderOverrideDefault

    # Lower case alias for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.taskitem.ReminderOverrideDefault = value

    # Lower case alias for ReminderOverrideDefault setter
    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.taskitem.ReminderPlaySound

    # Lower case alias for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.taskitem.ReminderPlaySound = value

    # Lower case alias for ReminderPlaySound setter
    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.taskitem.ReminderSet

    # Lower case alias for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.taskitem.ReminderSet = value

    # Lower case alias for ReminderSet setter
    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.taskitem.ReminderSoundFile

    # Lower case alias for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.taskitem.ReminderSoundFile = value

    # Lower case alias for ReminderSoundFile setter
    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.taskitem.ReminderTime

    # Lower case alias for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.taskitem.ReminderTime = value

    # Lower case alias for ReminderTime setter
    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def ResponseState(self):
        return OlTaskResponse(self.taskitem.ResponseState)

    # Lower case alias for ResponseState
    @property
    def responsestate(self):
        return self.ResponseState

    @property
    def Role(self):
        return self.taskitem.Role

    # Lower case alias for Role
    @property
    def role(self):
        return self.Role

    @Role.setter
    def Role(self, value):
        self.taskitem.Role = value

    # Lower case alias for Role setter
    @role.setter
    def role(self, value):
        self.Role = value

    @property
    def RTFBody(self):
        return self.taskitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.taskitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SchedulePlusPriority(self):
        return self.taskitem.SchedulePlusPriority

    # Lower case alias for SchedulePlusPriority
    @property
    def schedulepluspriority(self):
        return self.SchedulePlusPriority

    @SchedulePlusPriority.setter
    def SchedulePlusPriority(self, value):
        self.taskitem.SchedulePlusPriority = value

    # Lower case alias for SchedulePlusPriority setter
    @schedulepluspriority.setter
    def schedulepluspriority(self, value):
        self.SchedulePlusPriority = value

    @property
    def SendUsingAccount(self):
        return Account(self.taskitem.SendUsingAccount)

    # Lower case alias for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.taskitem.SendUsingAccount = value

    # Lower case alias for SendUsingAccount setter
    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.taskitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def StartDate(self):
        return self.taskitem.StartDate

    # Lower case alias for StartDate
    @property
    def startdate(self):
        return self.StartDate

    @StartDate.setter
    def StartDate(self, value):
        self.taskitem.StartDate = value

    # Lower case alias for StartDate setter
    @startdate.setter
    def startdate(self, value):
        self.StartDate = value

    @property
    def Status(self):
        return OlTaskStatus(self.taskitem.Status)

    # Lower case alias for Status
    @property
    def status(self):
        return self.Status

    @Status.setter
    def Status(self, value):
        self.taskitem.Status = value

    # Lower case alias for Status setter
    @status.setter
    def status(self, value):
        self.Status = value

    @property
    def StatusOnCompletionRecipients(self):
        return self.taskitem.StatusOnCompletionRecipients

    # Lower case alias for StatusOnCompletionRecipients
    @property
    def statusoncompletionrecipients(self):
        return self.StatusOnCompletionRecipients

    @StatusOnCompletionRecipients.setter
    def StatusOnCompletionRecipients(self, value):
        self.taskitem.StatusOnCompletionRecipients = value

    # Lower case alias for StatusOnCompletionRecipients setter
    @statusoncompletionrecipients.setter
    def statusoncompletionrecipients(self, value):
        self.StatusOnCompletionRecipients = value

    @property
    def StatusUpdateRecipients(self):
        return self.taskitem.StatusUpdateRecipients

    # Lower case alias for StatusUpdateRecipients
    @property
    def statusupdaterecipients(self):
        return self.StatusUpdateRecipients

    @StatusUpdateRecipients.setter
    def StatusUpdateRecipients(self, value):
        self.taskitem.StatusUpdateRecipients = value

    # Lower case alias for StatusUpdateRecipients setter
    @statusupdaterecipients.setter
    def statusupdaterecipients(self, value):
        self.StatusUpdateRecipients = value

    @property
    def Subject(self):
        return self.taskitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def TeamTask(self):
        return self.taskitem.TeamTask

    # Lower case alias for TeamTask
    @property
    def teamtask(self):
        return self.TeamTask

    @TeamTask.setter
    def TeamTask(self, value):
        self.taskitem.TeamTask = value

    # Lower case alias for TeamTask setter
    @teamtask.setter
    def teamtask(self, value):
        self.TeamTask = value

    @property
    def ToDoTaskOrdinal(self):
        return TaskItem(self.taskitem.ToDoTaskOrdinal)

    # Lower case alias for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.taskitem.ToDoTaskOrdinal = value

    # Lower case alias for ToDoTaskOrdinal setter
    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def TotalWork(self):
        return self.taskitem.TotalWork

    # Lower case alias for TotalWork
    @property
    def totalwork(self):
        return self.TotalWork

    @TotalWork.setter
    def TotalWork(self, value):
        self.taskitem.TotalWork = value

    # Lower case alias for TotalWork setter
    @totalwork.setter
    def totalwork(self, value):
        self.TotalWork = value

    @property
    def UnRead(self):
        return self.taskitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Assign(self):
        return self.taskitem.Assign()

    def CancelResponseState(self):
        self.taskitem.CancelResponseState()

    def ClearRecurrencePattern(self):
        self.taskitem.ClearRecurrencePattern()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.taskitem.Close(*arguments)

    def Copy(self):
        self.taskitem.Copy()

    def Delete(self):
        self.taskitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.taskitem.Display(*arguments)

    def GetConversation(self):
        return self.taskitem.GetConversation()

    def GetRecurrencePattern(self):
        return self.taskitem.GetRecurrencePattern()

    def MarkComplete(self):
        self.taskitem.MarkComplete()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.taskitem.Move(*arguments)

    def PrintOut(self):
        self.taskitem.PrintOut()

    def Respond(self, Response=None, fNoUI=None, fAdditionalTextDialog=None):
        arguments = com_arguments([Response, fNoUI, fAdditionalTextDialog])
        return TaskItem(self.taskitem.Respond(*arguments))

    def Save(self):
        self.taskitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.taskitem.SaveAs(*arguments)

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

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.taskrequestacceptitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskrequestacceptitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.taskrequestacceptitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskrequestacceptitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskrequestacceptitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.taskrequestacceptitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.taskrequestacceptitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.taskrequestacceptitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskrequestacceptitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskrequestacceptitem.Class)

    @property
    def Companies(self):
        return self.taskrequestacceptitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskrequestacceptitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.taskrequestacceptitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.taskrequestacceptitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.taskrequestacceptitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskrequestacceptitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskrequestacceptitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskrequestacceptitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.taskrequestacceptitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskrequestacceptitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.taskrequestacceptitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.taskrequestacceptitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.taskrequestacceptitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.taskrequestacceptitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskrequestacceptitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.taskrequestacceptitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskrequestacceptitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskrequestacceptitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskrequestacceptitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskrequestacceptitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.taskrequestacceptitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskrequestacceptitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.taskrequestacceptitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskrequestacceptitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.taskrequestacceptitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskrequestacceptitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.taskrequestacceptitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskrequestacceptitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.taskrequestacceptitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskrequestacceptitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.taskrequestacceptitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskrequestacceptitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskrequestacceptitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskrequestacceptitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.taskrequestacceptitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.taskrequestacceptitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskrequestacceptitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.taskrequestacceptitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskrequestacceptitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskrequestacceptitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.taskrequestacceptitem.Close(*arguments)

    def Copy(self):
        self.taskrequestacceptitem.Copy()

    def Delete(self):
        self.taskrequestacceptitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.taskrequestacceptitem.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([AddToTaskList])
        return self.taskrequestacceptitem.GetAssociatedTask(*arguments)

    def GetConversation(self):
        return self.taskrequestacceptitem.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.taskrequestacceptitem.Move(*arguments)

    def PrintOut(self):
        self.taskrequestacceptitem.PrintOut()

    def Save(self):
        self.taskrequestacceptitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.taskrequestacceptitem.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.taskrequestacceptitem.ShowCategoriesDialog()


class TaskRequestDeclineItem:

    def __init__(self, taskrequestdeclineitem=None):
        self.taskrequestdeclineitem = taskrequestdeclineitem

    @property
    def Actions(self):
        return Actions(self.taskrequestdeclineitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.taskrequestdeclineitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskrequestdeclineitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.taskrequestdeclineitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskrequestdeclineitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskrequestdeclineitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.taskrequestdeclineitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.taskrequestdeclineitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.taskrequestdeclineitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskrequestdeclineitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskrequestdeclineitem.Class)

    @property
    def Companies(self):
        return self.taskrequestdeclineitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskrequestdeclineitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.taskrequestdeclineitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.taskrequestdeclineitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.taskrequestdeclineitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskrequestdeclineitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskrequestdeclineitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskrequestdeclineitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.taskrequestdeclineitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskrequestdeclineitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.taskrequestdeclineitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.taskrequestdeclineitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.taskrequestdeclineitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.taskrequestdeclineitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskrequestdeclineitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.taskrequestdeclineitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskrequestdeclineitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskrequestdeclineitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskrequestdeclineitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskrequestdeclineitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.taskrequestdeclineitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskrequestdeclineitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.taskrequestdeclineitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskrequestdeclineitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.taskrequestdeclineitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskrequestdeclineitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.taskrequestdeclineitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskrequestdeclineitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.taskrequestdeclineitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskrequestdeclineitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.taskrequestdeclineitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskrequestdeclineitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskrequestdeclineitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskrequestdeclineitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.taskrequestdeclineitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.taskrequestdeclineitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskrequestdeclineitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.taskrequestdeclineitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskrequestdeclineitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskrequestdeclineitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.taskrequestdeclineitem.Close(*arguments)

    def Copy(self):
        self.taskrequestdeclineitem.Copy()

    def Delete(self):
        self.taskrequestdeclineitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.taskrequestdeclineitem.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([AddToTaskList])
        return self.taskrequestdeclineitem.GetAssociatedTask(*arguments)

    def GetConversation(self):
        return self.taskrequestdeclineitem.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.taskrequestdeclineitem.Move(*arguments)

    def PrintOut(self):
        self.taskrequestdeclineitem.PrintOut()

    def Save(self):
        self.taskrequestdeclineitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.taskrequestdeclineitem.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.taskrequestdeclineitem.ShowCategoriesDialog()


class TaskRequestItem:

    def __init__(self, taskrequestitem=None):
        self.taskrequestitem = taskrequestitem

    @property
    def Actions(self):
        return Actions(self.taskrequestitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.taskrequestitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskrequestitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.taskrequestitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskrequestitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskrequestitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.taskrequestitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.taskrequestitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.taskrequestitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskrequestitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskrequestitem.Class)

    @property
    def Companies(self):
        return self.taskrequestitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskrequestitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.taskrequestitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.taskrequestitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.taskrequestitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskrequestitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskrequestitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskrequestitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.taskrequestitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskrequestitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.taskrequestitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.taskrequestitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.taskrequestitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.taskrequestitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskrequestitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.taskrequestitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskrequestitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskrequestitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskrequestitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskrequestitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.taskrequestitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskrequestitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.taskrequestitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskrequestitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.taskrequestitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskrequestitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.taskrequestitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskrequestitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.taskrequestitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskrequestitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.taskrequestitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskrequestitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskrequestitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskrequestitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.taskrequestitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.taskrequestitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskrequestitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.taskrequestitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskrequestitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskrequestitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.taskrequestitem.Close(*arguments)

    def Copy(self):
        self.taskrequestitem.Copy()

    def Delete(self):
        self.taskrequestitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.taskrequestitem.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([AddToTaskList])
        return self.taskrequestitem.GetAssociatedTask(*arguments)

    def GetConversation(self):
        return self.taskrequestitem.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.taskrequestitem.Move(*arguments)

    def PrintOut(self):
        self.taskrequestitem.PrintOut()

    def Save(self):
        self.taskrequestitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.taskrequestitem.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.taskrequestitem.ShowCategoriesDialog()


class TaskRequestUpdateItem:

    def __init__(self, taskrequestupdateitem=None):
        self.taskrequestupdateitem = taskrequestupdateitem

    @property
    def Actions(self):
        return Actions(self.taskrequestupdateitem.Actions)

    # Lower case alias for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.taskrequestupdateitem.Application)

    @property
    def Attachments(self):
        return Attachments(self.taskrequestupdateitem.Attachments)

    # Lower case alias for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.taskrequestupdateitem.AutoResolvedWinner

    # Lower case alias for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.taskrequestupdateitem.BillingInformation

    # Lower case alias for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.taskrequestupdateitem.BillingInformation = value

    # Lower case alias for BillingInformation setter
    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.taskrequestupdateitem.Body

    # Lower case alias for Body
    @property
    def body(self):
        return self.Body

    @Body.setter
    def Body(self, value):
        self.taskrequestupdateitem.Body = value

    # Lower case alias for Body setter
    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.taskrequestupdateitem.Categories

    # Lower case alias for Categories
    @property
    def categories(self):
        return self.Categories

    @Categories.setter
    def Categories(self, value):
        self.taskrequestupdateitem.Categories = value

    # Lower case alias for Categories setter
    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.taskrequestupdateitem.Class)

    @property
    def Companies(self):
        return self.taskrequestupdateitem.Companies

    # Lower case alias for Companies
    @property
    def companies(self):
        return self.Companies

    @Companies.setter
    def Companies(self, value):
        self.taskrequestupdateitem.Companies = value

    # Lower case alias for Companies setter
    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.taskrequestupdateitem.Conflicts

    # Lower case alias for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.taskrequestupdateitem.ConversationID)

    # Lower case alias for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.taskrequestupdateitem.ConversationIndex

    # Lower case alias for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.taskrequestupdateitem.ConversationTopic

    # Lower case alias for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.taskrequestupdateitem.CreationTime

    # Lower case alias for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.taskrequestupdateitem.DownloadState)

    # Lower case alias for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.taskrequestupdateitem.EntryID

    # Lower case alias for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.taskrequestupdateitem.FormDescription)

    # Lower case alias for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.taskrequestupdateitem.GetInspector)

    # Lower case alias for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.taskrequestupdateitem.Importance)

    # Lower case alias for Importance
    @property
    def importance(self):
        return self.Importance

    @Importance.setter
    def Importance(self, value):
        self.taskrequestupdateitem.Importance = value

    # Lower case alias for Importance setter
    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.taskrequestupdateitem.IsConflict

    # Lower case alias for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.taskrequestupdateitem.ItemProperties)

    # Lower case alias for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.taskrequestupdateitem.LastModificationTime

    # Lower case alias for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.taskrequestupdateitem.MarkForDownload)

    # Lower case alias for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.taskrequestupdateitem.MarkForDownload = value

    # Lower case alias for MarkForDownload setter
    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.taskrequestupdateitem.MessageClass

    # Lower case alias for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.taskrequestupdateitem.MessageClass = value

    # Lower case alias for MessageClass setter
    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.taskrequestupdateitem.Mileage

    # Lower case alias for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.taskrequestupdateitem.Mileage = value

    # Lower case alias for Mileage setter
    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.taskrequestupdateitem.NoAging

    # Lower case alias for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.taskrequestupdateitem.NoAging = value

    # Lower case alias for NoAging setter
    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.taskrequestupdateitem.OutlookInternalVersion

    # Lower case alias for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.taskrequestupdateitem.OutlookVersion

    # Lower case alias for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.taskrequestupdateitem.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.taskrequestupdateitem.PropertyAccessor)

    # Lower case alias for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.taskrequestupdateitem.RTFBody

    # Lower case alias for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.taskrequestupdateitem.RTFBody = value

    # Lower case alias for RTFBody setter
    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.taskrequestupdateitem.Saved

    # Lower case alias for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.taskrequestupdateitem.Sensitivity)

    # Lower case alias for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.taskrequestupdateitem.Sensitivity = value

    # Lower case alias for Sensitivity setter
    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.taskrequestupdateitem.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.taskrequestupdateitem.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.taskrequestupdateitem.Subject

    # Lower case alias for Subject
    @property
    def subject(self):
        return self.Subject

    @Subject.setter
    def Subject(self, value):
        self.taskrequestupdateitem.Subject = value

    # Lower case alias for Subject setter
    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.taskrequestupdateitem.UnRead

    # Lower case alias for UnRead
    @property
    def unread(self):
        return self.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.taskrequestupdateitem.UnRead = value

    # Lower case alias for UnRead setter
    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.taskrequestupdateitem.UserProperties)

    # Lower case alias for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.taskrequestupdateitem.Close(*arguments)

    def Copy(self):
        self.taskrequestupdateitem.Copy()

    def Delete(self):
        self.taskrequestupdateitem.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.taskrequestupdateitem.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([AddToTaskList])
        return self.taskrequestupdateitem.GetAssociatedTask(*arguments)

    def GetConversation(self):
        return self.taskrequestupdateitem.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return self.taskrequestupdateitem.Move(*arguments)

    def PrintOut(self):
        self.taskrequestupdateitem.PrintOut()

    def Save(self):
        self.taskrequestupdateitem.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.taskrequestupdateitem.SaveAs(*arguments)

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

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.tasksmodule.NavigationGroups)

    # Lower case alias for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.tasksmodule.NavigationModuleType)

    # Lower case alias for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.tasksmodule.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return TasksModule(self.tasksmodule.Position)

    # Lower case alias for Position
    @property
    def position(self):
        return self.Position

    @Position.setter
    def Position(self, value):
        self.tasksmodule.Position = value

    # Lower case alias for Position setter
    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.tasksmodule.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return TasksModule(self.tasksmodule.Visible)

    # Lower case alias for Visible
    @property
    def visible(self):
        return self.Visible

    @Visible.setter
    def Visible(self, value):
        self.tasksmodule.Visible = value

    # Lower case alias for Visible setter
    @visible.setter
    def visible(self, value):
        self.Visible = value


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

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.textrulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.textrulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.textrulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.textrulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Text(self):
        return self.textrulecondition.Text

    # Lower case alias for Text
    @property
    def text(self):
        return self.Text

    @Text.setter
    def Text(self, value):
        self.textrulecondition.Text = value

    # Lower case alias for Text setter
    @text.setter
    def text(self, value):
        self.Text = value


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

    # Lower case alias for DefaultExpandCollapseSetting
    @property
    def defaultexpandcollapsesetting(self):
        return self.DefaultExpandCollapseSetting

    @DefaultExpandCollapseSetting.setter
    def DefaultExpandCollapseSetting(self, value):
        self.timelineview.DefaultExpandCollapseSetting = value

    # Lower case alias for DefaultExpandCollapseSetting setter
    @defaultexpandcollapsesetting.setter
    def defaultexpandcollapsesetting(self, value):
        self.DefaultExpandCollapseSetting = value

    @property
    def EndField(self):
        return TimelineView(self.timelineview.EndField)

    # Lower case alias for EndField
    @property
    def endfield(self):
        return self.EndField

    @EndField.setter
    def EndField(self, value):
        self.timelineview.EndField = value

    # Lower case alias for EndField setter
    @endfield.setter
    def endfield(self, value):
        self.EndField = value

    @property
    def Filter(self):
        return self.timelineview.Filter

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @Filter.setter
    def Filter(self, value):
        self.timelineview.Filter = value

    # Lower case alias for Filter setter
    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def GroupByFields(self):
        return OrderFields(self.timelineview.GroupByFields)

    # Lower case alias for GroupByFields
    @property
    def groupbyfields(self):
        return self.GroupByFields

    @property
    def ItemFont(self):
        return ViewFont(self.timelineview.ItemFont)

    # Lower case alias for ItemFont
    @property
    def itemfont(self):
        return self.ItemFont

    @property
    def Language(self):
        return self.timelineview.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @Language.setter
    def Language(self, value):
        self.timelineview.Language = value

    # Lower case alias for Language setter
    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.timelineview.LockUserChanges

    # Lower case alias for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.timelineview.LockUserChanges = value

    # Lower case alias for LockUserChanges setter
    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def LowerScaleFont(self):
        return ViewFont(self.timelineview.LowerScaleFont)

    # Lower case alias for LowerScaleFont
    @property
    def lowerscalefont(self):
        return self.LowerScaleFont

    @property
    def MaxLabelWidth(self):
        return TimelineView(self.timelineview.MaxLabelWidth)

    # Lower case alias for MaxLabelWidth
    @property
    def maxlabelwidth(self):
        return self.MaxLabelWidth

    @MaxLabelWidth.setter
    def MaxLabelWidth(self, value):
        self.timelineview.MaxLabelWidth = value

    # Lower case alias for MaxLabelWidth setter
    @maxlabelwidth.setter
    def maxlabelwidth(self, value):
        self.MaxLabelWidth = value

    @property
    def Name(self):
        return self.timelineview.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.timelineview.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.timelineview.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.timelineview.SaveOption)

    # Lower case alias for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.timelineview.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowLabelWhenViewingByMonth(self):
        return TimelineView(self.timelineview.ShowLabelWhenViewingByMonth)

    # Lower case alias for ShowLabelWhenViewingByMonth
    @property
    def showlabelwhenviewingbymonth(self):
        return self.ShowLabelWhenViewingByMonth

    @ShowLabelWhenViewingByMonth.setter
    def ShowLabelWhenViewingByMonth(self, value):
        self.timelineview.ShowLabelWhenViewingByMonth = value

    # Lower case alias for ShowLabelWhenViewingByMonth setter
    @showlabelwhenviewingbymonth.setter
    def showlabelwhenviewingbymonth(self, value):
        self.ShowLabelWhenViewingByMonth = value

    @property
    def ShowWeekNumbers(self):
        return TimelineView(self.timelineview.ShowWeekNumbers)

    # Lower case alias for ShowWeekNumbers
    @property
    def showweeknumbers(self):
        return self.ShowWeekNumbers

    @ShowWeekNumbers.setter
    def ShowWeekNumbers(self, value):
        self.timelineview.ShowWeekNumbers = value

    # Lower case alias for ShowWeekNumbers setter
    @showweeknumbers.setter
    def showweeknumbers(self, value):
        self.ShowWeekNumbers = value

    @property
    def Standard(self):
        return TimelineView(self.timelineview.Standard)

    # Lower case alias for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def StartField(self):
        return TimelineView(self.timelineview.StartField)

    # Lower case alias for StartField
    @property
    def startfield(self):
        return self.StartField

    @StartField.setter
    def StartField(self, value):
        self.timelineview.StartField = value

    # Lower case alias for StartField setter
    @startfield.setter
    def startfield(self, value):
        self.StartField = value

    @property
    def TimelineViewMode(self):
        return OlTimelineViewMode(self.timelineview.TimelineViewMode)

    # Lower case alias for TimelineViewMode
    @property
    def timelineviewmode(self):
        return self.TimelineViewMode

    @TimelineViewMode.setter
    def TimelineViewMode(self, value):
        self.timelineview.TimelineViewMode = value

    # Lower case alias for TimelineViewMode setter
    @timelineviewmode.setter
    def timelineviewmode(self, value):
        self.TimelineViewMode = value

    @property
    def UpperScaleFont(self):
        return ViewFont(self.timelineview.UpperScaleFont)

    # Lower case alias for UpperScaleFont
    @property
    def upperscalefont(self):
        return self.UpperScaleFont

    @property
    def ViewType(self):
        return OlViewType(self.timelineview.ViewType)

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.timelineview.XML

    # Lower case alias for XML
    @property
    def xml(self):
        return self.XML

    @XML.setter
    def XML(self, value):
        self.timelineview.XML = value

    # Lower case alias for XML setter
    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.timelineview.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return self.timelineview.Copy(*arguments)

    def Delete(self):
        self.timelineview.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.timelineview.GoToDate(*arguments)

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

    # Lower case alias for Bias
    @property
    def bias(self):
        return self.Bias

    @property
    def Class(self):
        return OlObjectClass(self.timezone.Class)

    @property
    def DaylightBias(self):
        return self.timezone.DaylightBias

    # Lower case alias for DaylightBias
    @property
    def daylightbias(self):
        return self.DaylightBias

    @property
    def DaylightDate(self):
        return self.timezone.DaylightDate

    # Lower case alias for DaylightDate
    @property
    def daylightdate(self):
        return self.DaylightDate

    @property
    def DaylightDesignation(self):
        return self.timezone.DaylightDesignation

    # Lower case alias for DaylightDesignation
    @property
    def daylightdesignation(self):
        return self.DaylightDesignation

    @property
    def ID(self):
        return self.timezone.ID

    # Lower case alias for ID
    @property
    def id(self):
        return self.ID

    @property
    def Name(self):
        return self.timezone.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.timezone.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.timezone.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def StandardBias(self):
        return self.timezone.StandardBias

    # Lower case alias for StandardBias
    @property
    def standardbias(self):
        return self.StandardBias

    @property
    def StandardDate(self):
        return self.timezone.StandardDate

    # Lower case alias for StandardDate
    @property
    def standarddate(self):
        return self.StandardDate

    @property
    def StandardDesignation(self):
        return self.timezone.StandardDesignation

    # Lower case alias for StandardDesignation
    @property
    def standarddesignation(self):
        return self.StandardDesignation


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def CurrentTimeZone(self):
        return TimeZone(self.timezones.CurrentTimeZone)

    # Lower case alias for CurrentTimeZone
    @property
    def currenttimezone(self):
        return self.CurrentTimeZone

    @property
    def Parent(self):
        return self.timezones.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.timezones.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def ConvertTime(self, SourceDateTime=None, SourceTimeZone=None, DestinationTimeZone=None):
        arguments = com_arguments([SourceDateTime, SourceTimeZone, DestinationTimeZone])
        return self.timezones.ConvertTime(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.timezones.Item(*arguments)


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

    # Lower case alias for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.toorfromrulecondition.Enabled

    # Lower case alias for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.toorfromrulecondition.Enabled = value

    # Lower case alias for Enabled setter
    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.toorfromrulecondition.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.toorfromrulecondition.Recipients)

    # Lower case alias for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def Session(self):
        return NameSpace(self.toorfromrulecondition.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.userdefinedproperties.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.userdefinedproperties.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Type=None, DisplayFormat=None, Formula=None):
        arguments = com_arguments([Name, Type, DisplayFormat, Formula])
        return self.userdefinedproperties.Add(*arguments)

    def Find(self, Name=None):
        arguments = com_arguments([Name])
        return self.userdefinedproperties.Find(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return UserDefinedProperty(self.userdefinedproperties.Item(*arguments))

    def Refresh(self):
        self.userdefinedproperties.Refresh()

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.userdefinedproperties.Remove(*arguments)


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

    # Lower case alias for DisplayFormat
    @property
    def displayformat(self):
        return self.DisplayFormat

    @property
    def Formula(self):
        return self.userdefinedproperty.Formula

    # Lower case alias for Formula
    @property
    def formula(self):
        return self.Formula

    @property
    def Name(self):
        return self.userdefinedproperty.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.userdefinedproperty.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.userdefinedproperty.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.userdefinedproperty.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.userproperties.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.userproperties.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = com_arguments([Name, Type, AddToFolderFields, DisplayFormat])
        return UserProperty(self.userproperties.Add(*arguments))

    def Find(self, Name=None, Custom=None):
        arguments = com_arguments([Name, Custom])
        return self.userproperties.Find(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.userproperties.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.userproperties.Remove(*arguments)


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

    # Lower case alias for Formula
    @property
    def formula(self):
        return self.Formula

    @Formula.setter
    def Formula(self, value):
        self.userproperty.Formula = value

    # Lower case alias for Formula setter
    @formula.setter
    def formula(self, value):
        self.Formula = value

    @property
    def Name(self):
        return self.userproperty.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.userproperty.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.userproperty.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.userproperty.Type)

    # Lower case alias for Type
    @property
    def type(self):
        return self.Type

    @property
    def ValidationFormula(self):
        return self.userproperty.ValidationFormula

    # Lower case alias for ValidationFormula
    @property
    def validationformula(self):
        return self.ValidationFormula

    @ValidationFormula.setter
    def ValidationFormula(self, value):
        self.userproperty.ValidationFormula = value

    # Lower case alias for ValidationFormula setter
    @validationformula.setter
    def validationformula(self, value):
        self.ValidationFormula = value

    @property
    def ValidationText(self):
        return self.userproperty.ValidationText

    # Lower case alias for ValidationText
    @property
    def validationtext(self):
        return self.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.userproperty.ValidationText = value

    # Lower case alias for ValidationText setter
    @validationtext.setter
    def validationtext(self, value):
        self.ValidationText = value

    @property
    def Value(self):
        return self.userproperty.Value

    # Lower case alias for Value
    @property
    def value(self):
        return self.Value

    @Value.setter
    def Value(self, value):
        self.userproperty.Value = value

    # Lower case alias for Value setter
    @value.setter
    def value(self, value):
        self.Value = value

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

    # Lower case alias for Filter
    @property
    def filter(self):
        return self.Filter

    @Filter.setter
    def Filter(self, value):
        self.view.Filter = value

    # Lower case alias for Filter setter
    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def Language(self):
        return self.view.Language

    # Lower case alias for Language
    @property
    def language(self):
        return self.Language

    @Language.setter
    def Language(self, value):
        self.view.Language = value

    # Lower case alias for Language setter
    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.view.LockUserChanges

    # Lower case alias for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.view.LockUserChanges = value

    # Lower case alias for LockUserChanges setter
    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.view.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.view.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.view.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.view.SaveOption)

    # Lower case alias for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.view.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Standard(self):
        return self.view.Standard

    # Lower case alias for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.view.ViewType)

    # Lower case alias for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.view.XML

    # Lower case alias for XML
    @property
    def xml(self):
        return self.XML

    @XML.setter
    def XML(self, value):
        self.view.XML = value

    # Lower case alias for XML setter
    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.view.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        self.view.Copy(*arguments)

    def Delete(self):
        self.view.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.view.GoToDate(*arguments)

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

    # Lower case alias for ColumnFormat
    @property
    def columnformat(self):
        return self.ColumnFormat

    @property
    def Parent(self):
        return self.viewfield.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.viewfield.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def ViewXMLSchemaName(self):
        return ViewField(self.viewfield.ViewXMLSchemaName)

    # Lower case alias for ViewXMLSchemaName
    @property
    def viewxmlschemaname(self):
        return self.ViewXMLSchemaName


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.viewfields.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.viewfields.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, PropertyName=None):
        arguments = com_arguments([PropertyName])
        return self.viewfields.Add(*arguments)

    def Insert(self, PropertyName=None, Index=None):
        arguments = com_arguments([PropertyName, Index])
        return self.viewfields.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.viewfields.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.viewfields.Remove(*arguments)


class ViewFont:

    def __init__(self, viewfont=None):
        self.viewfont = viewfont

    @property
    def Application(self):
        return Application(self.viewfont.Application)

    @property
    def Bold(self):
        return ViewFont(self.viewfont.Bold)

    # Lower case alias for Bold
    @property
    def bold(self):
        return self.Bold

    @Bold.setter
    def Bold(self, value):
        self.viewfont.Bold = value

    # Lower case alias for Bold setter
    @bold.setter
    def bold(self, value):
        self.Bold = value

    @property
    def Class(self):
        return OlObjectClass(self.viewfont.Class)

    @property
    def Color(self):
        return OlColor(self.viewfont.Color)

    # Lower case alias for Color
    @property
    def color(self):
        return self.Color

    @Color.setter
    def Color(self, value):
        self.viewfont.Color = value

    # Lower case alias for Color setter
    @color.setter
    def color(self, value):
        self.Color = value

    @property
    def ExtendedColor(self):
        return OlCategoryColor(self.viewfont.ExtendedColor)

    # Lower case alias for ExtendedColor
    @property
    def extendedcolor(self):
        return self.ExtendedColor

    @ExtendedColor.setter
    def ExtendedColor(self, value):
        self.viewfont.ExtendedColor = value

    # Lower case alias for ExtendedColor setter
    @extendedcolor.setter
    def extendedcolor(self, value):
        self.ExtendedColor = value

    @property
    def Italic(self):
        return ViewFont(self.viewfont.Italic)

    # Lower case alias for Italic
    @property
    def italic(self):
        return self.Italic

    @Italic.setter
    def Italic(self, value):
        self.viewfont.Italic = value

    # Lower case alias for Italic setter
    @italic.setter
    def italic(self, value):
        self.Italic = value

    @property
    def Name(self):
        return self.viewfont.Name

    # Lower case alias for Name
    @property
    def name(self):
        return self.Name

    @Name.setter
    def Name(self, value):
        self.viewfont.Name = value

    # Lower case alias for Name setter
    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.viewfont.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.viewfont.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.viewfont.Size

    # Lower case alias for Size
    @property
    def size(self):
        return self.Size

    @Size.setter
    def Size(self, value):
        self.viewfont.Size = value

    # Lower case alias for Size setter
    @size.setter
    def size(self, value):
        self.Size = value

    @property
    def Strikethrough(self):
        return ViewFont(self.viewfont.Strikethrough)

    # Lower case alias for Strikethrough
    @property
    def strikethrough(self):
        return self.Strikethrough

    @Strikethrough.setter
    def Strikethrough(self, value):
        self.viewfont.Strikethrough = value

    # Lower case alias for Strikethrough setter
    @strikethrough.setter
    def strikethrough(self, value):
        self.Strikethrough = value

    @property
    def Underline(self):
        return ViewFont(self.viewfont.Underline)

    # Lower case alias for Underline
    @property
    def underline(self):
        return self.Underline

    @Underline.setter
    def Underline(self, value):
        self.viewfont.Underline = value

    # Lower case alias for Underline setter
    @underline.setter
    def underline(self, value):
        self.Underline = value


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

    # Lower case alias for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.views.Parent

    # Lower case alias for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.views.Session)

    # Lower case alias for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, ViewType=None, SaveOption=None):
        arguments = com_arguments([Name, ViewType, SaveOption])
        return View(self.views.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.views.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.views.Remove(*arguments)

