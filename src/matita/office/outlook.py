from . import com_arguments
from .office import *

import win32com.client

class Account:

    def __init__(self, account=None):
        self.com_object= account

    @property
    def AccountType(self):
        return OlAccountType(self.com_object.AccountType)

    # Lower case aliases for AccountType
    @property
    def accounttype(self):
        return self.AccountType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoDiscoverConnectionMode(self):
        return OlAutoDiscoverConnectionMode(self.com_object.AutoDiscoverConnectionMode)

    # Lower case aliases for AutoDiscoverConnectionMode
    @property
    def autodiscoverconnectionmode(self):
        return self.AutoDiscoverConnectionMode

    @property
    def AutoDiscoverXml(self):
        return self.com_object.AutoDiscoverXml

    # Lower case aliases for AutoDiscoverXml
    @property
    def autodiscoverxml(self):
        return self.AutoDiscoverXml

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentUser(self):
        return Recipient(self.com_object.CurrentUser)

    # Lower case aliases for CurrentUser
    @property
    def currentuser(self):
        return self.CurrentUser

    @property
    def DeliveryStore(self):
        return Store(self.com_object.DeliveryStore)

    # Lower case aliases for DeliveryStore
    @property
    def deliverystore(self):
        return self.DeliveryStore

    @property
    def DisplayName(self):
        return Account(self.com_object.DisplayName)

    # Lower case aliases for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def ExchangeConnectionMode(self):
        return OlExchangeConnectionMode(self.com_object.ExchangeConnectionMode)

    # Lower case aliases for ExchangeConnectionMode
    @property
    def exchangeconnectionmode(self):
        return self.ExchangeConnectionMode

    @property
    def ExchangeMailboxServerName(self):
        return self.com_object.ExchangeMailboxServerName

    # Lower case aliases for ExchangeMailboxServerName
    @property
    def exchangemailboxservername(self):
        return self.ExchangeMailboxServerName

    @property
    def ExchangeMailboxServerVersion(self):
        return self.com_object.ExchangeMailboxServerVersion

    # Lower case aliases for ExchangeMailboxServerVersion
    @property
    def exchangemailboxserverversion(self):
        return self.ExchangeMailboxServerVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def SmtpAddress(self):
        return Account(self.com_object.SmtpAddress)

    # Lower case aliases for SmtpAddress
    @property
    def smtpaddress(self):
        return self.SmtpAddress

    @property
    def UserName(self):
        return Account(self.com_object.UserName)

    # Lower case aliases for UserName
    @property
    def username(self):
        return self.UserName

    def GetAddressEntryFromID(self, ID=None):
        arguments = com_arguments([ID])
        return ID(self.com_object.GetAddressEntryFromID(*arguments))

    # Lower case alias for GetAddressEntryFromID
    def getaddressentryfromid(self, ID=None):
        arguments = [ID]
        return self.GetAddressEntryFromID(*arguments)

    def GetRecipientFromID(self, EntryID=None):
        arguments = com_arguments([EntryID])
        return Recipient(self.com_object.GetRecipientFromID(*arguments))

    # Lower case alias for GetRecipientFromID
    def getrecipientfromid(self, EntryID=None):
        arguments = [EntryID]
        return self.GetRecipientFromID(*arguments)


class AccountRuleCondition:

    def __init__(self, accountrulecondition=None):
        self.com_object= accountrulecondition

    @property
    def Account(self):
        return Account(self.com_object.Account)

    @Account.setter
    def Account(self, value):
        self.com_object.Account = value

    # Lower case aliases for Account
    @property
    def account(self):
        return self.Account

    @account.setter
    def account(self, value):
        self.Account = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class Accounts:

    def __init__(self, accounts=None):
        self.com_object= accounts

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Account(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class AccountSelector:

    def __init__(self, accountselector=None):
        self.com_object= accountselector

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SelectedAccount(self):
        return Account(self.com_object.SelectedAccount)

    # Lower case aliases for SelectedAccount
    @property
    def selectedaccount(self):
        return self.SelectedAccount

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class Action:

    def __init__(self, action=None):
        self.com_object= action

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CopyLike(self):
        return OlActionCopyLike(self.com_object.CopyLike)

    @CopyLike.setter
    def CopyLike(self, value):
        self.com_object.CopyLike = value

    # Lower case aliases for CopyLike
    @property
    def copylike(self):
        return self.CopyLike

    @copylike.setter
    def copylike(self, value):
        self.CopyLike = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def MessageClass(self):
        return Action(self.com_object.MessageClass)

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Prefix(self):
        return self.com_object.Prefix

    @Prefix.setter
    def Prefix(self, value):
        self.com_object.Prefix = value

    # Lower case aliases for Prefix
    @property
    def prefix(self):
        return self.Prefix

    @prefix.setter
    def prefix(self, value):
        self.Prefix = value

    @property
    def ReplyStyle(self):
        return OlActionReplyStyle(self.com_object.ReplyStyle)

    @ReplyStyle.setter
    def ReplyStyle(self, value):
        self.com_object.ReplyStyle = value

    # Lower case aliases for ReplyStyle
    @property
    def replystyle(self):
        return self.ReplyStyle

    @replystyle.setter
    def replystyle(self, value):
        self.ReplyStyle = value

    @property
    def ResponseStyle(self):
        return OlActionResponseStyle(self.com_object.ResponseStyle)

    @ResponseStyle.setter
    def ResponseStyle(self, value):
        self.com_object.ResponseStyle = value

    # Lower case aliases for ResponseStyle
    @property
    def responsestyle(self):
        return self.ResponseStyle

    @responsestyle.setter
    def responsestyle(self, value):
        self.ResponseStyle = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowOn(self):
        return OlActionShowOn(self.com_object.ShowOn)

    @ShowOn.setter
    def ShowOn(self, value):
        self.com_object.ShowOn = value

    # Lower case aliases for ShowOn
    @property
    def showon(self):
        return self.ShowOn

    @showon.setter
    def showon(self, value):
        self.ShowOn = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Execute(self):
        return Object(self.com_object.Execute())

    # Lower case alias for Execute
    def execute(self):
        return self.Execute()


class Actions:

    def __init__(self, actions=None):
        self.com_object= actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self):
        return Action(self.com_object.Add())

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Action(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class AddressEntries:

    def __init__(self, addressentries=None):
        self.com_object= addressentries

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Type=None, Name=None, Address=None):
        arguments = com_arguments([Type, Name, Address])
        return AddressEntry(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Type=None, Name=None, Address=None):
        arguments = [Type, Name, Address]
        return self.Add(*arguments)

    def GetFirst(self):
        return AddressEntry(self.com_object.GetFirst())

    # Lower case alias for GetFirst
    def getfirst(self):
        return self.GetFirst()

    def GetLast(self):
        return AddressEntry(self.com_object.GetLast())

    # Lower case alias for GetLast
    def getlast(self):
        return self.GetLast()

    def GetNext(self):
        return AddressEntry(self.com_object.GetNext())

    # Lower case alias for GetNext
    def getnext(self):
        return self.GetNext()

    def GetPrevious(self):
        return AddressEntry(self.com_object.GetPrevious())

    # Lower case alias for GetPrevious
    def getprevious(self):
        return self.GetPrevious()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return AddressEntry(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Sort(self, Property=None, Order=None):
        arguments = com_arguments([Property, Order])
        self.com_object.Sort(*arguments)

    # Lower case alias for Sort
    def sort(self, Property=None, Order=None):
        arguments = [Property, Order]
        return self.Sort(*arguments)


class AddressEntry:

    def __init__(self, addressentry=None):
        self.com_object= addressentry

    @property
    def Address(self):
        return AddressEntry(self.com_object.Address)

    @Address.setter
    def Address(self, value):
        self.com_object.Address = value

    # Lower case aliases for Address
    @property
    def address(self):
        return self.Address

    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.com_object.AddressEntryUserType)

    # Lower case aliases for AddressEntryUserType
    @property
    def addressentryusertype(self):
        return self.AddressEntryUserType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DisplayType(self):
        return OlDisplayType(self.com_object.DisplayType)

    # Lower case aliases for DisplayType
    @property
    def displaytype(self):
        return self.DisplayType

    @property
    def ID(self):
        return self.com_object.ID

    # Lower case aliases for ID
    @property
    def id(self):
        return self.ID

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return self.com_object.Type

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @type.setter
    def type(self, value):
        self.Type = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([HWnd])
        self.com_object.Details(*arguments)

    # Lower case alias for Details
    def details(self, HWnd=None):
        arguments = [HWnd]
        return self.Details(*arguments)

    def GetContact(self):
        return self.com_object.GetContact()

    # Lower case alias for GetContact
    def getcontact(self):
        return self.GetContact()

    def GetExchangeDistributionList(self):
        return self.com_object.GetExchangeDistributionList()

    # Lower case alias for GetExchangeDistributionList
    def getexchangedistributionlist(self):
        return self.GetExchangeDistributionList()

    def GetExchangeUser(self):
        return self.com_object.GetExchangeUser()

    # Lower case alias for GetExchangeUser
    def getexchangeuser(self):
        return self.GetExchangeUser()

    def GetFreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([Start, MinPerChar, CompleteFormat])
        return self.com_object.GetFreeBusy(*arguments)

    # Lower case alias for GetFreeBusy
    def getfreebusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = [Start, MinPerChar, CompleteFormat]
        return self.GetFreeBusy(*arguments)

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([MakePermanent, Refresh])
        self.com_object.Update(*arguments)

    # Lower case alias for Update
    def update(self, MakePermanent=None, Refresh=None):
        arguments = [MakePermanent, Refresh]
        return self.Update(*arguments)


class AddressList:

    def __init__(self, addresslist=None):
        self.com_object= addresslist

    @property
    def AddressEntries(self):
        return AddressEntries(self.com_object.AddressEntries)

    # Lower case aliases for AddressEntries
    @property
    def addressentries(self):
        return self.AddressEntries

    @property
    def AddressListType(self):
        return OlAddressListType(self.com_object.AddressListType)

    # Lower case aliases for AddressListType
    @property
    def addresslisttype(self):
        return self.AddressListType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ID(self):
        return self.com_object.ID

    # Lower case aliases for ID
    @property
    def id(self):
        return self.ID

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def IsInitialAddressList(self):
        return AddressList(self.com_object.IsInitialAddressList)

    # Lower case aliases for IsInitialAddressList
    @property
    def isinitialaddresslist(self):
        return self.IsInitialAddressList

    @property
    def IsReadOnly(self):
        return AddressList(self.com_object.IsReadOnly)

    # Lower case aliases for IsReadOnly
    @property
    def isreadonly(self):
        return self.IsReadOnly

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ResolutionOrder(self):
        return AddressList(self.com_object.ResolutionOrder)

    # Lower case aliases for ResolutionOrder
    @property
    def resolutionorder(self):
        return self.ResolutionOrder

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def GetContactsFolder(self):
        return self.com_object.GetContactsFolder()

    # Lower case alias for GetContactsFolder
    def getcontactsfolder(self):
        return self.GetContactsFolder()


class AddressLists:

    def __init__(self, addresslists=None):
        self.com_object= addresslists

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return AddressList(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class AddressRuleCondition:

    def __init__(self, addressrulecondition=None):
        self.com_object= addressrulecondition

    @property
    def Address(self):
        return self.com_object.Address

    @Address.setter
    def Address(self, value):
        self.com_object.Address = value

    # Lower case aliases for Address
    @property
    def address(self):
        return self.Address

    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class Application:

    def __init__(self, application=None):
        self.com_object= application

    def new(self):
        self.com_object = win32com.client.Dispatch("Outlook.Application")
        return self

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Assistance(self):
        return self.com_object.Assistance

    # Lower case aliases for Assistance
    @property
    def assistance(self):
        return self.Assistance

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def COMAddIns(self):
        return self.com_object.COMAddIns

    # Lower case aliases for COMAddIns
    @property
    def comaddins(self):
        return self.COMAddIns

    @property
    def DefaultProfileName(self):
        return self.com_object.DefaultProfileName

    # Lower case aliases for DefaultProfileName
    @property
    def defaultprofilename(self):
        return self.DefaultProfileName

    @property
    def Explorers(self):
        return Explorers(self.com_object.Explorers)

    # Lower case aliases for Explorers
    @property
    def explorers(self):
        return self.Explorers

    @property
    def Inspectors(self):
        return Inspectors(self.com_object.Inspectors)

    # Lower case aliases for Inspectors
    @property
    def inspectors(self):
        return self.Inspectors

    @property
    def IsTrusted(self):
        return self.com_object.IsTrusted

    # Lower case aliases for IsTrusted
    @property
    def istrusted(self):
        return self.IsTrusted

    @property
    def LanguageSettings(self):
        return self.com_object.LanguageSettings

    # Lower case aliases for LanguageSettings
    @property
    def languagesettings(self):
        return self.LanguageSettings

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PickerDialog(self):
        return self.com_object.PickerDialog

    # Lower case aliases for PickerDialog
    @property
    def pickerdialog(self):
        return self.PickerDialog

    @property
    def ProductCode(self):
        return self.com_object.ProductCode

    # Lower case aliases for ProductCode
    @property
    def productcode(self):
        return self.ProductCode

    @property
    def Reminders(self):
        return Reminders(self.com_object.Reminders)

    # Lower case aliases for Reminders
    @property
    def reminders(self):
        return self.Reminders

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def TimeZones(self):
        return TimeZones(self.com_object.TimeZones)

    # Lower case aliases for TimeZones
    @property
    def timezones(self):
        return self.TimeZones

    @property
    def Version(self):
        return self.com_object.Version

    @Version.setter
    def Version(self, value):
        self.com_object.Version = value

    # Lower case aliases for Version
    @property
    def version(self):
        return self.Version

    @version.setter
    def version(self, value):
        self.Version = value

    def ActiveExplorer(self):
        return self.com_object.ActiveExplorer()

    # Lower case alias for ActiveExplorer
    def activeexplorer(self):
        return self.ActiveExplorer()

    def ActiveInspector(self):
        return Inspector(self.com_object.ActiveInspector())

    # Lower case alias for ActiveInspector
    def activeinspector(self):
        return self.ActiveInspector()

    def ActiveWindow(self):
        return self.com_object.ActiveWindow()

    # Lower case alias for ActiveWindow
    def activewindow(self):
        return self.ActiveWindow()

    def AdvancedSearch(self, Scope=None, Filter=None, SearchSubFolders=None, Tag=None):
        arguments = com_arguments([Scope, Filter, SearchSubFolders, Tag])
        return Search(self.com_object.AdvancedSearch(*arguments))

    # Lower case alias for AdvancedSearch
    def advancedsearch(self, Scope=None, Filter=None, SearchSubFolders=None, Tag=None):
        arguments = [Scope, Filter, SearchSubFolders, Tag]
        return self.AdvancedSearch(*arguments)

    def CopyFile(self, FilePath=None, DestFolderPath=None):
        arguments = com_arguments([FilePath, DestFolderPath])
        return Object(self.com_object.CopyFile(*arguments))

    # Lower case alias for CopyFile
    def copyfile(self, FilePath=None, DestFolderPath=None):
        arguments = [FilePath, DestFolderPath]
        return self.CopyFile(*arguments)

    def CreateItem(self, ItemType=None):
        arguments = com_arguments([ItemType])
        return Object(self.com_object.CreateItem(*arguments))

    # Lower case alias for CreateItem
    def createitem(self, ItemType=None):
        arguments = [ItemType]
        return self.CreateItem(*arguments)

    def CreateItemFromTemplate(self, TemplatePath=None, InFolder=None):
        arguments = com_arguments([TemplatePath, InFolder])
        return Object(self.com_object.CreateItemFromTemplate(*arguments))

    # Lower case alias for CreateItemFromTemplate
    def createitemfromtemplate(self, TemplatePath=None, InFolder=None):
        arguments = [TemplatePath, InFolder]
        return self.CreateItemFromTemplate(*arguments)

    def CreateObject(self, ObjectName=None):
        arguments = com_arguments([ObjectName])
        return CreateObject(self.com_object.CreateObject(*arguments))

    # Lower case alias for CreateObject
    def createobject(self, ObjectName=None):
        arguments = [ObjectName]
        return self.CreateObject(*arguments)

    def GetNamespace(self, Type=None):
        arguments = com_arguments([Type])
        return NameSpace(self.com_object.GetNamespace(*arguments))

    # Lower case alias for GetNamespace
    def getnamespace(self, Type=None):
        arguments = [Type]
        return self.GetNamespace(*arguments)

    def GetObjectReference(self, Item=None, ReferenceType=None):
        arguments = com_arguments([Item, ReferenceType])
        return Object(self.com_object.GetObjectReference(*arguments))

    # Lower case alias for GetObjectReference
    def getobjectreference(self, Item=None, ReferenceType=None):
        arguments = [Item, ReferenceType]
        return self.GetObjectReference(*arguments)

    def IsSearchSynchronous(self, LookInFolders=None):
        arguments = com_arguments([LookInFolders])
        return self.com_object.IsSearchSynchronous(*arguments)

    # Lower case alias for IsSearchSynchronous
    def issearchsynchronous(self, LookInFolders=None):
        arguments = [LookInFolders]
        return self.IsSearchSynchronous(*arguments)

    def Quit(self):
        self.com_object.Quit()

    # Lower case alias for Quit
    def quit(self):
        return self.Quit()

    def RefreshFormRegionDefinition(self, RegionName=None):
        arguments = com_arguments([RegionName])
        self.com_object.RefreshFormRegionDefinition(*arguments)

    # Lower case alias for RefreshFormRegionDefinition
    def refreshformregiondefinition(self, RegionName=None):
        arguments = [RegionName]
        return self.RefreshFormRegionDefinition(*arguments)


class AppointmentItem:

    def __init__(self, appointmentitem=None):
        self.com_object= appointmentitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def AllDayEvent(self):
        return self.com_object.AllDayEvent

    @AllDayEvent.setter
    def AllDayEvent(self, value):
        self.com_object.AllDayEvent = value

    # Lower case aliases for AllDayEvent
    @property
    def alldayevent(self):
        return self.AllDayEvent

    @alldayevent.setter
    def alldayevent(self, value):
        self.AllDayEvent = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def BusyStatus(self):
        return OlBusyStatus(self.com_object.BusyStatus)

    @BusyStatus.setter
    def BusyStatus(self, value):
        self.com_object.BusyStatus = value

    # Lower case aliases for BusyStatus
    @property
    def busystatus(self):
        return self.BusyStatus

    @busystatus.setter
    def busystatus(self, value):
        self.BusyStatus = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def Duration(self):
        return AppointmentItem(self.com_object.Duration)

    @Duration.setter
    def Duration(self, value):
        self.com_object.Duration = value

    # Lower case aliases for Duration
    @property
    def duration(self):
        return self.Duration

    @duration.setter
    def duration(self, value):
        self.Duration = value

    @property
    def End(self):
        return AppointmentItem(self.com_object.End)

    @End.setter
    def End(self, value):
        self.com_object.End = value

    # Lower case aliases for End
    @property
    def end(self):
        return self.End

    @end.setter
    def end(self, value):
        self.End = value

    @property
    def EndInEndTimeZone(self):
        return AppointmentItem.EndTimeZone(self.com_object.EndInEndTimeZone)

    @EndInEndTimeZone.setter
    def EndInEndTimeZone(self, value):
        self.com_object.EndInEndTimeZone = value

    # Lower case aliases for EndInEndTimeZone
    @property
    def endinendtimezone(self):
        return self.EndInEndTimeZone

    @endinendtimezone.setter
    def endinendtimezone(self, value):
        self.EndInEndTimeZone = value

    @property
    def EndTimeZone(self):
        return TimeZone(self.com_object.EndTimeZone)

    @EndTimeZone.setter
    def EndTimeZone(self, value):
        self.com_object.EndTimeZone = value

    # Lower case aliases for EndTimeZone
    @property
    def endtimezone(self):
        return self.EndTimeZone

    @endtimezone.setter
    def endtimezone(self, value):
        self.EndTimeZone = value

    @property
    def EndUTC(self):
        return self.com_object.EndUTC

    @EndUTC.setter
    def EndUTC(self, value):
        self.com_object.EndUTC = value

    # Lower case aliases for EndUTC
    @property
    def endutc(self):
        return self.EndUTC

    @endutc.setter
    def endutc(self, value):
        self.EndUTC = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ForceUpdateToAllAttendees(self):
        return self.com_object.ForceUpdateToAllAttendees

    @ForceUpdateToAllAttendees.setter
    def ForceUpdateToAllAttendees(self, value):
        self.com_object.ForceUpdateToAllAttendees = value

    # Lower case aliases for ForceUpdateToAllAttendees
    @property
    def forceupdatetoallattendees(self):
        return self.ForceUpdateToAllAttendees

    @forceupdatetoallattendees.setter
    def forceupdatetoallattendees(self, value):
        self.ForceUpdateToAllAttendees = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def GlobalAppointmentID(self):
        return AppointmentItem(self.com_object.GlobalAppointmentID)

    # Lower case aliases for GlobalAppointmentID
    @property
    def globalappointmentid(self):
        return self.GlobalAppointmentID

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    # Lower case aliases for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsRecurring(self):
        return self.com_object.IsRecurring

    # Lower case aliases for IsRecurring
    @property
    def isrecurring(self):
        return self.IsRecurring

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def Location(self):
        return self.com_object.Location

    @Location.setter
    def Location(self, value):
        self.com_object.Location = value

    # Lower case aliases for Location
    @property
    def location(self):
        return self.Location

    @location.setter
    def location(self, value):
        self.Location = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MeetingStatus(self):
        return OlMeetingStatus(self.com_object.MeetingStatus)

    @MeetingStatus.setter
    def MeetingStatus(self, value):
        self.com_object.MeetingStatus = value

    # Lower case aliases for MeetingStatus
    @property
    def meetingstatus(self):
        return self.MeetingStatus

    @meetingstatus.setter
    def meetingstatus(self, value):
        self.MeetingStatus = value

    @property
    def MeetingWorkspaceURL(self):
        return self.com_object.MeetingWorkspaceURL

    # Lower case aliases for MeetingWorkspaceURL
    @property
    def meetingworkspaceurl(self):
        return self.MeetingWorkspaceURL

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OptionalAttendees(self):
        return self.com_object.OptionalAttendees

    @OptionalAttendees.setter
    def OptionalAttendees(self, value):
        self.com_object.OptionalAttendees = value

    # Lower case aliases for OptionalAttendees
    @property
    def optionalattendees(self):
        return self.OptionalAttendees

    @optionalattendees.setter
    def optionalattendees(self, value):
        self.OptionalAttendees = value

    @property
    def Organizer(self):
        return self.com_object.Organizer

    # Lower case aliases for Organizer
    @property
    def organizer(self):
        return self.Organizer

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def RecurrenceState(self):
        return OlRecurrenceState(self.com_object.RecurrenceState)

    # Lower case aliases for RecurrenceState
    @property
    def recurrencestate(self):
        return self.RecurrenceState

    @property
    def ReminderMinutesBeforeStart(self):
        return self.com_object.ReminderMinutesBeforeStart

    @ReminderMinutesBeforeStart.setter
    def ReminderMinutesBeforeStart(self, value):
        self.com_object.ReminderMinutesBeforeStart = value

    # Lower case aliases for ReminderMinutesBeforeStart
    @property
    def reminderminutesbeforestart(self):
        return self.ReminderMinutesBeforeStart

    @reminderminutesbeforestart.setter
    def reminderminutesbeforestart(self, value):
        self.ReminderMinutesBeforeStart = value

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    # Lower case aliases for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    # Lower case aliases for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    # Lower case aliases for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    # Lower case aliases for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReplyTime(self):
        return self.com_object.ReplyTime

    @ReplyTime.setter
    def ReplyTime(self, value):
        self.com_object.ReplyTime = value

    # Lower case aliases for ReplyTime
    @property
    def replytime(self):
        return self.ReplyTime

    @replytime.setter
    def replytime(self, value):
        self.ReplyTime = value

    @property
    def RequiredAttendees(self):
        return self.com_object.RequiredAttendees

    @RequiredAttendees.setter
    def RequiredAttendees(self, value):
        self.com_object.RequiredAttendees = value

    # Lower case aliases for RequiredAttendees
    @property
    def requiredattendees(self):
        return self.RequiredAttendees

    @requiredattendees.setter
    def requiredattendees(self, value):
        self.RequiredAttendees = value

    @property
    def Resources(self):
        return self.com_object.Resources

    @Resources.setter
    def Resources(self, value):
        self.com_object.Resources = value

    # Lower case aliases for Resources
    @property
    def resources(self):
        return self.Resources

    @resources.setter
    def resources(self, value):
        self.Resources = value

    @property
    def ResponseRequested(self):
        return self.com_object.ResponseRequested

    @ResponseRequested.setter
    def ResponseRequested(self, value):
        self.com_object.ResponseRequested = value

    # Lower case aliases for ResponseRequested
    @property
    def responserequested(self):
        return self.ResponseRequested

    @responserequested.setter
    def responserequested(self, value):
        self.ResponseRequested = value

    @property
    def ResponseStatus(self):
        return OlResponseStatus(self.com_object.ResponseStatus)

    # Lower case aliases for ResponseStatus
    @property
    def responsestatus(self):
        return self.ResponseStatus

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    # Lower case aliases for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Start(self):
        return self.com_object.Start

    @Start.setter
    def Start(self, value):
        self.com_object.Start = value

    # Lower case aliases for Start
    @property
    def start(self):
        return self.Start

    @start.setter
    def start(self, value):
        self.Start = value

    @property
    def StartInStartTimeZone(self):
        return AppointmentItem.StartTimeZone(self.com_object.StartInStartTimeZone)

    @StartInStartTimeZone.setter
    def StartInStartTimeZone(self, value):
        self.com_object.StartInStartTimeZone = value

    # Lower case aliases for StartInStartTimeZone
    @property
    def startinstarttimezone(self):
        return self.StartInStartTimeZone

    @startinstarttimezone.setter
    def startinstarttimezone(self, value):
        self.StartInStartTimeZone = value

    @property
    def StartTimeZone(self):
        return TimeZone(self.com_object.StartTimeZone)

    @StartTimeZone.setter
    def StartTimeZone(self, value):
        self.com_object.StartTimeZone = value

    # Lower case aliases for StartTimeZone
    @property
    def starttimezone(self):
        return self.StartTimeZone

    @starttimezone.setter
    def starttimezone(self, value):
        self.StartTimeZone = value

    @property
    def StartUTC(self):
        return self.com_object.StartUTC

    @StartUTC.setter
    def StartUTC(self, value):
        self.com_object.StartUTC = value

    # Lower case aliases for StartUTC
    @property
    def startutc(self):
        return self.StartUTC

    @startutc.setter
    def startutc(self, value):
        self.StartUTC = value

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def ClearRecurrencePattern(self):
        self.com_object.ClearRecurrencePattern()

    # Lower case alias for ClearRecurrencePattern
    def clearrecurrencepattern(self):
        return self.ClearRecurrencePattern()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def CopyTo(self, DestinationFolder=None, CopyOptions=None):
        arguments = com_arguments([DestinationFolder, CopyOptions])
        return AppointmentItem(self.com_object.CopyTo(*arguments))

    # Lower case alias for CopyTo
    def copyto(self, DestinationFolder=None, CopyOptions=None):
        arguments = [DestinationFolder, CopyOptions]
        return self.CopyTo(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def ForwardAsVcal(self):
        return MailItem(self.com_object.ForwardAsVcal())

    # Lower case alias for ForwardAsVcal
    def forwardasvcal(self):
        return self.ForwardAsVcal()

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def GetOrganizer(self):
        return AddressEntry(self.com_object.GetOrganizer())

    # Lower case alias for GetOrganizer
    def getorganizer(self):
        return self.GetOrganizer()

    def GetRecurrencePattern(self):
        return RecurrencePattern(self.com_object.GetRecurrencePattern())

    # Lower case alias for GetRecurrencePattern
    def getrecurrencepattern(self):
        return self.GetRecurrencePattern()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Respond(self, Response=None, fNoUI=None, fAdditionalTextDialog=None):
        arguments = com_arguments([Response, fNoUI, fAdditionalTextDialog])
        return MeetingItem(self.com_object.Respond(*arguments))

    # Lower case alias for Respond
    def respond(self, Response=None, fNoUI=None, fAdditionalTextDialog=None):
        arguments = [Response, fNoUI, fAdditionalTextDialog]
        return self.Respond(*arguments)

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def Send(self):
        self.com_object.Send()

    # Lower case alias for Send
    def send(self):
        return self.Send()

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class AssignToCategoryRuleAction:

    def __init__(self, assigntocategoryruleaction=None):
        self.com_object= assigntocategoryruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    # Lower case aliases for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class Attachment:

    def __init__(self, attachment=None):
        self.com_object= attachment

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BlockLevel(self):
        return OlAttachmentBlockLevel(self.com_object.BlockLevel)

    # Lower case aliases for BlockLevel
    @property
    def blocklevel(self):
        return self.BlockLevel

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DisplayName(self):
        return self.com_object.DisplayName

    @DisplayName.setter
    def DisplayName(self, value):
        self.com_object.DisplayName = value

    # Lower case aliases for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @displayname.setter
    def displayname(self, value):
        self.DisplayName = value

    @property
    def FileName(self):
        return self.com_object.FileName

    # Lower case aliases for FileName
    @property
    def filename(self):
        return self.FileName

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PathName(self):
        return self.com_object.PathName

    # Lower case aliases for PathName
    @property
    def pathname(self):
        return self.PathName

    @property
    def Position(self):
        return self.com_object.Position

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Type(self):
        return OlAttachmentType(self.com_object.Type)

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GetTemporaryFilePath(self):
        return String(self.com_object.GetTemporaryFilePath())

    # Lower case alias for GetTemporaryFilePath
    def gettemporaryfilepath(self):
        return self.GetTemporaryFilePath()

    def SaveAsFile(self, Path=None):
        arguments = com_arguments([Path])
        self.com_object.SaveAsFile(*arguments)

    # Lower case alias for SaveAsFile
    def saveasfile(self, Path=None):
        arguments = [Path]
        return self.SaveAsFile(*arguments)


class Attachments:

    def __init__(self, attachments=None):
        self.com_object= attachments

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Source=None, Type=None, Position=None, DisplayName=None):
        arguments = com_arguments([Source, Type, Position, DisplayName])
        return Attachment(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Source=None, Type=None, Position=None, DisplayName=None):
        arguments = [Source, Type, Position, DisplayName]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Attachment(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class AttachmentSelection:

    def __init__(self, attachmentselection=None):
        self.com_object= attachmentselection

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Location(self):
        return OlSelectionLocation(self.com_object.Location)

    # Lower case aliases for Location
    @property
    def location(self):
        return self.Location

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def GetSelection(self, SelectionContents=None):
        arguments = com_arguments([SelectionContents])
        return Selection(self.com_object.GetSelection(*arguments))

    # Lower case alias for GetSelection
    def getselection(self, SelectionContents=None):
        arguments = [SelectionContents]
        return self.GetSelection(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Attachment(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class AutoFormatRule:

    def __init__(self, autoformatrule=None):
        self.com_object= autoformatrule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Enabled(self):
        return AutoFormatRule(self.com_object.Enabled)

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def Font(self):
        return ViewFont(self.com_object.Font)

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Standard(self):
        return AutoFormatRule(self.com_object.Standard)

    # Lower case aliases for Standard
    @property
    def standard(self):
        return self.Standard


class AutoFormatRules:

    def __init__(self, autoformatrules=None):
        self.com_object= autoformatrules

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return AutoFormatRule(self.com_object.Count)

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([Name])
        return AutoFormatRule(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None):
        arguments = [Name]
        return self.Add(*arguments)

    def Insert(self, Name=None, Index=None):
        arguments = com_arguments([Name, Index])
        return AutoFormatRule(self.com_object.Insert(*arguments))

    # Lower case alias for Insert
    def insert(self, Name=None, Index=None):
        arguments = [Name, Index]
        return self.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return AutoFormatRule(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

    def RemoveAll(self):
        self.com_object.RemoveAll()

    # Lower case alias for RemoveAll
    def removeall(self):
        return self.RemoveAll()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class BusinessCardView:

    def __init__(self, businesscardview=None):
        self.com_object= businesscardview

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def CardSize(self):
        return self.com_object.CardSize

    @CardSize.setter
    def CardSize(self, value):
        self.com_object.CardSize = value

    # Lower case aliases for CardSize
    @property
    def cardsize(self):
        return self.CardSize

    @cardsize.setter
    def cardsize(self, value):
        self.CardSize = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def HeadingsFont(self):
        return ViewFont(self.com_object.HeadingsFont)

    # Lower case aliases for HeadingsFont
    @property
    def headingsfont(self):
        return self.HeadingsFont

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    # Lower case aliases for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    # Lower case aliases for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def SortFields(self):
        return OrderFields(self.com_object.SortFields)

    # Lower case aliases for SortFields
    @property
    def sortfields(self):
        return self.SortFields

    @property
    def Standard(self):
        return BusinessCardView(self.com_object.Standard)

    # Lower case aliases for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    # Lower case aliases for XML
    @property
    def xml(self):
        return self.XML

    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return View(self.com_object.Copy(*arguments))

    # Lower case alias for Copy
    def copy(self, Name=None, SaveOption=None):
        arguments = [Name, SaveOption]
        return self.Copy(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.com_object.GoToDate(*arguments)

    # Lower case alias for GoToDate
    def gotodate(self, Date=None):
        arguments = [Date]
        return self.GoToDate(*arguments)

    def Reset(self):
        self.com_object.Reset()

    # Lower case alias for Reset
    def reset(self):
        return self.Reset()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class CalendarModule:

    def __init__(self, calendarmodule=None):
        self.com_object= calendarmodule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return CalendarModule(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    # Lower case aliases for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.com_object.NavigationModuleType)

    # Lower case aliases for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return CalendarModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return CalendarModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value


class CalendarSharing:

    def __init__(self, calendarsharing=None):
        self.com_object= calendarsharing

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def CalendarDetail(self):
        return OlCalendarDetail(self.com_object.CalendarDetail)

    @CalendarDetail.setter
    def CalendarDetail(self, value):
        self.com_object.CalendarDetail = value

    # Lower case aliases for CalendarDetail
    @property
    def calendardetail(self):
        return self.CalendarDetail

    @calendardetail.setter
    def calendardetail(self, value):
        self.CalendarDetail = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def EndDate(self):
        return CalendarSharing(self.com_object.EndDate)

    @EndDate.setter
    def EndDate(self, value):
        self.com_object.EndDate = value

    # Lower case aliases for EndDate
    @property
    def enddate(self):
        return self.EndDate

    @enddate.setter
    def enddate(self, value):
        self.EndDate = value

    @property
    def Folder(self):
        return Folder(self.com_object.Folder)

    # Lower case aliases for Folder
    @property
    def folder(self):
        return self.Folder

    @property
    def IncludeAttachments(self):
        return self.com_object.IncludeAttachments

    @IncludeAttachments.setter
    def IncludeAttachments(self, value):
        self.com_object.IncludeAttachments = value

    # Lower case aliases for IncludeAttachments
    @property
    def includeattachments(self):
        return self.IncludeAttachments

    @includeattachments.setter
    def includeattachments(self, value):
        self.IncludeAttachments = value

    @property
    def IncludePrivateDetails(self):
        return self.com_object.IncludePrivateDetails

    @IncludePrivateDetails.setter
    def IncludePrivateDetails(self, value):
        self.com_object.IncludePrivateDetails = value

    # Lower case aliases for IncludePrivateDetails
    @property
    def includeprivatedetails(self):
        return self.IncludePrivateDetails

    @includeprivatedetails.setter
    def includeprivatedetails(self, value):
        self.IncludePrivateDetails = value

    @property
    def IncludeWholeCalendar(self):
        return self.com_object.IncludeWholeCalendar

    @IncludeWholeCalendar.setter
    def IncludeWholeCalendar(self, value):
        self.com_object.IncludeWholeCalendar = value

    # Lower case aliases for IncludeWholeCalendar
    @property
    def includewholecalendar(self):
        return self.IncludeWholeCalendar

    @includewholecalendar.setter
    def includewholecalendar(self, value):
        self.IncludeWholeCalendar = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RestrictToWorkingHours(self):
        return self.com_object.RestrictToWorkingHours

    @RestrictToWorkingHours.setter
    def RestrictToWorkingHours(self, value):
        self.com_object.RestrictToWorkingHours = value

    # Lower case aliases for RestrictToWorkingHours
    @property
    def restricttoworkinghours(self):
        return self.RestrictToWorkingHours

    @restricttoworkinghours.setter
    def restricttoworkinghours(self, value):
        self.RestrictToWorkingHours = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def StartDate(self):
        return CalendarSharing(self.com_object.StartDate)

    @StartDate.setter
    def StartDate(self, value):
        self.com_object.StartDate = value

    # Lower case aliases for StartDate
    @property
    def startdate(self):
        return self.StartDate

    @startdate.setter
    def startdate(self, value):
        self.StartDate = value

    def ForwardAsICal(self, MailFormat=None):
        arguments = com_arguments([MailFormat])
        return MailItem(self.com_object.ForwardAsICal(*arguments))

    # Lower case alias for ForwardAsICal
    def forwardasical(self, MailFormat=None):
        arguments = [MailFormat]
        return self.ForwardAsICal(*arguments)

    def SaveAsICal(self, Path=None):
        arguments = com_arguments([Path])
        self.com_object.SaveAsICal(*arguments)

    # Lower case alias for SaveAsICal
    def saveasical(self, Path=None):
        arguments = [Path]
        return self.SaveAsICal(*arguments)


class CalendarView:

    def __init__(self, calendarview=None):
        self.com_object= calendarview

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.com_object.AutoFormatRules)

    # Lower case aliases for AutoFormatRules
    @property
    def autoformatrules(self):
        return self.AutoFormatRules

    @property
    def BoldDatesWithItems(self):
        return CalendarView(self.com_object.BoldDatesWithItems)

    @BoldDatesWithItems.setter
    def BoldDatesWithItems(self, value):
        self.com_object.BoldDatesWithItems = value

    # Lower case aliases for BoldDatesWithItems
    @property
    def bolddateswithitems(self):
        return self.BoldDatesWithItems

    @bolddateswithitems.setter
    def bolddateswithitems(self, value):
        self.BoldDatesWithItems = value

    @property
    def BoldSubjects(self):
        return CalendarView(self.com_object.BoldSubjects)

    @BoldSubjects.setter
    def BoldSubjects(self, value):
        self.com_object.BoldSubjects = value

    # Lower case aliases for BoldSubjects
    @property
    def boldsubjects(self):
        return self.BoldSubjects

    @boldsubjects.setter
    def boldsubjects(self, value):
        self.BoldSubjects = value

    @property
    def CalendarViewMode(self):
        return OlCalendarViewMode(self.com_object.CalendarViewMode)

    @CalendarViewMode.setter
    def CalendarViewMode(self, value):
        self.com_object.CalendarViewMode = value

    # Lower case aliases for CalendarViewMode
    @property
    def calendarviewmode(self):
        return self.CalendarViewMode

    @calendarviewmode.setter
    def calendarviewmode(self, value):
        self.CalendarViewMode = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DaysInMultiDayMode(self):
        return CalendarView(self.com_object.DaysInMultiDayMode)

    @DaysInMultiDayMode.setter
    def DaysInMultiDayMode(self, value):
        self.com_object.DaysInMultiDayMode = value

    # Lower case aliases for DaysInMultiDayMode
    @property
    def daysinmultidaymode(self):
        return self.DaysInMultiDayMode

    @daysinmultidaymode.setter
    def daysinmultidaymode(self, value):
        self.DaysInMultiDayMode = value

    @property
    def DayWeekTimeScale(self):
        return OlDayWeekTimeScale(self.com_object.DayWeekTimeScale)

    @DayWeekTimeScale.setter
    def DayWeekTimeScale(self, value):
        self.com_object.DayWeekTimeScale = value

    # Lower case aliases for DayWeekTimeScale
    @property
    def dayweektimescale(self):
        return self.DayWeekTimeScale

    @dayweektimescale.setter
    def dayweektimescale(self, value):
        self.DayWeekTimeScale = value

    @property
    def DisplayedDates(self):
        return CalendarView(self.com_object.DisplayedDates)

    # Lower case aliases for DisplayedDates
    @property
    def displayeddates(self):
        return self.DisplayedDates

    @property
    def EndField(self):
        return CalendarView(self.com_object.EndField)

    @EndField.setter
    def EndField(self, value):
        self.com_object.EndField = value

    # Lower case aliases for EndField
    @property
    def endfield(self):
        return self.EndField

    @endfield.setter
    def endfield(self, value):
        self.EndField = value

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    # Lower case aliases for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def MonthShowEndTime(self):
        return CalendarView(self.com_object.MonthShowEndTime)

    @MonthShowEndTime.setter
    def MonthShowEndTime(self, value):
        self.com_object.MonthShowEndTime = value

    # Lower case aliases for MonthShowEndTime
    @property
    def monthshowendtime(self):
        return self.MonthShowEndTime

    @monthshowendtime.setter
    def monthshowendtime(self, value):
        self.MonthShowEndTime = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    # Lower case aliases for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def SelectedEndTime(self):
        return CalendarView(self.com_object.SelectedEndTime)

    # Lower case aliases for SelectedEndTime
    @property
    def selectedendtime(self):
        return self.SelectedEndTime

    @property
    def SelectedStartTime(self):
        return CalendarView(self.com_object.SelectedStartTime)

    # Lower case aliases for SelectedStartTime
    @property
    def selectedstarttime(self):
        return self.SelectedStartTime

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Standard(self):
        return CalendarView(self.com_object.Standard)

    # Lower case aliases for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def StartField(self):
        return CalendarView(self.com_object.StartField)

    @StartField.setter
    def StartField(self, value):
        self.com_object.StartField = value

    # Lower case aliases for StartField
    @property
    def startfield(self):
        return self.StartField

    @startfield.setter
    def startfield(self, value):
        self.StartField = value

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    # Lower case aliases for XML
    @property
    def xml(self):
        return self.XML

    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return View(self.com_object.Copy(*arguments))

    # Lower case alias for Copy
    def copy(self, Name=None, SaveOption=None):
        arguments = [Name, SaveOption]
        return self.Copy(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.com_object.GoToDate(*arguments)

    # Lower case alias for GoToDate
    def gotodate(self, Date=None):
        arguments = [Date]
        return self.GoToDate(*arguments)

    def Reset(self):
        self.com_object.Reset()

    # Lower case alias for Reset
    def reset(self):
        return self.Reset()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class CardView:

    def __init__(self, cardview=None):
        self.com_object= cardview

    @property
    def AllowInCellEditing(self):
        return CardView(self.com_object.AllowInCellEditing)

    @AllowInCellEditing.setter
    def AllowInCellEditing(self, value):
        self.com_object.AllowInCellEditing = value

    # Lower case aliases for AllowInCellEditing
    @property
    def allowincellediting(self):
        return self.AllowInCellEditing

    @allowincellediting.setter
    def allowincellediting(self, value):
        self.AllowInCellEditing = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.com_object.AutoFormatRules)

    # Lower case aliases for AutoFormatRules
    @property
    def autoformatrules(self):
        return self.AutoFormatRules

    @property
    def BodyFont(self):
        return ViewFont(self.com_object.BodyFont)

    # Lower case aliases for BodyFont
    @property
    def bodyfont(self):
        return self.BodyFont

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def HeadingsFont(self):
        return ViewFont(self.com_object.HeadingsFont)

    # Lower case aliases for HeadingsFont
    @property
    def headingsfont(self):
        return self.HeadingsFont

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    # Lower case aliases for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def MultiLineFieldHeight(self):
        return CardView(self.com_object.MultiLineFieldHeight)

    @MultiLineFieldHeight.setter
    def MultiLineFieldHeight(self, value):
        self.com_object.MultiLineFieldHeight = value

    # Lower case aliases for MultiLineFieldHeight
    @property
    def multilinefieldheight(self):
        return self.MultiLineFieldHeight

    @multilinefieldheight.setter
    def multilinefieldheight(self, value):
        self.MultiLineFieldHeight = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    # Lower case aliases for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowEmptyFields(self):
        return CardView(self.com_object.ShowEmptyFields)

    @ShowEmptyFields.setter
    def ShowEmptyFields(self, value):
        self.com_object.ShowEmptyFields = value

    # Lower case aliases for ShowEmptyFields
    @property
    def showemptyfields(self):
        return self.ShowEmptyFields

    @showemptyfields.setter
    def showemptyfields(self, value):
        self.ShowEmptyFields = value

    @property
    def SortFields(self):
        return OrderFields(self.com_object.SortFields)

    # Lower case aliases for SortFields
    @property
    def sortfields(self):
        return self.SortFields

    @property
    def Standard(self):
        return CardView(self.com_object.Standard)

    # Lower case aliases for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewFields(self):
        return ViewFields(self.com_object.ViewFields)

    # Lower case aliases for ViewFields
    @property
    def viewfields(self):
        return self.ViewFields

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def Width(self):
        return CardView(self.com_object.Width)

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    # Lower case aliases for XML
    @property
    def xml(self):
        return self.XML

    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return View(self.com_object.Copy(*arguments))

    # Lower case alias for Copy
    def copy(self, Name=None, SaveOption=None):
        arguments = [Name, SaveOption]
        return self.Copy(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.com_object.GoToDate(*arguments)

    # Lower case alias for GoToDate
    def gotodate(self, Date=None):
        arguments = [Date]
        return self.GoToDate(*arguments)

    def Reset(self):
        self.com_object.Reset()

    # Lower case alias for Reset
    def reset(self):
        return self.Reset()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class Categories:

    def __init__(self, categories=None):
        self.com_object= categories

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return Category(self.com_object.Count)

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Color=None, ShortcutKey=None):
        arguments = com_arguments([Name, Color, ShortcutKey])
        return Category(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Color=None, ShortcutKey=None):
        arguments = [Name, Color, ShortcutKey]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Category(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class Category:

    def __init__(self, category=None):
        self.com_object= category

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def CategoryBorderColor(self):
        return Category(self.com_object.CategoryBorderColor)

    # Lower case aliases for CategoryBorderColor
    @property
    def categorybordercolor(self):
        return self.CategoryBorderColor

    @property
    def CategoryGradientBottomColor(self):
        return Category(self.com_object.CategoryGradientBottomColor)

    # Lower case aliases for CategoryGradientBottomColor
    @property
    def categorygradientbottomcolor(self):
        return self.CategoryGradientBottomColor

    @property
    def CategoryGradientTopColor(self):
        return Category(self.com_object.CategoryGradientTopColor)

    # Lower case aliases for CategoryGradientTopColor
    @property
    def categorygradienttopcolor(self):
        return self.CategoryGradientTopColor

    @property
    def CategoryID(self):
        return Category(self.com_object.CategoryID)

    # Lower case aliases for CategoryID
    @property
    def categoryid(self):
        return self.CategoryID

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Color(self):
        return OlCategoryColor(self.com_object.Color)

    @Color.setter
    def Color(self, value):
        self.com_object.Color = value

    # Lower case aliases for Color
    @property
    def color(self):
        return self.Color

    @color.setter
    def color(self, value):
        self.Color = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShortcutKey(self):
        return OlCategoryShortcutKey(self.com_object.ShortcutKey)

    @ShortcutKey.setter
    def ShortcutKey(self, value):
        self.com_object.ShortcutKey = value

    # Lower case aliases for ShortcutKey
    @property
    def shortcutkey(self):
        return self.ShortcutKey

    @shortcutkey.setter
    def shortcutkey(self, value):
        self.ShortcutKey = value


class CategoryRuleCondition:

    def __init__(self, categoryrulecondition=None):
        self.com_object= categoryrulecondition

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class Column:

    def __init__(self, column=None):
        self.com_object= column

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return Column(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return Column(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class ColumnFormat:

    def __init__(self, columnformat=None):
        self.com_object= columnformat

    @property
    def Align(self):
        return OlAlign(self.com_object.Align)

    @Align.setter
    def Align(self, value):
        self.com_object.Align = value

    # Lower case aliases for Align
    @property
    def align(self):
        return self.Align

    @align.setter
    def align(self, value):
        self.Align = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def FieldFormat(self):
        return ColumnFormat(self.com_object.FieldFormat)

    @FieldFormat.setter
    def FieldFormat(self, value):
        self.com_object.FieldFormat = value

    # Lower case aliases for FieldFormat
    @property
    def fieldformat(self):
        return self.FieldFormat

    @fieldformat.setter
    def fieldformat(self, value):
        self.FieldFormat = value

    @property
    def FieldType(self):
        return OlUserPropertyType(self.com_object.FieldType)

    # Lower case aliases for FieldType
    @property
    def fieldtype(self):
        return self.FieldType

    @property
    def Label(self):
        return ColumnFormat(self.com_object.Label)

    @Label.setter
    def Label(self, value):
        self.com_object.Label = value

    # Lower case aliases for Label
    @property
    def label(self):
        return self.Label

    @label.setter
    def label(self, value):
        self.Label = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Width(self):
        return self.com_object.Width

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    @width.setter
    def width(self, value):
        self.Width = value


class Columns:

    def __init__(self, columns=None):
        self.com_object= columns

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return Column(self.com_object.Count)

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return Columns(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([Name])
        return Column(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None):
        arguments = [Name]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

    def RemoveAll(self):
        self.com_object.RemoveAll()

    # Lower case alias for RemoveAll
    def removeall(self):
        return self.RemoveAll()


class Conflict:

    def __init__(self, conflict=None):
        self.com_object= conflict

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Item(self):
        return self.com_object.Item

    # Lower case aliases for Item
    @property
    def item(self):
        return self.Item

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return OlObjectClass(self.com_object.Type)

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type


class Conflicts:

    def __init__(self, conflicts=None):
        self.com_object= conflicts

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def GetFirst(self):
        return Conflict(self.com_object.GetFirst())

    # Lower case alias for GetFirst
    def getfirst(self):
        return self.GetFirst()

    def GetLast(self):
        return Conflict(self.com_object.GetLast())

    # Lower case alias for GetLast
    def getlast(self):
        return self.GetLast()

    def GetNext(self):
        return Conflict(self.com_object.GetNext())

    # Lower case alias for GetNext
    def getnext(self):
        return self.GetNext()

    def GetPrevious(self):
        return Conflict(self.com_object.GetPrevious())

    # Lower case alias for GetPrevious
    def getprevious(self):
        return self.GetPrevious()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Conflict(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ContactItem:

    def __init__(self, contactitem=None):
        self.com_object= contactitem

    @property
    def Account(self):
        return self.com_object.Account

    @Account.setter
    def Account(self, value):
        self.com_object.Account = value

    # Lower case aliases for Account
    @property
    def account(self):
        return self.Account

    @account.setter
    def account(self, value):
        self.Account = value

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Anniversary(self):
        return self.com_object.Anniversary

    @Anniversary.setter
    def Anniversary(self, value):
        self.com_object.Anniversary = value

    # Lower case aliases for Anniversary
    @property
    def anniversary(self):
        return self.Anniversary

    @anniversary.setter
    def anniversary(self, value):
        self.Anniversary = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AssistantName(self):
        return self.com_object.AssistantName

    @AssistantName.setter
    def AssistantName(self, value):
        self.com_object.AssistantName = value

    # Lower case aliases for AssistantName
    @property
    def assistantname(self):
        return self.AssistantName

    @assistantname.setter
    def assistantname(self, value):
        self.AssistantName = value

    @property
    def AssistantTelephoneNumber(self):
        return self.com_object.AssistantTelephoneNumber

    @AssistantTelephoneNumber.setter
    def AssistantTelephoneNumber(self, value):
        self.com_object.AssistantTelephoneNumber = value

    # Lower case aliases for AssistantTelephoneNumber
    @property
    def assistanttelephonenumber(self):
        return self.AssistantTelephoneNumber

    @assistanttelephonenumber.setter
    def assistanttelephonenumber(self, value):
        self.AssistantTelephoneNumber = value

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Birthday(self):
        return self.com_object.Birthday

    @Birthday.setter
    def Birthday(self, value):
        self.com_object.Birthday = value

    # Lower case aliases for Birthday
    @property
    def birthday(self):
        return self.Birthday

    @birthday.setter
    def birthday(self, value):
        self.Birthday = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Business2TelephoneNumber(self):
        return self.com_object.Business2TelephoneNumber

    @Business2TelephoneNumber.setter
    def Business2TelephoneNumber(self, value):
        self.com_object.Business2TelephoneNumber = value

    # Lower case aliases for Business2TelephoneNumber
    @property
    def business2telephonenumber(self):
        return self.Business2TelephoneNumber

    @business2telephonenumber.setter
    def business2telephonenumber(self, value):
        self.Business2TelephoneNumber = value

    @property
    def BusinessAddress(self):
        return self.com_object.BusinessAddress

    @BusinessAddress.setter
    def BusinessAddress(self, value):
        self.com_object.BusinessAddress = value

    # Lower case aliases for BusinessAddress
    @property
    def businessaddress(self):
        return self.BusinessAddress

    @businessaddress.setter
    def businessaddress(self, value):
        self.BusinessAddress = value

    @property
    def BusinessAddressCity(self):
        return self.com_object.BusinessAddressCity

    @BusinessAddressCity.setter
    def BusinessAddressCity(self, value):
        self.com_object.BusinessAddressCity = value

    # Lower case aliases for BusinessAddressCity
    @property
    def businessaddresscity(self):
        return self.BusinessAddressCity

    @businessaddresscity.setter
    def businessaddresscity(self, value):
        self.BusinessAddressCity = value

    @property
    def BusinessAddressCountry(self):
        return self.com_object.BusinessAddressCountry

    @BusinessAddressCountry.setter
    def BusinessAddressCountry(self, value):
        self.com_object.BusinessAddressCountry = value

    # Lower case aliases for BusinessAddressCountry
    @property
    def businessaddresscountry(self):
        return self.BusinessAddressCountry

    @businessaddresscountry.setter
    def businessaddresscountry(self, value):
        self.BusinessAddressCountry = value

    @property
    def BusinessAddressPostalCode(self):
        return self.com_object.BusinessAddressPostalCode

    @BusinessAddressPostalCode.setter
    def BusinessAddressPostalCode(self, value):
        self.com_object.BusinessAddressPostalCode = value

    # Lower case aliases for BusinessAddressPostalCode
    @property
    def businessaddresspostalcode(self):
        return self.BusinessAddressPostalCode

    @businessaddresspostalcode.setter
    def businessaddresspostalcode(self, value):
        self.BusinessAddressPostalCode = value

    @property
    def BusinessAddressPostOfficeBox(self):
        return self.com_object.BusinessAddressPostOfficeBox

    @BusinessAddressPostOfficeBox.setter
    def BusinessAddressPostOfficeBox(self, value):
        self.com_object.BusinessAddressPostOfficeBox = value

    # Lower case aliases for BusinessAddressPostOfficeBox
    @property
    def businessaddresspostofficebox(self):
        return self.BusinessAddressPostOfficeBox

    @businessaddresspostofficebox.setter
    def businessaddresspostofficebox(self, value):
        self.BusinessAddressPostOfficeBox = value

    @property
    def BusinessAddressState(self):
        return self.com_object.BusinessAddressState

    @BusinessAddressState.setter
    def BusinessAddressState(self, value):
        self.com_object.BusinessAddressState = value

    # Lower case aliases for BusinessAddressState
    @property
    def businessaddressstate(self):
        return self.BusinessAddressState

    @businessaddressstate.setter
    def businessaddressstate(self, value):
        self.BusinessAddressState = value

    @property
    def BusinessAddressStreet(self):
        return self.com_object.BusinessAddressStreet

    @BusinessAddressStreet.setter
    def BusinessAddressStreet(self, value):
        self.com_object.BusinessAddressStreet = value

    # Lower case aliases for BusinessAddressStreet
    @property
    def businessaddressstreet(self):
        return self.BusinessAddressStreet

    @businessaddressstreet.setter
    def businessaddressstreet(self, value):
        self.BusinessAddressStreet = value

    @property
    def BusinessCardLayoutXml(self):
        return self.com_object.BusinessCardLayoutXml

    @BusinessCardLayoutXml.setter
    def BusinessCardLayoutXml(self, value):
        self.com_object.BusinessCardLayoutXml = value

    # Lower case aliases for BusinessCardLayoutXml
    @property
    def businesscardlayoutxml(self):
        return self.BusinessCardLayoutXml

    @businesscardlayoutxml.setter
    def businesscardlayoutxml(self, value):
        self.BusinessCardLayoutXml = value

    @property
    def BusinessCardType(self):
        return OlBusinessCardType(self.com_object.BusinessCardType)

    # Lower case aliases for BusinessCardType
    @property
    def businesscardtype(self):
        return self.BusinessCardType

    @property
    def BusinessFaxNumber(self):
        return self.com_object.BusinessFaxNumber

    @BusinessFaxNumber.setter
    def BusinessFaxNumber(self, value):
        self.com_object.BusinessFaxNumber = value

    # Lower case aliases for BusinessFaxNumber
    @property
    def businessfaxnumber(self):
        return self.BusinessFaxNumber

    @businessfaxnumber.setter
    def businessfaxnumber(self, value):
        self.BusinessFaxNumber = value

    @property
    def BusinessHomePage(self):
        return self.com_object.BusinessHomePage

    @BusinessHomePage.setter
    def BusinessHomePage(self, value):
        self.com_object.BusinessHomePage = value

    # Lower case aliases for BusinessHomePage
    @property
    def businesshomepage(self):
        return self.BusinessHomePage

    @businesshomepage.setter
    def businesshomepage(self, value):
        self.BusinessHomePage = value

    @property
    def BusinessTelephoneNumber(self):
        return self.com_object.BusinessTelephoneNumber

    @BusinessTelephoneNumber.setter
    def BusinessTelephoneNumber(self, value):
        self.com_object.BusinessTelephoneNumber = value

    # Lower case aliases for BusinessTelephoneNumber
    @property
    def businesstelephonenumber(self):
        return self.BusinessTelephoneNumber

    @businesstelephonenumber.setter
    def businesstelephonenumber(self, value):
        self.BusinessTelephoneNumber = value

    @property
    def CallbackTelephoneNumber(self):
        return self.com_object.CallbackTelephoneNumber

    @CallbackTelephoneNumber.setter
    def CallbackTelephoneNumber(self, value):
        self.com_object.CallbackTelephoneNumber = value

    # Lower case aliases for CallbackTelephoneNumber
    @property
    def callbacktelephonenumber(self):
        return self.CallbackTelephoneNumber

    @callbacktelephonenumber.setter
    def callbacktelephonenumber(self, value):
        self.CallbackTelephoneNumber = value

    @property
    def CarTelephoneNumber(self):
        return self.com_object.CarTelephoneNumber

    @CarTelephoneNumber.setter
    def CarTelephoneNumber(self, value):
        self.com_object.CarTelephoneNumber = value

    # Lower case aliases for CarTelephoneNumber
    @property
    def cartelephonenumber(self):
        return self.CarTelephoneNumber

    @cartelephonenumber.setter
    def cartelephonenumber(self, value):
        self.CarTelephoneNumber = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Children(self):
        return self.com_object.Children

    @Children.setter
    def Children(self, value):
        self.com_object.Children = value

    # Lower case aliases for Children
    @property
    def children(self):
        return self.Children

    @children.setter
    def children(self, value):
        self.Children = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def CompanyAndFullName(self):
        return self.com_object.CompanyAndFullName

    # Lower case aliases for CompanyAndFullName
    @property
    def companyandfullname(self):
        return self.CompanyAndFullName

    @property
    def CompanyLastFirstNoSpace(self):
        return self.com_object.CompanyLastFirstNoSpace

    # Lower case aliases for CompanyLastFirstNoSpace
    @property
    def companylastfirstnospace(self):
        return self.CompanyLastFirstNoSpace

    @property
    def CompanyLastFirstSpaceOnly(self):
        return self.com_object.CompanyLastFirstSpaceOnly

    # Lower case aliases for CompanyLastFirstSpaceOnly
    @property
    def companylastfirstspaceonly(self):
        return self.CompanyLastFirstSpaceOnly

    @property
    def CompanyMainTelephoneNumber(self):
        return self.com_object.CompanyMainTelephoneNumber

    @CompanyMainTelephoneNumber.setter
    def CompanyMainTelephoneNumber(self, value):
        self.com_object.CompanyMainTelephoneNumber = value

    # Lower case aliases for CompanyMainTelephoneNumber
    @property
    def companymaintelephonenumber(self):
        return self.CompanyMainTelephoneNumber

    @companymaintelephonenumber.setter
    def companymaintelephonenumber(self, value):
        self.CompanyMainTelephoneNumber = value

    @property
    def CompanyName(self):
        return self.com_object.CompanyName

    @CompanyName.setter
    def CompanyName(self, value):
        self.com_object.CompanyName = value

    # Lower case aliases for CompanyName
    @property
    def companyname(self):
        return self.CompanyName

    @companyname.setter
    def companyname(self, value):
        self.CompanyName = value

    @property
    def ComputerNetworkName(self):
        return self.com_object.ComputerNetworkName

    @ComputerNetworkName.setter
    def ComputerNetworkName(self, value):
        self.com_object.ComputerNetworkName = value

    # Lower case aliases for ComputerNetworkName
    @property
    def computernetworkname(self):
        return self.ComputerNetworkName

    @computernetworkname.setter
    def computernetworkname(self, value):
        self.ComputerNetworkName = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def CustomerID(self):
        return self.com_object.CustomerID

    @CustomerID.setter
    def CustomerID(self, value):
        self.com_object.CustomerID = value

    # Lower case aliases for CustomerID
    @property
    def customerid(self):
        return self.CustomerID

    @customerid.setter
    def customerid(self, value):
        self.CustomerID = value

    @property
    def Department(self):
        return self.com_object.Department

    @Department.setter
    def Department(self, value):
        self.com_object.Department = value

    # Lower case aliases for Department
    @property
    def department(self):
        return self.Department

    @department.setter
    def department(self, value):
        self.Department = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def Email1Address(self):
        return self.com_object.Email1Address

    @Email1Address.setter
    def Email1Address(self, value):
        self.com_object.Email1Address = value

    # Lower case aliases for Email1Address
    @property
    def email1address(self):
        return self.Email1Address

    @email1address.setter
    def email1address(self, value):
        self.Email1Address = value

    @property
    def Email1AddressType(self):
        return self.com_object.Email1AddressType

    @Email1AddressType.setter
    def Email1AddressType(self, value):
        self.com_object.Email1AddressType = value

    # Lower case aliases for Email1AddressType
    @property
    def email1addresstype(self):
        return self.Email1AddressType

    @email1addresstype.setter
    def email1addresstype(self, value):
        self.Email1AddressType = value

    @property
    def Email1DisplayName(self):
        return self.com_object.Email1DisplayName

    @Email1DisplayName.setter
    def Email1DisplayName(self, value):
        self.com_object.Email1DisplayName = value

    # Lower case aliases for Email1DisplayName
    @property
    def email1displayname(self):
        return self.Email1DisplayName

    @email1displayname.setter
    def email1displayname(self, value):
        self.Email1DisplayName = value

    @property
    def Email1EntryID(self):
        return self.com_object.Email1EntryID

    # Lower case aliases for Email1EntryID
    @property
    def email1entryid(self):
        return self.Email1EntryID

    @property
    def Email2Address(self):
        return self.com_object.Email2Address

    @Email2Address.setter
    def Email2Address(self, value):
        self.com_object.Email2Address = value

    # Lower case aliases for Email2Address
    @property
    def email2address(self):
        return self.Email2Address

    @email2address.setter
    def email2address(self, value):
        self.Email2Address = value

    @property
    def Email2AddressType(self):
        return self.com_object.Email2AddressType

    @Email2AddressType.setter
    def Email2AddressType(self, value):
        self.com_object.Email2AddressType = value

    # Lower case aliases for Email2AddressType
    @property
    def email2addresstype(self):
        return self.Email2AddressType

    @email2addresstype.setter
    def email2addresstype(self, value):
        self.Email2AddressType = value

    @property
    def Email2DisplayName(self):
        return self.com_object.Email2DisplayName

    @Email2DisplayName.setter
    def Email2DisplayName(self, value):
        self.com_object.Email2DisplayName = value

    # Lower case aliases for Email2DisplayName
    @property
    def email2displayname(self):
        return self.Email2DisplayName

    @email2displayname.setter
    def email2displayname(self, value):
        self.Email2DisplayName = value

    @property
    def Email2EntryID(self):
        return self.com_object.Email2EntryID

    # Lower case aliases for Email2EntryID
    @property
    def email2entryid(self):
        return self.Email2EntryID

    @property
    def Email3Address(self):
        return self.com_object.Email3Address

    @Email3Address.setter
    def Email3Address(self, value):
        self.com_object.Email3Address = value

    # Lower case aliases for Email3Address
    @property
    def email3address(self):
        return self.Email3Address

    @email3address.setter
    def email3address(self, value):
        self.Email3Address = value

    @property
    def Email3AddressType(self):
        return self.com_object.Email3AddressType

    @Email3AddressType.setter
    def Email3AddressType(self, value):
        self.com_object.Email3AddressType = value

    # Lower case aliases for Email3AddressType
    @property
    def email3addresstype(self):
        return self.Email3AddressType

    @email3addresstype.setter
    def email3addresstype(self, value):
        self.Email3AddressType = value

    @property
    def Email3DisplayName(self):
        return self.com_object.Email3DisplayName

    @Email3DisplayName.setter
    def Email3DisplayName(self, value):
        self.com_object.Email3DisplayName = value

    # Lower case aliases for Email3DisplayName
    @property
    def email3displayname(self):
        return self.Email3DisplayName

    @email3displayname.setter
    def email3displayname(self, value):
        self.Email3DisplayName = value

    @property
    def Email3EntryID(self):
        return self.com_object.Email3EntryID

    # Lower case aliases for Email3EntryID
    @property
    def email3entryid(self):
        return self.Email3EntryID

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FileAs(self):
        return self.com_object.FileAs

    @FileAs.setter
    def FileAs(self, value):
        self.com_object.FileAs = value

    # Lower case aliases for FileAs
    @property
    def fileas(self):
        return self.FileAs

    @fileas.setter
    def fileas(self, value):
        self.FileAs = value

    @property
    def FirstName(self):
        return self.com_object.FirstName

    @FirstName.setter
    def FirstName(self, value):
        self.com_object.FirstName = value

    # Lower case aliases for FirstName
    @property
    def firstname(self):
        return self.FirstName

    @firstname.setter
    def firstname(self, value):
        self.FirstName = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def FTPSite(self):
        return self.com_object.FTPSite

    @FTPSite.setter
    def FTPSite(self, value):
        self.com_object.FTPSite = value

    # Lower case aliases for FTPSite
    @property
    def ftpsite(self):
        return self.FTPSite

    @ftpsite.setter
    def ftpsite(self, value):
        self.FTPSite = value

    @property
    def FullName(self):
        return self.com_object.FullName

    @FullName.setter
    def FullName(self, value):
        self.com_object.FullName = value

    # Lower case aliases for FullName
    @property
    def fullname(self):
        return self.FullName

    @fullname.setter
    def fullname(self, value):
        self.FullName = value

    @property
    def FullNameAndCompany(self):
        return self.com_object.FullNameAndCompany

    # Lower case aliases for FullNameAndCompany
    @property
    def fullnameandcompany(self):
        return self.FullNameAndCompany

    @property
    def Gender(self):
        return OlGender(self.com_object.Gender)

    @Gender.setter
    def Gender(self, value):
        self.com_object.Gender = value

    # Lower case aliases for Gender
    @property
    def gender(self):
        return self.Gender

    @gender.setter
    def gender(self, value):
        self.Gender = value

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def GovernmentIDNumber(self):
        return self.com_object.GovernmentIDNumber

    @GovernmentIDNumber.setter
    def GovernmentIDNumber(self, value):
        self.com_object.GovernmentIDNumber = value

    # Lower case aliases for GovernmentIDNumber
    @property
    def governmentidnumber(self):
        return self.GovernmentIDNumber

    @governmentidnumber.setter
    def governmentidnumber(self, value):
        self.GovernmentIDNumber = value

    @property
    def HasPicture(self):
        return self.com_object.HasPicture

    # Lower case aliases for HasPicture
    @property
    def haspicture(self):
        return self.HasPicture

    @property
    def Hobby(self):
        return self.com_object.Hobby

    @Hobby.setter
    def Hobby(self, value):
        self.com_object.Hobby = value

    # Lower case aliases for Hobby
    @property
    def hobby(self):
        return self.Hobby

    @hobby.setter
    def hobby(self, value):
        self.Hobby = value

    @property
    def Home2TelephoneNumber(self):
        return self.com_object.Home2TelephoneNumber

    @Home2TelephoneNumber.setter
    def Home2TelephoneNumber(self, value):
        self.com_object.Home2TelephoneNumber = value

    # Lower case aliases for Home2TelephoneNumber
    @property
    def home2telephonenumber(self):
        return self.Home2TelephoneNumber

    @home2telephonenumber.setter
    def home2telephonenumber(self, value):
        self.Home2TelephoneNumber = value

    @property
    def HomeAddress(self):
        return self.com_object.HomeAddress

    @HomeAddress.setter
    def HomeAddress(self, value):
        self.com_object.HomeAddress = value

    # Lower case aliases for HomeAddress
    @property
    def homeaddress(self):
        return self.HomeAddress

    @homeaddress.setter
    def homeaddress(self, value):
        self.HomeAddress = value

    @property
    def HomeAddressCity(self):
        return self.com_object.HomeAddressCity

    @HomeAddressCity.setter
    def HomeAddressCity(self, value):
        self.com_object.HomeAddressCity = value

    # Lower case aliases for HomeAddressCity
    @property
    def homeaddresscity(self):
        return self.HomeAddressCity

    @homeaddresscity.setter
    def homeaddresscity(self, value):
        self.HomeAddressCity = value

    @property
    def HomeAddressCountry(self):
        return self.com_object.HomeAddressCountry

    @HomeAddressCountry.setter
    def HomeAddressCountry(self, value):
        self.com_object.HomeAddressCountry = value

    # Lower case aliases for HomeAddressCountry
    @property
    def homeaddresscountry(self):
        return self.HomeAddressCountry

    @homeaddresscountry.setter
    def homeaddresscountry(self, value):
        self.HomeAddressCountry = value

    @property
    def HomeAddressPostalCode(self):
        return self.com_object.HomeAddressPostalCode

    @HomeAddressPostalCode.setter
    def HomeAddressPostalCode(self, value):
        self.com_object.HomeAddressPostalCode = value

    # Lower case aliases for HomeAddressPostalCode
    @property
    def homeaddresspostalcode(self):
        return self.HomeAddressPostalCode

    @homeaddresspostalcode.setter
    def homeaddresspostalcode(self, value):
        self.HomeAddressPostalCode = value

    @property
    def HomeAddressPostOfficeBox(self):
        return self.com_object.HomeAddressPostOfficeBox

    @HomeAddressPostOfficeBox.setter
    def HomeAddressPostOfficeBox(self, value):
        self.com_object.HomeAddressPostOfficeBox = value

    # Lower case aliases for HomeAddressPostOfficeBox
    @property
    def homeaddresspostofficebox(self):
        return self.HomeAddressPostOfficeBox

    @homeaddresspostofficebox.setter
    def homeaddresspostofficebox(self, value):
        self.HomeAddressPostOfficeBox = value

    @property
    def HomeAddressState(self):
        return self.com_object.HomeAddressState

    @HomeAddressState.setter
    def HomeAddressState(self, value):
        self.com_object.HomeAddressState = value

    # Lower case aliases for HomeAddressState
    @property
    def homeaddressstate(self):
        return self.HomeAddressState

    @homeaddressstate.setter
    def homeaddressstate(self, value):
        self.HomeAddressState = value

    @property
    def HomeAddressStreet(self):
        return self.com_object.HomeAddressStreet

    @HomeAddressStreet.setter
    def HomeAddressStreet(self, value):
        self.com_object.HomeAddressStreet = value

    # Lower case aliases for HomeAddressStreet
    @property
    def homeaddressstreet(self):
        return self.HomeAddressStreet

    @homeaddressstreet.setter
    def homeaddressstreet(self, value):
        self.HomeAddressStreet = value

    @property
    def HomeFaxNumber(self):
        return self.com_object.HomeFaxNumber

    @HomeFaxNumber.setter
    def HomeFaxNumber(self, value):
        self.com_object.HomeFaxNumber = value

    # Lower case aliases for HomeFaxNumber
    @property
    def homefaxnumber(self):
        return self.HomeFaxNumber

    @homefaxnumber.setter
    def homefaxnumber(self, value):
        self.HomeFaxNumber = value

    @property
    def HomeTelephoneNumber(self):
        return self.com_object.HomeTelephoneNumber

    @HomeTelephoneNumber.setter
    def HomeTelephoneNumber(self, value):
        self.com_object.HomeTelephoneNumber = value

    # Lower case aliases for HomeTelephoneNumber
    @property
    def hometelephonenumber(self):
        return self.HomeTelephoneNumber

    @hometelephonenumber.setter
    def hometelephonenumber(self, value):
        self.HomeTelephoneNumber = value

    @property
    def IMAddress(self):
        return self.com_object.IMAddress

    @IMAddress.setter
    def IMAddress(self, value):
        self.com_object.IMAddress = value

    # Lower case aliases for IMAddress
    @property
    def imaddress(self):
        return self.IMAddress

    @imaddress.setter
    def imaddress(self, value):
        self.IMAddress = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def Initials(self):
        return self.com_object.Initials

    @Initials.setter
    def Initials(self, value):
        self.com_object.Initials = value

    # Lower case aliases for Initials
    @property
    def initials(self):
        return self.Initials

    @initials.setter
    def initials(self, value):
        self.Initials = value

    @property
    def InternetFreeBusyAddress(self):
        return self.com_object.InternetFreeBusyAddress

    @InternetFreeBusyAddress.setter
    def InternetFreeBusyAddress(self, value):
        self.com_object.InternetFreeBusyAddress = value

    # Lower case aliases for InternetFreeBusyAddress
    @property
    def internetfreebusyaddress(self):
        return self.InternetFreeBusyAddress

    @internetfreebusyaddress.setter
    def internetfreebusyaddress(self, value):
        self.InternetFreeBusyAddress = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ISDNNumber(self):
        return self.com_object.ISDNNumber

    @ISDNNumber.setter
    def ISDNNumber(self, value):
        self.com_object.ISDNNumber = value

    # Lower case aliases for ISDNNumber
    @property
    def isdnnumber(self):
        return self.ISDNNumber

    @isdnnumber.setter
    def isdnnumber(self, value):
        self.ISDNNumber = value

    @property
    def IsMarkedAsTask(self):
        return ContactItem(self.com_object.IsMarkedAsTask)

    # Lower case aliases for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def JobTitle(self):
        return self.com_object.JobTitle

    @JobTitle.setter
    def JobTitle(self, value):
        self.com_object.JobTitle = value

    # Lower case aliases for JobTitle
    @property
    def jobtitle(self):
        return self.JobTitle

    @jobtitle.setter
    def jobtitle(self, value):
        self.JobTitle = value

    @property
    def Journal(self):
        return self.com_object.Journal

    @Journal.setter
    def Journal(self, value):
        self.com_object.Journal = value

    # Lower case aliases for Journal
    @property
    def journal(self):
        return self.Journal

    @journal.setter
    def journal(self, value):
        self.Journal = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LastFirstAndSuffix(self):
        return self.com_object.LastFirstAndSuffix

    # Lower case aliases for LastFirstAndSuffix
    @property
    def lastfirstandsuffix(self):
        return self.LastFirstAndSuffix

    @property
    def LastFirstNoSpace(self):
        return self.com_object.LastFirstNoSpace

    # Lower case aliases for LastFirstNoSpace
    @property
    def lastfirstnospace(self):
        return self.LastFirstNoSpace

    @property
    def LastFirstNoSpaceAndSuffix(self):
        return self.com_object.LastFirstNoSpaceAndSuffix

    # Lower case aliases for LastFirstNoSpaceAndSuffix
    @property
    def lastfirstnospaceandsuffix(self):
        return self.LastFirstNoSpaceAndSuffix

    @property
    def LastFirstNoSpaceCompany(self):
        return self.com_object.LastFirstNoSpaceCompany

    # Lower case aliases for LastFirstNoSpaceCompany
    @property
    def lastfirstnospacecompany(self):
        return self.LastFirstNoSpaceCompany

    @property
    def LastFirstSpaceOnly(self):
        return self.com_object.LastFirstSpaceOnly

    # Lower case aliases for LastFirstSpaceOnly
    @property
    def lastfirstspaceonly(self):
        return self.LastFirstSpaceOnly

    @property
    def LastFirstSpaceOnlyCompany(self):
        return self.com_object.LastFirstSpaceOnlyCompany

    # Lower case aliases for LastFirstSpaceOnlyCompany
    @property
    def lastfirstspaceonlycompany(self):
        return self.LastFirstSpaceOnlyCompany

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def LastName(self):
        return self.com_object.LastName

    @LastName.setter
    def LastName(self, value):
        self.com_object.LastName = value

    # Lower case aliases for LastName
    @property
    def lastname(self):
        return self.LastName

    @lastname.setter
    def lastname(self, value):
        self.LastName = value

    @property
    def LastNameAndFirstName(self):
        return self.com_object.LastNameAndFirstName

    # Lower case aliases for LastNameAndFirstName
    @property
    def lastnameandfirstname(self):
        return self.LastNameAndFirstName

    @property
    def MailingAddress(self):
        return self.com_object.MailingAddress

    @MailingAddress.setter
    def MailingAddress(self, value):
        self.com_object.MailingAddress = value

    # Lower case aliases for MailingAddress
    @property
    def mailingaddress(self):
        return self.MailingAddress

    @mailingaddress.setter
    def mailingaddress(self, value):
        self.MailingAddress = value

    @property
    def MailingAddressCity(self):
        return self.com_object.MailingAddressCity

    @MailingAddressCity.setter
    def MailingAddressCity(self, value):
        self.com_object.MailingAddressCity = value

    # Lower case aliases for MailingAddressCity
    @property
    def mailingaddresscity(self):
        return self.MailingAddressCity

    @mailingaddresscity.setter
    def mailingaddresscity(self, value):
        self.MailingAddressCity = value

    @property
    def MailingAddressCountry(self):
        return self.com_object.MailingAddressCountry

    @MailingAddressCountry.setter
    def MailingAddressCountry(self, value):
        self.com_object.MailingAddressCountry = value

    # Lower case aliases for MailingAddressCountry
    @property
    def mailingaddresscountry(self):
        return self.MailingAddressCountry

    @mailingaddresscountry.setter
    def mailingaddresscountry(self, value):
        self.MailingAddressCountry = value

    @property
    def MailingAddressPostalCode(self):
        return self.com_object.MailingAddressPostalCode

    @MailingAddressPostalCode.setter
    def MailingAddressPostalCode(self, value):
        self.com_object.MailingAddressPostalCode = value

    # Lower case aliases for MailingAddressPostalCode
    @property
    def mailingaddresspostalcode(self):
        return self.MailingAddressPostalCode

    @mailingaddresspostalcode.setter
    def mailingaddresspostalcode(self, value):
        self.MailingAddressPostalCode = value

    @property
    def MailingAddressPostOfficeBox(self):
        return self.com_object.MailingAddressPostOfficeBox

    @MailingAddressPostOfficeBox.setter
    def MailingAddressPostOfficeBox(self, value):
        self.com_object.MailingAddressPostOfficeBox = value

    # Lower case aliases for MailingAddressPostOfficeBox
    @property
    def mailingaddresspostofficebox(self):
        return self.MailingAddressPostOfficeBox

    @mailingaddresspostofficebox.setter
    def mailingaddresspostofficebox(self, value):
        self.MailingAddressPostOfficeBox = value

    @property
    def MailingAddressState(self):
        return self.com_object.MailingAddressState

    @MailingAddressState.setter
    def MailingAddressState(self, value):
        self.com_object.MailingAddressState = value

    # Lower case aliases for MailingAddressState
    @property
    def mailingaddressstate(self):
        return self.MailingAddressState

    @mailingaddressstate.setter
    def mailingaddressstate(self, value):
        self.MailingAddressState = value

    @property
    def MailingAddressStreet(self):
        return self.com_object.MailingAddressStreet

    @MailingAddressStreet.setter
    def MailingAddressStreet(self, value):
        self.com_object.MailingAddressStreet = value

    # Lower case aliases for MailingAddressStreet
    @property
    def mailingaddressstreet(self):
        return self.MailingAddressStreet

    @mailingaddressstreet.setter
    def mailingaddressstreet(self, value):
        self.MailingAddressStreet = value

    @property
    def ManagerName(self):
        return self.com_object.ManagerName

    @ManagerName.setter
    def ManagerName(self, value):
        self.com_object.ManagerName = value

    # Lower case aliases for ManagerName
    @property
    def managername(self):
        return self.ManagerName

    @managername.setter
    def managername(self, value):
        self.ManagerName = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def MiddleName(self):
        return self.com_object.MiddleName

    @MiddleName.setter
    def MiddleName(self, value):
        self.com_object.MiddleName = value

    # Lower case aliases for MiddleName
    @property
    def middlename(self):
        return self.MiddleName

    @middlename.setter
    def middlename(self, value):
        self.MiddleName = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def MobileTelephoneNumber(self):
        return self.com_object.MobileTelephoneNumber

    @MobileTelephoneNumber.setter
    def MobileTelephoneNumber(self, value):
        self.com_object.MobileTelephoneNumber = value

    # Lower case aliases for MobileTelephoneNumber
    @property
    def mobiletelephonenumber(self):
        return self.MobileTelephoneNumber

    @mobiletelephonenumber.setter
    def mobiletelephonenumber(self, value):
        self.MobileTelephoneNumber = value

    @property
    def NetMeetingAlias(self):
        return self.com_object.NetMeetingAlias

    @NetMeetingAlias.setter
    def NetMeetingAlias(self, value):
        self.com_object.NetMeetingAlias = value

    # Lower case aliases for NetMeetingAlias
    @property
    def netmeetingalias(self):
        return self.NetMeetingAlias

    @netmeetingalias.setter
    def netmeetingalias(self, value):
        self.NetMeetingAlias = value

    @property
    def NetMeetingServer(self):
        return self.com_object.NetMeetingServer

    @NetMeetingServer.setter
    def NetMeetingServer(self, value):
        self.com_object.NetMeetingServer = value

    # Lower case aliases for NetMeetingServer
    @property
    def netmeetingserver(self):
        return self.NetMeetingServer

    @netmeetingserver.setter
    def netmeetingserver(self, value):
        self.NetMeetingServer = value

    @property
    def NickName(self):
        return self.com_object.NickName

    @NickName.setter
    def NickName(self, value):
        self.com_object.NickName = value

    # Lower case aliases for NickName
    @property
    def nickname(self):
        return self.NickName

    @nickname.setter
    def nickname(self, value):
        self.NickName = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OfficeLocation(self):
        return self.com_object.OfficeLocation

    @OfficeLocation.setter
    def OfficeLocation(self, value):
        self.com_object.OfficeLocation = value

    # Lower case aliases for OfficeLocation
    @property
    def officelocation(self):
        return self.OfficeLocation

    @officelocation.setter
    def officelocation(self, value):
        self.OfficeLocation = value

    @property
    def OrganizationalIDNumber(self):
        return self.com_object.OrganizationalIDNumber

    @OrganizationalIDNumber.setter
    def OrganizationalIDNumber(self, value):
        self.com_object.OrganizationalIDNumber = value

    # Lower case aliases for OrganizationalIDNumber
    @property
    def organizationalidnumber(self):
        return self.OrganizationalIDNumber

    @organizationalidnumber.setter
    def organizationalidnumber(self, value):
        self.OrganizationalIDNumber = value

    @property
    def OtherAddress(self):
        return self.com_object.OtherAddress

    @OtherAddress.setter
    def OtherAddress(self, value):
        self.com_object.OtherAddress = value

    # Lower case aliases for OtherAddress
    @property
    def otheraddress(self):
        return self.OtherAddress

    @otheraddress.setter
    def otheraddress(self, value):
        self.OtherAddress = value

    @property
    def OtherAddressCity(self):
        return self.com_object.OtherAddressCity

    @OtherAddressCity.setter
    def OtherAddressCity(self, value):
        self.com_object.OtherAddressCity = value

    # Lower case aliases for OtherAddressCity
    @property
    def otheraddresscity(self):
        return self.OtherAddressCity

    @otheraddresscity.setter
    def otheraddresscity(self, value):
        self.OtherAddressCity = value

    @property
    def OtherAddressCountry(self):
        return self.com_object.OtherAddressCountry

    @OtherAddressCountry.setter
    def OtherAddressCountry(self, value):
        self.com_object.OtherAddressCountry = value

    # Lower case aliases for OtherAddressCountry
    @property
    def otheraddresscountry(self):
        return self.OtherAddressCountry

    @otheraddresscountry.setter
    def otheraddresscountry(self, value):
        self.OtherAddressCountry = value

    @property
    def OtherAddressPostalCode(self):
        return self.com_object.OtherAddressPostalCode

    @OtherAddressPostalCode.setter
    def OtherAddressPostalCode(self, value):
        self.com_object.OtherAddressPostalCode = value

    # Lower case aliases for OtherAddressPostalCode
    @property
    def otheraddresspostalcode(self):
        return self.OtherAddressPostalCode

    @otheraddresspostalcode.setter
    def otheraddresspostalcode(self, value):
        self.OtherAddressPostalCode = value

    @property
    def OtherAddressPostOfficeBox(self):
        return self.com_object.OtherAddressPostOfficeBox

    @OtherAddressPostOfficeBox.setter
    def OtherAddressPostOfficeBox(self, value):
        self.com_object.OtherAddressPostOfficeBox = value

    # Lower case aliases for OtherAddressPostOfficeBox
    @property
    def otheraddresspostofficebox(self):
        return self.OtherAddressPostOfficeBox

    @otheraddresspostofficebox.setter
    def otheraddresspostofficebox(self, value):
        self.OtherAddressPostOfficeBox = value

    @property
    def OtherAddressState(self):
        return self.com_object.OtherAddressState

    @OtherAddressState.setter
    def OtherAddressState(self, value):
        self.com_object.OtherAddressState = value

    # Lower case aliases for OtherAddressState
    @property
    def otheraddressstate(self):
        return self.OtherAddressState

    @otheraddressstate.setter
    def otheraddressstate(self, value):
        self.OtherAddressState = value

    @property
    def OtherAddressStreet(self):
        return self.com_object.OtherAddressStreet

    @OtherAddressStreet.setter
    def OtherAddressStreet(self, value):
        self.com_object.OtherAddressStreet = value

    # Lower case aliases for OtherAddressStreet
    @property
    def otheraddressstreet(self):
        return self.OtherAddressStreet

    @otheraddressstreet.setter
    def otheraddressstreet(self, value):
        self.OtherAddressStreet = value

    @property
    def OtherFaxNumber(self):
        return self.com_object.OtherFaxNumber

    @OtherFaxNumber.setter
    def OtherFaxNumber(self, value):
        self.com_object.OtherFaxNumber = value

    # Lower case aliases for OtherFaxNumber
    @property
    def otherfaxnumber(self):
        return self.OtherFaxNumber

    @otherfaxnumber.setter
    def otherfaxnumber(self, value):
        self.OtherFaxNumber = value

    @property
    def OtherTelephoneNumber(self):
        return self.com_object.OtherTelephoneNumber

    @OtherTelephoneNumber.setter
    def OtherTelephoneNumber(self, value):
        self.com_object.OtherTelephoneNumber = value

    # Lower case aliases for OtherTelephoneNumber
    @property
    def othertelephonenumber(self):
        return self.OtherTelephoneNumber

    @othertelephonenumber.setter
    def othertelephonenumber(self, value):
        self.OtherTelephoneNumber = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def PagerNumber(self):
        return self.com_object.PagerNumber

    @PagerNumber.setter
    def PagerNumber(self, value):
        self.com_object.PagerNumber = value

    # Lower case aliases for PagerNumber
    @property
    def pagernumber(self):
        return self.PagerNumber

    @pagernumber.setter
    def pagernumber(self, value):
        self.PagerNumber = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PersonalHomePage(self):
        return self.com_object.PersonalHomePage

    @PersonalHomePage.setter
    def PersonalHomePage(self, value):
        self.com_object.PersonalHomePage = value

    # Lower case aliases for PersonalHomePage
    @property
    def personalhomepage(self):
        return self.PersonalHomePage

    @personalhomepage.setter
    def personalhomepage(self, value):
        self.PersonalHomePage = value

    @property
    def PrimaryTelephoneNumber(self):
        return self.com_object.PrimaryTelephoneNumber

    @PrimaryTelephoneNumber.setter
    def PrimaryTelephoneNumber(self, value):
        self.com_object.PrimaryTelephoneNumber = value

    # Lower case aliases for PrimaryTelephoneNumber
    @property
    def primarytelephonenumber(self):
        return self.PrimaryTelephoneNumber

    @primarytelephonenumber.setter
    def primarytelephonenumber(self, value):
        self.PrimaryTelephoneNumber = value

    @property
    def Profession(self):
        return self.com_object.Profession

    @Profession.setter
    def Profession(self, value):
        self.com_object.Profession = value

    # Lower case aliases for Profession
    @property
    def profession(self):
        return self.Profession

    @profession.setter
    def profession(self, value):
        self.Profession = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RadioTelephoneNumber(self):
        return self.com_object.RadioTelephoneNumber

    @RadioTelephoneNumber.setter
    def RadioTelephoneNumber(self, value):
        self.com_object.RadioTelephoneNumber = value

    # Lower case aliases for RadioTelephoneNumber
    @property
    def radiotelephonenumber(self):
        return self.RadioTelephoneNumber

    @radiotelephonenumber.setter
    def radiotelephonenumber(self, value):
        self.RadioTelephoneNumber = value

    @property
    def ReferredBy(self):
        return self.com_object.ReferredBy

    @ReferredBy.setter
    def ReferredBy(self, value):
        self.com_object.ReferredBy = value

    # Lower case aliases for ReferredBy
    @property
    def referredby(self):
        return self.ReferredBy

    @referredby.setter
    def referredby(self, value):
        self.ReferredBy = value

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    # Lower case aliases for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    # Lower case aliases for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    # Lower case aliases for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    # Lower case aliases for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    # Lower case aliases for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SelectedMailingAddress(self):
        return OlMailingAddress(self.com_object.SelectedMailingAddress)

    @SelectedMailingAddress.setter
    def SelectedMailingAddress(self, value):
        self.com_object.SelectedMailingAddress = value

    # Lower case aliases for SelectedMailingAddress
    @property
    def selectedmailingaddress(self):
        return self.SelectedMailingAddress

    @selectedmailingaddress.setter
    def selectedmailingaddress(self, value):
        self.SelectedMailingAddress = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Spouse(self):
        return self.com_object.Spouse

    @Spouse.setter
    def Spouse(self, value):
        self.com_object.Spouse = value

    # Lower case aliases for Spouse
    @property
    def spouse(self):
        return self.Spouse

    @spouse.setter
    def spouse(self, value):
        self.Spouse = value

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Suffix(self):
        return self.com_object.Suffix

    @Suffix.setter
    def Suffix(self, value):
        self.com_object.Suffix = value

    # Lower case aliases for Suffix
    @property
    def suffix(self):
        return self.Suffix

    @suffix.setter
    def suffix(self, value):
        self.Suffix = value

    @property
    def TaskCompletedDate(self):
        return ContactItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    # Lower case aliases for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return ContactItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    # Lower case aliases for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return ContactItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    # Lower case aliases for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return ContactItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    # Lower case aliases for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def TelexNumber(self):
        return self.com_object.TelexNumber

    @TelexNumber.setter
    def TelexNumber(self, value):
        self.com_object.TelexNumber = value

    # Lower case aliases for TelexNumber
    @property
    def telexnumber(self):
        return self.TelexNumber

    @telexnumber.setter
    def telexnumber(self, value):
        self.TelexNumber = value

    @property
    def Title(self):
        return self.com_object.Title

    @Title.setter
    def Title(self, value):
        self.com_object.Title = value

    # Lower case aliases for Title
    @property
    def title(self):
        return self.Title

    @title.setter
    def title(self, value):
        self.Title = value

    @property
    def ToDoTaskOrdinal(self):
        return ContactItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    # Lower case aliases for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def TTYTDDTelephoneNumber(self):
        return self.com_object.TTYTDDTelephoneNumber

    @TTYTDDTelephoneNumber.setter
    def TTYTDDTelephoneNumber(self, value):
        self.com_object.TTYTDDTelephoneNumber = value

    # Lower case aliases for TTYTDDTelephoneNumber
    @property
    def ttytddtelephonenumber(self):
        return self.TTYTDDTelephoneNumber

    @ttytddtelephonenumber.setter
    def ttytddtelephonenumber(self, value):
        self.TTYTDDTelephoneNumber = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def User1(self):
        return self.com_object.User1

    @User1.setter
    def User1(self, value):
        self.com_object.User1 = value

    # Lower case aliases for User1
    @property
    def user1(self):
        return self.User1

    @user1.setter
    def user1(self, value):
        self.User1 = value

    @property
    def User2(self):
        return self.com_object.User2

    @User2.setter
    def User2(self, value):
        self.com_object.User2 = value

    # Lower case aliases for User2
    @property
    def user2(self):
        return self.User2

    @user2.setter
    def user2(self, value):
        self.User2 = value

    @property
    def User3(self):
        return self.com_object.User3

    @User3.setter
    def User3(self, value):
        self.com_object.User3 = value

    # Lower case aliases for User3
    @property
    def user3(self):
        return self.User3

    @user3.setter
    def user3(self, value):
        self.User3 = value

    @property
    def User4(self):
        return self.com_object.User4

    @User4.setter
    def User4(self, value):
        self.com_object.User4 = value

    # Lower case aliases for User4
    @property
    def user4(self):
        return self.User4

    @user4.setter
    def user4(self, value):
        self.User4 = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    @property
    def WebPage(self):
        return self.com_object.WebPage

    @WebPage.setter
    def WebPage(self, value):
        self.com_object.WebPage = value

    # Lower case aliases for WebPage
    @property
    def webpage(self):
        return self.WebPage

    @webpage.setter
    def webpage(self, value):
        self.WebPage = value

    @property
    def YomiCompanyName(self):
        return self.com_object.YomiCompanyName

    @YomiCompanyName.setter
    def YomiCompanyName(self, value):
        self.com_object.YomiCompanyName = value

    # Lower case aliases for YomiCompanyName
    @property
    def yomicompanyname(self):
        return self.YomiCompanyName

    @yomicompanyname.setter
    def yomicompanyname(self, value):
        self.YomiCompanyName = value

    @property
    def YomiFirstName(self):
        return self.com_object.YomiFirstName

    @YomiFirstName.setter
    def YomiFirstName(self, value):
        self.com_object.YomiFirstName = value

    # Lower case aliases for YomiFirstName
    @property
    def yomifirstname(self):
        return self.YomiFirstName

    @yomifirstname.setter
    def yomifirstname(self, value):
        self.YomiFirstName = value

    @property
    def YomiLastName(self):
        return self.com_object.YomiLastName

    @YomiLastName.setter
    def YomiLastName(self, value):
        self.com_object.YomiLastName = value

    # Lower case aliases for YomiLastName
    @property
    def yomilastname(self):
        return self.YomiLastName

    @yomilastname.setter
    def yomilastname(self, value):
        self.YomiLastName = value

    def AddBusinessCardLogoPicture(self, Path=None):
        arguments = com_arguments([Path])
        self.com_object.AddBusinessCardLogoPicture(*arguments)

    # Lower case alias for AddBusinessCardLogoPicture
    def addbusinesscardlogopicture(self, Path=None):
        arguments = [Path]
        return self.AddBusinessCardLogoPicture(*arguments)

    def AddPicture(self, Path=None):
        arguments = com_arguments([Path])
        self.com_object.AddPicture(*arguments)

    # Lower case alias for AddPicture
    def addpicture(self, Path=None):
        arguments = [Path]
        return self.AddPicture(*arguments)

    def ClearTaskFlag(self):
        self.com_object.ClearTaskFlag()

    # Lower case alias for ClearTaskFlag
    def cleartaskflag(self):
        return self.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def ForwardAsBusinessCard(self):
        return MailItem(self.com_object.ForwardAsBusinessCard())

    # Lower case alias for ForwardAsBusinessCard
    def forwardasbusinesscard(self):
        return self.ForwardAsBusinessCard()

    def ForwardAsVcard(self):
        return MailItem(self.com_object.ForwardAsVcard())

    # Lower case alias for ForwardAsVcard
    def forwardasvcard(self):
        return self.ForwardAsVcard()

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def RemovePicture(self):
        self.com_object.RemovePicture()

    # Lower case alias for RemovePicture
    def removepicture(self):
        return self.RemovePicture()

    def ResetBusinessCard(self):
        self.com_object.ResetBusinessCard()

    # Lower case alias for ResetBusinessCard
    def resetbusinesscard(self):
        return self.ResetBusinessCard()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def SaveBusinessCardImage(self, Path=None):
        arguments = com_arguments([Path])
        self.com_object.SaveBusinessCardImage(*arguments)

    # Lower case alias for SaveBusinessCardImage
    def savebusinesscardimage(self, Path=None):
        arguments = [Path]
        return self.SaveBusinessCardImage(*arguments)

    def ShowBusinessCardEditor(self):
        self.com_object.ShowBusinessCardEditor()

    # Lower case alias for ShowBusinessCardEditor
    def showbusinesscardeditor(self):
        return self.ShowBusinessCardEditor()

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()

    def ShowCheckPhoneDialog(self, PhoneNumber=None):
        arguments = com_arguments([PhoneNumber])
        self.com_object.ShowCheckPhoneDialog(*arguments)

    # Lower case alias for ShowCheckPhoneDialog
    def showcheckphonedialog(self, PhoneNumber=None):
        arguments = [PhoneNumber]
        return self.ShowCheckPhoneDialog(*arguments)


class ContactsModule:

    def __init__(self, contactsmodule=None):
        self.com_object= contactsmodule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return ContactsModule(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    # Lower case aliases for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    # Lower case aliases for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return ContactsModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return ContactsModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value


class Conversation:

    def __init__(self, conversation=None):
        self.com_object= conversation

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def Parent(self):
        return Conversation(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def ClearAlwaysAssignCategories(self, Store=None):
        arguments = com_arguments([Store])
        self.com_object.ClearAlwaysAssignCategories(*arguments)

    # Lower case alias for ClearAlwaysAssignCategories
    def clearalwaysassigncategories(self, Store=None):
        arguments = [Store]
        return self.ClearAlwaysAssignCategories(*arguments)

    def GetAlwaysAssignCategories(self, Store=None):
        arguments = com_arguments([Store])
        return String(self.com_object.GetAlwaysAssignCategories(*arguments))

    # Lower case alias for GetAlwaysAssignCategories
    def getalwaysassigncategories(self, Store=None):
        arguments = [Store]
        return self.GetAlwaysAssignCategories(*arguments)

    def GetAlwaysDelete(self, Store=None):
        arguments = com_arguments([Store])
        return OlAlwaysDeleteConversation(self.com_object.GetAlwaysDelete(*arguments))

    # Lower case alias for GetAlwaysDelete
    def getalwaysdelete(self, Store=None):
        arguments = [Store]
        return self.GetAlwaysDelete(*arguments)

    def GetAlwaysMoveToFolder(self, Store=None):
        arguments = com_arguments([Store])
        return Folder(self.com_object.GetAlwaysMoveToFolder(*arguments))

    # Lower case alias for GetAlwaysMoveToFolder
    def getalwaysmovetofolder(self, Store=None):
        arguments = [Store]
        return self.GetAlwaysMoveToFolder(*arguments)

    def GetChildren(self, Item=None):
        arguments = com_arguments([Item])
        return SimpleItems(self.com_object.GetChildren(*arguments))

    # Lower case alias for GetChildren
    def getchildren(self, Item=None):
        arguments = [Item]
        return self.GetChildren(*arguments)

    def GetParent(self, Item=None):
        arguments = com_arguments([Item])
        return Object(self.com_object.GetParent(*arguments))

    # Lower case alias for GetParent
    def getparent(self, Item=None):
        arguments = [Item]
        return self.GetParent(*arguments)

    def GetRootItems(self):
        return SimpleItems(self.com_object.GetRootItems())

    # Lower case alias for GetRootItems
    def getrootitems(self):
        return self.GetRootItems()

    def GetTable(self):
        return Table(self.com_object.GetTable())

    # Lower case alias for GetTable
    def gettable(self):
        return self.GetTable()

    def MarkAsRead(self):
        self.com_object.MarkAsRead()

    # Lower case alias for MarkAsRead
    def markasread(self):
        return self.MarkAsRead()

    def MarkAsUnread(self):
        self.com_object.MarkAsUnread()

    # Lower case alias for MarkAsUnread
    def markasunread(self):
        return self.MarkAsUnread()

    def SetAlwaysAssignCategories(self, Categories=None, Store=None):
        arguments = com_arguments([Categories, Store])
        self.com_object.SetAlwaysAssignCategories(*arguments)

    # Lower case alias for SetAlwaysAssignCategories
    def setalwaysassigncategories(self, Categories=None, Store=None):
        arguments = [Categories, Store]
        return self.SetAlwaysAssignCategories(*arguments)

    def SetAlwaysDelete(self, AlwaysDelete=None, Store=None):
        arguments = com_arguments([AlwaysDelete, Store])
        self.com_object.SetAlwaysDelete(*arguments)

    # Lower case alias for SetAlwaysDelete
    def setalwaysdelete(self, AlwaysDelete=None, Store=None):
        arguments = [AlwaysDelete, Store]
        return self.SetAlwaysDelete(*arguments)

    def SetAlwaysMoveToFolder(self, MoveToFolder=None, Store=None):
        arguments = com_arguments([MoveToFolder, Store])
        self.com_object.SetAlwaysMoveToFolder(*arguments)

    # Lower case alias for SetAlwaysMoveToFolder
    def setalwaysmovetofolder(self, MoveToFolder=None, Store=None):
        arguments = [MoveToFolder, Store]
        return self.SetAlwaysMoveToFolder(*arguments)

    def StopAlwaysDelete(self, Store=None):
        arguments = com_arguments([Store])
        self.com_object.StopAlwaysDelete(*arguments)

    # Lower case alias for StopAlwaysDelete
    def stopalwaysdelete(self, Store=None):
        arguments = [Store]
        return self.StopAlwaysDelete(*arguments)

    def StopAlwaysMoveToFolder(self, Store=None):
        arguments = com_arguments([Store])
        self.com_object.StopAlwaysMoveToFolder(*arguments)

    # Lower case alias for StopAlwaysMoveToFolder
    def stopalwaysmovetofolder(self, Store=None):
        arguments = [Store]
        return self.StopAlwaysMoveToFolder(*arguments)


class ConversationHeader:

    def __init__(self, conversationheader=None):
        self.com_object= conversationheader

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def GetItems(self):
        return SimpleItems(self.com_object.GetItems())

    # Lower case alias for GetItems
    def getitems(self):
        return self.GetItems()


class DistListItem:

    def __init__(self, distlistitem=None):
        self.com_object= distlistitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DLName(self):
        return self.com_object.DLName

    @DLName.setter
    def DLName(self, value):
        self.com_object.DLName = value

    # Lower case aliases for DLName
    @property
    def dlname(self):
        return self.DLName

    @dlname.setter
    def dlname(self, value):
        self.DLName = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return DistListItem(self.com_object.IsMarkedAsTask)

    # Lower case aliases for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MemberCount(self):
        return self.com_object.MemberCount

    # Lower case aliases for MemberCount
    @property
    def membercount(self):
        return self.MemberCount

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    # Lower case aliases for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    # Lower case aliases for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    # Lower case aliases for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    # Lower case aliases for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    # Lower case aliases for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def TaskCompletedDate(self):
        return DistListItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    # Lower case aliases for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return DistListItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    # Lower case aliases for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return DistListItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    # Lower case aliases for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return DistListItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    # Lower case aliases for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def ToDoTaskOrdinal(self):
        return DistListItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    # Lower case aliases for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def AddMember(self, Recipient=None):
        arguments = com_arguments([Recipient])
        self.com_object.AddMember(*arguments)

    # Lower case alias for AddMember
    def addmember(self, Recipient=None):
        arguments = [Recipient]
        return self.AddMember(*arguments)

    def AddMembers(self, Recipients=None):
        arguments = com_arguments([Recipients])
        self.com_object.AddMembers(*arguments)

    # Lower case alias for AddMembers
    def addmembers(self, Recipients=None):
        arguments = [Recipients]
        return self.AddMembers(*arguments)

    def ClearTaskFlag(self):
        self.com_object.ClearTaskFlag()

    # Lower case alias for ClearTaskFlag
    def cleartaskflag(self):
        return self.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def GetMember(self, Index=None):
        arguments = com_arguments([Index])
        return Recipient(self.com_object.GetMember(*arguments))

    # Lower case alias for GetMember
    def getmember(self, Index=None):
        arguments = [Index]
        return self.GetMember(*arguments)

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def RemoveMember(self, Recipient=None):
        arguments = com_arguments([Recipient])
        self.com_object.RemoveMember(*arguments)

    # Lower case alias for RemoveMember
    def removemember(self, Recipient=None):
        arguments = [Recipient]
        return self.RemoveMember(*arguments)

    def RemoveMembers(self, Recipients=None):
        arguments = com_arguments([Recipients])
        self.com_object.RemoveMembers(*arguments)

    # Lower case alias for RemoveMembers
    def removemembers(self, Recipients=None):
        arguments = [Recipients]
        return self.RemoveMembers(*arguments)

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class DocumentItem:

    def __init__(self, documentitem=None):
        self.com_object= documentitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return self.com_object.GetInspector

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return self.com_object.MarkForDownload

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class Exception:

    def __init__(self, exception=None):
        self.com_object= exception

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AppointmentItem(self):
        return AppointmentItem(self.com_object.AppointmentItem)

    # Lower case aliases for AppointmentItem
    @property
    def appointmentitem(self):
        return self.AppointmentItem

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Deleted(self):
        return AppointmentItem(self.com_object.Deleted)

    # Lower case aliases for Deleted
    @property
    def deleted(self):
        return self.Deleted

    @property
    def OriginalDate(self):
        return AppointmentItem(self.com_object.OriginalDate)

    # Lower case aliases for OriginalDate
    @property
    def originaldate(self):
        return self.OriginalDate

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class Exceptions:

    def __init__(self, exceptions=None):
        self.com_object= exceptions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Exception(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ExchangeDistributionList:

    def __init__(self, exchangedistributionlist=None):
        self.com_object= exchangedistributionlist

    @property
    def Address(self):
        return ExchangeDistributionList(self.com_object.Address)

    @Address.setter
    def Address(self, value):
        self.com_object.Address = value

    # Lower case aliases for Address
    @property
    def address(self):
        return self.Address

    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.com_object.AddressEntryUserType)

    # Lower case aliases for AddressEntryUserType
    @property
    def addressentryusertype(self):
        return self.AddressEntryUserType

    @property
    def Alias(self):
        return ExchangeDistributionList(self.com_object.Alias)

    # Lower case aliases for Alias
    @property
    def alias(self):
        return self.Alias

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Comments(self):
        return self.com_object.Comments

    @Comments.setter
    def Comments(self, value):
        self.com_object.Comments = value

    # Lower case aliases for Comments
    @property
    def comments(self):
        return self.Comments

    @comments.setter
    def comments(self, value):
        self.Comments = value

    @property
    def DisplayType(self):
        return OlDisplayType(self.com_object.DisplayType)

    # Lower case aliases for DisplayType
    @property
    def displaytype(self):
        return self.DisplayType

    @property
    def ID(self):
        return ExchangeDistributionList(self.com_object.ID)

    # Lower case aliases for ID
    @property
    def id(self):
        return self.ID

    @property
    def Name(self):
        return ExchangeDistributionList(self.com_object.Name)

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return ExchangeDistributionList(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PrimarySmtpAddress(self):
        return ExchangeDistributionList(self.com_object.PrimarySmtpAddress)

    # Lower case aliases for PrimarySmtpAddress
    @property
    def primarysmtpaddress(self):
        return self.PrimarySmtpAddress

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return ExchangeDistributionList(self.com_object.Type)

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @type.setter
    def type(self, value):
        self.Type = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([HWnd])
        self.com_object.Details(*arguments)

    # Lower case alias for Details
    def details(self, HWnd=None):
        arguments = [HWnd]
        return self.Details(*arguments)

    def GetContact(self):
        return self.com_object.GetContact()

    # Lower case alias for GetContact
    def getcontact(self):
        return self.GetContact()

    def GetExchangeDistributionList(self):
        return ExchangeDistributionList(self.com_object.GetExchangeDistributionList())

    # Lower case alias for GetExchangeDistributionList
    def getexchangedistributionlist(self):
        return self.GetExchangeDistributionList()

    def GetExchangeDistributionListMembers(self):
        return AddressEntry(self.com_object.GetExchangeDistributionListMembers())

    # Lower case alias for GetExchangeDistributionListMembers
    def getexchangedistributionlistmembers(self):
        return self.GetExchangeDistributionListMembers()

    def GetExchangeUser(self):
        return self.com_object.GetExchangeUser()

    # Lower case alias for GetExchangeUser
    def getexchangeuser(self):
        return self.GetExchangeUser()

    def GetFreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([Start, MinPerChar, CompleteFormat])
        self.com_object.GetFreeBusy(*arguments)

    # Lower case alias for GetFreeBusy
    def getfreebusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = [Start, MinPerChar, CompleteFormat]
        return self.GetFreeBusy(*arguments)

    def GetMemberOfList(self):
        return AddressEntries(self.com_object.GetMemberOfList())

    # Lower case alias for GetMemberOfList
    def getmemberoflist(self):
        return self.GetMemberOfList()

    def GetOwners(self):
        return AddressEntry(self.com_object.GetOwners())

    # Lower case alias for GetOwners
    def getowners(self):
        return self.GetOwners()

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([MakePermanent, Refresh])
        self.com_object.Update(*arguments)

    # Lower case alias for Update
    def update(self, MakePermanent=None, Refresh=None):
        arguments = [MakePermanent, Refresh]
        return self.Update(*arguments)


class ExchangeUser:

    def __init__(self, exchangeuser=None):
        self.com_object= exchangeuser

    @property
    def Address(self):
        return ExchangeUser(self.com_object.Address)

    @Address.setter
    def Address(self, value):
        self.com_object.Address = value

    # Lower case aliases for Address
    @property
    def address(self):
        return self.Address

    @address.setter
    def address(self, value):
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.com_object.AddressEntryUserType)

    # Lower case aliases for AddressEntryUserType
    @property
    def addressentryusertype(self):
        return self.AddressEntryUserType

    @property
    def Alias(self):
        return ExchangeUser(self.com_object.Alias)

    # Lower case aliases for Alias
    @property
    def alias(self):
        return self.Alias

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AssistantName(self):
        return ExchangeUser(self.com_object.AssistantName)

    @AssistantName.setter
    def AssistantName(self, value):
        self.com_object.AssistantName = value

    # Lower case aliases for AssistantName
    @property
    def assistantname(self):
        return self.AssistantName

    @assistantname.setter
    def assistantname(self, value):
        self.AssistantName = value

    @property
    def BusinessTelephoneNumber(self):
        return ExchangeUser(self.com_object.BusinessTelephoneNumber)

    @BusinessTelephoneNumber.setter
    def BusinessTelephoneNumber(self, value):
        self.com_object.BusinessTelephoneNumber = value

    # Lower case aliases for BusinessTelephoneNumber
    @property
    def businesstelephonenumber(self):
        return self.BusinessTelephoneNumber

    @businesstelephonenumber.setter
    def businesstelephonenumber(self, value):
        self.BusinessTelephoneNumber = value

    @property
    def City(self):
        return ExchangeUser(self.com_object.City)

    @City.setter
    def City(self, value):
        self.com_object.City = value

    # Lower case aliases for City
    @property
    def city(self):
        return self.City

    @city.setter
    def city(self, value):
        self.City = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Comments(self):
        return self.com_object.Comments

    @Comments.setter
    def Comments(self, value):
        self.com_object.Comments = value

    # Lower case aliases for Comments
    @property
    def comments(self):
        return self.Comments

    @comments.setter
    def comments(self, value):
        self.Comments = value

    @property
    def CompanyName(self):
        return ExchangeUser(self.com_object.CompanyName)

    @CompanyName.setter
    def CompanyName(self, value):
        self.com_object.CompanyName = value

    # Lower case aliases for CompanyName
    @property
    def companyname(self):
        return self.CompanyName

    @companyname.setter
    def companyname(self, value):
        self.CompanyName = value

    @property
    def Department(self):
        return ExchangeUser(self.com_object.Department)

    @Department.setter
    def Department(self, value):
        self.com_object.Department = value

    # Lower case aliases for Department
    @property
    def department(self):
        return self.Department

    @department.setter
    def department(self, value):
        self.Department = value

    @property
    def DisplayType(self):
        return OlDisplayType(self.com_object.DisplayType)

    # Lower case aliases for DisplayType
    @property
    def displaytype(self):
        return self.DisplayType

    @property
    def FirstName(self):
        return ExchangeUser(self.com_object.FirstName)

    @FirstName.setter
    def FirstName(self, value):
        self.com_object.FirstName = value

    # Lower case aliases for FirstName
    @property
    def firstname(self):
        return self.FirstName

    @firstname.setter
    def firstname(self, value):
        self.FirstName = value

    @property
    def ID(self):
        return ExchangeUser(self.com_object.ID)

    # Lower case aliases for ID
    @property
    def id(self):
        return self.ID

    @property
    def JobTitle(self):
        return ExchangeUser(self.com_object.JobTitle)

    @JobTitle.setter
    def JobTitle(self, value):
        self.com_object.JobTitle = value

    # Lower case aliases for JobTitle
    @property
    def jobtitle(self):
        return self.JobTitle

    @jobtitle.setter
    def jobtitle(self, value):
        self.JobTitle = value

    @property
    def LastName(self):
        return ExchangeUser(self.com_object.LastName)

    @LastName.setter
    def LastName(self, value):
        self.com_object.LastName = value

    # Lower case aliases for LastName
    @property
    def lastname(self):
        return self.LastName

    @lastname.setter
    def lastname(self, value):
        self.LastName = value

    @property
    def MobileTelephoneNumber(self):
        return ExchangeUser(self.com_object.MobileTelephoneNumber)

    @MobileTelephoneNumber.setter
    def MobileTelephoneNumber(self, value):
        self.com_object.MobileTelephoneNumber = value

    # Lower case aliases for MobileTelephoneNumber
    @property
    def mobiletelephonenumber(self):
        return self.MobileTelephoneNumber

    @mobiletelephonenumber.setter
    def mobiletelephonenumber(self, value):
        self.MobileTelephoneNumber = value

    @property
    def Name(self):
        return ExchangeUser(self.com_object.Name)

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def OfficeLocation(self):
        return ExchangeUser(self.com_object.OfficeLocation)

    @OfficeLocation.setter
    def OfficeLocation(self, value):
        self.com_object.OfficeLocation = value

    # Lower case aliases for OfficeLocation
    @property
    def officelocation(self):
        return self.OfficeLocation

    @officelocation.setter
    def officelocation(self, value):
        self.OfficeLocation = value

    @property
    def Parent(self):
        return ExchangeUser(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PostalCode(self):
        return ExchangeUser(self.com_object.PostalCode)

    @PostalCode.setter
    def PostalCode(self, value):
        self.com_object.PostalCode = value

    # Lower case aliases for PostalCode
    @property
    def postalcode(self):
        return self.PostalCode

    @postalcode.setter
    def postalcode(self, value):
        self.PostalCode = value

    @property
    def PrimarySmtpAddress(self):
        return ExchangeUser(self.com_object.PrimarySmtpAddress)

    # Lower case aliases for PrimarySmtpAddress
    @property
    def primarysmtpaddress(self):
        return self.PrimarySmtpAddress

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def StateOrProvince(self):
        return ExchangeUser(self.com_object.StateOrProvince)

    @StateOrProvince.setter
    def StateOrProvince(self, value):
        self.com_object.StateOrProvince = value

    # Lower case aliases for StateOrProvince
    @property
    def stateorprovince(self):
        return self.StateOrProvince

    @stateorprovince.setter
    def stateorprovince(self, value):
        self.StateOrProvince = value

    @property
    def StreetAddress(self):
        return ExchangeUser(self.com_object.StreetAddress)

    @StreetAddress.setter
    def StreetAddress(self, value):
        self.com_object.StreetAddress = value

    # Lower case aliases for StreetAddress
    @property
    def streetaddress(self):
        return self.StreetAddress

    @streetaddress.setter
    def streetaddress(self, value):
        self.StreetAddress = value

    @property
    def Type(self):
        return ExchangeUser(self.com_object.Type)

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def YomiCompanyName(self):
        return ExchangeUser(self.com_object.YomiCompanyName)

    @YomiCompanyName.setter
    def YomiCompanyName(self, value):
        self.com_object.YomiCompanyName = value

    # Lower case aliases for YomiCompanyName
    @property
    def yomicompanyname(self):
        return self.YomiCompanyName

    @yomicompanyname.setter
    def yomicompanyname(self, value):
        self.YomiCompanyName = value

    @property
    def YomiDepartment(self):
        return ExchangeUser(self.com_object.YomiDepartment)

    @YomiDepartment.setter
    def YomiDepartment(self, value):
        self.com_object.YomiDepartment = value

    # Lower case aliases for YomiDepartment
    @property
    def yomidepartment(self):
        return self.YomiDepartment

    @yomidepartment.setter
    def yomidepartment(self, value):
        self.YomiDepartment = value

    @property
    def YomiDisplayName(self):
        return ExchangeUser(self.com_object.YomiDisplayName)

    @YomiDisplayName.setter
    def YomiDisplayName(self, value):
        self.com_object.YomiDisplayName = value

    # Lower case aliases for YomiDisplayName
    @property
    def yomidisplayname(self):
        return self.YomiDisplayName

    @yomidisplayname.setter
    def yomidisplayname(self, value):
        self.YomiDisplayName = value

    @property
    def YomiFirstName(self):
        return ExchangeUser(self.com_object.YomiFirstName)

    @YomiFirstName.setter
    def YomiFirstName(self, value):
        self.com_object.YomiFirstName = value

    # Lower case aliases for YomiFirstName
    @property
    def yomifirstname(self):
        return self.YomiFirstName

    @yomifirstname.setter
    def yomifirstname(self, value):
        self.YomiFirstName = value

    @property
    def YomiLastName(self):
        return ExchangeUser(self.com_object.YomiLastName)

    @YomiLastName.setter
    def YomiLastName(self, value):
        self.com_object.YomiLastName = value

    # Lower case aliases for YomiLastName
    @property
    def yomilastname(self):
        return self.YomiLastName

    @yomilastname.setter
    def yomilastname(self, value):
        self.YomiLastName = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([HWnd])
        self.com_object.Details(*arguments)

    # Lower case alias for Details
    def details(self, HWnd=None):
        arguments = [HWnd]
        return self.Details(*arguments)

    def GetContact(self):
        return self.com_object.GetContact()

    # Lower case alias for GetContact
    def getcontact(self):
        return self.GetContact()

    def GetDirectReports(self):
        return AddressEntry(self.com_object.GetDirectReports())

    # Lower case alias for GetDirectReports
    def getdirectreports(self):
        return self.GetDirectReports()

    def GetExchangeDistributionList(self):
        return self.com_object.GetExchangeDistributionList()

    # Lower case alias for GetExchangeDistributionList
    def getexchangedistributionlist(self):
        return self.GetExchangeDistributionList()

    def GetExchangeUser(self):
        return ExchangeUser(self.com_object.GetExchangeUser())

    # Lower case alias for GetExchangeUser
    def getexchangeuser(self):
        return self.GetExchangeUser()

    def GetExchangeUserManager(self):
        return ExchangeUser(self.com_object.GetExchangeUserManager())

    # Lower case alias for GetExchangeUserManager
    def getexchangeusermanager(self):
        return self.GetExchangeUserManager()

    def GetFreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([Start, MinPerChar, CompleteFormat])
        return String(self.com_object.GetFreeBusy(*arguments))

    # Lower case alias for GetFreeBusy
    def getfreebusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = [Start, MinPerChar, CompleteFormat]
        return self.GetFreeBusy(*arguments)

    def GetMemberOfList(self):
        return ExchangeUser(self.com_object.GetMemberOfList())

    # Lower case alias for GetMemberOfList
    def getmemberoflist(self):
        return self.GetMemberOfList()

    def GetPicture(self):
        return IPictureDisp(self.com_object.GetPicture())

    # Lower case alias for GetPicture
    def getpicture(self):
        return self.GetPicture()

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([MakePermanent, Refresh])
        self.com_object.Update(*arguments)

    # Lower case alias for Update
    def update(self, MakePermanent=None, Refresh=None):
        arguments = [MakePermanent, Refresh]
        return self.Update(*arguments)


class Explorer:

    def __init__(self, explorer=None):
        self.com_object= explorer

    @property
    def AccountSelector(self):
        return AccountSelector(self.com_object.AccountSelector)

    # Lower case aliases for AccountSelector
    @property
    def accountselector(self):
        return self.AccountSelector

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AttachmentSelection(self):
        return AttachmentSelection(self.com_object.AttachmentSelection)

    # Lower case aliases for AttachmentSelection
    @property
    def attachmentselection(self):
        return self.AttachmentSelection

    @property
    def Caption(self):
        return self.com_object.Caption

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentFolder(self):
        return Folder(self.com_object.CurrentFolder)

    @CurrentFolder.setter
    def CurrentFolder(self, value):
        self.com_object.CurrentFolder = value

    # Lower case aliases for CurrentFolder
    @property
    def currentfolder(self):
        return self.CurrentFolder

    @currentfolder.setter
    def currentfolder(self, value):
        self.CurrentFolder = value

    @property
    def CurrentView(self):
        return self.com_object.CurrentView

    @CurrentView.setter
    def CurrentView(self, value):
        self.com_object.CurrentView = value

    # Lower case aliases for CurrentView
    @property
    def currentview(self):
        return self.CurrentView

    @currentview.setter
    def currentview(self, value):
        self.CurrentView = value

    @property
    def Height(self):
        return self.com_object.Height

    @Height.setter
    def Height(self, value):
        self.com_object.Height = value

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def HTMLDocument(self):
        return self.com_object.HTMLDocument

    # Lower case aliases for HTMLDocument
    @property
    def htmldocument(self):
        return self.HTMLDocument

    @property
    def Left(self):
        return self.com_object.Left

    @Left.setter
    def Left(self, value):
        self.com_object.Left = value

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def NavigationPane(self):
        return NavigationPane(self.com_object.NavigationPane)

    # Lower case aliases for NavigationPane
    @property
    def navigationpane(self):
        return self.NavigationPane

    @property
    def Panes(self):
        return Panes(self.com_object.Panes)

    # Lower case aliases for Panes
    @property
    def panes(self):
        return self.Panes

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Selection(self):
        return Selection(self.com_object.Selection)

    # Lower case aliases for Selection
    @property
    def selection(self):
        return self.Selection

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Top(self):
        return self.com_object.Top

    @Top.setter
    def Top(self, value):
        self.com_object.Top = value

    # Lower case aliases for Top
    @property
    def top(self):
        return self.Top

    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.com_object.Width

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def WindowState(self):
        return OlWindowState(self.com_object.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.com_object.WindowState = value

    # Lower case aliases for WindowState
    @property
    def windowstate(self):
        return self.WindowState

    @windowstate.setter
    def windowstate(self, value):
        self.WindowState = value

    def Activate(self):
        self.com_object.Activate()

    # Lower case alias for Activate
    def activate(self):
        return self.Activate()

    def AddToSelection(self, Item=None):
        arguments = com_arguments([Item])
        self.com_object.AddToSelection(*arguments)

    # Lower case alias for AddToSelection
    def addtoselection(self, Item=None):
        arguments = [Item]
        return self.AddToSelection(*arguments)

    def ClearSearch(self):
        self.com_object.ClearSearch()

    # Lower case alias for ClearSearch
    def clearsearch(self):
        return self.ClearSearch()

    def ClearSelection(self):
        self.com_object.ClearSelection()

    # Lower case alias for ClearSelection
    def clearselection(self):
        return self.ClearSelection()

    def Close(self):
        self.com_object.Close()

    # Lower case alias for Close
    def close(self):
        return self.Close()

    def Display(self):
        self.com_object.Display()

    # Lower case alias for Display
    def display(self):
        return self.Display()

    def IsItemSelectableInView(self, Item=None):
        arguments = com_arguments([Item])
        return Boolean(self.com_object.IsItemSelectableInView(*arguments))

    # Lower case alias for IsItemSelectableInView
    def isitemselectableinview(self, Item=None):
        arguments = [Item]
        return self.IsItemSelectableInView(*arguments)

    def IsPaneVisible(self, Pane=None):
        arguments = com_arguments([Pane])
        return self.com_object.IsPaneVisible(*arguments)

    # Lower case alias for IsPaneVisible
    def ispanevisible(self, Pane=None):
        arguments = [Pane]
        return self.IsPaneVisible(*arguments)

    def RemoveFromSelection(self, Item=None):
        arguments = com_arguments([Item])
        self.com_object.RemoveFromSelection(*arguments)

    # Lower case alias for RemoveFromSelection
    def removefromselection(self, Item=None):
        arguments = [Item]
        return self.RemoveFromSelection(*arguments)

    def Search(self, Query=None, SearchScope=None):
        arguments = com_arguments([Query, SearchScope])
        self.com_object.Search(*arguments)

    # Lower case alias for Search
    def search(self, Query=None, SearchScope=None):
        arguments = [Query, SearchScope]
        return self.Search(*arguments)

    def SelectAllItems(self):
        self.com_object.SelectAllItems()

    # Lower case alias for SelectAllItems
    def selectallitems(self):
        return self.SelectAllItems()

    def ShowPane(self, Pane=None, Visible=None):
        arguments = com_arguments([Pane, Visible])
        self.com_object.ShowPane(*arguments)

    # Lower case alias for ShowPane
    def showpane(self, Pane=None, Visible=None):
        arguments = [Pane, Visible]
        return self.ShowPane(*arguments)


class Explorers:

    def __init__(self, explorers=None):
        self.com_object= explorers

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Folder=None, DisplayMode=None):
        arguments = com_arguments([Folder, DisplayMode])
        return Explorer(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Folder=None, DisplayMode=None):
        arguments = [Folder, DisplayMode]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Explorer(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Folder:

    def __init__(self, folder=None):
        self.com_object= folder

    @property
    def AddressBookName(self):
        return Folder(self.com_object.AddressBookName)

    @AddressBookName.setter
    def AddressBookName(self, value):
        self.com_object.AddressBookName = value

    # Lower case aliases for AddressBookName
    @property
    def addressbookname(self):
        return self.AddressBookName

    @addressbookname.setter
    def addressbookname(self, value):
        self.AddressBookName = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentView(self):
        return View(self.com_object.CurrentView)

    # Lower case aliases for CurrentView
    @property
    def currentview(self):
        return self.CurrentView

    @property
    def CustomViewsOnly(self):
        return self.com_object.CustomViewsOnly

    @CustomViewsOnly.setter
    def CustomViewsOnly(self, value):
        self.com_object.CustomViewsOnly = value

    # Lower case aliases for CustomViewsOnly
    @property
    def customviewsonly(self):
        return self.CustomViewsOnly

    @customviewsonly.setter
    def customviewsonly(self, value):
        self.CustomViewsOnly = value

    @property
    def DefaultItemType(self):
        return OlItemType(self.com_object.DefaultItemType)

    # Lower case aliases for DefaultItemType
    @property
    def defaultitemtype(self):
        return self.DefaultItemType

    @property
    def DefaultMessageClass(self):
        return self.com_object.DefaultMessageClass

    # Lower case aliases for DefaultMessageClass
    @property
    def defaultmessageclass(self):
        return self.DefaultMessageClass

    @property
    def Description(self):
        return self.com_object.Description

    @Description.setter
    def Description(self, value):
        self.com_object.Description = value

    # Lower case aliases for Description
    @property
    def description(self):
        return self.Description

    @description.setter
    def description(self, value):
        self.Description = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FolderPath(self):
        return self.com_object.FolderPath

    # Lower case aliases for FolderPath
    @property
    def folderpath(self):
        return self.FolderPath

    @property
    def Folders(self):
        return Folders(self.com_object.Folders)

    # Lower case aliases for Folders
    @property
    def folders(self):
        return self.Folders

    @property
    def InAppFolderSyncObject(self):
        return self.com_object.InAppFolderSyncObject

    @InAppFolderSyncObject.setter
    def InAppFolderSyncObject(self, value):
        self.com_object.InAppFolderSyncObject = value

    # Lower case aliases for InAppFolderSyncObject
    @property
    def inappfoldersyncobject(self):
        return self.InAppFolderSyncObject

    @inappfoldersyncobject.setter
    def inappfoldersyncobject(self, value):
        self.InAppFolderSyncObject = value

    @property
    def IsSharePointFolder(self):
        return self.com_object.IsSharePointFolder

    # Lower case aliases for IsSharePointFolder
    @property
    def issharepointfolder(self):
        return self.IsSharePointFolder

    @property
    def Items(self):
        return Items(self.com_object.Items)

    # Lower case aliases for Items
    @property
    def items(self):
        return self.Items

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowAsOutlookAB(self):
        return self.com_object.ShowAsOutlookAB

    @ShowAsOutlookAB.setter
    def ShowAsOutlookAB(self, value):
        self.com_object.ShowAsOutlookAB = value

    # Lower case aliases for ShowAsOutlookAB
    @property
    def showasoutlookab(self):
        return self.ShowAsOutlookAB

    @showasoutlookab.setter
    def showasoutlookab(self, value):
        self.ShowAsOutlookAB = value

    @property
    def ShowItemCount(self):
        return self.com_object.ShowItemCount

    @ShowItemCount.setter
    def ShowItemCount(self, value):
        self.com_object.ShowItemCount = value

    # Lower case aliases for ShowItemCount
    @property
    def showitemcount(self):
        return self.ShowItemCount

    @showitemcount.setter
    def showitemcount(self, value):
        self.ShowItemCount = value

    @property
    def Store(self):
        return Store(self.com_object.Store)

    # Lower case aliases for Store
    @property
    def store(self):
        return self.Store

    @property
    def StoreID(self):
        return self.com_object.StoreID

    # Lower case aliases for StoreID
    @property
    def storeid(self):
        return self.StoreID

    @property
    def UnReadItemCount(self):
        return self.com_object.UnReadItemCount

    # Lower case aliases for UnReadItemCount
    @property
    def unreaditemcount(self):
        return self.UnReadItemCount

    @property
    def UserDefinedProperties(self):
        return UserDefinedProperties(self.com_object.UserDefinedProperties)

    # Lower case aliases for UserDefinedProperties
    @property
    def userdefinedproperties(self):
        return self.UserDefinedProperties

    @property
    def Views(self):
        return Views(self.com_object.Views)

    # Lower case aliases for Views
    @property
    def views(self):
        return self.Views

    @property
    def WebViewOn(self):
        return self.com_object.WebViewOn

    @WebViewOn.setter
    def WebViewOn(self, value):
        self.com_object.WebViewOn = value

    # Lower case aliases for WebViewOn
    @property
    def webviewon(self):
        return self.WebViewOn

    @webviewon.setter
    def webviewon(self, value):
        self.WebViewOn = value

    @property
    def WebViewURL(self):
        return self.com_object.WebViewURL

    @WebViewURL.setter
    def WebViewURL(self, value):
        self.com_object.WebViewURL = value

    # Lower case aliases for WebViewURL
    @property
    def webviewurl(self):
        return self.WebViewURL

    @webviewurl.setter
    def webviewurl(self, value):
        self.WebViewURL = value

    def AddToPFFavorites(self):
        self.com_object.AddToPFFavorites()

    # Lower case alias for AddToPFFavorites
    def addtopffavorites(self):
        return self.AddToPFFavorites()

    def CopyTo(self, DestinationFolder=None):
        arguments = com_arguments([DestinationFolder])
        return Folder(self.com_object.CopyTo(*arguments))

    # Lower case alias for CopyTo
    def copyto(self, DestinationFolder=None):
        arguments = [DestinationFolder]
        return self.CopyTo(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self):
        self.com_object.Display()

    # Lower case alias for Display
    def display(self):
        return self.Display()

    def GetCalendarExporter(self):
        return CalendarSharing(self.com_object.GetCalendarExporter())

    # Lower case alias for GetCalendarExporter
    def getcalendarexporter(self):
        return self.GetCalendarExporter()

    def GetCustomIcon(self):
        return IPictureDisp(self.com_object.GetCustomIcon())

    # Lower case alias for GetCustomIcon
    def getcustomicon(self):
        return self.GetCustomIcon()

    def GetExplorer(self, DisplayMode=None):
        arguments = com_arguments([DisplayMode])
        return Explorer(self.com_object.GetExplorer(*arguments))

    # Lower case alias for GetExplorer
    def getexplorer(self, DisplayMode=None):
        arguments = [DisplayMode]
        return self.GetExplorer(*arguments)

    def GetOrganizer(self):
        return AddressEntry(self.com_object.GetOrganizer())

    # Lower case alias for GetOrganizer
    def getorganizer(self):
        return self.GetOrganizer()

    def GetStorage(self, StorageIdentifier=None, StorageIdentifierType=None):
        arguments = com_arguments([StorageIdentifier, StorageIdentifierType])
        return StorageItem(self.com_object.GetStorage(*arguments))

    # Lower case alias for GetStorage
    def getstorage(self, StorageIdentifier=None, StorageIdentifierType=None):
        arguments = [StorageIdentifier, StorageIdentifierType]
        return self.GetStorage(*arguments)

    def GetTable(self, Filter=None, TableContents=None):
        arguments = com_arguments([Filter, TableContents])
        return Folder(self.com_object.GetTable(*arguments))

    # Lower case alias for GetTable
    def gettable(self, Filter=None, TableContents=None):
        arguments = [Filter, TableContents]
        return self.GetTable(*arguments)

    def MoveTo(self, DestinationFolder=None):
        arguments = com_arguments([DestinationFolder])
        self.com_object.MoveTo(*arguments)

    # Lower case alias for MoveTo
    def moveto(self, DestinationFolder=None):
        arguments = [DestinationFolder]
        return self.MoveTo(*arguments)

    def SetCustomIcon(self, Picture=None):
        arguments = com_arguments([Picture])
        self.com_object.SetCustomIcon(*arguments)

    # Lower case alias for SetCustomIcon
    def setcustomicon(self, Picture=None):
        arguments = [Picture]
        return self.SetCustomIcon(*arguments)


class Folders:

    def __init__(self, folders=None):
        self.com_object= folders

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Type=None):
        arguments = com_arguments([Name, Type])
        return Folder(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Type=None):
        arguments = [Name, Type]
        return self.Add(*arguments)

    def GetFirst(self):
        return Folder(self.com_object.GetFirst())

    # Lower case alias for GetFirst
    def getfirst(self):
        return self.GetFirst()

    def GetLast(self):
        return Folder(self.com_object.GetLast())

    # Lower case alias for GetLast
    def getlast(self):
        return self.GetLast()

    def GetNext(self):
        return Folder(self.com_object.GetNext())

    # Lower case alias for GetNext
    def getnext(self):
        return self.GetNext()

    def GetPrevious(self):
        return Folder(self.com_object.GetPrevious())

    # Lower case alias for GetPrevious
    def getprevious(self):
        return self.GetPrevious()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Folder(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class FormDescription:

    def __init__(self, formdescription=None):
        self.com_object= formdescription

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Category(self):
        return self.com_object.Category

    @Category.setter
    def Category(self, value):
        self.com_object.Category = value

    # Lower case aliases for Category
    @property
    def category(self):
        return self.Category

    @category.setter
    def category(self, value):
        self.Category = value

    @property
    def CategorySub(self):
        return self.com_object.CategorySub

    @CategorySub.setter
    def CategorySub(self, value):
        self.com_object.CategorySub = value

    # Lower case aliases for CategorySub
    @property
    def categorysub(self):
        return self.CategorySub

    @categorysub.setter
    def categorysub(self, value):
        self.CategorySub = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Comment(self):
        return self.com_object.Comment

    @Comment.setter
    def Comment(self, value):
        self.com_object.Comment = value

    # Lower case aliases for Comment
    @property
    def comment(self):
        return self.Comment

    @comment.setter
    def comment(self, value):
        self.Comment = value

    @property
    def ContactName(self):
        return FormDescription(self.com_object.ContactName)

    @ContactName.setter
    def ContactName(self, value):
        self.com_object.ContactName = value

    # Lower case aliases for ContactName
    @property
    def contactname(self):
        return self.ContactName

    @contactname.setter
    def contactname(self, value):
        self.ContactName = value

    @property
    def DisplayName(self):
        return self.com_object.DisplayName

    @DisplayName.setter
    def DisplayName(self, value):
        self.com_object.DisplayName = value

    # Lower case aliases for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @displayname.setter
    def displayname(self, value):
        self.DisplayName = value

    @property
    def Hidden(self):
        return self.com_object.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.com_object.Hidden = value

    # Lower case aliases for Hidden
    @property
    def hidden(self):
        return self.Hidden

    @hidden.setter
    def hidden(self, value):
        self.Hidden = value

    @property
    def Icon(self):
        return self.com_object.Icon

    @Icon.setter
    def Icon(self, value):
        self.com_object.Icon = value

    # Lower case aliases for Icon
    @property
    def icon(self):
        return self.Icon

    @icon.setter
    def icon(self, value):
        self.Icon = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    # Lower case aliases for Locked
    @property
    def locked(self):
        return self.Locked

    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MessageClass(self):
        return FormDescription(self.com_object.MessageClass)

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @property
    def MiniIcon(self):
        return self.com_object.MiniIcon

    @MiniIcon.setter
    def MiniIcon(self, value):
        self.com_object.MiniIcon = value

    # Lower case aliases for MiniIcon
    @property
    def miniicon(self):
        return self.MiniIcon

    @miniicon.setter
    def miniicon(self, value):
        self.MiniIcon = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Number(self):
        return self.com_object.Number

    @Number.setter
    def Number(self, value):
        self.com_object.Number = value

    # Lower case aliases for Number
    @property
    def number(self):
        return self.Number

    @number.setter
    def number(self, value):
        self.Number = value

    @property
    def OneOff(self):
        return self.com_object.OneOff

    @OneOff.setter
    def OneOff(self, value):
        self.com_object.OneOff = value

    # Lower case aliases for OneOff
    @property
    def oneoff(self):
        return self.OneOff

    @oneoff.setter
    def oneoff(self, value):
        self.OneOff = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def ScriptText(self):
        return self.com_object.ScriptText

    # Lower case aliases for ScriptText
    @property
    def scripttext(self):
        return self.ScriptText

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Template(self):
        return self.com_object.Template

    @Template.setter
    def Template(self, value):
        self.com_object.Template = value

    # Lower case aliases for Template
    @property
    def template(self):
        return self.Template

    @template.setter
    def template(self, value):
        self.Template = value

    @property
    def UseWordMail(self):
        return self.com_object.UseWordMail

    @UseWordMail.setter
    def UseWordMail(self, value):
        self.com_object.UseWordMail = value

    # Lower case aliases for UseWordMail
    @property
    def usewordmail(self):
        return self.UseWordMail

    @usewordmail.setter
    def usewordmail(self, value):
        self.UseWordMail = value

    @property
    def Version(self):
        return self.com_object.Version

    @Version.setter
    def Version(self, value):
        self.com_object.Version = value

    # Lower case aliases for Version
    @property
    def version(self):
        return self.Version

    @version.setter
    def version(self, value):
        self.Version = value

    def PublishForm(self, Registry=None, Folder=None):
        arguments = com_arguments([Registry, Folder])
        self.com_object.PublishForm(*arguments)

    # Lower case alias for PublishForm
    def publishform(self, Registry=None, Folder=None):
        arguments = [Registry, Folder]
        return self.PublishForm(*arguments)


class FormNameRuleCondition:

    def __init__(self, formnamerulecondition=None):
        self.com_object= formnamerulecondition

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def FormName(self):
        return self.com_object.FormName

    @FormName.setter
    def FormName(self, value):
        self.com_object.FormName = value

    # Lower case aliases for FormName
    @property
    def formname(self):
        return self.FormName

    @formname.setter
    def formname(self, value):
        self.FormName = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class FormRegion:

    def __init__(self, formregion=None):
        self.com_object= formregion

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Detail(self):
        return self.com_object.Detail

    @Detail.setter
    def Detail(self, value):
        self.com_object.Detail = value

    # Lower case aliases for Detail
    @property
    def detail(self):
        return self.Detail

    @detail.setter
    def detail(self, value):
        self.Detail = value

    @property
    def DisplayName(self):
        return self.com_object.DisplayName

    # Lower case aliases for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def EnableAutoLayout(self):
        return self.com_object.EnableAutoLayout

    @EnableAutoLayout.setter
    def EnableAutoLayout(self, value):
        self.com_object.EnableAutoLayout = value

    # Lower case aliases for EnableAutoLayout
    @property
    def enableautolayout(self):
        return self.EnableAutoLayout

    @enableautolayout.setter
    def enableautolayout(self, value):
        self.EnableAutoLayout = value

    @property
    def Form(self):
        return self.com_object.Form

    # Lower case aliases for Form
    @property
    def form(self):
        return self.Form

    @property
    def FormRegionMode(self):
        return OlFormRegionMode(self.com_object.FormRegionMode)

    # Lower case aliases for FormRegionMode
    @property
    def formregionmode(self):
        return self.FormRegionMode

    @property
    def Inspector(self):
        return Inspector(self.com_object.Inspector)

    # Lower case aliases for Inspector
    @property
    def inspector(self):
        return self.Inspector

    @property
    def InternalName(self):
        return self.com_object.InternalName

    # Lower case aliases for InternalName
    @property
    def internalname(self):
        return self.InternalName

    @property
    def IsExpanded(self):
        return self.com_object.IsExpanded

    # Lower case aliases for IsExpanded
    @property
    def isexpanded(self):
        return self.IsExpanded

    @property
    def Item(self):
        return self.com_object.Item

    # Lower case aliases for Item
    @property
    def item(self):
        return self.Item

    @property
    def Language(self):
        return self.com_object.Language

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def SuppressControlReplacement(self):
        return self.com_object.SuppressControlReplacement

    @SuppressControlReplacement.setter
    def SuppressControlReplacement(self, value):
        self.com_object.SuppressControlReplacement = value

    # Lower case aliases for SuppressControlReplacement
    @property
    def suppresscontrolreplacement(self):
        return self.SuppressControlReplacement

    @suppresscontrolreplacement.setter
    def suppresscontrolreplacement(self, value):
        self.SuppressControlReplacement = value

    @property
    def Visible(self):
        return self.com_object.Visible

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value

    def Reflow(self):
        self.com_object.Reflow()

    # Lower case alias for Reflow
    def reflow(self):
        return self.Reflow()

    def Select(self):
        self.com_object.Select()

    # Lower case alias for Select
    def select(self):
        return self.Select()

    def SetControlItemProperty(self, Control=None, PropertyName=None):
        arguments = com_arguments([Control, PropertyName])
        self.com_object.SetControlItemProperty(*arguments)

    # Lower case alias for SetControlItemProperty
    def setcontrolitemproperty(self, Control=None, PropertyName=None):
        arguments = [Control, PropertyName]
        return self.SetControlItemProperty(*arguments)


class FromRssFeedRuleCondition:

    def __init__(self, fromrssfeedrulecondition=None):
        self.com_object= fromrssfeedrulecondition

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def FromRssFeed(self):
        return self.com_object.FromRssFeed

    @FromRssFeed.setter
    def FromRssFeed(self, value):
        self.com_object.FromRssFeed = value

    # Lower case aliases for FromRssFeed
    @property
    def fromrssfeed(self):
        return self.FromRssFeed

    @fromrssfeed.setter
    def fromrssfeed(self, value):
        self.FromRssFeed = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class IconView:

    def __init__(self, iconview=None):
        self.com_object= iconview

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def IconPlacement(self):
        return OlIconViewPlacement(self.com_object.IconPlacement)

    @IconPlacement.setter
    def IconPlacement(self, value):
        self.com_object.IconPlacement = value

    # Lower case aliases for IconPlacement
    @property
    def iconplacement(self):
        return self.IconPlacement

    @iconplacement.setter
    def iconplacement(self, value):
        self.IconPlacement = value

    @property
    def IconViewType(self):
        return OlIconViewType(self.com_object.IconViewType)

    @IconViewType.setter
    def IconViewType(self, value):
        self.com_object.IconViewType = value

    # Lower case aliases for IconViewType
    @property
    def iconviewtype(self):
        return self.IconViewType

    @iconviewtype.setter
    def iconviewtype(self, value):
        self.IconViewType = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    # Lower case aliases for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    # Lower case aliases for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def SortFields(self):
        return OrderFields(self.com_object.SortFields)

    # Lower case aliases for SortFields
    @property
    def sortfields(self):
        return self.SortFields

    @property
    def Standard(self):
        return IconView(self.com_object.Standard)

    # Lower case aliases for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    # Lower case aliases for XML
    @property
    def xml(self):
        return self.XML

    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return View(self.com_object.Copy(*arguments))

    # Lower case alias for Copy
    def copy(self, Name=None, SaveOption=None):
        arguments = [Name, SaveOption]
        return self.Copy(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.com_object.GoToDate(*arguments)

    # Lower case alias for GoToDate
    def gotodate(self, Date=None):
        arguments = [Date]
        return self.GoToDate(*arguments)

    def Reset(self):
        self.com_object.Reset()

    # Lower case alias for Reset
    def reset(self):
        return self.Reset()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class ImportanceRuleCondition:

    def __init__(self, importancerulecondition=None):
        self.com_object= importancerulecondition

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class Inspector:

    def __init__(self, inspector=None):
        self.com_object= inspector

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AttachmentSelection(self):
        return AttachmentSelection(self.com_object.AttachmentSelection)

    # Lower case aliases for AttachmentSelection
    @property
    def attachmentselection(self):
        return self.AttachmentSelection

    @property
    def Caption(self):
        return self.com_object.Caption

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentItem(self):
        return self.com_object.CurrentItem

    # Lower case aliases for CurrentItem
    @property
    def currentitem(self):
        return self.CurrentItem

    @property
    def EditorType(self):
        return OlEditorType(self.com_object.EditorType)

    # Lower case aliases for EditorType
    @property
    def editortype(self):
        return self.EditorType

    @property
    def Height(self):
        return self.com_object.Height

    @Height.setter
    def Height(self, value):
        self.com_object.Height = value

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def Left(self):
        return self.com_object.Left

    @Left.setter
    def Left(self, value):
        self.com_object.Left = value

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def ModifiedFormPages(self):
        return Pages(self.com_object.ModifiedFormPages)

    # Lower case aliases for ModifiedFormPages
    @property
    def modifiedformpages(self):
        return self.ModifiedFormPages

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Top(self):
        return self.com_object.Top

    @Top.setter
    def Top(self, value):
        self.com_object.Top = value

    # Lower case aliases for Top
    @property
    def top(self):
        return self.Top

    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.com_object.Width

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    @width.setter
    def width(self, value):
        self.Width = value

    @property
    def WindowState(self):
        return OlWindowState(self.com_object.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.com_object.WindowState = value

    # Lower case aliases for WindowState
    @property
    def windowstate(self):
        return self.WindowState

    @windowstate.setter
    def windowstate(self, value):
        self.WindowState = value

    @property
    def WordEditor(self):
        return self.com_object.WordEditor

    # Lower case aliases for WordEditor
    @property
    def wordeditor(self):
        return self.WordEditor

    def Activate(self):
        self.com_object.Activate()

    # Lower case alias for Activate
    def activate(self):
        return self.Activate()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def HideFormPage(self, PageName=None):
        arguments = com_arguments([PageName])
        self.com_object.HideFormPage(*arguments)

    # Lower case alias for HideFormPage
    def hideformpage(self, PageName=None):
        arguments = [PageName]
        return self.HideFormPage(*arguments)

    def IsWordMail(self):
        return self.com_object.IsWordMail()

    # Lower case alias for IsWordMail
    def iswordmail(self):
        return self.IsWordMail()

    def NewFormRegion(self):
        return Object(self.com_object.NewFormRegion())

    # Lower case alias for NewFormRegion
    def newformregion(self):
        return self.NewFormRegion()

    def OpenFormRegion(self, Path=None):
        arguments = com_arguments([Path])
        return Object(self.com_object.OpenFormRegion(*arguments))

    # Lower case alias for OpenFormRegion
    def openformregion(self, Path=None):
        arguments = [Path]
        return self.OpenFormRegion(*arguments)

    def SaveFormRegion(self, Page=None, FileName=None):
        arguments = com_arguments([Page, FileName])
        self.com_object.SaveFormRegion(*arguments)

    # Lower case alias for SaveFormRegion
    def saveformregion(self, Page=None, FileName=None):
        arguments = [Page, FileName]
        return self.SaveFormRegion(*arguments)

    def SetControlItemProperty(self, Control=None, PropertyName=None):
        arguments = com_arguments([Control, PropertyName])
        self.com_object.SetControlItemProperty(*arguments)

    # Lower case alias for SetControlItemProperty
    def setcontrolitemproperty(self, Control=None, PropertyName=None):
        arguments = [Control, PropertyName]
        return self.SetControlItemProperty(*arguments)

    def SetCurrentFormPage(self, PageName=None):
        arguments = com_arguments([PageName])
        self.com_object.SetCurrentFormPage(*arguments)

    # Lower case alias for SetCurrentFormPage
    def setcurrentformpage(self, PageName=None):
        arguments = [PageName]
        return self.SetCurrentFormPage(*arguments)

    def SetSchedulingStartTime(self, Start=None):
        arguments = com_arguments([Start])
        self.com_object.SetSchedulingStartTime(*arguments)

    # Lower case alias for SetSchedulingStartTime
    def setschedulingstarttime(self, Start=None):
        arguments = [Start]
        return self.SetSchedulingStartTime(*arguments)

    def ShowFormPage(self, PageName=None):
        arguments = com_arguments([PageName])
        self.com_object.ShowFormPage(*arguments)

    # Lower case alias for ShowFormPage
    def showformpage(self, PageName=None):
        arguments = [PageName]
        return self.ShowFormPage(*arguments)


class Inspectors:

    def __init__(self, inspectors=None):
        self.com_object= inspectors

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self):
        return Inspector(self.com_object.Add())

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Inspector(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ItemProperties:

    def __init__(self, itemproperties=None):
        self.com_object= itemproperties

    def __call__(self, item):
        return ItemPropertie(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = com_arguments([Name, Type, AddToFolderFields, DisplayFormat])
        self.com_object.Add(*arguments)

    # Lower case alias for Add
    def add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = [Name, Type, AddToFolderFields, DisplayFormat]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return ItemProperty(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class ItemProperty:

    def __init__(self, itemproperty=None):
        self.com_object= itemproperty

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def IsUserProperty(self):
        return self.com_object.IsUserProperty

    # Lower case aliases for IsUserProperty
    @property
    def isuserproperty(self):
        return self.IsUserProperty

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.com_object.Type)

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value


class Items:

    def __init__(self, items=None):
        self.com_object= items

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def IncludeRecurrences(self):
        return Items(self.com_object.IncludeRecurrences)

    @IncludeRecurrences.setter
    def IncludeRecurrences(self, value):
        self.com_object.IncludeRecurrences = value

    # Lower case aliases for IncludeRecurrences
    @property
    def includerecurrences(self):
        return self.IncludeRecurrences

    @includerecurrences.setter
    def includerecurrences(self, value):
        self.IncludeRecurrences = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self):
        return Object(self.com_object.Add())

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Find(self, Filter=None):
        arguments = com_arguments([Filter])
        return self.com_object.Find(*arguments)

    # Lower case alias for Find
    def find(self, Filter=None):
        arguments = [Filter]
        return self.Find(*arguments)

    def FindNext(self):
        return Object(self.com_object.FindNext())

    # Lower case alias for FindNext
    def findnext(self):
        return self.FindNext()

    def GetFirst(self):
        return Object(self.com_object.GetFirst())

    # Lower case alias for GetFirst
    def getfirst(self):
        return self.GetFirst()

    def GetLast(self):
        return Object(self.com_object.GetLast())

    # Lower case alias for GetLast
    def getlast(self):
        return self.GetLast()

    def GetNext(self):
        return Object(self.com_object.GetNext())

    # Lower case alias for GetNext
    def getnext(self):
        return self.GetNext()

    def GetPrevious(self):
        return Object(self.com_object.GetPrevious())

    # Lower case alias for GetPrevious
    def getprevious(self):
        return self.GetPrevious()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Object(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

    def ResetColumns(self):
        self.com_object.ResetColumns()

    # Lower case alias for ResetColumns
    def resetcolumns(self):
        return self.ResetColumns()

    def Restrict(self, Filter=None):
        arguments = com_arguments([Filter])
        return Items(self.com_object.Restrict(*arguments))

    # Lower case alias for Restrict
    def restrict(self, Filter=None):
        arguments = [Filter]
        return self.Restrict(*arguments)

    def SetColumns(self, Columns=None):
        arguments = com_arguments([Columns])
        self.com_object.SetColumns(*arguments)

    # Lower case alias for SetColumns
    def setcolumns(self, Columns=None):
        arguments = [Columns]
        return self.SetColumns(*arguments)

    def Sort(self, Property=None, Descending=None):
        arguments = com_arguments([Property, Descending])
        self.com_object.Sort(*arguments)

    # Lower case alias for Sort
    def sort(self, Property=None, Descending=None):
        arguments = [Property, Descending]
        return self.Sort(*arguments)


class JournalItem:

    def __init__(self, journalitem=None):
        self.com_object= journalitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return Conflicts(self.com_object.Conflicts)

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ContactNames(self):
        return self.com_object.ContactNames

    @ContactNames.setter
    def ContactNames(self, value):
        self.com_object.ContactNames = value

    # Lower case aliases for ContactNames
    @property
    def contactnames(self):
        return self.ContactNames

    @contactnames.setter
    def contactnames(self, value):
        self.ContactNames = value

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DocPosted(self):
        return self.com_object.DocPosted

    @DocPosted.setter
    def DocPosted(self, value):
        self.com_object.DocPosted = value

    # Lower case aliases for DocPosted
    @property
    def docposted(self):
        return self.DocPosted

    @docposted.setter
    def docposted(self, value):
        self.DocPosted = value

    @property
    def DocPrinted(self):
        return self.com_object.DocPrinted

    @DocPrinted.setter
    def DocPrinted(self, value):
        self.com_object.DocPrinted = value

    # Lower case aliases for DocPrinted
    @property
    def docprinted(self):
        return self.DocPrinted

    @docprinted.setter
    def docprinted(self, value):
        self.DocPrinted = value

    @property
    def DocRouted(self):
        return self.com_object.DocRouted

    @DocRouted.setter
    def DocRouted(self, value):
        self.com_object.DocRouted = value

    # Lower case aliases for DocRouted
    @property
    def docrouted(self):
        return self.DocRouted

    @docrouted.setter
    def docrouted(self, value):
        self.DocRouted = value

    @property
    def DocSaved(self):
        return self.com_object.DocSaved

    @DocSaved.setter
    def DocSaved(self, value):
        self.com_object.DocSaved = value

    # Lower case aliases for DocSaved
    @property
    def docsaved(self):
        return self.DocSaved

    @docsaved.setter
    def docsaved(self, value):
        self.DocSaved = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def Duration(self):
        return JournalItem(self.com_object.Duration)

    @Duration.setter
    def Duration(self, value):
        self.com_object.Duration = value

    # Lower case aliases for Duration
    @property
    def duration(self):
        return self.Duration

    @duration.setter
    def duration(self, value):
        self.Duration = value

    @property
    def End(self):
        return self.com_object.End

    @End.setter
    def End(self, value):
        self.com_object.End = value

    # Lower case aliases for End
    @property
    def end(self):
        return self.End

    @end.setter
    def end(self, value):
        self.End = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Start(self):
        return self.com_object.Start

    @Start.setter
    def Start(self, value):
        self.com_object.Start = value

    # Lower case aliases for Start
    @property
    def start(self):
        return self.Start

    @start.setter
    def start(self, value):
        self.Start = value

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Type(self):
        return self.com_object.Type

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Forward(self):
        return MailItem(self.com_object.Forward())

    # Lower case alias for Forward
    def forward(self):
        return self.Forward()

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Reply(self):
        return MailItem(self.com_object.Reply())

    # Lower case alias for Reply
    def reply(self):
        return self.Reply()

    def ReplyAll(self):
        return MailItem(self.com_object.ReplyAll())

    # Lower case alias for ReplyAll
    def replyall(self):
        return self.ReplyAll()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()

    def StartTimer(self):
        self.com_object.StartTimer()

    # Lower case alias for StartTimer
    def starttimer(self):
        return self.StartTimer()

    def StopTimer(self):
        self.com_object.StopTimer()

    # Lower case alias for StopTimer
    def stoptimer(self):
        return self.StopTimer()


class JournalModule:

    def __init__(self, journalmodule=None):
        self.com_object= journalmodule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return JournalModule(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    # Lower case aliases for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    # Lower case aliases for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return JournalModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return JournalModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value


class MailItem:

    def __init__(self, mailitem=None):
        self.com_object= mailitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def AlternateRecipientAllowed(self):
        return self.com_object.AlternateRecipientAllowed

    @AlternateRecipientAllowed.setter
    def AlternateRecipientAllowed(self, value):
        self.com_object.AlternateRecipientAllowed = value

    # Lower case aliases for AlternateRecipientAllowed
    @property
    def alternaterecipientallowed(self):
        return self.AlternateRecipientAllowed

    @alternaterecipientallowed.setter
    def alternaterecipientallowed(self, value):
        self.AlternateRecipientAllowed = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.com_object.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.com_object.AutoForwarded = value

    # Lower case aliases for AutoForwarded
    @property
    def autoforwarded(self):
        return self.AutoForwarded

    @autoforwarded.setter
    def autoforwarded(self, value):
        self.AutoForwarded = value

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BCC(self):
        return MailItem(self.com_object.BCC)

    @BCC.setter
    def BCC(self, value):
        self.com_object.BCC = value

    # Lower case aliases for BCC
    @property
    def bcc(self):
        return self.BCC

    @bcc.setter
    def bcc(self, value):
        self.BCC = value

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.com_object.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.com_object.BodyFormat = value

    # Lower case aliases for BodyFormat
    @property
    def bodyformat(self):
        return self.BodyFormat

    @bodyformat.setter
    def bodyformat(self, value):
        self.BodyFormat = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def CC(self):
        return MailItem(self.com_object.CC)

    @CC.setter
    def CC(self, value):
        self.com_object.CC = value

    # Lower case aliases for CC
    @property
    def cc(self):
        return self.CC

    @cc.setter
    def cc(self, value):
        self.CC = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return self.com_object.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.com_object.DeferredDeliveryTime = value

    # Lower case aliases for DeferredDeliveryTime
    @property
    def deferreddeliverytime(self):
        return self.DeferredDeliveryTime

    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.com_object.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.com_object.DeleteAfterSubmit = value

    # Lower case aliases for DeleteAfterSubmit
    @property
    def deleteaftersubmit(self):
        return self.DeleteAfterSubmit

    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.com_object.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.com_object.ExpiryTime = value

    # Lower case aliases for ExpiryTime
    @property
    def expirytime(self):
        return self.ExpiryTime

    @expirytime.setter
    def expirytime(self, value):
        self.ExpiryTime = value

    @property
    def FlagRequest(self):
        return self.com_object.FlagRequest

    @FlagRequest.setter
    def FlagRequest(self, value):
        self.com_object.FlagRequest = value

    # Lower case aliases for FlagRequest
    @property
    def flagrequest(self):
        return self.FlagRequest

    @flagrequest.setter
    def flagrequest(self, value):
        self.FlagRequest = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def HTMLBody(self):
        return self.com_object.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.com_object.HTMLBody = value

    # Lower case aliases for HTMLBody
    @property
    def htmlbody(self):
        return self.HTMLBody

    @htmlbody.setter
    def htmlbody(self, value):
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    # Lower case aliases for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return MailItem(self.com_object.IsMarkedAsTask)

    # Lower case aliases for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return self.com_object.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.com_object.OriginatorDeliveryReportRequested = value

    # Lower case aliases for OriginatorDeliveryReportRequested
    @property
    def originatordeliveryreportrequested(self):
        return self.OriginatorDeliveryReportRequested

    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Permission(self):
        return self.com_object.Permission

    @Permission.setter
    def Permission(self, value):
        self.com_object.Permission = value

    # Lower case aliases for Permission
    @property
    def permission(self):
        return self.Permission

    @permission.setter
    def permission(self, value):
        self.Permission = value

    @property
    def PermissionService(self):
        return self.com_object.PermissionService

    @PermissionService.setter
    def PermissionService(self, value):
        self.com_object.PermissionService = value

    # Lower case aliases for PermissionService
    @property
    def permissionservice(self):
        return self.PermissionService

    @permissionservice.setter
    def permissionservice(self, value):
        self.PermissionService = value

    @property
    def PermissionTemplateGuid(self):
        return MailItem(self.com_object.PermissionTemplateGuid)

    @PermissionTemplateGuid.setter
    def PermissionTemplateGuid(self, value):
        self.com_object.PermissionTemplateGuid = value

    # Lower case aliases for PermissionTemplateGuid
    @property
    def permissiontemplateguid(self):
        return self.PermissionTemplateGuid

    @permissiontemplateguid.setter
    def permissiontemplateguid(self, value):
        self.PermissionTemplateGuid = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReadReceiptRequested(self):
        return self.com_object.ReadReceiptRequested

    # Lower case aliases for ReadReceiptRequested
    @property
    def readreceiptrequested(self):
        return self.ReadReceiptRequested

    @property
    def ReceivedByEntryID(self):
        return self.com_object.ReceivedByEntryID

    # Lower case aliases for ReceivedByEntryID
    @property
    def receivedbyentryid(self):
        return self.ReceivedByEntryID

    @property
    def ReceivedByName(self):
        return self.com_object.ReceivedByName

    # Lower case aliases for ReceivedByName
    @property
    def receivedbyname(self):
        return self.ReceivedByName

    @property
    def ReceivedOnBehalfOfEntryID(self):
        return self.com_object.ReceivedOnBehalfOfEntryID

    # Lower case aliases for ReceivedOnBehalfOfEntryID
    @property
    def receivedonbehalfofentryid(self):
        return self.ReceivedOnBehalfOfEntryID

    @property
    def ReceivedOnBehalfOfName(self):
        return self.com_object.ReceivedOnBehalfOfName

    # Lower case aliases for ReceivedOnBehalfOfName
    @property
    def receivedonbehalfofname(self):
        return self.ReceivedOnBehalfOfName

    @property
    def ReceivedTime(self):
        return self.com_object.ReceivedTime

    # Lower case aliases for ReceivedTime
    @property
    def receivedtime(self):
        return self.ReceivedTime

    @property
    def RecipientReassignmentProhibited(self):
        return self.com_object.RecipientReassignmentProhibited

    @RecipientReassignmentProhibited.setter
    def RecipientReassignmentProhibited(self, value):
        self.com_object.RecipientReassignmentProhibited = value

    # Lower case aliases for RecipientReassignmentProhibited
    @property
    def recipientreassignmentprohibited(self):
        return self.RecipientReassignmentProhibited

    @recipientreassignmentprohibited.setter
    def recipientreassignmentprohibited(self, value):
        self.RecipientReassignmentProhibited = value

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    # Lower case aliases for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    # Lower case aliases for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    # Lower case aliases for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    # Lower case aliases for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    # Lower case aliases for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RemoteStatus(self):
        return OlRemoteStatus(self.com_object.RemoteStatus)

    @RemoteStatus.setter
    def RemoteStatus(self, value):
        self.com_object.RemoteStatus = value

    # Lower case aliases for RemoteStatus
    @property
    def remotestatus(self):
        return self.RemoteStatus

    @remotestatus.setter
    def remotestatus(self, value):
        self.RemoteStatus = value

    @property
    def ReplyRecipientNames(self):
        return self.com_object.ReplyRecipientNames

    # Lower case aliases for ReplyRecipientNames
    @property
    def replyrecipientnames(self):
        return self.ReplyRecipientNames

    @property
    def ReplyRecipients(self):
        return Recipients(self.com_object.ReplyRecipients)

    # Lower case aliases for ReplyRecipients
    @property
    def replyrecipients(self):
        return self.ReplyRecipients

    @property
    def RetentionExpirationDate(self):
        return MailItem(self.com_object.RetentionExpirationDate)

    # Lower case aliases for RetentionExpirationDate
    @property
    def retentionexpirationdate(self):
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.com_object.RetentionPolicyName

    # Lower case aliases for RetentionPolicyName
    @property
    def retentionpolicyname(self):
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return Folder(self.com_object.SaveSentMessageFolder)

    @SaveSentMessageFolder.setter
    def SaveSentMessageFolder(self, value):
        self.com_object.SaveSentMessageFolder = value

    # Lower case aliases for SaveSentMessageFolder
    @property
    def savesentmessagefolder(self):
        return self.SaveSentMessageFolder

    @savesentmessagefolder.setter
    def savesentmessagefolder(self, value):
        self.SaveSentMessageFolder = value

    @property
    def Sender(self):
        return self.com_object.Sender

    @Sender.setter
    def Sender(self, value):
        self.com_object.Sender = value

    # Lower case aliases for Sender
    @property
    def sender(self):
        return self.Sender

    @sender.setter
    def sender(self, value):
        self.Sender = value

    @property
    def SenderEmailAddress(self):
        return self.com_object.SenderEmailAddress

    # Lower case aliases for SenderEmailAddress
    @property
    def senderemailaddress(self):
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.com_object.SenderEmailType

    # Lower case aliases for SenderEmailType
    @property
    def senderemailtype(self):
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.com_object.SenderName

    # Lower case aliases for SenderName
    @property
    def sendername(self):
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    # Lower case aliases for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Sent(self):
        return self.com_object.Sent

    # Lower case aliases for Sent
    @property
    def sent(self):
        return self.Sent

    @property
    def SentOn(self):
        return self.com_object.SentOn

    # Lower case aliases for SentOn
    @property
    def senton(self):
        return self.SentOn

    @property
    def SentOnBehalfOfName(self):
        return self.com_object.SentOnBehalfOfName

    @SentOnBehalfOfName.setter
    def SentOnBehalfOfName(self, value):
        self.com_object.SentOnBehalfOfName = value

    # Lower case aliases for SentOnBehalfOfName
    @property
    def sentonbehalfofname(self):
        return self.SentOnBehalfOfName

    @sentonbehalfofname.setter
    def sentonbehalfofname(self, value):
        self.SentOnBehalfOfName = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Submitted(self):
        return self.com_object.Submitted

    # Lower case aliases for Submitted
    @property
    def submitted(self):
        return self.Submitted

    @property
    def TaskCompletedDate(self):
        return MailItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    # Lower case aliases for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return MailItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    # Lower case aliases for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return MailItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    # Lower case aliases for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return MailItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    # Lower case aliases for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def To(self):
        return self.com_object.To

    @To.setter
    def To(self, value):
        self.com_object.To = value

    # Lower case aliases for To
    @property
    def to(self):
        return self.To

    @to.setter
    def to(self, value):
        self.To = value

    @property
    def ToDoTaskOrdinal(self):
        return MailItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    # Lower case aliases for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    @property
    def VotingOptions(self):
        return self.com_object.VotingOptions

    @VotingOptions.setter
    def VotingOptions(self, value):
        self.com_object.VotingOptions = value

    # Lower case aliases for VotingOptions
    @property
    def votingoptions(self):
        return self.VotingOptions

    @votingoptions.setter
    def votingoptions(self, value):
        self.VotingOptions = value

    @property
    def VotingResponse(self):
        return self.com_object.VotingResponse

    @VotingResponse.setter
    def VotingResponse(self, value):
        self.com_object.VotingResponse = value

    # Lower case aliases for VotingResponse
    @property
    def votingresponse(self):
        return self.VotingResponse

    @votingresponse.setter
    def votingresponse(self, value):
        self.VotingResponse = value

    def AddBusinessCard(self, contact=None):
        arguments = com_arguments([contact])
        self.com_object.AddBusinessCard(*arguments)

    # Lower case alias for AddBusinessCard
    def addbusinesscard(self, contact=None):
        arguments = [contact]
        return self.AddBusinessCard(*arguments)

    def ClearConversationIndex(self):
        self.com_object.ClearConversationIndex()

    # Lower case alias for ClearConversationIndex
    def clearconversationindex(self):
        return self.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.com_object.ClearTaskFlag()

    # Lower case alias for ClearTaskFlag
    def cleartaskflag(self):
        return self.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Forward(self):
        return MailItem(self.com_object.Forward())

    # Lower case alias for Forward
    def forward(self):
        return self.Forward()

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Reply(self):
        return MailItem(self.com_object.Reply())

    # Lower case alias for Reply
    def reply(self):
        return self.Reply()

    def ReplyAll(self):
        return MailItem(self.com_object.ReplyAll())

    # Lower case alias for ReplyAll
    def replyall(self):
        return self.ReplyAll()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def Send(self):
        self.com_object.Send()

    # Lower case alias for Send
    def send(self):
        return self.Send()

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class MailModule:

    def __init__(self, mailmodule=None):
        self.com_object= mailmodule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return MailModule(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    # Lower case aliases for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    # Lower case aliases for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return MailModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return MailModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value


class MarkAsTaskRuleAction:

    def __init__(self, markastaskruleaction=None):
        self.com_object= markastaskruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    # Lower case aliases for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def FlagTo(self):
        return self.com_object.FlagTo

    @FlagTo.setter
    def FlagTo(self, value):
        self.com_object.FlagTo = value

    # Lower case aliases for FlagTo
    @property
    def flagto(self):
        return self.FlagTo

    @flagto.setter
    def flagto(self, value):
        self.FlagTo = value

    @property
    def MarkInterval(self):
        return OlMarkInterval(self.com_object.MarkInterval)

    @MarkInterval.setter
    def MarkInterval(self, value):
        self.com_object.MarkInterval = value

    # Lower case aliases for MarkInterval
    @property
    def markinterval(self):
        return self.MarkInterval

    @markinterval.setter
    def markinterval(self, value):
        self.MarkInterval = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class MeetingItem:

    def __init__(self, meetingitem=None):
        self.com_object= meetingitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.com_object.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.com_object.AutoForwarded = value

    # Lower case aliases for AutoForwarded
    @property
    def autoforwarded(self):
        return self.AutoForwarded

    @autoforwarded.setter
    def autoforwarded(self, value):
        self.AutoForwarded = value

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return self.com_object.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.com_object.DeferredDeliveryTime = value

    # Lower case aliases for DeferredDeliveryTime
    @property
    def deferreddeliverytime(self):
        return self.DeferredDeliveryTime

    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.com_object.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.com_object.DeleteAfterSubmit = value

    # Lower case aliases for DeleteAfterSubmit
    @property
    def deleteaftersubmit(self):
        return self.DeleteAfterSubmit

    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.com_object.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.com_object.ExpiryTime = value

    # Lower case aliases for ExpiryTime
    @property
    def expirytime(self):
        return self.ExpiryTime

    @expirytime.setter
    def expirytime(self, value):
        self.ExpiryTime = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsLatestVersion(self):
        return MeetingItem(self.com_object.IsLatestVersion)

    # Lower case aliases for IsLatestVersion
    @property
    def islatestversion(self):
        return self.IsLatestVersion

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MeetingWorkspaceURL(self):
        return self.com_object.MeetingWorkspaceURL

    # Lower case aliases for MeetingWorkspaceURL
    @property
    def meetingworkspaceurl(self):
        return self.MeetingWorkspaceURL

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return self.com_object.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.com_object.OriginatorDeliveryReportRequested = value

    # Lower case aliases for OriginatorDeliveryReportRequested
    @property
    def originatordeliveryreportrequested(self):
        return self.OriginatorDeliveryReportRequested

    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReceivedTime(self):
        return self.com_object.ReceivedTime

    @ReceivedTime.setter
    def ReceivedTime(self, value):
        self.com_object.ReceivedTime = value

    # Lower case aliases for ReceivedTime
    @property
    def receivedtime(self):
        return self.ReceivedTime

    @receivedtime.setter
    def receivedtime(self, value):
        self.ReceivedTime = value

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    # Lower case aliases for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    # Lower case aliases for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def ReplyRecipients(self):
        return Recipients(self.com_object.ReplyRecipients)

    # Lower case aliases for ReplyRecipients
    @property
    def replyrecipients(self):
        return self.ReplyRecipients

    @property
    def RetentionExpirationDate(self):
        return MeetingItem(self.com_object.RetentionExpirationDate)

    # Lower case aliases for RetentionExpirationDate
    @property
    def retentionexpirationdate(self):
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.com_object.RetentionPolicyName

    # Lower case aliases for RetentionPolicyName
    @property
    def retentionpolicyname(self):
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return self.com_object.SaveSentMessageFolder

    # Lower case aliases for SaveSentMessageFolder
    @property
    def savesentmessagefolder(self):
        return self.SaveSentMessageFolder

    @property
    def SenderEmailAddress(self):
        return self.com_object.SenderEmailAddress

    # Lower case aliases for SenderEmailAddress
    @property
    def senderemailaddress(self):
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.com_object.SenderEmailType

    # Lower case aliases for SenderEmailType
    @property
    def senderemailtype(self):
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.com_object.SenderName

    # Lower case aliases for SenderName
    @property
    def sendername(self):
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    # Lower case aliases for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Sent(self):
        return self.com_object.Sent

    # Lower case aliases for Sent
    @property
    def sent(self):
        return self.Sent

    @property
    def SentOn(self):
        return self.com_object.SentOn

    # Lower case aliases for SentOn
    @property
    def senton(self):
        return self.SentOn

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Submitted(self):
        return self.com_object.Submitted

    # Lower case aliases for Submitted
    @property
    def submitted(self):
        return self.Submitted

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Forward(self):
        return MeetingItem(self.com_object.Forward())

    # Lower case alias for Forward
    def forward(self):
        return self.Forward()

    def GetAssociatedAppointment(self, AddToCalendar=None):
        arguments = com_arguments([AddToCalendar])
        return AppointmentItem(self.com_object.GetAssociatedAppointment(*arguments))

    # Lower case alias for GetAssociatedAppointment
    def getassociatedappointment(self, AddToCalendar=None):
        arguments = [AddToCalendar]
        return self.GetAssociatedAppointment(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Reply(self):
        return MailItem(self.com_object.Reply())

    # Lower case alias for Reply
    def reply(self):
        return self.Reply()

    def ReplyAll(self):
        return MailItem(self.com_object.ReplyAll())

    # Lower case alias for ReplyAll
    def replyall(self):
        return self.ReplyAll()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def Send(self):
        self.com_object.Send()

    # Lower case alias for Send
    def send(self):
        return self.Send()

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class MoveOrCopyRuleAction:

    def __init__(self, moveorcopyruleaction=None):
        self.com_object= moveorcopyruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    # Lower case aliases for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Folder(self):
        return Folder(self.com_object.Folder)

    @Folder.setter
    def Folder(self, value):
        self.com_object.Folder = value

    # Lower case aliases for Folder
    @property
    def folder(self):
        return self.Folder

    @folder.setter
    def folder(self, value):
        self.Folder = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class NameSpace:

    def __init__(self, namespace=None):
        self.com_object= namespace

    @property
    def Accounts(self):
        return Accounts(self.com_object.Accounts)

    # Lower case aliases for Accounts
    @property
    def accounts(self):
        return self.Accounts

    @property
    def AddressLists(self):
        return AddressLists(self.com_object.AddressLists)

    # Lower case aliases for AddressLists
    @property
    def addresslists(self):
        return self.AddressLists

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoDiscoverConnectionMode(self):
        return OlAutoDiscoverConnectionMode(self.com_object.AutoDiscoverConnectionMode)

    # Lower case aliases for AutoDiscoverConnectionMode
    @property
    def autodiscoverconnectionmode(self):
        return self.AutoDiscoverConnectionMode

    @property
    def AutoDiscoverXml(self):
        return self.com_object.AutoDiscoverXml

    # Lower case aliases for AutoDiscoverXml
    @property
    def autodiscoverxml(self):
        return self.AutoDiscoverXml

    @property
    def Categories(self):
        return Categories(self.com_object.Categories)

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentProfileName(self):
        return self.com_object.CurrentProfileName

    # Lower case aliases for CurrentProfileName
    @property
    def currentprofilename(self):
        return self.CurrentProfileName

    @property
    def CurrentUser(self):
        return Recipient(self.com_object.CurrentUser)

    # Lower case aliases for CurrentUser
    @property
    def currentuser(self):
        return self.CurrentUser

    @property
    def DefaultStore(self):
        return Store(self.com_object.DefaultStore)

    # Lower case aliases for DefaultStore
    @property
    def defaultstore(self):
        return self.DefaultStore

    @property
    def ExchangeConnectionMode(self):
        return OlExchangeConnectionMode(self.com_object.ExchangeConnectionMode)

    # Lower case aliases for ExchangeConnectionMode
    @property
    def exchangeconnectionmode(self):
        return self.ExchangeConnectionMode

    @property
    def ExchangeMailboxServerName(self):
        return self.com_object.ExchangeMailboxServerName

    # Lower case aliases for ExchangeMailboxServerName
    @property
    def exchangemailboxservername(self):
        return self.ExchangeMailboxServerName

    @property
    def ExchangeMailboxServerVersion(self):
        return self.com_object.ExchangeMailboxServerVersion

    # Lower case aliases for ExchangeMailboxServerVersion
    @property
    def exchangemailboxserverversion(self):
        return self.ExchangeMailboxServerVersion

    @property
    def Folders(self):
        return Folders(self.com_object.Folders)

    # Lower case aliases for Folders
    @property
    def folders(self):
        return self.Folders

    @property
    def Offline(self):
        return self.com_object.Offline

    # Lower case aliases for Offline
    @property
    def offline(self):
        return self.Offline

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Stores(self):
        return Stores(self.com_object.Stores)

    # Lower case aliases for Stores
    @property
    def stores(self):
        return self.Stores

    @property
    def SyncObjects(self):
        return SyncObjects(self.com_object.SyncObjects)

    # Lower case aliases for SyncObjects
    @property
    def syncobjects(self):
        return self.SyncObjects

    @property
    def Type(self):
        return self.com_object.Type

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    def AddStore(self, Store=None):
        arguments = com_arguments([Store])
        self.com_object.AddStore(*arguments)

    # Lower case alias for AddStore
    def addstore(self, Store=None):
        arguments = [Store]
        return self.AddStore(*arguments)

    def AddStoreEx(self, Store=None, Type=None):
        arguments = com_arguments([Store, Type])
        self.com_object.AddStoreEx(*arguments)

    # Lower case alias for AddStoreEx
    def addstoreex(self, Store=None, Type=None):
        arguments = [Store, Type]
        return self.AddStoreEx(*arguments)

    def CompareEntryIDs(self, FirstEntryID=None, SecondEntryID=None):
        arguments = com_arguments([FirstEntryID, SecondEntryID])
        return self.com_object.CompareEntryIDs(*arguments)

    # Lower case alias for CompareEntryIDs
    def compareentryids(self, FirstEntryID=None, SecondEntryID=None):
        arguments = [FirstEntryID, SecondEntryID]
        return self.CompareEntryIDs(*arguments)

    def CreateContactCard(self, Address=None):
        arguments = com_arguments([Address])
        return Office.ContactCard(self.com_object.CreateContactCard(*arguments))

    # Lower case alias for CreateContactCard
    def createcontactcard(self, Address=None):
        arguments = [Address]
        return self.CreateContactCard(*arguments)

    def CreateRecipient(self, RecipientName=None):
        arguments = com_arguments([RecipientName])
        return Recipient(self.com_object.CreateRecipient(*arguments))

    # Lower case alias for CreateRecipient
    def createrecipient(self, RecipientName=None):
        arguments = [RecipientName]
        return self.CreateRecipient(*arguments)

    def CreateSharingItem(self, Context=None, Provider=None):
        arguments = com_arguments([Context, Provider])
        return SharingItem(self.com_object.CreateSharingItem(*arguments))

    # Lower case alias for CreateSharingItem
    def createsharingitem(self, Context=None, Provider=None):
        arguments = [Context, Provider]
        return self.CreateSharingItem(*arguments)

    def Dial(self, ContactItem=None):
        arguments = com_arguments([ContactItem])
        self.com_object.Dial(*arguments)

    # Lower case alias for Dial
    def dial(self, ContactItem=None):
        arguments = [ContactItem]
        return self.Dial(*arguments)

    def GetAddressEntryFromID(self, ID=None):
        arguments = com_arguments([ID])
        return ID(self.com_object.GetAddressEntryFromID(*arguments))

    # Lower case alias for GetAddressEntryFromID
    def getaddressentryfromid(self, ID=None):
        arguments = [ID]
        return self.GetAddressEntryFromID(*arguments)

    def GetDefaultFolder(self, FolderType=None):
        arguments = com_arguments([FolderType])
        return Folder(self.com_object.GetDefaultFolder(*arguments))

    # Lower case alias for GetDefaultFolder
    def getdefaultfolder(self, FolderType=None):
        arguments = [FolderType]
        return self.GetDefaultFolder(*arguments)

    def GetFolderFromID(self, EntryIDFolder=None, EntryIDStore=None):
        arguments = com_arguments([EntryIDFolder, EntryIDStore])
        return Folder(self.com_object.GetFolderFromID(*arguments))

    # Lower case alias for GetFolderFromID
    def getfolderfromid(self, EntryIDFolder=None, EntryIDStore=None):
        arguments = [EntryIDFolder, EntryIDStore]
        return self.GetFolderFromID(*arguments)

    def GetGlobalAddressList(self):
        return AddressList(self.com_object.GetGlobalAddressList())

    # Lower case alias for GetGlobalAddressList
    def getglobaladdresslist(self):
        return self.GetGlobalAddressList()

    def GetItemFromID(self, EntryIDItem=None, EntryIDStore=None):
        arguments = com_arguments([EntryIDItem, EntryIDStore])
        return Object(self.com_object.GetItemFromID(*arguments))

    # Lower case alias for GetItemFromID
    def getitemfromid(self, EntryIDItem=None, EntryIDStore=None):
        arguments = [EntryIDItem, EntryIDStore]
        return self.GetItemFromID(*arguments)

    def GetRecipientFromID(self, EntryID=None):
        arguments = com_arguments([EntryID])
        return Recipient(self.com_object.GetRecipientFromID(*arguments))

    # Lower case alias for GetRecipientFromID
    def getrecipientfromid(self, EntryID=None):
        arguments = [EntryID]
        return self.GetRecipientFromID(*arguments)

    def GetSelectNamesDialog(self):
        return SelectNamesDialog(self.com_object.GetSelectNamesDialog())

    # Lower case alias for GetSelectNamesDialog
    def getselectnamesdialog(self):
        return self.GetSelectNamesDialog()

    def GetSharedDefaultFolder(self, Recipient=None, FolderType=None):
        arguments = com_arguments([Recipient, FolderType])
        return Folder(self.com_object.GetSharedDefaultFolder(*arguments))

    # Lower case alias for GetSharedDefaultFolder
    def getshareddefaultfolder(self, Recipient=None, FolderType=None):
        arguments = [Recipient, FolderType]
        return self.GetSharedDefaultFolder(*arguments)

    def GetStoreFromID(self, ID=None):
        arguments = com_arguments([ID])
        return StoreID(self.com_object.GetStoreFromID(*arguments))

    # Lower case alias for GetStoreFromID
    def getstorefromid(self, ID=None):
        arguments = [ID]
        return self.GetStoreFromID(*arguments)

    def Logoff(self):
        self.com_object.Logoff()

    # Lower case alias for Logoff
    def logoff(self):
        return self.Logoff()

    def Logon(self, Profile=None, Password=None, ShowDialog=None, NewSession=None):
        arguments = com_arguments([Profile, Password, ShowDialog, NewSession])
        self.com_object.Logon(*arguments)

    # Lower case alias for Logon
    def logon(self, Profile=None, Password=None, ShowDialog=None, NewSession=None):
        arguments = [Profile, Password, ShowDialog, NewSession]
        return self.Logon(*arguments)

    def OpenSharedFolder(self, Path=None, Name=None, DownloadAttachments=None, UseTTL=None):
        arguments = com_arguments([Path, Name, DownloadAttachments, UseTTL])
        return Folder(self.com_object.OpenSharedFolder(*arguments))

    # Lower case alias for OpenSharedFolder
    def opensharedfolder(self, Path=None, Name=None, DownloadAttachments=None, UseTTL=None):
        arguments = [Path, Name, DownloadAttachments, UseTTL]
        return self.OpenSharedFolder(*arguments)

    def OpenSharedItem(self, Path=None):
        arguments = com_arguments([Path])
        return Object(self.com_object.OpenSharedItem(*arguments))

    # Lower case alias for OpenSharedItem
    def openshareditem(self, Path=None):
        arguments = [Path]
        return self.OpenSharedItem(*arguments)

    def PickFolder(self):
        return self.com_object.PickFolder()

    # Lower case alias for PickFolder
    def pickfolder(self):
        return self.PickFolder()

    def RemoveStore(self, Folder=None):
        arguments = com_arguments([Folder])
        self.com_object.RemoveStore(*arguments)

    # Lower case alias for RemoveStore
    def removestore(self, Folder=None):
        arguments = [Folder]
        return self.RemoveStore(*arguments)

    def SendAndReceive(self, showProgressDialog=None):
        arguments = com_arguments([showProgressDialog])
        self.com_object.SendAndReceive(*arguments)

    # Lower case alias for SendAndReceive
    def sendandreceive(self, showProgressDialog=None):
        arguments = [showProgressDialog]
        return self.SendAndReceive(*arguments)


class NavigationFolder:

    def __init__(self, navigationfolder=None):
        self.com_object= navigationfolder

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DisplayName(self):
        return NavigationFolder(self.com_object.DisplayName)

    # Lower case aliases for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def Folder(self):
        return Folder(self.com_object.Folder)

    # Lower case aliases for Folder
    @property
    def folder(self):
        return self.Folder

    @property
    def IsRemovable(self):
        return NavigationFolder(self.com_object.IsRemovable)

    # Lower case aliases for IsRemovable
    @property
    def isremovable(self):
        return self.IsRemovable

    @property
    def IsSelected(self):
        return NavigationFolder(self.com_object.IsSelected)

    @IsSelected.setter
    def IsSelected(self, value):
        self.com_object.IsSelected = value

    # Lower case aliases for IsSelected
    @property
    def isselected(self):
        return self.IsSelected

    @isselected.setter
    def isselected(self, value):
        self.IsSelected = value

    @property
    def IsSideBySide(self):
        return NavigationFolder(self.com_object.IsSideBySide)

    @IsSideBySide.setter
    def IsSideBySide(self, value):
        self.com_object.IsSideBySide = value

    # Lower case aliases for IsSideBySide
    @property
    def issidebyside(self):
        return self.IsSideBySide

    @issidebyside.setter
    def issidebyside(self, value):
        self.IsSideBySide = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return NavigationFolder(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class NavigationFolders:

    def __init__(self, navigationfolders=None):
        self.com_object= navigationfolders

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Folder=None):
        arguments = com_arguments([Folder])
        return NavigationFolder(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Folder=None):
        arguments = [Folder]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return NavigationFolder(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, RemovableFolder=None):
        arguments = com_arguments([RemovableFolder])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, RemovableFolder=None):
        arguments = [RemovableFolder]
        return self.Remove(*arguments)


class NavigationGroup:

    def __init__(self, navigationgroup=None):
        self.com_object= navigationgroup

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def GroupType(self):
        return OlGroupType(self.com_object.GroupType)

    # Lower case aliases for GroupType
    @property
    def grouptype(self):
        return self.GroupType

    @property
    def Name(self):
        return NavigationGroup(self.com_object.Name)

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def NavigationFolders(self):
        return NavigationFolders(self.com_object.NavigationFolders)

    # Lower case aliases for NavigationFolders
    @property
    def navigationfolders(self):
        return self.NavigationFolders

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return NavigationGroup(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class NavigationGroups:

    def __init__(self, navigationgroups=None):
        self.com_object= navigationgroups

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Create(self, GroupDisplayName=None):
        arguments = com_arguments([GroupDisplayName])
        return NavigationGroup(self.com_object.Create(*arguments))

    # Lower case alias for Create
    def create(self, GroupDisplayName=None):
        arguments = [GroupDisplayName]
        return self.Create(*arguments)

    def Delete(self, Group=None):
        arguments = com_arguments([Group])
        self.com_object.Delete(*arguments)

    # Lower case alias for Delete
    def delete(self, Group=None):
        arguments = [Group]
        return self.Delete(*arguments)

    def GetDefaultNavigationGroup(self, DefaultFolderGroup=None):
        arguments = com_arguments([DefaultFolderGroup])
        return NavigationGroup(self.com_object.GetDefaultNavigationGroup(*arguments))

    # Lower case alias for GetDefaultNavigationGroup
    def getdefaultnavigationgroup(self, DefaultFolderGroup=None):
        arguments = [DefaultFolderGroup]
        return self.GetDefaultNavigationGroup(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return NavigationGroup(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class NavigationModule:

    def __init__(self, navigationmodule=None):
        self.com_object= navigationmodule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return NavigationModule(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    # Lower case aliases for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return NavigationModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return NavigationModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value


class NavigationModules:

    def __init__(self, navigationmodules=None):
        self.com_object= navigationmodules

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def GetNavigationModule(self, ModuleType=None):
        arguments = com_arguments([ModuleType])
        return NavigationModule(self.com_object.GetNavigationModule(*arguments))

    # Lower case alias for GetNavigationModule
    def getnavigationmodule(self, ModuleType=None):
        arguments = [ModuleType]
        return self.GetNavigationModule(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return NavigationModule(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class NavigationPane:

    def __init__(self, navigationpane=None):
        self.com_object= navigationpane

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentModule(self):
        return NavigationModule(self.com_object.CurrentModule)

    @CurrentModule.setter
    def CurrentModule(self, value):
        self.com_object.CurrentModule = value

    # Lower case aliases for CurrentModule
    @property
    def currentmodule(self):
        return self.CurrentModule

    @currentmodule.setter
    def currentmodule(self, value):
        self.CurrentModule = value

    @property
    def DisplayedModuleCount(self):
        return NavigationModule(self.com_object.DisplayedModuleCount)

    @DisplayedModuleCount.setter
    def DisplayedModuleCount(self, value):
        self.com_object.DisplayedModuleCount = value

    # Lower case aliases for DisplayedModuleCount
    @property
    def displayedmodulecount(self):
        return self.DisplayedModuleCount

    @displayedmodulecount.setter
    def displayedmodulecount(self, value):
        self.DisplayedModuleCount = value

    @property
    def IsCollapsed(self):
        return self.com_object.IsCollapsed

    @IsCollapsed.setter
    def IsCollapsed(self, value):
        self.com_object.IsCollapsed = value

    # Lower case aliases for IsCollapsed
    @property
    def iscollapsed(self):
        return self.IsCollapsed

    @iscollapsed.setter
    def iscollapsed(self, value):
        self.IsCollapsed = value

    @property
    def Modules(self):
        return NavigationModules(self.com_object.Modules)

    # Lower case aliases for Modules
    @property
    def modules(self):
        return self.Modules

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class NewItemAlertRuleAction:

    def __init__(self, newitemalertruleaction=None):
        self.com_object= newitemalertruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    # Lower case aliases for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    # Lower case aliases for Text
    @property
    def text(self):
        return self.Text

    @text.setter
    def text(self, value):
        self.Text = value


class NoteItem:

    def __init__(self, noteitem=None):
        self.com_object= noteitem

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Height(self):
        return self.com_object.Height

    @Height.setter
    def Height(self, value):
        self.com_object.Height = value

    # Lower case aliases for Height
    @property
    def height(self):
        return self.Height

    @height.setter
    def height(self, value):
        self.Height = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def Left(self):
        return self.com_object.Left

    @Left.setter
    def Left(self, value):
        self.com_object.Left = value

    # Lower case aliases for Left
    @property
    def left(self):
        return self.Left

    @left.setter
    def left(self, value):
        self.Left = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Top(self):
        return self.com_object.Top

    @Top.setter
    def Top(self, value):
        self.com_object.Top = value

    # Lower case aliases for Top
    @property
    def top(self):
        return self.Top

    @top.setter
    def top(self, value):
        self.Top = value

    @property
    def Width(self):
        return self.com_object.Width

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    # Lower case aliases for Width
    @property
    def width(self):
        return self.Width

    @width.setter
    def width(self, value):
        self.Width = value

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        return NoteItem(self.com_object.Copy())

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)


class NotesModule:

    def __init__(self, notesmodule=None):
        self.com_object= notesmodule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return NotesModule(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    # Lower case aliases for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    # Lower case aliases for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return NotesModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return NotesModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

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
        self.com_object= olkbusinesscardcontrol

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkCategory:

    def __init__(self, olkcategory=None):
        self.com_object= olkcategory

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    # Lower case aliases for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    # Lower case aliases for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkCheckBox:

    def __init__(self, olkcheckbox=None):
        self.com_object= olkcheckbox

    @property
    def Accelerator(self):
        return self.com_object.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.com_object.Accelerator = value

    # Lower case aliases for Accelerator
    @property
    def accelerator(self):
        return self.Accelerator

    @accelerator.setter
    def accelerator(self, value):
        self.Accelerator = value

    @property
    def Alignment(self):
        return olAlignment(self.com_object.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.com_object.Alignment = value

    # Lower case aliases for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    # Lower case aliases for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def TripleState(self):
        return self.com_object.TripleState

    @TripleState.setter
    def TripleState(self, value):
        self.com_object.TripleState = value

    # Lower case aliases for TripleState
    @property
    def triplestate(self):
        return self.TripleState

    @triplestate.setter
    def triplestate(self, value):
        self.TripleState = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    # Lower case aliases for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class OlkComboBox:

    def __init__(self, olkcombobox=None):
        self.com_object= olkcombobox

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    # Lower case aliases for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def AutoTab(self):
        return self.com_object.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.com_object.AutoTab = value

    # Lower case aliases for AutoTab
    @property
    def autotab(self):
        return self.AutoTab

    @autotab.setter
    def autotab(self, value):
        self.AutoTab = value

    @property
    def AutoWordSelect(self):
        return self.com_object.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.com_object.AutoWordSelect = value

    # Lower case aliases for AutoWordSelect
    @property
    def autowordselect(self):
        return self.AutoWordSelect

    @autowordselect.setter
    def autowordselect(self, value):
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    # Lower case aliases for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def DragBehavior(self):
        return self.com_object.DragBehavior

    @DragBehavior.setter
    def DragBehavior(self, value):
        self.com_object.DragBehavior = value

    # Lower case aliases for DragBehavior
    @property
    def dragbehavior(self):
        return self.DragBehavior

    @dragbehavior.setter
    def dragbehavior(self, value):
        self.DragBehavior = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.com_object.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.com_object.EnterFieldBehavior = value

    # Lower case aliases for EnterFieldBehavior
    @property
    def enterfieldbehavior(self):
        return self.EnterFieldBehavior

    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.com_object.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.com_object.HideSelection = value

    # Lower case aliases for HideSelection
    @property
    def hideselection(self):
        return self.HideSelection

    @hideselection.setter
    def hideselection(self, value):
        self.HideSelection = value

    @property
    def ListCount(self):
        return self.com_object.ListCount

    # Lower case aliases for ListCount
    @property
    def listcount(self):
        return self.ListCount

    @property
    def ListIndex(self):
        return self.com_object.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.com_object.ListIndex = value

    # Lower case aliases for ListIndex
    @property
    def listindex(self):
        return self.ListIndex

    @listindex.setter
    def listindex(self, value):
        self.ListIndex = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    # Lower case aliases for Locked
    @property
    def locked(self):
        return self.Locked

    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MaxLength(self):
        return self.com_object.MaxLength

    @MaxLength.setter
    def MaxLength(self, value):
        self.com_object.MaxLength = value

    # Lower case aliases for MaxLength
    @property
    def maxlength(self):
        return self.MaxLength

    @maxlength.setter
    def maxlength(self, value):
        self.MaxLength = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def SelectionMargin(self):
        return self.com_object.SelectionMargin

    @SelectionMargin.setter
    def SelectionMargin(self, value):
        self.com_object.SelectionMargin = value

    # Lower case aliases for SelectionMargin
    @property
    def selectionmargin(self):
        return self.SelectionMargin

    @selectionmargin.setter
    def selectionmargin(self, value):
        self.SelectionMargin = value

    @property
    def SelLength(self):
        return self.com_object.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.com_object.SelLength = value

    # Lower case aliases for SelLength
    @property
    def sellength(self):
        return self.SelLength

    @sellength.setter
    def sellength(self, value):
        self.SelLength = value

    @property
    def SelStart(self):
        return self.com_object.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.com_object.SelStart = value

    # Lower case aliases for SelStart
    @property
    def selstart(self):
        return self.SelStart

    @selstart.setter
    def selstart(self, value):
        self.SelStart = value

    @property
    def SelText(self):
        return self.com_object.SelText

    # Lower case aliases for SelText
    @property
    def seltext(self):
        return self.SelText

    @property
    def Style(self):
        return OlComboBoxStyle(self.com_object.Style)

    @Style.setter
    def Style(self, value):
        self.com_object.Style = value

    # Lower case aliases for Style
    @property
    def style(self):
        return self.Style

    @style.setter
    def style(self, value):
        self.Style = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    # Lower case aliases for Text
    @property
    def text(self):
        return self.Text

    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    # Lower case aliases for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def TopIndex(self):
        return self.com_object.TopIndex

    @TopIndex.setter
    def TopIndex(self, value):
        self.com_object.TopIndex = value

    # Lower case aliases for TopIndex
    @property
    def topindex(self):
        return self.TopIndex

    @topindex.setter
    def topindex(self, value):
        self.TopIndex = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    def AddItem(self, ItemText=None, Index=None):
        arguments = com_arguments([ItemText, Index])
        self.com_object.AddItem(*arguments)

    # Lower case alias for AddItem
    def additem(self, ItemText=None, Index=None):
        arguments = [ItemText, Index]
        return self.AddItem(*arguments)

    def Clear(self):
        self.com_object.Clear()

    # Lower case alias for Clear
    def clear(self):
        return self.Clear()

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Cut(self):
        self.com_object.Cut()

    # Lower case alias for Cut
    def cut(self):
        return self.Cut()

    def DropDown(self):
        self.com_object.DropDown()

    # Lower case alias for DropDown
    def dropdown(self):
        return self.DropDown()

    def GetItem(self, Index=None):
        arguments = com_arguments([Index])
        return String(self.com_object.GetItem(*arguments))

    # Lower case alias for GetItem
    def getitem(self, Index=None):
        arguments = [Index]
        return self.GetItem(*arguments)

    def Paste(self):
        self.com_object.Paste()

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()

    def RemoveItem(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.RemoveItem(*arguments)

    # Lower case alias for RemoveItem
    def removeitem(self, Index=None):
        arguments = [Index]
        return self.RemoveItem(*arguments)

    def SetItem(self, Index=None, Item=None):
        arguments = com_arguments([Index, Item])
        self.com_object.SetItem(*arguments)

    # Lower case alias for SetItem
    def setitem(self, Index=None, Item=None):
        arguments = [Index, Item]
        return self.SetItem(*arguments)


class OlkCommandButton:

    def __init__(self, olkcommandbutton=None):
        self.com_object= olkcommandbutton

    @property
    def Accelerator(self):
        return self.com_object.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.com_object.Accelerator = value

    # Lower case aliases for Accelerator
    @property
    def accelerator(self):
        return self.Accelerator

    @accelerator.setter
    def accelerator(self, value):
        self.Accelerator = value

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    # Lower case aliases for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def DisplayDropArrow(self):
        return self.com_object.DisplayDropArrow

    @DisplayDropArrow.setter
    def DisplayDropArrow(self, value):
        self.com_object.DisplayDropArrow = value

    # Lower case aliases for DisplayDropArrow
    @property
    def displaydroparrow(self):
        return self.DisplayDropArrow

    @displaydroparrow.setter
    def displaydroparrow(self, value):
        self.DisplayDropArrow = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def Picture(self):
        return self.com_object.Picture

    @Picture.setter
    def Picture(self, value):
        self.com_object.Picture = value

    # Lower case aliases for Picture
    @property
    def picture(self):
        return self.Picture

    @picture.setter
    def picture(self, value):
        self.Picture = value

    @property
    def PictureAlignment(self):
        return OlPictureAlignment(self.com_object.PictureAlignment)

    @PictureAlignment.setter
    def PictureAlignment(self, value):
        self.com_object.PictureAlignment = value

    # Lower case aliases for PictureAlignment
    @property
    def picturealignment(self):
        return self.PictureAlignment

    @picturealignment.setter
    def picturealignment(self, value):
        self.PictureAlignment = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    # Lower case aliases for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    # Lower case aliases for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class OlkContactPhoto:

    def __init__(self, olkcontactphoto=None):
        self.com_object= olkcontactphoto

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkDateControl:

    def __init__(self, olkdatecontrol=None):
        self.com_object= olkdatecontrol

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    # Lower case aliases for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def AutoWordSelect(self):
        return self.com_object.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.com_object.AutoWordSelect = value

    # Lower case aliases for AutoWordSelect
    @property
    def autowordselect(self):
        return self.AutoWordSelect

    @autowordselect.setter
    def autowordselect(self, value):
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    # Lower case aliases for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Date(self):
        return self.com_object.Date

    @Date.setter
    def Date(self, value):
        self.com_object.Date = value

    # Lower case aliases for Date
    @property
    def date(self):
        return self.Date

    @date.setter
    def date(self, value):
        self.Date = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.com_object.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.com_object.EnterFieldBehavior = value

    # Lower case aliases for EnterFieldBehavior
    @property
    def enterfieldbehavior(self):
        return self.EnterFieldBehavior

    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.com_object.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.com_object.HideSelection = value

    # Lower case aliases for HideSelection
    @property
    def hideselection(self):
        return self.HideSelection

    @hideselection.setter
    def hideselection(self, value):
        self.HideSelection = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    # Lower case aliases for Locked
    @property
    def locked(self):
        return self.Locked

    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def ShowNoneButton(self):
        return self.com_object.ShowNoneButton

    @ShowNoneButton.setter
    def ShowNoneButton(self, value):
        self.com_object.ShowNoneButton = value

    # Lower case aliases for ShowNoneButton
    @property
    def shownonebutton(self):
        return self.ShowNoneButton

    @shownonebutton.setter
    def shownonebutton(self, value):
        self.ShowNoneButton = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    # Lower case aliases for Text
    @property
    def text(self):
        return self.Text

    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    # Lower case aliases for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    def DropDown(self):
        self.com_object.DropDown()

    # Lower case alias for DropDown
    def dropdown(self):
        return self.DropDown()


class OlkFrameHeader:

    def __init__(self, olkframeheader=None):
        self.com_object= olkframeheader

    @property
    def Alignment(self):
        return olAlignment(self.com_object.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.com_object.Alignment = value

    # Lower case aliases for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkInfoBar:

    def __init__(self, olkinfobar=None):
        self.com_object= olkinfobar

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value


class OlkLabel:

    def __init__(self, olklabel=None):
        self.com_object= olklabel

    @property
    def Accelerator(self):
        return self.com_object.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.com_object.Accelerator = value

    # Lower case aliases for Accelerator
    @property
    def accelerator(self):
        return self.Accelerator

    @accelerator.setter
    def accelerator(self, value):
        self.Accelerator = value

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    # Lower case aliases for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    # Lower case aliases for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    # Lower case aliases for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    # Lower case aliases for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def UseHeaderColor(self):
        return self.com_object.UseHeaderColor

    @UseHeaderColor.setter
    def UseHeaderColor(self, value):
        self.com_object.UseHeaderColor = value

    # Lower case aliases for UseHeaderColor
    @property
    def useheadercolor(self):
        return self.UseHeaderColor

    @useheadercolor.setter
    def useheadercolor(self, value):
        self.UseHeaderColor = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    # Lower case aliases for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class OlkListBox:

    def __init__(self, olklistbox=None):
        self.com_object= olklistbox

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    # Lower case aliases for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def ListCount(self):
        return self.com_object.ListCount

    # Lower case aliases for ListCount
    @property
    def listcount(self):
        return self.ListCount

    @property
    def ListIndex(self):
        return self.com_object.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.com_object.ListIndex = value

    # Lower case aliases for ListIndex
    @property
    def listindex(self):
        return self.ListIndex

    @listindex.setter
    def listindex(self, value):
        self.ListIndex = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    # Lower case aliases for Locked
    @property
    def locked(self):
        return self.Locked

    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MatchEntry(self):
        return olMatchEntry(self.com_object.MatchEntry)

    @MatchEntry.setter
    def MatchEntry(self, value):
        self.com_object.MatchEntry = value

    # Lower case aliases for MatchEntry
    @property
    def matchentry(self):
        return self.MatchEntry

    @matchentry.setter
    def matchentry(self, value):
        self.MatchEntry = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def MultiSelect(self):
        return OlMultiSelect(self.com_object.MultiSelect)

    @MultiSelect.setter
    def MultiSelect(self, value):
        self.com_object.MultiSelect = value

    # Lower case aliases for MultiSelect
    @property
    def multiselect(self):
        return self.MultiSelect

    @multiselect.setter
    def multiselect(self, value):
        self.MultiSelect = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    # Lower case aliases for Text
    @property
    def text(self):
        return self.Text

    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    # Lower case aliases for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def TopIndex(self):
        return self.com_object.TopIndex

    @TopIndex.setter
    def TopIndex(self, value):
        self.com_object.TopIndex = value

    # Lower case aliases for TopIndex
    @property
    def topindex(self):
        return self.TopIndex

    @topindex.setter
    def topindex(self, value):
        self.TopIndex = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    def AddItem(self, ItemText=None, Index=None):
        arguments = com_arguments([ItemText, Index])
        self.com_object.AddItem(*arguments)

    # Lower case alias for AddItem
    def additem(self, ItemText=None, Index=None):
        arguments = [ItemText, Index]
        return self.AddItem(*arguments)

    def Clear(self):
        self.com_object.Clear()

    # Lower case alias for Clear
    def clear(self):
        return self.Clear()

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def GetItem(self, Index=None):
        arguments = com_arguments([Index])
        return String(self.com_object.GetItem(*arguments))

    # Lower case alias for GetItem
    def getitem(self, Index=None):
        arguments = [Index]
        return self.GetItem(*arguments)

    def GetSelected(self, Index=None):
        arguments = com_arguments([Index])
        return self.com_object.GetSelected(*arguments)

    # Lower case alias for GetSelected
    def getselected(self, Index=None):
        arguments = [Index]
        return self.GetSelected(*arguments)

    def RemoveItem(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.RemoveItem(*arguments)

    # Lower case alias for RemoveItem
    def removeitem(self, Index=None):
        arguments = [Index]
        return self.RemoveItem(*arguments)

    def SetItem(self, Index=None, Item=None):
        arguments = com_arguments([Index, Item])
        self.com_object.SetItem(*arguments)

    # Lower case alias for SetItem
    def setitem(self, Index=None, Item=None):
        arguments = [Index, Item]
        return self.SetItem(*arguments)

    def SetSelected(self, Index=None, Selected=None):
        arguments = com_arguments([Index, Selected])
        self.com_object.SetSelected(*arguments)

    # Lower case alias for SetSelected
    def setselected(self, Index=None, Selected=None):
        arguments = [Index, Selected]
        return self.SetSelected(*arguments)


class OlkOptionButton:

    def __init__(self, olkoptionbutton=None):
        self.com_object= olkoptionbutton

    @property
    def Accelerator(self):
        return self.com_object.Accelerator

    @Accelerator.setter
    def Accelerator(self, value):
        self.com_object.Accelerator = value

    # Lower case aliases for Accelerator
    @property
    def accelerator(self):
        return self.Accelerator

    @accelerator.setter
    def accelerator(self, value):
        self.Accelerator = value

    @property
    def Alignment(self):
        return olAlignment(self.com_object.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.com_object.Alignment = value

    # Lower case aliases for Alignment
    @property
    def alignment(self):
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        self.Alignment = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    # Lower case aliases for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def GroupName(self):
        return self.com_object.GroupName

    @GroupName.setter
    def GroupName(self, value):
        self.com_object.GroupName = value

    # Lower case aliases for GroupName
    @property
    def groupname(self):
        return self.GroupName

    @groupname.setter
    def groupname(self, value):
        self.GroupName = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    # Lower case aliases for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value


class OlkPageControl:

    def __init__(self, olkpagecontrol=None):
        self.com_object= olkpagecontrol

    @property
    def Page(self):
        return OlPageType(self.com_object.Page)

    @Page.setter
    def Page(self, value):
        self.com_object.Page = value

    # Lower case aliases for Page
    @property
    def page(self):
        return self.Page

    @page.setter
    def page(self, value):
        self.Page = value


class OlkSenderPhoto:

    def __init__(self, olksenderphoto=None):
        self.com_object= olksenderphoto

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def PreferredHeight(self):
        return self.com_object.PreferredHeight

    # Lower case aliases for PreferredHeight
    @property
    def preferredheight(self):
        return self.PreferredHeight

    @property
    def PreferredWidth(self):
        return self.com_object.PreferredWidth

    # Lower case aliases for PreferredWidth
    @property
    def preferredwidth(self):
        return self.PreferredWidth


class OlkTextBox:

    def __init__(self, olktextbox=None):
        self.com_object= olktextbox

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    # Lower case aliases for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def AutoTab(self):
        return self.com_object.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.com_object.AutoTab = value

    # Lower case aliases for AutoTab
    @property
    def autotab(self):
        return self.AutoTab

    @autotab.setter
    def autotab(self, value):
        self.AutoTab = value

    @property
    def AutoWordSelect(self):
        return self.com_object.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.com_object.AutoWordSelect = value

    # Lower case aliases for AutoWordSelect
    @property
    def autowordselect(self):
        return self.AutoWordSelect

    @autowordselect.setter
    def autowordselect(self, value):
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    # Lower case aliases for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def DragBehavior(self):
        return self.com_object.DragBehavior

    @DragBehavior.setter
    def DragBehavior(self, value):
        self.com_object.DragBehavior = value

    # Lower case aliases for DragBehavior
    @property
    def dragbehavior(self):
        return self.DragBehavior

    @dragbehavior.setter
    def dragbehavior(self, value):
        self.DragBehavior = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.com_object.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.com_object.EnterFieldBehavior = value

    # Lower case aliases for EnterFieldBehavior
    @property
    def enterfieldbehavior(self):
        return self.EnterFieldBehavior

    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        self.EnterFieldBehavior = value

    @property
    def EnterKeyBehavior(self):
        return self.com_object.EnterKeyBehavior

    @EnterKeyBehavior.setter
    def EnterKeyBehavior(self, value):
        self.com_object.EnterKeyBehavior = value

    # Lower case aliases for EnterKeyBehavior
    @property
    def enterkeybehavior(self):
        return self.EnterKeyBehavior

    @enterkeybehavior.setter
    def enterkeybehavior(self, value):
        self.EnterKeyBehavior = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.com_object.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.com_object.HideSelection = value

    # Lower case aliases for HideSelection
    @property
    def hideselection(self):
        return self.HideSelection

    @hideselection.setter
    def hideselection(self, value):
        self.HideSelection = value

    @property
    def IntegralHeight(self):
        return self.com_object.IntegralHeight

    @IntegralHeight.setter
    def IntegralHeight(self, value):
        self.com_object.IntegralHeight = value

    # Lower case aliases for IntegralHeight
    @property
    def integralheight(self):
        return self.IntegralHeight

    @integralheight.setter
    def integralheight(self, value):
        self.IntegralHeight = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    # Lower case aliases for Locked
    @property
    def locked(self):
        return self.Locked

    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MaxLength(self):
        return self.com_object.MaxLength

    @MaxLength.setter
    def MaxLength(self, value):
        self.com_object.MaxLength = value

    # Lower case aliases for MaxLength
    @property
    def maxlength(self):
        return self.MaxLength

    @maxlength.setter
    def maxlength(self, value):
        self.MaxLength = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def Multiline(self):
        return self.com_object.Multiline

    @Multiline.setter
    def Multiline(self, value):
        self.com_object.Multiline = value

    # Lower case aliases for Multiline
    @property
    def multiline(self):
        return self.Multiline

    @multiline.setter
    def multiline(self, value):
        self.Multiline = value

    @property
    def PasswordChar(self):
        return self.com_object.PasswordChar

    @PasswordChar.setter
    def PasswordChar(self, value):
        self.com_object.PasswordChar = value

    # Lower case aliases for PasswordChar
    @property
    def passwordchar(self):
        return self.PasswordChar

    @passwordchar.setter
    def passwordchar(self, value):
        self.PasswordChar = value

    @property
    def Scrollbars(self):
        return olScrollBars(self.com_object.Scrollbars)

    @Scrollbars.setter
    def Scrollbars(self, value):
        self.com_object.Scrollbars = value

    # Lower case aliases for Scrollbars
    @property
    def scrollbars(self):
        return self.Scrollbars

    @scrollbars.setter
    def scrollbars(self, value):
        self.Scrollbars = value

    @property
    def SelectionMargin(self):
        return self.com_object.SelectionMargin

    @SelectionMargin.setter
    def SelectionMargin(self, value):
        self.com_object.SelectionMargin = value

    # Lower case aliases for SelectionMargin
    @property
    def selectionmargin(self):
        return self.SelectionMargin

    @selectionmargin.setter
    def selectionmargin(self, value):
        self.SelectionMargin = value

    @property
    def SelLength(self):
        return self.com_object.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.com_object.SelLength = value

    # Lower case aliases for SelLength
    @property
    def sellength(self):
        return self.SelLength

    @sellength.setter
    def sellength(self, value):
        self.SelLength = value

    @property
    def SelStart(self):
        return self.com_object.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.com_object.SelStart = value

    # Lower case aliases for SelStart
    @property
    def selstart(self):
        return self.SelStart

    @selstart.setter
    def selstart(self, value):
        self.SelStart = value

    @property
    def SelText(self):
        return self.com_object.SelText

    # Lower case aliases for SelText
    @property
    def seltext(self):
        return self.SelText

    @property
    def TabKeyBehavior(self):
        return self.com_object.TabKeyBehavior

    @TabKeyBehavior.setter
    def TabKeyBehavior(self, value):
        self.com_object.TabKeyBehavior = value

    # Lower case aliases for TabKeyBehavior
    @property
    def tabkeybehavior(self):
        return self.TabKeyBehavior

    @tabkeybehavior.setter
    def tabkeybehavior(self, value):
        self.TabKeyBehavior = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    # Lower case aliases for Text
    @property
    def text(self):
        return self.Text

    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    # Lower case aliases for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    # Lower case aliases for WordWrap
    @property
    def wordwrap(self):
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        self.WordWrap = value

    def Clear(self):
        self.com_object.Clear()

    # Lower case alias for Clear
    def clear(self):
        return self.Clear()

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Cut(self):
        self.com_object.Cut()

    # Lower case alias for Cut
    def cut(self):
        return self.Cut()

    def Paste(self):
        self.com_object.Paste()

    # Lower case alias for Paste
    def paste(self):
        return self.Paste()


class OlkTimeControl:

    def __init__(self, olktimecontrol=None):
        self.com_object= olktimecontrol

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    # Lower case aliases for AutoSize
    @property
    def autosize(self):
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        self.AutoSize = value

    @property
    def AutoWordSelect(self):
        return self.com_object.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.com_object.AutoWordSelect = value

    # Lower case aliases for AutoWordSelect
    @property
    def autowordselect(self):
        return self.AutoWordSelect

    @autowordselect.setter
    def autowordselect(self, value):
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    # Lower case aliases for BackColor
    @property
    def backcolor(self):
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    # Lower case aliases for BackStyle
    @property
    def backstyle(self):
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        self.BackStyle = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.com_object.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.com_object.EnterFieldBehavior = value

    # Lower case aliases for EnterFieldBehavior
    @property
    def enterfieldbehavior(self):
        return self.EnterFieldBehavior

    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.com_object.Font

    # Lower case aliases for Font
    @property
    def font(self):
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    # Lower case aliases for ForeColor
    @property
    def forecolor(self):
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.com_object.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.com_object.HideSelection = value

    # Lower case aliases for HideSelection
    @property
    def hideselection(self):
        return self.HideSelection

    @hideselection.setter
    def hideselection(self, value):
        self.HideSelection = value

    @property
    def IntervalTime(self):
        return self.com_object.IntervalTime

    @IntervalTime.setter
    def IntervalTime(self, value):
        self.com_object.IntervalTime = value

    # Lower case aliases for IntervalTime
    @property
    def intervaltime(self):
        return self.IntervalTime

    @intervaltime.setter
    def intervaltime(self, value):
        self.IntervalTime = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    # Lower case aliases for Locked
    @property
    def locked(self):
        return self.Locked

    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def ReferenceTime(self):
        return self.com_object.ReferenceTime

    @ReferenceTime.setter
    def ReferenceTime(self, value):
        self.com_object.ReferenceTime = value

    # Lower case aliases for ReferenceTime
    @property
    def referencetime(self):
        return self.ReferenceTime

    @referencetime.setter
    def referencetime(self, value):
        self.ReferenceTime = value

    @property
    def Style(self):
        return OlTimeStyle(self.com_object.Style)

    @Style.setter
    def Style(self, value):
        self.com_object.Style = value

    # Lower case aliases for Style
    @property
    def style(self):
        return self.Style

    @style.setter
    def style(self, value):
        self.Style = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    # Lower case aliases for Text
    @property
    def text(self):
        return self.Text

    @text.setter
    def text(self, value):
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    # Lower case aliases for TextAlign
    @property
    def textalign(self):
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        self.TextAlign = value

    @property
    def Time(self):
        return self.com_object.Time

    @Time.setter
    def Time(self, value):
        self.com_object.Time = value

    # Lower case aliases for Time
    @property
    def time(self):
        return self.Time

    @time.setter
    def time(self, value):
        self.Time = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    def DropDown(self):
        self.com_object.DropDown()

    # Lower case alias for DropDown
    def dropdown(self):
        return self.DropDown()


class OlkTimeZoneControl:

    def __init__(self, olktimezonecontrol=None):
        self.com_object= olktimezonecontrol

    @property
    def AppointmentTimeField(self):
        return OlAppointmentTimeField(self.com_object.AppointmentTimeField)

    @AppointmentTimeField.setter
    def AppointmentTimeField(self, value):
        self.com_object.AppointmentTimeField = value

    # Lower case aliases for AppointmentTimeField
    @property
    def appointmenttimefield(self):
        return self.AppointmentTimeField

    @appointmenttimefield.setter
    def appointmenttimefield(self, value):
        self.AppointmentTimeField = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    # Lower case aliases for BorderStyle
    @property
    def borderstyle(self):
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        self.BorderStyle = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    # Lower case aliases for Locked
    @property
    def locked(self):
        return self.Locked

    @locked.setter
    def locked(self, value):
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    # Lower case aliases for MouseIcon
    @property
    def mouseicon(self):
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    # Lower case aliases for MousePointer
    @property
    def mousepointer(self):
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        self.MousePointer = value

    @property
    def SelectedTimeZoneIndex(self):
        return Application.TimeZones(self.com_object.SelectedTimeZoneIndex)

    @SelectedTimeZoneIndex.setter
    def SelectedTimeZoneIndex(self, value):
        self.com_object.SelectedTimeZoneIndex = value

    # Lower case aliases for SelectedTimeZoneIndex
    @property
    def selectedtimezoneindex(self):
        return self.SelectedTimeZoneIndex

    @selectedtimezoneindex.setter
    def selectedtimezoneindex(self, value):
        self.SelectedTimeZoneIndex = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    def DropDown(self):
        self.com_object.DropDown()

    # Lower case alias for DropDown
    def dropdown(self):
        return self.DropDown()


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
        self.com_object= orderfield

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def IsDescending(self):
        return OrderField(self.com_object.IsDescending)

    @IsDescending.setter
    def IsDescending(self, value):
        self.com_object.IsDescending = value

    # Lower case aliases for IsDescending
    @property
    def isdescending(self):
        return self.IsDescending

    @isdescending.setter
    def isdescending(self, value):
        self.IsDescending = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ViewXMLSchemaName(self):
        return OrderField(self.com_object.ViewXMLSchemaName)

    # Lower case aliases for ViewXMLSchemaName
    @property
    def viewxmlschemaname(self):
        return self.ViewXMLSchemaName


class OrderFields:

    def __init__(self, orderfields=None):
        self.com_object= orderfields

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return OrderField(self.com_object.Count)

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, PropertyName=None, IsDescending=None):
        arguments = com_arguments([PropertyName, IsDescending])
        return OrderField(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, PropertyName=None, IsDescending=None):
        arguments = [PropertyName, IsDescending]
        return self.Add(*arguments)

    def Insert(self, PropertyName=None, Index=None, IsDescending=None):
        arguments = com_arguments([PropertyName, Index, IsDescending])
        return OrderField(self.com_object.Insert(*arguments))

    # Lower case alias for Insert
    def insert(self, PropertyName=None, Index=None, IsDescending=None):
        arguments = [PropertyName, Index, IsDescending]
        return self.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return OrderField(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

    def RemoveAll(self):
        self.com_object.RemoveAll()

    # Lower case alias for RemoveAll
    def removeall(self):
        return self.RemoveAll()


class OutlookBarGroup:

    def __init__(self, outlookbargroup=None):
        self.com_object= outlookbargroup

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Shortcuts(self):
        return OutlookBarShortcuts(self.com_object.Shortcuts)

    # Lower case aliases for Shortcuts
    @property
    def shortcuts(self):
        return self.Shortcuts

    @property
    def ViewType(self):
        return OlOutlookBarViewType(self.com_object.ViewType)

    @ViewType.setter
    def ViewType(self, value):
        self.com_object.ViewType = value

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @viewtype.setter
    def viewtype(self, value):
        self.ViewType = value


class OutlookBarGroups:

    def __init__(self, outlookbargroups=None):
        self.com_object= outlookbargroups

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Index=None):
        arguments = com_arguments([Name, Index])
        return OutlookBarGroup(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Index=None):
        arguments = [Name, Index]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return OutlookBarGroup(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class OutlookBarPane:

    def __init__(self, outlookbarpane=None):
        self.com_object= outlookbarpane

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Contents(self):
        return OutlookBarStorage(self.com_object.Contents)

    # Lower case aliases for Contents
    @property
    def contents(self):
        return self.Contents

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return self.com_object.Visible

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value


class OutlookBarShortcut:

    def __init__(self, outlookbarshortcut=None):
        self.com_object= outlookbarshortcut

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Target(self):
        return self.com_object.Target

    # Lower case aliases for Target
    @property
    def target(self):
        return self.Target

    def SetIcon(self, Icon=None):
        arguments = com_arguments([Icon])
        self.com_object.SetIcon(*arguments)

    # Lower case alias for SetIcon
    def seticon(self, Icon=None):
        arguments = [Icon]
        return self.SetIcon(*arguments)


class OutlookBarShortcuts:

    def __init__(self, outlookbarshortcuts=None):
        self.com_object= outlookbarshortcuts

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Target=None, Name=None, Index=None):
        arguments = com_arguments([Target, Name, Index])
        return OutlookBarShortcut(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Target=None, Name=None, Index=None):
        arguments = [Target, Name, Index]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return OutlookBarShortcut(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class OutlookBarStorage:

    def __init__(self, outlookbarstorage=None):
        self.com_object= outlookbarstorage

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Groups(self):
        return OutlookBarGroups(self.com_object.Groups)

    # Lower case aliases for Groups
    @property
    def groups(self):
        return self.Groups

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class Pages:

    def __init__(self, pages=None):
        self.com_object= pages

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self):
        return Page(self.com_object.Add())

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Object(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class Panes:

    def __init__(self, panes=None):
        self.com_object= panes

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Object(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class PlaySoundRuleAction:

    def __init__(self, playsoundruleaction=None):
        self.com_object= playsoundruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    # Lower case aliases for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def FilePath(self):
        return self.com_object.FilePath

    @FilePath.setter
    def FilePath(self, value):
        self.com_object.FilePath = value

    # Lower case aliases for FilePath
    @property
    def filepath(self):
        return self.FilePath

    @filepath.setter
    def filepath(self, value):
        self.FilePath = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class PostItem:

    def __init__(self, postitem=None):
        self.com_object= postitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.com_object.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.com_object.BodyFormat = value

    # Lower case aliases for BodyFormat
    @property
    def bodyformat(self):
        return self.BodyFormat

    @bodyformat.setter
    def bodyformat(self, value):
        self.BodyFormat = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.com_object.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.com_object.ExpiryTime = value

    # Lower case aliases for ExpiryTime
    @property
    def expirytime(self):
        return self.ExpiryTime

    @expirytime.setter
    def expirytime(self, value):
        self.ExpiryTime = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def HTMLBody(self):
        return self.com_object.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.com_object.HTMLBody = value

    # Lower case aliases for HTMLBody
    @property
    def htmlbody(self):
        return self.HTMLBody

    @htmlbody.setter
    def htmlbody(self, value):
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    # Lower case aliases for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return PostItem(self.com_object.IsMarkedAsTask)

    # Lower case aliases for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReceivedTime(self):
        return self.com_object.ReceivedTime

    # Lower case aliases for ReceivedTime
    @property
    def receivedtime(self):
        return self.ReceivedTime

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    # Lower case aliases for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    # Lower case aliases for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    # Lower case aliases for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    # Lower case aliases for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    # Lower case aliases for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SenderEmailAddress(self):
        return self.com_object.SenderEmailAddress

    # Lower case aliases for SenderEmailAddress
    @property
    def senderemailaddress(self):
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.com_object.SenderEmailType

    # Lower case aliases for SenderEmailType
    @property
    def senderemailtype(self):
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.com_object.SenderName

    # Lower case aliases for SenderName
    @property
    def sendername(self):
        return self.SenderName

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def SentOn(self):
        return self.com_object.SentOn

    # Lower case aliases for SentOn
    @property
    def senton(self):
        return self.SentOn

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def TaskCompletedDate(self):
        return PostItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    # Lower case aliases for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return PostItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    # Lower case aliases for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return PostItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    # Lower case aliases for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return PostItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    # Lower case aliases for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def ToDoTaskOrdinal(self):
        return PostItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    # Lower case aliases for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def ClearConversationIndex(self):
        self.com_object.ClearConversationIndex()

    # Lower case alias for ClearConversationIndex
    def clearconversationindex(self):
        return self.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.com_object.ClearTaskFlag()

    # Lower case alias for ClearTaskFlag
    def cleartaskflag(self):
        return self.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Forward(self):
        return MailItem(self.com_object.Forward())

    # Lower case alias for Forward
    def forward(self):
        return self.Forward()

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def Post(self):
        self.com_object.Post()

    # Lower case alias for Post
    def post(self):
        return self.Post()

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Reply(self):
        return MailItem(self.com_object.Reply())

    # Lower case alias for Reply
    def reply(self):
        return self.Reply()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class PropertyAccessor:

    def __init__(self, propertyaccessor=None):
        self.com_object= propertyaccessor

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return PropertyAccessor(self.com_object.Class)

    @property
    def Parent(self):
        return PropertyAccessor(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def BinaryToString(self, Value=None):
        arguments = com_arguments([Value])
        return String(self.com_object.BinaryToString(*arguments))

    # Lower case alias for BinaryToString
    def binarytostring(self, Value=None):
        arguments = [Value]
        return self.BinaryToString(*arguments)

    def DeleteProperties(self, SchemaNames=None):
        arguments = com_arguments([SchemaNames])
        return self.com_object.DeleteProperties(*arguments)

    # Lower case alias for DeleteProperties
    def deleteproperties(self, SchemaNames=None):
        arguments = [SchemaNames]
        return self.DeleteProperties(*arguments)

    def DeleteProperty(self, SchemaName=None):
        arguments = com_arguments([SchemaName])
        self.com_object.DeleteProperty(*arguments)

    # Lower case alias for DeleteProperty
    def deleteproperty(self, SchemaName=None):
        arguments = [SchemaName]
        return self.DeleteProperty(*arguments)

    def GetProperties(self, SchemaNames=None):
        arguments = com_arguments([SchemaNames])
        return Err(self.com_object.GetProperties(*arguments))

    # Lower case alias for GetProperties
    def getproperties(self, SchemaNames=None):
        arguments = [SchemaNames]
        return self.GetProperties(*arguments)

    def GetProperty(self, SchemaName=None):
        arguments = com_arguments([SchemaName])
        return Variant(self.com_object.GetProperty(*arguments))

    # Lower case alias for GetProperty
    def getproperty(self, SchemaName=None):
        arguments = [SchemaName]
        return self.GetProperty(*arguments)

    def LocalTimeToUTC(self, Value=None):
        arguments = com_arguments([Value])
        return Date(self.com_object.LocalTimeToUTC(*arguments))

    # Lower case alias for LocalTimeToUTC
    def localtimetoutc(self, Value=None):
        arguments = [Value]
        return self.LocalTimeToUTC(*arguments)

    def SetProperties(self, SchemaNames=None, Values=None):
        arguments = com_arguments([SchemaNames, Values])
        return self.com_object.SetProperties(*arguments)

    # Lower case alias for SetProperties
    def setproperties(self, SchemaNames=None, Values=None):
        arguments = [SchemaNames, Values]
        return self.SetProperties(*arguments)

    def SetProperty(self, SchemaName=None, Value=None):
        arguments = com_arguments([SchemaName, Value])
        self.com_object.SetProperty(*arguments)

    # Lower case alias for SetProperty
    def setproperty(self, SchemaName=None, Value=None):
        arguments = [SchemaName, Value]
        return self.SetProperty(*arguments)

    def StringToBinary(self, Value=None):
        arguments = com_arguments([Value])
        return Variant(self.com_object.StringToBinary(*arguments))

    # Lower case alias for StringToBinary
    def stringtobinary(self, Value=None):
        arguments = [Value]
        return self.StringToBinary(*arguments)

    def UTCToLocalTime(self, Value=None):
        arguments = com_arguments([Value])
        return Date(self.com_object.UTCToLocalTime(*arguments))

    # Lower case alias for UTCToLocalTime
    def utctolocaltime(self, Value=None):
        arguments = [Value]
        return self.UTCToLocalTime(*arguments)


class PropertyPage:

    def __init__(self, propertypage=None):
        self.com_object= propertypage

    def Dirty(self, Dirty=None):
        arguments = com_arguments([Dirty])
        if callable(self.com_object.Dirty):
            return self.com_object.Dirty(*arguments)
        else:
            return self.com_object.GetDirty(*arguments)

    # Lower case aliases for Dirty
    def dirty(self, Dirty=None):
        arguments = [Dirty]
        return self.Dirty(*arguments)

    def Apply(self):
        return self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def GetPageInfo(self, HelpFile=None, HelpContext=None):
        arguments = com_arguments([HelpFile, HelpContext])
        return HRESULT(self.com_object.GetPageInfo(*arguments))

    # Lower case alias for GetPageInfo
    def getpageinfo(self, HelpFile=None, HelpContext=None):
        arguments = [HelpFile, HelpContext]
        return self.GetPageInfo(*arguments)


class PropertyPages:

    def __init__(self, propertypages=None):
        self.com_object= propertypages

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Page=None, Title=None):
        arguments = com_arguments([Page, Title])
        self.com_object.Add(*arguments)

    # Lower case alias for Add
    def add(self, Page=None, Title=None):
        arguments = [Page, Title]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Object(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class PropertyPageSite:

    def __init__(self, propertypagesite=None):
        self.com_object= propertypagesite

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def OnStatusChange(self):
        self.com_object.OnStatusChange()

    # Lower case alias for OnStatusChange
    def onstatuschange(self):
        return self.OnStatusChange()


class Recipient:

    def __init__(self, recipient=None):
        self.com_object= recipient

    @property
    def Address(self):
        return self.com_object.Address

    # Lower case aliases for Address
    @property
    def address(self):
        return self.Address

    @property
    def AddressEntry(self):
        return AddressEntry(self.com_object.AddressEntry)

    @AddressEntry.setter
    def AddressEntry(self, value):
        self.com_object.AddressEntry = value

    # Lower case aliases for AddressEntry
    @property
    def addressentry(self):
        return self.AddressEntry

    @addressentry.setter
    def addressentry(self, value):
        self.AddressEntry = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoResponse(self):
        return Recipient(self.com_object.AutoResponse)

    @AutoResponse.setter
    def AutoResponse(self, value):
        self.com_object.AutoResponse = value

    # Lower case aliases for AutoResponse
    @property
    def autoresponse(self):
        return self.AutoResponse

    @autoresponse.setter
    def autoresponse(self, value):
        self.AutoResponse = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DisplayType(self):
        return OlDisplayType(self.com_object.DisplayType)

    # Lower case aliases for DisplayType
    @property
    def displaytype(self):
        return self.DisplayType

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def Index(self):
        return self.com_object.Index

    # Lower case aliases for Index
    @property
    def index(self):
        return self.Index

    @property
    def MeetingResponseStatus(self):
        return OlResponseStatus(self.com_object.MeetingResponseStatus)

    # Lower case aliases for MeetingResponseStatus
    @property
    def meetingresponsestatus(self):
        return self.MeetingResponseStatus

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Resolved(self):
        return self.com_object.Resolved

    # Lower case aliases for Resolved
    @property
    def resolved(self):
        return self.Resolved

    @property
    def Sendable(self):
        return Recipient(self.com_object.Sendable)

    @Sendable.setter
    def Sendable(self, value):
        self.com_object.Sendable = value

    # Lower case aliases for Sendable
    @property
    def sendable(self):
        return self.Sendable

    @sendable.setter
    def sendable(self, value):
        self.Sendable = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def TrackingStatus(self):
        return OlTrackingStatus(self.com_object.TrackingStatus)

    @TrackingStatus.setter
    def TrackingStatus(self, value):
        self.com_object.TrackingStatus = value

    # Lower case aliases for TrackingStatus
    @property
    def trackingstatus(self):
        return self.TrackingStatus

    @trackingstatus.setter
    def trackingstatus(self, value):
        self.TrackingStatus = value

    @property
    def TrackingStatusTime(self):
        return self.com_object.TrackingStatusTime

    @TrackingStatusTime.setter
    def TrackingStatusTime(self, value):
        self.com_object.TrackingStatusTime = value

    # Lower case aliases for TrackingStatusTime
    @property
    def trackingstatustime(self):
        return self.TrackingStatusTime

    @trackingstatustime.setter
    def trackingstatustime(self, value):
        self.TrackingStatusTime = value

    @property
    def Type(self):
        return self.com_object.Type

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @type.setter
    def type(self, value):
        self.Type = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def FreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([Start, MinPerChar, CompleteFormat])
        return String(self.com_object.FreeBusy(*arguments))

    # Lower case alias for FreeBusy
    def freebusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = [Start, MinPerChar, CompleteFormat]
        return self.FreeBusy(*arguments)

    def Resolve(self):
        return self.com_object.Resolve()

    # Lower case alias for Resolve
    def resolve(self):
        return self.Resolve()


class Recipients:

    def __init__(self, recipients=None):
        self.com_object= recipients

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([Name])
        return Recipient(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None):
        arguments = [Name]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Recipient(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

    def ResolveAll(self):
        return self.com_object.ResolveAll()

    # Lower case alias for ResolveAll
    def resolveall(self):
        return self.ResolveAll()


class RecurrencePattern:

    def __init__(self, recurrencepattern=None):
        self.com_object= recurrencepattern

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DayOfMonth(self):
        return self.com_object.DayOfMonth

    @DayOfMonth.setter
    def DayOfMonth(self, value):
        self.com_object.DayOfMonth = value

    # Lower case aliases for DayOfMonth
    @property
    def dayofmonth(self):
        return self.DayOfMonth

    @dayofmonth.setter
    def dayofmonth(self, value):
        self.DayOfMonth = value

    @property
    def DayOfWeekMask(self):
        return OlDaysOfWeek(self.com_object.DayOfWeekMask)

    @DayOfWeekMask.setter
    def DayOfWeekMask(self, value):
        self.com_object.DayOfWeekMask = value

    # Lower case aliases for DayOfWeekMask
    @property
    def dayofweekmask(self):
        return self.DayOfWeekMask

    @dayofweekmask.setter
    def dayofweekmask(self, value):
        self.DayOfWeekMask = value

    @property
    def Duration(self):
        return RecurrencePattern(self.com_object.Duration)

    @Duration.setter
    def Duration(self, value):
        self.com_object.Duration = value

    # Lower case aliases for Duration
    @property
    def duration(self):
        return self.Duration

    @duration.setter
    def duration(self, value):
        self.Duration = value

    @property
    def EndTime(self):
        return self.com_object.EndTime

    @EndTime.setter
    def EndTime(self, value):
        self.com_object.EndTime = value

    # Lower case aliases for EndTime
    @property
    def endtime(self):
        return self.EndTime

    @endtime.setter
    def endtime(self, value):
        self.EndTime = value

    @property
    def Exceptions(self):
        return Exceptions(self.com_object.Exceptions)

    # Lower case aliases for Exceptions
    @property
    def exceptions(self):
        return self.Exceptions

    @property
    def Instance(self):
        return self.com_object.Instance

    @Instance.setter
    def Instance(self, value):
        self.com_object.Instance = value

    # Lower case aliases for Instance
    @property
    def instance(self):
        return self.Instance

    @instance.setter
    def instance(self, value):
        self.Instance = value

    @property
    def Interval(self):
        return self.com_object.Interval

    @Interval.setter
    def Interval(self, value):
        self.com_object.Interval = value

    # Lower case aliases for Interval
    @property
    def interval(self):
        return self.Interval

    @interval.setter
    def interval(self, value):
        self.Interval = value

    @property
    def MonthOfYear(self):
        return self.com_object.MonthOfYear

    @MonthOfYear.setter
    def MonthOfYear(self, value):
        self.com_object.MonthOfYear = value

    # Lower case aliases for MonthOfYear
    @property
    def monthofyear(self):
        return self.MonthOfYear

    @monthofyear.setter
    def monthofyear(self, value):
        self.MonthOfYear = value

    @property
    def NoEndDate(self):
        return self.com_object.NoEndDate

    @NoEndDate.setter
    def NoEndDate(self, value):
        self.com_object.NoEndDate = value

    # Lower case aliases for NoEndDate
    @property
    def noenddate(self):
        return self.NoEndDate

    @noenddate.setter
    def noenddate(self, value):
        self.NoEndDate = value

    @property
    def Occurrences(self):
        return self.com_object.Occurrences

    @Occurrences.setter
    def Occurrences(self, value):
        self.com_object.Occurrences = value

    # Lower case aliases for Occurrences
    @property
    def occurrences(self):
        return self.Occurrences

    @occurrences.setter
    def occurrences(self, value):
        self.Occurrences = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PatternEndDate(self):
        return self.com_object.PatternEndDate

    @PatternEndDate.setter
    def PatternEndDate(self, value):
        self.com_object.PatternEndDate = value

    # Lower case aliases for PatternEndDate
    @property
    def patternenddate(self):
        return self.PatternEndDate

    @patternenddate.setter
    def patternenddate(self, value):
        self.PatternEndDate = value

    @property
    def PatternStartDate(self):
        return self.com_object.PatternStartDate

    @PatternStartDate.setter
    def PatternStartDate(self, value):
        self.com_object.PatternStartDate = value

    # Lower case aliases for PatternStartDate
    @property
    def patternstartdate(self):
        return self.PatternStartDate

    @patternstartdate.setter
    def patternstartdate(self, value):
        self.PatternStartDate = value

    @property
    def RecurrenceType(self):
        return OlRecurrenceType(self.com_object.RecurrenceType)

    @RecurrenceType.setter
    def RecurrenceType(self, value):
        self.com_object.RecurrenceType = value

    # Lower case aliases for RecurrenceType
    @property
    def recurrencetype(self):
        return self.RecurrenceType

    @recurrencetype.setter
    def recurrencetype(self, value):
        self.RecurrenceType = value

    @property
    def Regenerate(self):
        return self.com_object.Regenerate

    @Regenerate.setter
    def Regenerate(self, value):
        self.com_object.Regenerate = value

    # Lower case aliases for Regenerate
    @property
    def regenerate(self):
        return self.Regenerate

    @regenerate.setter
    def regenerate(self, value):
        self.Regenerate = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def StartTime(self):
        return self.com_object.StartTime

    @StartTime.setter
    def StartTime(self, value):
        self.com_object.StartTime = value

    # Lower case aliases for StartTime
    @property
    def starttime(self):
        return self.StartTime

    @starttime.setter
    def starttime(self, value):
        self.StartTime = value

    def GetOccurrence(self, StartDate=None):
        arguments = com_arguments([StartDate])
        return AppointmentItem(self.com_object.GetOccurrence(*arguments))

    # Lower case alias for GetOccurrence
    def getoccurrence(self, StartDate=None):
        arguments = [StartDate]
        return self.GetOccurrence(*arguments)


class Reminder:

    def __init__(self, reminder=None):
        self.com_object= reminder

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Caption(self):
        return self.com_object.Caption

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def IsVisible(self):
        return self.com_object.IsVisible

    # Lower case aliases for IsVisible
    @property
    def isvisible(self):
        return self.IsVisible

    @property
    def Item(self):
        return self.com_object.Item

    # Lower case aliases for Item
    @property
    def item(self):
        return self.Item

    @property
    def NextReminderDate(self):
        return self.com_object.NextReminderDate

    # Lower case aliases for NextReminderDate
    @property
    def nextreminderdate(self):
        return self.NextReminderDate

    @property
    def OriginalReminderDate(self):
        return self.com_object.OriginalReminderDate

    # Lower case aliases for OriginalReminderDate
    @property
    def originalreminderdate(self):
        return self.OriginalReminderDate

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Dismiss(self):
        self.com_object.Dismiss()

    # Lower case alias for Dismiss
    def dismiss(self):
        return self.Dismiss()

    def Snooze(self, SnoozeTime=None):
        arguments = com_arguments([SnoozeTime])
        self.com_object.Snooze(*arguments)

    # Lower case alias for Snooze
    def snooze(self, SnoozeTime=None):
        arguments = [SnoozeTime]
        return self.Snooze(*arguments)


class Reminders:

    def __init__(self, reminders=None):
        self.com_object= reminders

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Reminder(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class RemoteItem:

    def __init__(self, remoteitem=None):
        self.com_object= remoteitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def HasAttachment(self):
        return self.com_object.HasAttachment

    # Lower case aliases for HasAttachment
    @property
    def hasattachment(self):
        return self.HasAttachment

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RemoteMessageClass(self):
        return self.com_object.RemoteMessageClass

    # Lower case aliases for RemoteMessageClass
    @property
    def remotemessageclass(self):
        return self.RemoteMessageClass

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def TransferSize(self):
        return self.com_object.TransferSize

    # Lower case aliases for TransferSize
    @property
    def transfersize(self):
        return self.TransferSize

    @property
    def TransferTime(self):
        return self.com_object.TransferTime

    # Lower case aliases for TransferTime
    @property
    def transfertime(self):
        return self.TransferTime

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class ReportItem:

    def __init__(self, reportitem=None):
        self.com_object= reportitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RetentionExpirationDate(self):
        return ReportItem(self.com_object.RetentionExpirationDate)

    # Lower case aliases for RetentionExpirationDate
    @property
    def retentionexpirationdate(self):
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.com_object.RetentionPolicyName

    # Lower case aliases for RetentionPolicyName
    @property
    def retentionpolicyname(self):
        return self.RetentionPolicyName

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class Results:

    def __init__(self, results=None):
        self.com_object= results

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def DefaultItemType(self):
        return OlItemType(self.com_object.DefaultItemType)

    @DefaultItemType.setter
    def DefaultItemType(self, value):
        self.com_object.DefaultItemType = value

    # Lower case aliases for DefaultItemType
    @property
    def defaultitemtype(self):
        return self.DefaultItemType

    @defaultitemtype.setter
    def defaultitemtype(self, value):
        self.DefaultItemType = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def GetFirst(self):
        return Object(self.com_object.GetFirst())

    # Lower case alias for GetFirst
    def getfirst(self):
        return self.GetFirst()

    def GetLast(self):
        return Object(self.com_object.GetLast())

    # Lower case alias for GetLast
    def getlast(self):
        return self.GetLast()

    def GetNext(self):
        return Object(self.com_object.GetNext())

    # Lower case alias for GetNext
    def getnext(self):
        return self.GetNext()

    def GetPrevious(self):
        return Object(self.com_object.GetPrevious())

    # Lower case alias for GetPrevious
    def getprevious(self):
        return self.GetPrevious()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Object(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def ResetColumns(self):
        self.com_object.ResetColumns()

    # Lower case alias for ResetColumns
    def resetcolumns(self):
        return self.ResetColumns()

    def SetColumns(self, Columns=None):
        arguments = com_arguments([Columns])
        self.com_object.SetColumns(*arguments)

    # Lower case alias for SetColumns
    def setcolumns(self, Columns=None):
        arguments = [Columns]
        return self.SetColumns(*arguments)

    def Sort(self, Property=None, Descending=None):
        arguments = com_arguments([Property, Descending])
        self.com_object.Sort(*arguments)

    # Lower case alias for Sort
    def sort(self, Property=None, Descending=None):
        arguments = [Property, Descending]
        return self.Sort(*arguments)


class Row:

    def __init__(self, row=None):
        self.com_object= row

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Parent(self):
        return Row(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def BinaryToString(self, Index=None):
        arguments = com_arguments([Index])
        return String(self.com_object.BinaryToString(*arguments))

    # Lower case alias for BinaryToString
    def binarytostring(self, Index=None):
        arguments = [Index]
        return self.BinaryToString(*arguments)

    def GetValues(self):
        return Variant(self.com_object.GetValues())

    # Lower case alias for GetValues
    def getvalues(self):
        return self.GetValues()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Variant(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def LocalTimeToUTC(self, Index=None):
        arguments = com_arguments([Index])
        return Date(self.com_object.LocalTimeToUTC(*arguments))

    # Lower case alias for LocalTimeToUTC
    def localtimetoutc(self, Index=None):
        arguments = [Index]
        return self.LocalTimeToUTC(*arguments)

    def UTCToLocalTime(self, Index=None):
        arguments = com_arguments([Index])
        return Date(self.com_object.UTCToLocalTime(*arguments))

    # Lower case alias for UTCToLocalTime
    def utctolocaltime(self, Index=None):
        arguments = [Index]
        return self.UTCToLocalTime(*arguments)


class Rule:

    def __init__(self, rule=None):
        self.com_object= rule

    @property
    def Actions(self):
        return RuleActions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Conditions(self):
        return RuleConditions(self.com_object.Conditions)

    # Lower case aliases for Conditions
    @property
    def conditions(self):
        return self.Conditions

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Exceptions(self):
        return RuleConditions(self.com_object.Exceptions)

    # Lower case aliases for Exceptions
    @property
    def exceptions(self):
        return self.Exceptions

    @property
    def ExecutionOrder(self):
        return Rules(self.com_object.ExecutionOrder)

    @ExecutionOrder.setter
    def ExecutionOrder(self, value):
        self.com_object.ExecutionOrder = value

    # Lower case aliases for ExecutionOrder
    @property
    def executionorder(self):
        return self.ExecutionOrder

    @executionorder.setter
    def executionorder(self, value):
        self.ExecutionOrder = value

    @property
    def IsLocalRule(self):
        return self.com_object.IsLocalRule

    # Lower case aliases for IsLocalRule
    @property
    def islocalrule(self):
        return self.IsLocalRule

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RuleType(self):
        return OlRuleType(self.com_object.RuleType)

    # Lower case aliases for RuleType
    @property
    def ruletype(self):
        return self.RuleType

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Execute(self, ShowProgress=None, Folder=None, IncludeSubfolders=None, RuleExecuteOption=None):
        arguments = com_arguments([ShowProgress, Folder, IncludeSubfolders, RuleExecuteOption])
        self.com_object.Execute(*arguments)

    # Lower case alias for Execute
    def execute(self, ShowProgress=None, Folder=None, IncludeSubfolders=None, RuleExecuteOption=None):
        arguments = [ShowProgress, Folder, IncludeSubfolders, RuleExecuteOption]
        return self.Execute(*arguments)


class RuleAction:

    def __init__(self, ruleaction=None):
        self.com_object= ruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    # Lower case aliases for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Enabled(self):
        return RuleAction(self.com_object.Enabled)

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class RuleActions:

    def __init__(self, ruleactions=None):
        self.com_object= ruleactions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AssignToCategory(self):
        return AssignToCategoryRuleAction(self.com_object.AssignToCategory)

    # Lower case aliases for AssignToCategory
    @property
    def assigntocategory(self):
        return self.AssignToCategory

    @property
    def CC(self):
        return SendRuleAction(self.com_object.CC)

    # Lower case aliases for CC
    @property
    def cc(self):
        return self.CC

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ClearCategories(self):
        return RuleAction(self.com_object.ClearCategories)

    # Lower case aliases for ClearCategories
    @property
    def clearcategories(self):
        return self.ClearCategories

    @property
    def CopyToFolder(self):
        return MoveOrCopyRuleAction(self.com_object.CopyToFolder)

    # Lower case aliases for CopyToFolder
    @property
    def copytofolder(self):
        return self.CopyToFolder

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Delete(self):
        return RuleAction(self.com_object.Delete)

    # Lower case aliases for Delete
    @property
    def delete(self):
        return self.Delete

    @property
    def DeletePermanently(self):
        return RuleAction(self.com_object.DeletePermanently)

    # Lower case aliases for DeletePermanently
    @property
    def deletepermanently(self):
        return self.DeletePermanently

    @property
    def DesktopAlert(self):
        return RuleAction(self.com_object.DesktopAlert)

    # Lower case aliases for DesktopAlert
    @property
    def desktopalert(self):
        return self.DesktopAlert

    @property
    def Forward(self):
        return SendRuleAction(self.com_object.Forward)

    # Lower case aliases for Forward
    @property
    def forward(self):
        return self.Forward

    @property
    def ForwardAsAttachment(self):
        return SendRuleAction(self.com_object.ForwardAsAttachment)

    # Lower case aliases for ForwardAsAttachment
    @property
    def forwardasattachment(self):
        return self.ForwardAsAttachment

    @property
    def MarkAsTask(self):
        return MarkAsTaskRuleAction(self.com_object.MarkAsTask)

    # Lower case aliases for MarkAsTask
    @property
    def markastask(self):
        return self.MarkAsTask

    @property
    def MoveToFolder(self):
        return MoveOrCopyRuleAction(self.com_object.MoveToFolder)

    # Lower case aliases for MoveToFolder
    @property
    def movetofolder(self):
        return self.MoveToFolder

    @property
    def NewItemAlert(self):
        return NewItemAlertRuleAction(self.com_object.NewItemAlert)

    # Lower case aliases for NewItemAlert
    @property
    def newitemalert(self):
        return self.NewItemAlert

    @property
    def NotifyDelivery(self):
        return RuleAction(self.com_object.NotifyDelivery)

    # Lower case aliases for NotifyDelivery
    @property
    def notifydelivery(self):
        return self.NotifyDelivery

    @property
    def NotifyRead(self):
        return RuleAction(self.com_object.NotifyRead)

    # Lower case aliases for NotifyRead
    @property
    def notifyread(self):
        return self.NotifyRead

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PlaySound(self):
        return PlaySoundRuleAction(self.com_object.PlaySound)

    # Lower case aliases for PlaySound
    @property
    def playsound(self):
        return self.PlaySound

    @property
    def Redirect(self):
        return SendRuleAction(self.com_object.Redirect)

    # Lower case aliases for Redirect
    @property
    def redirect(self):
        return self.Redirect

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Stop(self):
        return RuleAction(self.com_object.Stop)

    # Lower case aliases for Stop
    @property
    def stop(self):
        return self.Stop

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return RuleAction(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class RuleCondition:

    def __init__(self, rulecondition=None):
        self.com_object= rulecondition

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return RuleCondition(self.com_object.Enabled)

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class RuleConditions:

    def __init__(self, ruleconditions=None):
        self.com_object= ruleconditions

    @property
    def Account(self):
        return AccountRuleCondition(self.com_object.Account)

    # Lower case aliases for Account
    @property
    def account(self):
        return self.Account

    @property
    def AnyCategory(self):
        return RuleCondition(self.com_object.AnyCategory)

    # Lower case aliases for AnyCategory
    @property
    def anycategory(self):
        return self.AnyCategory

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Body(self):
        return TextRuleCondition(self.com_object.Body)

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @property
    def BodyOrSubject(self):
        return TextRuleCondition(self.com_object.BodyOrSubject)

    # Lower case aliases for BodyOrSubject
    @property
    def bodyorsubject(self):
        return self.BodyOrSubject

    @property
    def Category(self):
        return CategoryRuleCondition(self.com_object.Category)

    # Lower case aliases for Category
    @property
    def category(self):
        return self.Category

    @property
    def CC(self):
        return RuleCondition(self.com_object.CC)

    # Lower case aliases for CC
    @property
    def cc(self):
        return self.CC

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def FormName(self):
        return FormNameRuleCondition(self.com_object.FormName)

    # Lower case aliases for FormName
    @property
    def formname(self):
        return self.FormName

    @property
    def From(self):
        return ToOrFromRuleCondition(self.com_object.From)

    @property
    def FromAnyRSSFeed(self):
        return RuleCondition(self.com_object.FromAnyRSSFeed)

    # Lower case aliases for FromAnyRSSFeed
    @property
    def fromanyrssfeed(self):
        return self.FromAnyRSSFeed

    @property
    def FromRssFeed(self):
        return FromRssFeedRuleCondition(self.com_object.FromRssFeed)

    # Lower case aliases for FromRssFeed
    @property
    def fromrssfeed(self):
        return self.FromRssFeed

    @property
    def HasAttachment(self):
        return RuleCondition(self.com_object.HasAttachment)

    # Lower case aliases for HasAttachment
    @property
    def hasattachment(self):
        return self.HasAttachment

    @property
    def Importance(self):
        return ImportanceRuleCondition(self.com_object.Importance)

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @property
    def MeetingInviteOrUpdate(self):
        return RuleCondition(self.com_object.MeetingInviteOrUpdate)

    # Lower case aliases for MeetingInviteOrUpdate
    @property
    def meetinginviteorupdate(self):
        return self.MeetingInviteOrUpdate

    @property
    def MessageHeader(self):
        return TextRuleCondition(self.com_object.MessageHeader)

    # Lower case aliases for MessageHeader
    @property
    def messageheader(self):
        return self.MessageHeader

    @property
    def NotTo(self):
        return RuleCondition(self.com_object.NotTo)

    # Lower case aliases for NotTo
    @property
    def notto(self):
        return self.NotTo

    @property
    def OnLocalMachine(self):
        return RuleCondition(self.com_object.OnLocalMachine)

    # Lower case aliases for OnLocalMachine
    @property
    def onlocalmachine(self):
        return self.OnLocalMachine

    @property
    def OnlyToMe(self):
        return RuleCondition(self.com_object.OnlyToMe)

    # Lower case aliases for OnlyToMe
    @property
    def onlytome(self):
        return self.OnlyToMe

    @property
    def OnOtherMachine(self):
        return RuleCondition(self.com_object.OnOtherMachine)

    # Lower case aliases for OnOtherMachine
    @property
    def onothermachine(self):
        return self.OnOtherMachine

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RecipientAddress(self):
        return AddressRuleCondition(self.com_object.RecipientAddress)

    # Lower case aliases for RecipientAddress
    @property
    def recipientaddress(self):
        return self.RecipientAddress

    @property
    def SenderAddress(self):
        return AddressRuleCondition(self.com_object.SenderAddress)

    # Lower case aliases for SenderAddress
    @property
    def senderaddress(self):
        return self.SenderAddress

    @property
    def SenderInAddressList(self):
        return SenderInAddressListRuleCondition(self.com_object.SenderInAddressList)

    # Lower case aliases for SenderInAddressList
    @property
    def senderinaddresslist(self):
        return self.SenderInAddressList

    @property
    def SentTo(self):
        return ToOrFromRuleCondition(self.com_object.SentTo)

    # Lower case aliases for SentTo
    @property
    def sentto(self):
        return self.SentTo

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Subject(self):
        return TextRuleCondition(self.com_object.Subject)

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @property
    def ToMe(self):
        return RuleCondition(self.com_object.ToMe)

    # Lower case aliases for ToMe
    @property
    def tome(self):
        return self.ToMe

    @property
    def ToOrCc(self):
        return RuleCondition(self.com_object.ToOrCc)

    # Lower case aliases for ToOrCc
    @property
    def toorcc(self):
        return self.ToOrCc

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return RuleCondition(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Rules:

    def __init__(self, rules=None):
        self.com_object= rules

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def IsRssRulesProcessingEnabled(self):
        return self.com_object.IsRssRulesProcessingEnabled

    @IsRssRulesProcessingEnabled.setter
    def IsRssRulesProcessingEnabled(self, value):
        self.com_object.IsRssRulesProcessingEnabled = value

    # Lower case aliases for IsRssRulesProcessingEnabled
    @property
    def isrssrulesprocessingenabled(self):
        return self.IsRssRulesProcessingEnabled

    @isrssrulesprocessingenabled.setter
    def isrssrulesprocessingenabled(self, value):
        self.IsRssRulesProcessingEnabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Create(self, Name=None, RuleType=None):
        arguments = com_arguments([Name, RuleType])
        return Rule(self.com_object.Create(*arguments))

    # Lower case alias for Create
    def create(self, Name=None, RuleType=None):
        arguments = [Name, RuleType]
        return self.Create(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Rule(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

    def Save(self, ShowProgress=None):
        arguments = com_arguments([ShowProgress])
        self.com_object.Save(*arguments)

    # Lower case alias for Save
    def save(self, ShowProgress=None):
        arguments = [ShowProgress]
        return self.Save(*arguments)


class Search:

    def __init__(self, search=None):
        self.com_object= search

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Filter(self):
        return self.com_object.Filter

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @property
    def IsSynchronous(self):
        return self.com_object.IsSynchronous

    # Lower case aliases for IsSynchronous
    @property
    def issynchronous(self):
        return self.IsSynchronous

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Results(self):
        return Results(self.com_object.Results)

    # Lower case aliases for Results
    @property
    def results(self):
        return self.Results

    @property
    def Scope(self):
        return self.com_object.Scope

    # Lower case aliases for Scope
    @property
    def scope(self):
        return self.Scope

    @property
    def SearchSubFolders(self):
        return self.com_object.SearchSubFolders

    # Lower case aliases for SearchSubFolders
    @property
    def searchsubfolders(self):
        return self.SearchSubFolders

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Tag(self):
        return self.com_object.Tag

    # Lower case aliases for Tag
    @property
    def tag(self):
        return self.Tag

    def GetTable(self):
        return Table(self.com_object.GetTable())

    # Lower case alias for GetTable
    def gettable(self):
        return self.GetTable()

    def Save(self, SchFldrName=None):
        arguments = com_arguments([SchFldrName])
        self.com_object.Save(*arguments)

    # Lower case alias for Save
    def save(self, SchFldrName=None):
        arguments = [SchFldrName]
        return self.Save(*arguments)

    def Stop(self):
        self.com_object.Stop()

    # Lower case alias for Stop
    def stop(self):
        return self.Stop()


class Selection:

    def __init__(self, selection=None):
        self.com_object= selection

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Location(self):
        return OlSelectionLocation(self.com_object.Location)

    # Lower case aliases for Location
    @property
    def location(self):
        return self.Location

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def GetSelection(self, SelectionContents=None):
        arguments = com_arguments([SelectionContents])
        return Selection(self.com_object.GetSelection(*arguments))

    # Lower case alias for GetSelection
    def getselection(self, SelectionContents=None):
        arguments = [SelectionContents]
        return self.GetSelection(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Object(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class SelectNamesDialog:

    def __init__(self, selectnamesdialog=None):
        self.com_object= selectnamesdialog

    @property
    def AllowMultipleSelection(self):
        return self.com_object.AllowMultipleSelection

    @AllowMultipleSelection.setter
    def AllowMultipleSelection(self, value):
        self.com_object.AllowMultipleSelection = value

    # Lower case aliases for AllowMultipleSelection
    @property
    def allowmultipleselection(self):
        return self.AllowMultipleSelection

    @allowmultipleselection.setter
    def allowmultipleselection(self, value):
        self.AllowMultipleSelection = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def BccLabel(self):
        return self.com_object.BccLabel

    @BccLabel.setter
    def BccLabel(self, value):
        self.com_object.BccLabel = value

    # Lower case aliases for BccLabel
    @property
    def bcclabel(self):
        return self.BccLabel

    @bcclabel.setter
    def bcclabel(self, value):
        self.BccLabel = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    # Lower case aliases for Caption
    @property
    def caption(self):
        return self.Caption

    @caption.setter
    def caption(self, value):
        self.Caption = value

    @property
    def CcLabel(self):
        return self.com_object.CcLabel

    @CcLabel.setter
    def CcLabel(self, value):
        self.com_object.CcLabel = value

    # Lower case aliases for CcLabel
    @property
    def cclabel(self):
        return self.CcLabel

    @cclabel.setter
    def cclabel(self, value):
        self.CcLabel = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ForceResolution(self):
        return SelectNamesDialog.Recipients(self.com_object.ForceResolution)

    @ForceResolution.setter
    def ForceResolution(self, value):
        self.com_object.ForceResolution = value

    # Lower case aliases for ForceResolution
    @property
    def forceresolution(self):
        return self.ForceResolution

    @forceresolution.setter
    def forceresolution(self, value):
        self.ForceResolution = value

    @property
    def InitialAddressList(self):
        return AddressList(self.com_object.InitialAddressList)

    @InitialAddressList.setter
    def InitialAddressList(self, value):
        self.com_object.InitialAddressList = value

    # Lower case aliases for InitialAddressList
    @property
    def initialaddresslist(self):
        return self.InitialAddressList

    @initialaddresslist.setter
    def initialaddresslist(self, value):
        self.InitialAddressList = value

    @property
    def NumberOfRecipientSelectors(self):
        return OlRecipientSelectors(self.com_object.NumberOfRecipientSelectors)

    @NumberOfRecipientSelectors.setter
    def NumberOfRecipientSelectors(self, value):
        self.com_object.NumberOfRecipientSelectors = value

    # Lower case aliases for NumberOfRecipientSelectors
    @property
    def numberofrecipientselectors(self):
        return self.NumberOfRecipientSelectors

    @numberofrecipientselectors.setter
    def numberofrecipientselectors(self, value):
        self.NumberOfRecipientSelectors = value

    @property
    def Parent(self):
        return SelectNamesDialog(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @Recipients.setter
    def Recipients(self, value):
        self.com_object.Recipients = value

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @recipients.setter
    def recipients(self, value):
        self.Recipients = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowOnlyInitialAddressList(self):
        return AddressList(self.com_object.ShowOnlyInitialAddressList)

    @ShowOnlyInitialAddressList.setter
    def ShowOnlyInitialAddressList(self, value):
        self.com_object.ShowOnlyInitialAddressList = value

    # Lower case aliases for ShowOnlyInitialAddressList
    @property
    def showonlyinitialaddresslist(self):
        return self.ShowOnlyInitialAddressList

    @showonlyinitialaddresslist.setter
    def showonlyinitialaddresslist(self, value):
        self.ShowOnlyInitialAddressList = value

    @property
    def ToLabel(self):
        return self.com_object.ToLabel

    @ToLabel.setter
    def ToLabel(self, value):
        self.com_object.ToLabel = value

    # Lower case aliases for ToLabel
    @property
    def tolabel(self):
        return self.ToLabel

    @tolabel.setter
    def tolabel(self, value):
        self.ToLabel = value

    def Display(self):
        return self.com_object.Display()

    # Lower case alias for Display
    def display(self):
        return self.Display()

    def SetDefaultDisplayMode(self, defaultMode=None):
        arguments = com_arguments([defaultMode])
        self.com_object.SetDefaultDisplayMode(*arguments)

    # Lower case alias for SetDefaultDisplayMode
    def setdefaultdisplaymode(self, defaultMode=None):
        arguments = [defaultMode]
        return self.SetDefaultDisplayMode(*arguments)


class SenderInAddressListRuleCondition:

    def __init__(self, senderinaddresslistrulecondition=None):
        self.com_object= senderinaddresslistrulecondition

    @property
    def AddressList(self):
        return AddressList(self.com_object.AddressList)

    @AddressList.setter
    def AddressList(self, value):
        self.com_object.AddressList = value

    # Lower case aliases for AddressList
    @property
    def addresslist(self):
        return self.AddressList

    @addresslist.setter
    def addresslist(self, value):
        self.AddressList = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class SendRuleAction:

    def __init__(self, sendruleaction=None):
        self.com_object= sendruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    # Lower case aliases for ActionType
    @property
    def actiontype(self):
        return self.ActionType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class SharingItem:

    def __init__(self, sharingitem=None):
        self.com_object= sharingitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def AllowWriteAccess(self):
        return self.com_object.AllowWriteAccess

    @AllowWriteAccess.setter
    def AllowWriteAccess(self, value):
        self.com_object.AllowWriteAccess = value

    # Lower case aliases for AllowWriteAccess
    @property
    def allowwriteaccess(self):
        return self.AllowWriteAccess

    @allowwriteaccess.setter
    def allowwriteaccess(self, value):
        self.AllowWriteAccess = value

    @property
    def AlternateRecipientAllowed(self):
        return self.com_object.AlternateRecipientAllowed

    @AlternateRecipientAllowed.setter
    def AlternateRecipientAllowed(self, value):
        self.com_object.AlternateRecipientAllowed = value

    # Lower case aliases for AlternateRecipientAllowed
    @property
    def alternaterecipientallowed(self):
        return self.AlternateRecipientAllowed

    @alternaterecipientallowed.setter
    def alternaterecipientallowed(self, value):
        self.AlternateRecipientAllowed = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.com_object.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.com_object.AutoForwarded = value

    # Lower case aliases for AutoForwarded
    @property
    def autoforwarded(self):
        return self.AutoForwarded

    @autoforwarded.setter
    def autoforwarded(self, value):
        self.AutoForwarded = value

    @property
    def BCC(self):
        return SharingItem(self.com_object.BCC)

    @BCC.setter
    def BCC(self, value):
        self.com_object.BCC = value

    # Lower case aliases for BCC
    @property
    def bcc(self):
        return self.BCC

    @bcc.setter
    def bcc(self, value):
        self.BCC = value

    @property
    def BillingInformation(self):
        return SharingItem(self.com_object.BillingInformation)

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return SharingItem(self.com_object.Body)

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.com_object.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.com_object.BodyFormat = value

    # Lower case aliases for BodyFormat
    @property
    def bodyformat(self):
        return self.BodyFormat

    @bodyformat.setter
    def bodyformat(self, value):
        self.BodyFormat = value

    @property
    def Categories(self):
        return SharingItem(self.com_object.Categories)

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def CC(self):
        return SharingItem(self.com_object.CC)

    @CC.setter
    def CC(self, value):
        self.com_object.CC = value

    # Lower case aliases for CC
    @property
    def cc(self):
        return self.CC

    @cc.setter
    def cc(self, value):
        self.CC = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return SharingItem(self.com_object.Companies)

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return SharingItem(self.com_object.ConversationIndex)

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return SharingItem(self.com_object.ConversationTopic)

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return SharingItem(self.com_object.CreationTime)

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return SharingItem(self.com_object.DeferredDeliveryTime)

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.com_object.DeferredDeliveryTime = value

    # Lower case aliases for DeferredDeliveryTime
    @property
    def deferreddeliverytime(self):
        return self.DeferredDeliveryTime

    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.com_object.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.com_object.DeleteAfterSubmit = value

    # Lower case aliases for DeleteAfterSubmit
    @property
    def deleteaftersubmit(self):
        return self.DeleteAfterSubmit

    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return SharingItem(self.com_object.EntryID)

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def ExpiryTime(self):
        return SharingItem(self.com_object.ExpiryTime)

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.com_object.ExpiryTime = value

    # Lower case aliases for ExpiryTime
    @property
    def expirytime(self):
        return self.ExpiryTime

    @expirytime.setter
    def expirytime(self, value):
        self.ExpiryTime = value

    @property
    def FlagRequest(self):
        return SharingItem(self.com_object.FlagRequest)

    @FlagRequest.setter
    def FlagRequest(self, value):
        self.com_object.FlagRequest = value

    # Lower case aliases for FlagRequest
    @property
    def flagrequest(self):
        return self.FlagRequest

    @flagrequest.setter
    def flagrequest(self, value):
        self.FlagRequest = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def HTMLBody(self):
        return SharingItem(self.com_object.HTMLBody)

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.com_object.HTMLBody = value

    # Lower case aliases for HTMLBody
    @property
    def htmlbody(self):
        return self.HTMLBody

    @htmlbody.setter
    def htmlbody(self, value):
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    # Lower case aliases for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return SharingItem(self.com_object.IsConflict)

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return SharingItem(self.com_object.IsMarkedAsTask)

    # Lower case aliases for IsMarkedAsTask
    @property
    def ismarkedastask(self):
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return SharingItem(self.com_object.LastModificationTime)

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return SharingItem(self.com_object.MessageClass)

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return SharingItem(self.com_object.Mileage)

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return SharingItem(self.com_object.NoAging)

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return SharingItem(self.com_object.OriginatorDeliveryReportRequested)

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.com_object.OriginatorDeliveryReportRequested = value

    # Lower case aliases for OriginatorDeliveryReportRequested
    @property
    def originatordeliveryreportrequested(self):
        return self.OriginatorDeliveryReportRequested

    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return SharingItem(self.com_object.OutlookInternalVersion)

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return SharingItem(self.com_object.OutlookVersion)

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return SharingItem(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Permission(self):
        return self.com_object.Permission

    @Permission.setter
    def Permission(self, value):
        self.com_object.Permission = value

    # Lower case aliases for Permission
    @property
    def permission(self):
        return self.Permission

    @permission.setter
    def permission(self, value):
        self.Permission = value

    @property
    def PermissionService(self):
        return self.com_object.PermissionService

    @PermissionService.setter
    def PermissionService(self, value):
        self.com_object.PermissionService = value

    # Lower case aliases for PermissionService
    @property
    def permissionservice(self):
        return self.PermissionService

    @permissionservice.setter
    def permissionservice(self, value):
        self.PermissionService = value

    @property
    def PermissionTemplateGuid(self):
        return SharingItem(self.com_object.PermissionTemplateGuid)

    @PermissionTemplateGuid.setter
    def PermissionTemplateGuid(self, value):
        self.com_object.PermissionTemplateGuid = value

    # Lower case aliases for PermissionTemplateGuid
    @property
    def permissiontemplateguid(self):
        return self.PermissionTemplateGuid

    @permissiontemplateguid.setter
    def permissiontemplateguid(self, value):
        self.PermissionTemplateGuid = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def ReadReceiptRequested(self):
        return self.com_object.ReadReceiptRequested

    # Lower case aliases for ReadReceiptRequested
    @property
    def readreceiptrequested(self):
        return self.ReadReceiptRequested

    @property
    def ReceivedByEntryID(self):
        return self.com_object.ReceivedByEntryID

    # Lower case aliases for ReceivedByEntryID
    @property
    def receivedbyentryid(self):
        return self.ReceivedByEntryID

    @property
    def ReceivedByName(self):
        return SharingItem(self.com_object.ReceivedByName)

    # Lower case aliases for ReceivedByName
    @property
    def receivedbyname(self):
        return self.ReceivedByName

    @property
    def ReceivedOnBehalfOfEntryID(self):
        return self.com_object.ReceivedOnBehalfOfEntryID

    # Lower case aliases for ReceivedOnBehalfOfEntryID
    @property
    def receivedonbehalfofentryid(self):
        return self.ReceivedOnBehalfOfEntryID

    @property
    def ReceivedOnBehalfOfName(self):
        return SharingItem(self.com_object.ReceivedOnBehalfOfName)

    # Lower case aliases for ReceivedOnBehalfOfName
    @property
    def receivedonbehalfofname(self):
        return self.ReceivedOnBehalfOfName

    @property
    def ReceivedTime(self):
        return SharingItem(self.com_object.ReceivedTime)

    # Lower case aliases for ReceivedTime
    @property
    def receivedtime(self):
        return self.ReceivedTime

    @property
    def RecipientReassignmentProhibited(self):
        return SharingItem(self.com_object.RecipientReassignmentProhibited)

    @RecipientReassignmentProhibited.setter
    def RecipientReassignmentProhibited(self, value):
        self.com_object.RecipientReassignmentProhibited = value

    # Lower case aliases for RecipientReassignmentProhibited
    @property
    def recipientreassignmentprohibited(self):
        return self.RecipientReassignmentProhibited

    @recipientreassignmentprohibited.setter
    def recipientreassignmentprohibited(self, value):
        self.RecipientReassignmentProhibited = value

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return SharingItem(self.com_object.ReminderOverrideDefault)

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    # Lower case aliases for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return SharingItem(self.com_object.ReminderPlaySound)

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    # Lower case aliases for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return SharingItem(self.com_object.ReminderSet)

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    # Lower case aliases for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    # Lower case aliases for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return SharingItem(self.com_object.ReminderTime)

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    # Lower case aliases for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def RemoteID(self):
        return SharingItem(self.com_object.RemoteID)

    # Lower case aliases for RemoteID
    @property
    def remoteid(self):
        return self.RemoteID

    @property
    def RemoteName(self):
        return SharingItem(self.com_object.RemoteName)

    # Lower case aliases for RemoteName
    @property
    def remotename(self):
        return self.RemoteName

    @property
    def RemotePath(self):
        return SharingItem(self.com_object.RemotePath)

    # Lower case aliases for RemotePath
    @property
    def remotepath(self):
        return self.RemotePath

    @property
    def RemoteStatus(self):
        return OlRemoteStatus(self.com_object.RemoteStatus)

    @RemoteStatus.setter
    def RemoteStatus(self, value):
        self.com_object.RemoteStatus = value

    # Lower case aliases for RemoteStatus
    @property
    def remotestatus(self):
        return self.RemoteStatus

    @remotestatus.setter
    def remotestatus(self, value):
        self.RemoteStatus = value

    @property
    def ReplyRecipientNames(self):
        return SharingItem(self.com_object.ReplyRecipientNames)

    # Lower case aliases for ReplyRecipientNames
    @property
    def replyrecipientnames(self):
        return self.ReplyRecipientNames

    @property
    def ReplyRecipients(self):
        return Recipients(self.com_object.ReplyRecipients)

    # Lower case aliases for ReplyRecipients
    @property
    def replyrecipients(self):
        return self.ReplyRecipients

    @property
    def RequestedFolder(self):
        return OlDefaultFolders(self.com_object.RequestedFolder)

    # Lower case aliases for RequestedFolder
    @property
    def requestedfolder(self):
        return self.RequestedFolder

    @property
    def RetentionExpirationDate(self):
        return SharingItem(self.com_object.RetentionExpirationDate)

    # Lower case aliases for RetentionExpirationDate
    @property
    def retentionexpirationdate(self):
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.com_object.RetentionPolicyName

    # Lower case aliases for RetentionPolicyName
    @property
    def retentionpolicyname(self):
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return SharingItem(self.com_object.Saved)

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return Folder(self.com_object.SaveSentMessageFolder)

    @SaveSentMessageFolder.setter
    def SaveSentMessageFolder(self, value):
        self.com_object.SaveSentMessageFolder = value

    # Lower case aliases for SaveSentMessageFolder
    @property
    def savesentmessagefolder(self):
        return self.SaveSentMessageFolder

    @savesentmessagefolder.setter
    def savesentmessagefolder(self, value):
        self.SaveSentMessageFolder = value

    @property
    def SenderEmailAddress(self):
        return SharingItem(self.com_object.SenderEmailAddress)

    # Lower case aliases for SenderEmailAddress
    @property
    def senderemailaddress(self):
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return SharingItem(self.com_object.SenderEmailType)

    # Lower case aliases for SenderEmailType
    @property
    def senderemailtype(self):
        return self.SenderEmailType

    @property
    def SenderName(self):
        return SharingItem(self.com_object.SenderName)

    # Lower case aliases for SenderName
    @property
    def sendername(self):
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    # Lower case aliases for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Sent(self):
        return SharingItem(self.com_object.Sent)

    # Lower case aliases for Sent
    @property
    def sent(self):
        return self.Sent

    @property
    def SentOn(self):
        return SharingItem(self.com_object.SentOn)

    # Lower case aliases for SentOn
    @property
    def senton(self):
        return self.SentOn

    @property
    def SentOnBehalfOfName(self):
        return SharingItem(self.com_object.SentOnBehalfOfName)

    @SentOnBehalfOfName.setter
    def SentOnBehalfOfName(self, value):
        self.com_object.SentOnBehalfOfName = value

    # Lower case aliases for SentOnBehalfOfName
    @property
    def sentonbehalfofname(self):
        return self.SentOnBehalfOfName

    @sentonbehalfofname.setter
    def sentonbehalfofname(self, value):
        self.SentOnBehalfOfName = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def SharingProvider(self):
        return OlSharingProvider(self.com_object.SharingProvider)

    # Lower case aliases for SharingProvider
    @property
    def sharingprovider(self):
        return self.SharingProvider

    @property
    def SharingProviderGuid(self):
        return SharingItem(self.com_object.SharingProviderGuid)

    # Lower case aliases for SharingProviderGuid
    @property
    def sharingproviderguid(self):
        return self.SharingProviderGuid

    @property
    def Size(self):
        return SharingItem(self.com_object.Size)

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return SharingItem(self.com_object.Subject)

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def Submitted(self):
        return SharingItem(self.com_object.Submitted)

    # Lower case aliases for Submitted
    @property
    def submitted(self):
        return self.Submitted

    @property
    def TaskCompletedDate(self):
        return SharingItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    # Lower case aliases for TaskCompletedDate
    @property
    def taskcompleteddate(self):
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return SharingItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    # Lower case aliases for TaskDueDate
    @property
    def taskduedate(self):
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return SharingItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    # Lower case aliases for TaskStartDate
    @property
    def taskstartdate(self):
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return SharingItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    # Lower case aliases for TaskSubject
    @property
    def tasksubject(self):
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        self.TaskSubject = value

    @property
    def To(self):
        return SharingItem(self.com_object.To)

    @To.setter
    def To(self, value):
        self.com_object.To = value

    # Lower case aliases for To
    @property
    def to(self):
        return self.To

    @to.setter
    def to(self, value):
        self.To = value

    @property
    def ToDoTaskOrdinal(self):
        return SharingItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    # Lower case aliases for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def Type(self):
        return OlSharingMsgType(self.com_object.Type)

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @type.setter
    def type(self, value):
        self.Type = value

    @property
    def UnRead(self):
        return SharingItem(self.com_object.UnRead)

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def AddBusinessCard(self, contact=None):
        arguments = com_arguments([contact])
        self.com_object.AddBusinessCard(*arguments)

    # Lower case alias for AddBusinessCard
    def addbusinesscard(self, contact=None):
        arguments = [contact]
        return self.AddBusinessCard(*arguments)

    def Allow(self):
        self.com_object.Allow()

    # Lower case alias for Allow
    def allow(self):
        return self.Allow()

    def ClearConversationIndex(self):
        self.com_object.ClearConversationIndex()

    # Lower case alias for ClearConversationIndex
    def clearconversationindex(self):
        return self.ClearConversationIndex()

    def ClearTaskFlag(self):
        self.com_object.ClearTaskFlag()

    # Lower case alias for ClearTaskFlag
    def cleartaskflag(self):
        return self.ClearTaskFlag()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Deny(self):
        return SharingItem(self.com_object.Deny())

    # Lower case alias for Deny
    def deny(self):
        return self.Deny()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Forward(self):
        return SharingItem(self.com_object.Forward())

    # Lower case alias for Forward
    def forward(self):
        return self.Forward()

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([MarkInterval])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def OpenSharedFolder(self):
        return Folder(self.com_object.OpenSharedFolder())

    # Lower case alias for OpenSharedFolder
    def opensharedfolder(self):
        return self.OpenSharedFolder()

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Reply(self):
        return MailItem(self.com_object.Reply())

    # Lower case alias for Reply
    def reply(self):
        return self.Reply()

    def ReplyAll(self):
        return MailItem(self.com_object.ReplyAll())

    # Lower case alias for ReplyAll
    def replyall(self):
        return self.ReplyAll()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def Send(self):
        self.com_object.Send()

    # Lower case alias for Send
    def send(self):
        return self.Send()

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class SimpleItems:

    def __init__(self, simpleitems=None):
        self.com_object= simpleitems

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return SimpleItems(self.com_object.Count)

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return SimpleItems(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Object(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class SolutionsModule:

    def __init__(self, solutionsmodule=None):
        self.com_object= solutionsmodule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return SolutionsModule(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.com_object.NavigationModuleType)

    # Lower case aliases for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return SolutionsModule(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return SolutionsModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return self.com_object.Visible

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value

    def AddSolution(self, Solution=None, Scope=None):
        arguments = com_arguments([Solution, Scope])
        self.com_object.AddSolution(*arguments)

    # Lower case alias for AddSolution
    def addsolution(self, Solution=None, Scope=None):
        arguments = [Solution, Scope]
        return self.AddSolution(*arguments)


class StorageItem:

    def __init__(self, storageitem=None):
        self.com_object= storageitem

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CreationTime(self):
        return StorageItem(self.com_object.CreationTime)

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def Creator(self):
        return StorageItem(self.com_object.Creator)

    @Creator.setter
    def Creator(self, value):
        self.com_object.Creator = value

    # Lower case aliases for Creator
    @property
    def creator(self):
        return self.Creator

    @creator.setter
    def creator(self, value):
        self.Creator = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return StorageItem(self.com_object.Size)

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class Store:

    def __init__(self, store=None):
        self.com_object= store

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Categories(self):
        return Categories(self.com_object.Categories)

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DisplayName(self):
        return Store(self.com_object.DisplayName)

    # Lower case aliases for DisplayName
    @property
    def displayname(self):
        return self.DisplayName

    @property
    def ExchangeStoreType(self):
        return OlExchangeStoreType(self.com_object.ExchangeStoreType)

    # Lower case aliases for ExchangeStoreType
    @property
    def exchangestoretype(self):
        return self.ExchangeStoreType

    @property
    def FilePath(self):
        return self.com_object.FilePath

    # Lower case aliases for FilePath
    @property
    def filepath(self):
        return self.FilePath

    @property
    def IsCachedExchange(self):
        return Store(self.com_object.IsCachedExchange)

    # Lower case aliases for IsCachedExchange
    @property
    def iscachedexchange(self):
        return self.IsCachedExchange

    @property
    def IsConversationEnabled(self):
        return self.com_object.IsConversationEnabled

    # Lower case aliases for IsConversationEnabled
    @property
    def isconversationenabled(self):
        return self.IsConversationEnabled

    @property
    def IsDataFileStore(self):
        return Store(self.com_object.IsDataFileStore)

    # Lower case aliases for IsDataFileStore
    @property
    def isdatafilestore(self):
        return self.IsDataFileStore

    @property
    def IsInstantSearchEnabled(self):
        return self.com_object.IsInstantSearchEnabled

    # Lower case aliases for IsInstantSearchEnabled
    @property
    def isinstantsearchenabled(self):
        return self.IsInstantSearchEnabled

    @property
    def IsOpen(self):
        return Store(self.com_object.IsOpen)

    # Lower case aliases for IsOpen
    @property
    def isopen(self):
        return self.IsOpen

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def StoreID(self):
        return Store(self.com_object.StoreID)

    # Lower case aliases for StoreID
    @property
    def storeid(self):
        return self.StoreID

    def GetDefaultFolder(self, FolderType=None):
        arguments = com_arguments([FolderType])
        return self.com_object.GetDefaultFolder(*arguments)

    # Lower case alias for GetDefaultFolder
    def getdefaultfolder(self, FolderType=None):
        arguments = [FolderType]
        return self.GetDefaultFolder(*arguments)

    def GetRootFolder(self):
        return Folder(self.com_object.GetRootFolder())

    # Lower case alias for GetRootFolder
    def getrootfolder(self):
        return self.GetRootFolder()

    def GetRules(self):
        return Rules(self.com_object.GetRules())

    # Lower case alias for GetRules
    def getrules(self):
        return self.GetRules()

    def GetSearchFolders(self):
        return Folders(self.com_object.GetSearchFolders())

    # Lower case alias for GetSearchFolders
    def getsearchfolders(self):
        return self.GetSearchFolders()

    def GetSpecialFolder(self, FolderType=None):
        arguments = com_arguments([FolderType])
        return Folder(self.com_object.GetSpecialFolder(*arguments))

    # Lower case alias for GetSpecialFolder
    def getspecialfolder(self, FolderType=None):
        arguments = [FolderType]
        return self.GetSpecialFolder(*arguments)

    def RefreshQuotaDisplay(self):
        self.com_object.RefreshQuotaDisplay()

    # Lower case alias for RefreshQuotaDisplay
    def refreshquotadisplay(self):
        return self.RefreshQuotaDisplay()


class Stores:

    def __init__(self, stores=None):
        self.com_object= stores

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Stores(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class SyncObject:

    def __init__(self, syncobject=None):
        self.com_object= syncobject

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Start(self):
        self.com_object.Start()

    # Lower case alias for Start
    def start(self):
        return self.Start()

    def Stop(self):
        self.com_object.Stop()

    # Lower case alias for Stop
    def stop(self):
        return self.Stop()


class SyncObjects:

    def __init__(self, syncobjects=None):
        self.com_object= syncobjects

    @property
    def AppFolders(self):
        return self.com_object.AppFolders

    # Lower case aliases for AppFolders
    @property
    def appfolders(self):
        return self.AppFolders

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return SyncObject(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class Table:

    def __init__(self, table=None):
        self.com_object= table

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Columns(self):
        return Columns(self.com_object.Columns)

    # Lower case aliases for Columns
    @property
    def columns(self):
        return self.Columns

    @property
    def EndOfTable(self):
        return Table(self.com_object.EndOfTable)

    # Lower case aliases for EndOfTable
    @property
    def endoftable(self):
        return self.EndOfTable

    @property
    def Parent(self):
        return Table(self.com_object.Parent)

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def FindNextRow(self):
        return self.com_object.FindNextRow()

    # Lower case alias for FindNextRow
    def findnextrow(self):
        return self.FindNextRow()

    def FindRow(self, Filter=None):
        arguments = com_arguments([Filter])
        return self.com_object.FindRow(*arguments)

    # Lower case alias for FindRow
    def findrow(self, Filter=None):
        arguments = [Filter]
        return self.FindRow(*arguments)

    def GetArray(self, MaxRows=None):
        arguments = com_arguments([MaxRows])
        return Variant(self.com_object.GetArray(*arguments))

    # Lower case alias for GetArray
    def getarray(self, MaxRows=None):
        arguments = [MaxRows]
        return self.GetArray(*arguments)

    def GetNextRow(self):
        return self.com_object.GetNextRow()

    # Lower case alias for GetNextRow
    def getnextrow(self):
        return self.GetNextRow()

    def GetRowCount(self):
        return self.com_object.GetRowCount()

    # Lower case alias for GetRowCount
    def getrowcount(self):
        return self.GetRowCount()

    def MoveToStart(self):
        self.com_object.MoveToStart()

    # Lower case alias for MoveToStart
    def movetostart(self):
        return self.MoveToStart()

    def Restrict(self, Filter=None):
        arguments = com_arguments([Filter])
        return Table(self.com_object.Restrict(*arguments))

    # Lower case alias for Restrict
    def restrict(self, Filter=None):
        arguments = [Filter]
        return self.Restrict(*arguments)

    def Sort(self, SortProperty=None, Descending=None):
        arguments = com_arguments([SortProperty, Descending])
        self.com_object.Sort(*arguments)

    # Lower case alias for Sort
    def sort(self, SortProperty=None, Descending=None):
        arguments = [SortProperty, Descending]
        return self.Sort(*arguments)


class TableView:

    def __init__(self, tableview=None):
        self.com_object= tableview

    @property
    def AllowInCellEditing(self):
        return TableView(self.com_object.AllowInCellEditing)

    @AllowInCellEditing.setter
    def AllowInCellEditing(self, value):
        self.com_object.AllowInCellEditing = value

    # Lower case aliases for AllowInCellEditing
    @property
    def allowincellediting(self):
        return self.AllowInCellEditing

    @allowincellediting.setter
    def allowincellediting(self, value):
        self.AllowInCellEditing = value

    @property
    def AlwaysExpandConversation(self):
        return self.com_object.AlwaysExpandConversation

    @AlwaysExpandConversation.setter
    def AlwaysExpandConversation(self, value):
        self.com_object.AlwaysExpandConversation = value

    # Lower case aliases for AlwaysExpandConversation
    @property
    def alwaysexpandconversation(self):
        return self.AlwaysExpandConversation

    @alwaysexpandconversation.setter
    def alwaysexpandconversation(self, value):
        self.AlwaysExpandConversation = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.com_object.AutoFormatRules)

    # Lower case aliases for AutoFormatRules
    @property
    def autoformatrules(self):
        return self.AutoFormatRules

    @property
    def AutomaticColumnSizing(self):
        return TableView(self.com_object.AutomaticColumnSizing)

    @AutomaticColumnSizing.setter
    def AutomaticColumnSizing(self, value):
        self.com_object.AutomaticColumnSizing = value

    # Lower case aliases for AutomaticColumnSizing
    @property
    def automaticcolumnsizing(self):
        return self.AutomaticColumnSizing

    @automaticcolumnsizing.setter
    def automaticcolumnsizing(self, value):
        self.AutomaticColumnSizing = value

    @property
    def AutomaticGrouping(self):
        return TableView(self.com_object.AutomaticGrouping)

    @AutomaticGrouping.setter
    def AutomaticGrouping(self, value):
        self.com_object.AutomaticGrouping = value

    # Lower case aliases for AutomaticGrouping
    @property
    def automaticgrouping(self):
        return self.AutomaticGrouping

    @automaticgrouping.setter
    def automaticgrouping(self, value):
        self.AutomaticGrouping = value

    @property
    def AutoPreview(self):
        return OlAutoPreview(self.com_object.AutoPreview)

    @AutoPreview.setter
    def AutoPreview(self, value):
        self.com_object.AutoPreview = value

    # Lower case aliases for AutoPreview
    @property
    def autopreview(self):
        return self.AutoPreview

    @autopreview.setter
    def autopreview(self, value):
        self.AutoPreview = value

    @property
    def AutoPreviewFont(self):
        return ViewFont(self.com_object.AutoPreviewFont)

    # Lower case aliases for AutoPreviewFont
    @property
    def autopreviewfont(self):
        return self.AutoPreviewFont

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ColumnFont(self):
        return ViewFont(self.com_object.ColumnFont)

    # Lower case aliases for ColumnFont
    @property
    def columnfont(self):
        return self.ColumnFont

    @property
    def DefaultExpandCollapseSetting(self):
        return OlDefaultExpandCollapseSetting(self.com_object.DefaultExpandCollapseSetting)

    @DefaultExpandCollapseSetting.setter
    def DefaultExpandCollapseSetting(self, value):
        self.com_object.DefaultExpandCollapseSetting = value

    # Lower case aliases for DefaultExpandCollapseSetting
    @property
    def defaultexpandcollapsesetting(self):
        return self.DefaultExpandCollapseSetting

    @defaultexpandcollapsesetting.setter
    def defaultexpandcollapsesetting(self, value):
        self.DefaultExpandCollapseSetting = value

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def GridLineStyle(self):
        return OlGridLineStyle(self.com_object.GridLineStyle)

    @GridLineStyle.setter
    def GridLineStyle(self, value):
        self.com_object.GridLineStyle = value

    # Lower case aliases for GridLineStyle
    @property
    def gridlinestyle(self):
        return self.GridLineStyle

    @gridlinestyle.setter
    def gridlinestyle(self, value):
        self.GridLineStyle = value

    @property
    def GroupByFields(self):
        return OrderFields(self.com_object.GroupByFields)

    # Lower case aliases for GroupByFields
    @property
    def groupbyfields(self):
        return self.GroupByFields

    @property
    def HideReadingPaneHeaderInfo(self):
        return TableView(self.com_object.HideReadingPaneHeaderInfo)

    @HideReadingPaneHeaderInfo.setter
    def HideReadingPaneHeaderInfo(self, value):
        self.com_object.HideReadingPaneHeaderInfo = value

    # Lower case aliases for HideReadingPaneHeaderInfo
    @property
    def hidereadingpaneheaderinfo(self):
        return self.HideReadingPaneHeaderInfo

    @hidereadingpaneheaderinfo.setter
    def hidereadingpaneheaderinfo(self, value):
        self.HideReadingPaneHeaderInfo = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    # Lower case aliases for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def MaxLinesInMultiLineView(self):
        return TableView(self.com_object.MaxLinesInMultiLineView)

    @MaxLinesInMultiLineView.setter
    def MaxLinesInMultiLineView(self, value):
        self.com_object.MaxLinesInMultiLineView = value

    # Lower case aliases for MaxLinesInMultiLineView
    @property
    def maxlinesinmultilineview(self):
        return self.MaxLinesInMultiLineView

    @maxlinesinmultilineview.setter
    def maxlinesinmultilineview(self, value):
        self.MaxLinesInMultiLineView = value

    @property
    def Multiline(self):
        return OlMultiLine(self.com_object.Multiline)

    @Multiline.setter
    def Multiline(self, value):
        self.com_object.Multiline = value

    # Lower case aliases for Multiline
    @property
    def multiline(self):
        return self.Multiline

    @multiline.setter
    def multiline(self, value):
        self.Multiline = value

    @property
    def MultiLineWidth(self):
        return TableView(self.com_object.MultiLineWidth)

    @MultiLineWidth.setter
    def MultiLineWidth(self, value):
        self.com_object.MultiLineWidth = value

    # Lower case aliases for MultiLineWidth
    @property
    def multilinewidth(self):
        return self.MultiLineWidth

    @multilinewidth.setter
    def multilinewidth(self, value):
        self.MultiLineWidth = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def RowFont(self):
        return ViewFont(self.com_object.RowFont)

    # Lower case aliases for RowFont
    @property
    def rowfont(self):
        return self.RowFont

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    # Lower case aliases for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowConversationByDate(self):
        return self.com_object.ShowConversationByDate

    @ShowConversationByDate.setter
    def ShowConversationByDate(self, value):
        self.com_object.ShowConversationByDate = value

    # Lower case aliases for ShowConversationByDate
    @property
    def showconversationbydate(self):
        return self.ShowConversationByDate

    @showconversationbydate.setter
    def showconversationbydate(self, value):
        self.ShowConversationByDate = value

    @property
    def ShowConversationSendersAboveSubject(self):
        return self.com_object.ShowConversationSendersAboveSubject

    @ShowConversationSendersAboveSubject.setter
    def ShowConversationSendersAboveSubject(self, value):
        self.com_object.ShowConversationSendersAboveSubject = value

    # Lower case aliases for ShowConversationSendersAboveSubject
    @property
    def showconversationsendersabovesubject(self):
        return self.ShowConversationSendersAboveSubject

    @showconversationsendersabovesubject.setter
    def showconversationsendersabovesubject(self, value):
        self.ShowConversationSendersAboveSubject = value

    @property
    def ShowFullConversations(self):
        return self.com_object.ShowFullConversations

    @ShowFullConversations.setter
    def ShowFullConversations(self, value):
        self.com_object.ShowFullConversations = value

    # Lower case aliases for ShowFullConversations
    @property
    def showfullconversations(self):
        return self.ShowFullConversations

    @showfullconversations.setter
    def showfullconversations(self, value):
        self.ShowFullConversations = value

    @property
    def ShowItemsInGroups(self):
        return TableView(self.com_object.ShowItemsInGroups)

    @ShowItemsInGroups.setter
    def ShowItemsInGroups(self, value):
        self.com_object.ShowItemsInGroups = value

    # Lower case aliases for ShowItemsInGroups
    @property
    def showitemsingroups(self):
        return self.ShowItemsInGroups

    @showitemsingroups.setter
    def showitemsingroups(self, value):
        self.ShowItemsInGroups = value

    @property
    def ShowNewItemRow(self):
        return TableView(self.com_object.ShowNewItemRow)

    @ShowNewItemRow.setter
    def ShowNewItemRow(self, value):
        self.com_object.ShowNewItemRow = value

    # Lower case aliases for ShowNewItemRow
    @property
    def shownewitemrow(self):
        return self.ShowNewItemRow

    @shownewitemrow.setter
    def shownewitemrow(self, value):
        self.ShowNewItemRow = value

    @property
    def ShowReadingPane(self):
        return TableView(self.com_object.ShowReadingPane)

    @ShowReadingPane.setter
    def ShowReadingPane(self, value):
        self.com_object.ShowReadingPane = value

    # Lower case aliases for ShowReadingPane
    @property
    def showreadingpane(self):
        return self.ShowReadingPane

    @showreadingpane.setter
    def showreadingpane(self, value):
        self.ShowReadingPane = value

    @property
    def SortFields(self):
        return OrderFields(self.com_object.SortFields)

    # Lower case aliases for SortFields
    @property
    def sortfields(self):
        return self.SortFields

    @property
    def Standard(self):
        return TableView(self.com_object.Standard)

    # Lower case aliases for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewFields(self):
        return ViewFields(self.com_object.ViewFields)

    # Lower case aliases for ViewFields
    @property
    def viewfields(self):
        return self.ViewFields

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    # Lower case aliases for XML
    @property
    def xml(self):
        return self.XML

    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return View(self.com_object.Copy(*arguments))

    # Lower case alias for Copy
    def copy(self, Name=None, SaveOption=None):
        arguments = [Name, SaveOption]
        return self.Copy(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GetTable(self):
        return Table(self.com_object.GetTable())

    # Lower case alias for GetTable
    def gettable(self):
        return self.GetTable()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.com_object.GoToDate(*arguments)

    # Lower case alias for GoToDate
    def gotodate(self, Date=None):
        arguments = [Date]
        return self.GoToDate(*arguments)

    def Reset(self):
        self.com_object.Reset()

    # Lower case alias for Reset
    def reset(self):
        return self.Reset()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class TaskItem:

    def __init__(self, taskitem=None):
        self.com_object= taskitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def ActualWork(self):
        return self.com_object.ActualWork

    @ActualWork.setter
    def ActualWork(self, value):
        self.com_object.ActualWork = value

    # Lower case aliases for ActualWork
    @property
    def actualwork(self):
        return self.ActualWork

    @actualwork.setter
    def actualwork(self, value):
        self.ActualWork = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def CardData(self):
        return self.com_object.CardData

    @CardData.setter
    def CardData(self, value):
        self.com_object.CardData = value

    # Lower case aliases for CardData
    @property
    def carddata(self):
        return self.CardData

    @carddata.setter
    def carddata(self, value):
        self.CardData = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Complete(self):
        return self.com_object.Complete

    @Complete.setter
    def Complete(self, value):
        self.com_object.Complete = value

    # Lower case aliases for Complete
    @property
    def complete(self):
        return self.Complete

    @complete.setter
    def complete(self, value):
        self.Complete = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ContactNames(self):
        return self.com_object.ContactNames

    @ContactNames.setter
    def ContactNames(self, value):
        self.com_object.ContactNames = value

    # Lower case aliases for ContactNames
    @property
    def contactnames(self):
        return self.ContactNames

    @contactnames.setter
    def contactnames(self, value):
        self.ContactNames = value

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DateCompleted(self):
        return self.com_object.DateCompleted

    @DateCompleted.setter
    def DateCompleted(self, value):
        self.com_object.DateCompleted = value

    # Lower case aliases for DateCompleted
    @property
    def datecompleted(self):
        return self.DateCompleted

    @datecompleted.setter
    def datecompleted(self, value):
        self.DateCompleted = value

    @property
    def DelegationState(self):
        return OlTaskDelegationState(self.com_object.DelegationState)

    # Lower case aliases for DelegationState
    @property
    def delegationstate(self):
        return self.DelegationState

    @property
    def Delegator(self):
        return self.com_object.Delegator

    # Lower case aliases for Delegator
    @property
    def delegator(self):
        return self.Delegator

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def DueDate(self):
        return self.com_object.DueDate

    @DueDate.setter
    def DueDate(self, value):
        self.com_object.DueDate = value

    # Lower case aliases for DueDate
    @property
    def duedate(self):
        return self.DueDate

    @duedate.setter
    def duedate(self, value):
        self.DueDate = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    # Lower case aliases for InternetCodepage
    @property
    def internetcodepage(self):
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def IsRecurring(self):
        return self.com_object.IsRecurring

    # Lower case aliases for IsRecurring
    @property
    def isrecurring(self):
        return self.IsRecurring

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def Ordinal(self):
        return self.com_object.Ordinal

    @Ordinal.setter
    def Ordinal(self, value):
        self.com_object.Ordinal = value

    # Lower case aliases for Ordinal
    @property
    def ordinal(self):
        return self.Ordinal

    @ordinal.setter
    def ordinal(self, value):
        self.Ordinal = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Owner(self):
        return self.com_object.Owner

    @Owner.setter
    def Owner(self, value):
        self.com_object.Owner = value

    # Lower case aliases for Owner
    @property
    def owner(self):
        return self.Owner

    @owner.setter
    def owner(self, value):
        self.Owner = value

    @property
    def Ownership(self):
        return OlTaskOwnership(self.com_object.Ownership)

    # Lower case aliases for Ownership
    @property
    def ownership(self):
        return self.Ownership

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PercentComplete(self):
        return self.com_object.PercentComplete

    @PercentComplete.setter
    def PercentComplete(self, value):
        self.com_object.PercentComplete = value

    # Lower case aliases for PercentComplete
    @property
    def percentcomplete(self):
        return self.PercentComplete

    @percentcomplete.setter
    def percentcomplete(self, value):
        self.PercentComplete = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    # Lower case aliases for ReminderOverrideDefault
    @property
    def reminderoverridedefault(self):
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    # Lower case aliases for ReminderPlaySound
    @property
    def reminderplaysound(self):
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    # Lower case aliases for ReminderSet
    @property
    def reminderset(self):
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    # Lower case aliases for ReminderSoundFile
    @property
    def remindersoundfile(self):
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    # Lower case aliases for ReminderTime
    @property
    def remindertime(self):
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        self.ReminderTime = value

    @property
    def ResponseState(self):
        return OlTaskResponse(self.com_object.ResponseState)

    # Lower case aliases for ResponseState
    @property
    def responsestate(self):
        return self.ResponseState

    @property
    def Role(self):
        return self.com_object.Role

    @Role.setter
    def Role(self, value):
        self.com_object.Role = value

    # Lower case aliases for Role
    @property
    def role(self):
        return self.Role

    @role.setter
    def role(self, value):
        self.Role = value

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def SchedulePlusPriority(self):
        return self.com_object.SchedulePlusPriority

    @SchedulePlusPriority.setter
    def SchedulePlusPriority(self, value):
        self.com_object.SchedulePlusPriority = value

    # Lower case aliases for SchedulePlusPriority
    @property
    def schedulepluspriority(self):
        return self.SchedulePlusPriority

    @schedulepluspriority.setter
    def schedulepluspriority(self, value):
        self.SchedulePlusPriority = value

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    # Lower case aliases for SendUsingAccount
    @property
    def sendusingaccount(self):
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def StartDate(self):
        return self.com_object.StartDate

    @StartDate.setter
    def StartDate(self, value):
        self.com_object.StartDate = value

    # Lower case aliases for StartDate
    @property
    def startdate(self):
        return self.StartDate

    @startdate.setter
    def startdate(self, value):
        self.StartDate = value

    @property
    def Status(self):
        return OlTaskStatus(self.com_object.Status)

    @Status.setter
    def Status(self, value):
        self.com_object.Status = value

    # Lower case aliases for Status
    @property
    def status(self):
        return self.Status

    @status.setter
    def status(self, value):
        self.Status = value

    @property
    def StatusOnCompletionRecipients(self):
        return self.com_object.StatusOnCompletionRecipients

    @StatusOnCompletionRecipients.setter
    def StatusOnCompletionRecipients(self, value):
        self.com_object.StatusOnCompletionRecipients = value

    # Lower case aliases for StatusOnCompletionRecipients
    @property
    def statusoncompletionrecipients(self):
        return self.StatusOnCompletionRecipients

    @statusoncompletionrecipients.setter
    def statusoncompletionrecipients(self, value):
        self.StatusOnCompletionRecipients = value

    @property
    def StatusUpdateRecipients(self):
        return self.com_object.StatusUpdateRecipients

    @StatusUpdateRecipients.setter
    def StatusUpdateRecipients(self, value):
        self.com_object.StatusUpdateRecipients = value

    # Lower case aliases for StatusUpdateRecipients
    @property
    def statusupdaterecipients(self):
        return self.StatusUpdateRecipients

    @statusupdaterecipients.setter
    def statusupdaterecipients(self, value):
        self.StatusUpdateRecipients = value

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def TeamTask(self):
        return self.com_object.TeamTask

    @TeamTask.setter
    def TeamTask(self, value):
        self.com_object.TeamTask = value

    # Lower case aliases for TeamTask
    @property
    def teamtask(self):
        return self.TeamTask

    @teamtask.setter
    def teamtask(self, value):
        self.TeamTask = value

    @property
    def ToDoTaskOrdinal(self):
        return TaskItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    # Lower case aliases for ToDoTaskOrdinal
    @property
    def todotaskordinal(self):
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        self.ToDoTaskOrdinal = value

    @property
    def TotalWork(self):
        return self.com_object.TotalWork

    @TotalWork.setter
    def TotalWork(self, value):
        self.com_object.TotalWork = value

    # Lower case aliases for TotalWork
    @property
    def totalwork(self):
        return self.TotalWork

    @totalwork.setter
    def totalwork(self, value):
        self.TotalWork = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Assign(self):
        return TaskItem(self.com_object.Assign())

    # Lower case alias for Assign
    def assign(self):
        return self.Assign()

    def CancelResponseState(self):
        self.com_object.CancelResponseState()

    # Lower case alias for CancelResponseState
    def cancelresponsestate(self):
        return self.CancelResponseState()

    def ClearRecurrencePattern(self):
        self.com_object.ClearRecurrencePattern()

    # Lower case alias for ClearRecurrencePattern
    def clearrecurrencepattern(self):
        return self.ClearRecurrencePattern()

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def GetRecurrencePattern(self):
        return RecurrencePattern(self.com_object.GetRecurrencePattern())

    # Lower case alias for GetRecurrencePattern
    def getrecurrencepattern(self):
        return self.GetRecurrencePattern()

    def MarkComplete(self):
        self.com_object.MarkComplete()

    # Lower case alias for MarkComplete
    def markcomplete(self):
        return self.MarkComplete()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Respond(self, Response=None, fNoUI=None, fAdditionalTextDialog=None):
        arguments = com_arguments([Response, fNoUI, fAdditionalTextDialog])
        return TaskItem(self.com_object.Respond(*arguments))

    # Lower case alias for Respond
    def respond(self, Response=None, fNoUI=None, fAdditionalTextDialog=None):
        arguments = [Response, fNoUI, fAdditionalTextDialog]
        return self.Respond(*arguments)

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def Send(self):
        self.com_object.Send()

    # Lower case alias for Send
    def send(self):
        return self.Send()

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()

    def SkipRecurrence(self):
        return self.com_object.SkipRecurrence()

    # Lower case alias for SkipRecurrence
    def skiprecurrence(self):
        return self.SkipRecurrence()

    def StatusReport(self):
        return Object(self.com_object.StatusReport())

    # Lower case alias for StatusReport
    def statusreport(self):
        return self.StatusReport()


class TaskRequestAcceptItem:

    def __init__(self, taskrequestacceptitem=None):
        self.com_object= taskrequestacceptitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([AddToTaskList])
        return TaskItem(self.com_object.GetAssociatedTask(*arguments))

    # Lower case alias for GetAssociatedTask
    def getassociatedtask(self, AddToTaskList=None):
        arguments = [AddToTaskList]
        return self.GetAssociatedTask(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class TaskRequestDeclineItem:

    def __init__(self, taskrequestdeclineitem=None):
        self.com_object= taskrequestdeclineitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([AddToTaskList])
        return TaskItem(self.com_object.GetAssociatedTask(*arguments))

    # Lower case alias for GetAssociatedTask
    def getassociatedtask(self, AddToTaskList=None):
        arguments = [AddToTaskList]
        return self.GetAssociatedTask(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class TaskRequestItem:

    def __init__(self, taskrequestitem=None):
        self.com_object= taskrequestitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([AddToTaskList])
        return TaskItem(self.com_object.GetAssociatedTask(*arguments))

    # Lower case alias for GetAssociatedTask
    def getassociatedtask(self, AddToTaskList=None):
        arguments = [AddToTaskList]
        return self.GetAssociatedTask(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class TaskRequestUpdateItem:

    def __init__(self, taskrequestupdateitem=None):
        self.com_object= taskrequestupdateitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    # Lower case aliases for Actions
    @property
    def actions(self):
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    # Lower case aliases for Attachments
    @property
    def attachments(self):
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    # Lower case aliases for AutoResolvedWinner
    @property
    def autoresolvedwinner(self):
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    # Lower case aliases for BillingInformation
    @property
    def billinginformation(self):
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    # Lower case aliases for Body
    @property
    def body(self):
        return self.Body

    @body.setter
    def body(self, value):
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    # Lower case aliases for Categories
    @property
    def categories(self):
        return self.Categories

    @categories.setter
    def categories(self, value):
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Companies(self):
        return self.com_object.Companies

    @Companies.setter
    def Companies(self, value):
        self.com_object.Companies = value

    # Lower case aliases for Companies
    @property
    def companies(self):
        return self.Companies

    @companies.setter
    def companies(self, value):
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    # Lower case aliases for Conflicts
    @property
    def conflicts(self):
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    # Lower case aliases for ConversationID
    @property
    def conversationid(self):
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    # Lower case aliases for ConversationIndex
    @property
    def conversationindex(self):
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    # Lower case aliases for ConversationTopic
    @property
    def conversationtopic(self):
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    # Lower case aliases for CreationTime
    @property
    def creationtime(self):
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    # Lower case aliases for DownloadState
    @property
    def downloadstate(self):
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    # Lower case aliases for EntryID
    @property
    def entryid(self):
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    # Lower case aliases for FormDescription
    @property
    def formdescription(self):
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    # Lower case aliases for GetInspector
    @property
    def getinspector(self):
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    # Lower case aliases for Importance
    @property
    def importance(self):
        return self.Importance

    @importance.setter
    def importance(self, value):
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    # Lower case aliases for IsConflict
    @property
    def isconflict(self):
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    # Lower case aliases for ItemProperties
    @property
    def itemproperties(self):
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    # Lower case aliases for LastModificationTime
    @property
    def lastmodificationtime(self):
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    # Lower case aliases for MarkForDownload
    @property
    def markfordownload(self):
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    # Lower case aliases for MessageClass
    @property
    def messageclass(self):
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    # Lower case aliases for Mileage
    @property
    def mileage(self):
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    # Lower case aliases for NoAging
    @property
    def noaging(self):
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    # Lower case aliases for OutlookInternalVersion
    @property
    def outlookinternalversion(self):
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    # Lower case aliases for OutlookVersion
    @property
    def outlookversion(self):
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    # Lower case aliases for PropertyAccessor
    @property
    def propertyaccessor(self):
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    # Lower case aliases for RTFBody
    @property
    def rtfbody(self):
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    # Lower case aliases for Saved
    @property
    def saved(self):
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    # Lower case aliases for Sensitivity
    @property
    def sensitivity(self):
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    # Lower case aliases for Subject
    @property
    def subject(self):
        return self.Subject

    @subject.setter
    def subject(self, value):
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    # Lower case aliases for UnRead
    @property
    def unread(self):
        return self.UnRead

    @unread.setter
    def unread(self, value):
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    # Lower case aliases for UserProperties
    @property
    def userproperties(self):
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([SaveMode])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Copy(self):
        self.com_object.Copy()

    # Lower case alias for Copy
    def copy(self):
        return self.Copy()

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Display(self, Modal=None):
        arguments = com_arguments([Modal])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([AddToTaskList])
        return TaskItem(self.com_object.GetAssociatedTask(*arguments))

    # Lower case alias for GetAssociatedTask
    def getassociatedtask(self, AddToTaskList=None):
        arguments = [AddToTaskList]
        return self.GetAssociatedTask(*arguments)

    def GetConversation(self):
        return Conversation(self.com_object.GetConversation())

    # Lower case alias for GetConversation
    def getconversation(self):
        return self.GetConversation()

    def Move(self, DestFldr=None):
        arguments = com_arguments([DestFldr])
        return Object(self.com_object.Move(*arguments))

    # Lower case alias for Move
    def move(self, DestFldr=None):
        arguments = [DestFldr]
        return self.Move(*arguments)

    def PrintOut(self):
        self.com_object.PrintOut()

    # Lower case alias for PrintOut
    def printout(self):
        return self.PrintOut()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()

    def SaveAs(self, Path=None, Type=None):
        arguments = com_arguments([Path, Type])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def ShowCategoriesDialog(self):
        self.com_object.ShowCategoriesDialog()

    # Lower case alias for ShowCategoriesDialog
    def showcategoriesdialog(self):
        return self.ShowCategoriesDialog()


class TasksModule:

    def __init__(self, tasksmodule=None):
        self.com_object= tasksmodule

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Name(self):
        return TasksModule(self.com_object.Name)

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    # Lower case aliases for NavigationGroups
    @property
    def navigationgroups(self):
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.com_object.NavigationModuleType)

    # Lower case aliases for NavigationModuleType
    @property
    def navigationmoduletype(self):
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Position(self):
        return TasksModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    # Lower case aliases for Position
    @property
    def position(self):
        return self.Position

    @position.setter
    def position(self, value):
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Visible(self):
        return TasksModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    # Lower case aliases for Visible
    @property
    def visible(self):
        return self.Visible

    @visible.setter
    def visible(self, value):
        self.Visible = value


class TextRuleCondition:

    def __init__(self, textrulecondition=None):
        self.com_object= textrulecondition

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    # Lower case aliases for Text
    @property
    def text(self):
        return self.Text

    @text.setter
    def text(self, value):
        self.Text = value


class TimelineView:

    def __init__(self, timelineview=None):
        self.com_object= timelineview

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DefaultExpandCollapseSetting(self):
        return OlDefaultExpandCollapseSetting(self.com_object.DefaultExpandCollapseSetting)

    @DefaultExpandCollapseSetting.setter
    def DefaultExpandCollapseSetting(self, value):
        self.com_object.DefaultExpandCollapseSetting = value

    # Lower case aliases for DefaultExpandCollapseSetting
    @property
    def defaultexpandcollapsesetting(self):
        return self.DefaultExpandCollapseSetting

    @defaultexpandcollapsesetting.setter
    def defaultexpandcollapsesetting(self, value):
        self.DefaultExpandCollapseSetting = value

    @property
    def EndField(self):
        return TimelineView(self.com_object.EndField)

    @EndField.setter
    def EndField(self, value):
        self.com_object.EndField = value

    # Lower case aliases for EndField
    @property
    def endfield(self):
        return self.EndField

    @endfield.setter
    def endfield(self, value):
        self.EndField = value

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def GroupByFields(self):
        return OrderFields(self.com_object.GroupByFields)

    # Lower case aliases for GroupByFields
    @property
    def groupbyfields(self):
        return self.GroupByFields

    @property
    def ItemFont(self):
        return ViewFont(self.com_object.ItemFont)

    # Lower case aliases for ItemFont
    @property
    def itemfont(self):
        return self.ItemFont

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    # Lower case aliases for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def LowerScaleFont(self):
        return ViewFont(self.com_object.LowerScaleFont)

    # Lower case aliases for LowerScaleFont
    @property
    def lowerscalefont(self):
        return self.LowerScaleFont

    @property
    def MaxLabelWidth(self):
        return TimelineView(self.com_object.MaxLabelWidth)

    @MaxLabelWidth.setter
    def MaxLabelWidth(self, value):
        self.com_object.MaxLabelWidth = value

    # Lower case aliases for MaxLabelWidth
    @property
    def maxlabelwidth(self):
        return self.MaxLabelWidth

    @maxlabelwidth.setter
    def maxlabelwidth(self, value):
        self.MaxLabelWidth = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    # Lower case aliases for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ShowLabelWhenViewingByMonth(self):
        return TimelineView(self.com_object.ShowLabelWhenViewingByMonth)

    @ShowLabelWhenViewingByMonth.setter
    def ShowLabelWhenViewingByMonth(self, value):
        self.com_object.ShowLabelWhenViewingByMonth = value

    # Lower case aliases for ShowLabelWhenViewingByMonth
    @property
    def showlabelwhenviewingbymonth(self):
        return self.ShowLabelWhenViewingByMonth

    @showlabelwhenviewingbymonth.setter
    def showlabelwhenviewingbymonth(self, value):
        self.ShowLabelWhenViewingByMonth = value

    @property
    def ShowWeekNumbers(self):
        return TimelineView(self.com_object.ShowWeekNumbers)

    @ShowWeekNumbers.setter
    def ShowWeekNumbers(self, value):
        self.com_object.ShowWeekNumbers = value

    # Lower case aliases for ShowWeekNumbers
    @property
    def showweeknumbers(self):
        return self.ShowWeekNumbers

    @showweeknumbers.setter
    def showweeknumbers(self, value):
        self.ShowWeekNumbers = value

    @property
    def Standard(self):
        return TimelineView(self.com_object.Standard)

    # Lower case aliases for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def StartField(self):
        return TimelineView(self.com_object.StartField)

    @StartField.setter
    def StartField(self, value):
        self.com_object.StartField = value

    # Lower case aliases for StartField
    @property
    def startfield(self):
        return self.StartField

    @startfield.setter
    def startfield(self, value):
        self.StartField = value

    @property
    def TimelineViewMode(self):
        return OlTimelineViewMode(self.com_object.TimelineViewMode)

    @TimelineViewMode.setter
    def TimelineViewMode(self, value):
        self.com_object.TimelineViewMode = value

    # Lower case aliases for TimelineViewMode
    @property
    def timelineviewmode(self):
        return self.TimelineViewMode

    @timelineviewmode.setter
    def timelineviewmode(self, value):
        self.TimelineViewMode = value

    @property
    def UpperScaleFont(self):
        return ViewFont(self.com_object.UpperScaleFont)

    # Lower case aliases for UpperScaleFont
    @property
    def upperscalefont(self):
        return self.UpperScaleFont

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    # Lower case aliases for XML
    @property
    def xml(self):
        return self.XML

    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        return View(self.com_object.Copy(*arguments))

    # Lower case alias for Copy
    def copy(self, Name=None, SaveOption=None):
        arguments = [Name, SaveOption]
        return self.Copy(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.com_object.GoToDate(*arguments)

    # Lower case alias for GoToDate
    def gotodate(self, Date=None):
        arguments = [Date]
        return self.GoToDate(*arguments)

    def Reset(self):
        self.com_object.Reset()

    # Lower case alias for Reset
    def reset(self):
        return self.Reset()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class TimeZone:

    def __init__(self, timezone=None):
        self.com_object= timezone

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Bias(self):
        return self.com_object.Bias

    # Lower case aliases for Bias
    @property
    def bias(self):
        return self.Bias

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DaylightBias(self):
        return self.com_object.DaylightBias

    # Lower case aliases for DaylightBias
    @property
    def daylightbias(self):
        return self.DaylightBias

    @property
    def DaylightDate(self):
        return self.com_object.DaylightDate

    # Lower case aliases for DaylightDate
    @property
    def daylightdate(self):
        return self.DaylightDate

    @property
    def DaylightDesignation(self):
        return self.com_object.DaylightDesignation

    # Lower case aliases for DaylightDesignation
    @property
    def daylightdesignation(self):
        return self.DaylightDesignation

    @property
    def ID(self):
        return self.com_object.ID

    # Lower case aliases for ID
    @property
    def id(self):
        return self.ID

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def StandardBias(self):
        return self.com_object.StandardBias

    # Lower case aliases for StandardBias
    @property
    def standardbias(self):
        return self.StandardBias

    @property
    def StandardDate(self):
        return self.com_object.StandardDate

    # Lower case aliases for StandardDate
    @property
    def standarddate(self):
        return self.StandardDate

    @property
    def StandardDesignation(self):
        return self.com_object.StandardDesignation

    # Lower case aliases for StandardDesignation
    @property
    def standarddesignation(self):
        return self.StandardDesignation


class TimeZones:

    def __init__(self, timezones=None):
        self.com_object= timezones

    def __call__(self, item):
        return TimeZone(self.com_object(item))

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def CurrentTimeZone(self):
        return TimeZone(self.com_object.CurrentTimeZone)

    # Lower case aliases for CurrentTimeZone
    @property
    def currenttimezone(self):
        return self.CurrentTimeZone

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def ConvertTime(self, SourceDateTime=None, SourceTimeZone=None, DestinationTimeZone=None):
        arguments = com_arguments([SourceDateTime, SourceTimeZone, DestinationTimeZone])
        return Date(self.com_object.ConvertTime(*arguments))

    # Lower case alias for ConvertTime
    def converttime(self, SourceDateTime=None, SourceTimeZone=None, DestinationTimeZone=None):
        arguments = [SourceDateTime, SourceTimeZone, DestinationTimeZone]
        return self.ConvertTime(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return TimeZone(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class ToOrFromRuleCondition:

    def __init__(self, toorfromrulecondition=None):
        self.com_object= toorfromrulecondition

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    # Lower case aliases for ConditionType
    @property
    def conditiontype(self):
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    # Lower case aliases for Enabled
    @property
    def enabled(self):
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    # Lower case aliases for Recipients
    @property
    def recipients(self):
        return self.Recipients

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session


class UserDefinedProperties:

    def __init__(self, userdefinedproperties=None):
        self.com_object= userdefinedproperties

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Type=None, DisplayFormat=None, Formula=None):
        arguments = com_arguments([Name, Type, DisplayFormat, Formula])
        return UserDefinedProperty(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Type=None, DisplayFormat=None, Formula=None):
        arguments = [Name, Type, DisplayFormat, Formula]
        return self.Add(*arguments)

    def Find(self, Name=None):
        arguments = com_arguments([Name])
        return self.com_object.Find(*arguments)

    # Lower case alias for Find
    def find(self, Name=None):
        arguments = [Name]
        return self.Find(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return UserDefinedProperty(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Refresh(self):
        self.com_object.Refresh()

    # Lower case alias for Refresh
    def refresh(self):
        return self.Refresh()

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class UserDefinedProperty:

    def __init__(self, userdefinedproperty=None):
        self.com_object= userdefinedproperty

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DisplayFormat(self):
        return UserDefinedProperty(self.com_object.DisplayFormat)

    # Lower case aliases for DisplayFormat
    @property
    def displayformat(self):
        return self.DisplayFormat

    @property
    def Formula(self):
        return self.com_object.Formula

    # Lower case aliases for Formula
    @property
    def formula(self):
        return self.Formula

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.com_object.Type)

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class UserProperties:

    def __init__(self, userproperties=None):
        self.com_object= userproperties

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = com_arguments([Name, Type, AddToFolderFields, DisplayFormat])
        return UserProperty(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = [Name, Type, AddToFolderFields, DisplayFormat]
        return self.Add(*arguments)

    def Find(self, Name=None, Custom=None):
        arguments = com_arguments([Name, Custom])
        return self.com_object.Find(*arguments)

    # Lower case alias for Find
    def find(self, Name=None, Custom=None):
        arguments = [Name, Custom]
        return self.Find(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return UserProperty(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class UserProperty:

    def __init__(self, userproperty=None):
        self.com_object= userproperty

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Formula(self):
        return self.com_object.Formula

    @Formula.setter
    def Formula(self, value):
        self.com_object.Formula = value

    # Lower case aliases for Formula
    @property
    def formula(self):
        return self.Formula

    @formula.setter
    def formula(self, value):
        self.Formula = value

    @property
    def Name(self):
        return self.com_object.Name

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.com_object.Type)

    # Lower case aliases for Type
    @property
    def type(self):
        return self.Type

    @property
    def ValidationFormula(self):
        return self.com_object.ValidationFormula

    @ValidationFormula.setter
    def ValidationFormula(self, value):
        self.com_object.ValidationFormula = value

    # Lower case aliases for ValidationFormula
    @property
    def validationformula(self):
        return self.ValidationFormula

    @validationformula.setter
    def validationformula(self, value):
        self.ValidationFormula = value

    @property
    def ValidationText(self):
        return self.com_object.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.com_object.ValidationText = value

    # Lower case aliases for ValidationText
    @property
    def validationtext(self):
        return self.ValidationText

    @validationtext.setter
    def validationtext(self, value):
        self.ValidationText = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    # Lower case aliases for Value
    @property
    def value(self):
        return self.Value

    @value.setter
    def value(self, value):
        self.Value = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()


class View:

    def __init__(self, view=None):
        self.com_object= view

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    # Lower case aliases for Filter
    @property
    def filter(self):
        return self.Filter

    @filter.setter
    def filter(self, value):
        self.Filter = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    # Lower case aliases for Language
    @property
    def language(self):
        return self.Language

    @language.setter
    def language(self, value):
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    # Lower case aliases for LockUserChanges
    @property
    def lockuserchanges(self):
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    # Lower case aliases for SaveOption
    @property
    def saveoption(self):
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Standard(self):
        return self.com_object.Standard

    # Lower case aliases for Standard
    @property
    def standard(self):
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    # Lower case aliases for ViewType
    @property
    def viewtype(self):
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    # Lower case aliases for XML
    @property
    def xml(self):
        return self.XML

    @xml.setter
    def xml(self, value):
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([Name, SaveOption])
        self.com_object.Copy(*arguments)

    # Lower case alias for Copy
    def copy(self, Name=None, SaveOption=None):
        arguments = [Name, SaveOption]
        return self.Copy(*arguments)

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GoToDate(self, Date=None):
        arguments = com_arguments([Date])
        self.com_object.GoToDate(*arguments)

    # Lower case alias for GoToDate
    def gotodate(self, Date=None):
        arguments = [Date]
        return self.GoToDate(*arguments)

    def Reset(self):
        self.com_object.Reset()

    # Lower case alias for Reset
    def reset(self):
        return self.Reset()

    def Save(self):
        self.com_object.Save()

    # Lower case alias for Save
    def save(self):
        return self.Save()


class ViewField:

    def __init__(self, viewfield=None):
        self.com_object= viewfield

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ColumnFormat(self):
        return ColumnFormat(self.com_object.ColumnFormat)

    # Lower case aliases for ColumnFormat
    @property
    def columnformat(self):
        return self.ColumnFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def ViewXMLSchemaName(self):
        return ViewField(self.com_object.ViewXMLSchemaName)

    # Lower case aliases for ViewXMLSchemaName
    @property
    def viewxmlschemaname(self):
        return self.ViewXMLSchemaName


class ViewFields:

    def __init__(self, viewfields=None):
        self.com_object= viewfields

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return ViewField(self.com_object.Count)

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, PropertyName=None):
        arguments = com_arguments([PropertyName])
        return ViewField(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, PropertyName=None):
        arguments = [PropertyName]
        return self.Add(*arguments)

    def Insert(self, PropertyName=None, Index=None):
        arguments = com_arguments([PropertyName, Index])
        return ViewField(self.com_object.Insert(*arguments))

    # Lower case alias for Insert
    def insert(self, PropertyName=None, Index=None):
        arguments = [PropertyName, Index]
        return self.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return ViewField(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class ViewFont:

    def __init__(self, viewfont=None):
        self.com_object= viewfont

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Bold(self):
        return ViewFont(self.com_object.Bold)

    @Bold.setter
    def Bold(self, value):
        self.com_object.Bold = value

    # Lower case aliases for Bold
    @property
    def bold(self):
        return self.Bold

    @bold.setter
    def bold(self, value):
        self.Bold = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Color(self):
        return OlColor(self.com_object.Color)

    @Color.setter
    def Color(self, value):
        self.com_object.Color = value

    # Lower case aliases for Color
    @property
    def color(self):
        return self.Color

    @color.setter
    def color(self, value):
        self.Color = value

    @property
    def ExtendedColor(self):
        return OlCategoryColor(self.com_object.ExtendedColor)

    @ExtendedColor.setter
    def ExtendedColor(self, value):
        self.com_object.ExtendedColor = value

    # Lower case aliases for ExtendedColor
    @property
    def extendedcolor(self):
        return self.ExtendedColor

    @extendedcolor.setter
    def extendedcolor(self, value):
        self.ExtendedColor = value

    @property
    def Italic(self):
        return ViewFont(self.com_object.Italic)

    @Italic.setter
    def Italic(self, value):
        self.com_object.Italic = value

    # Lower case aliases for Italic
    @property
    def italic(self):
        return self.Italic

    @italic.setter
    def italic(self, value):
        self.Italic = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    # Lower case aliases for Name
    @property
    def name(self):
        return self.Name

    @name.setter
    def name(self, value):
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @Size.setter
    def Size(self, value):
        self.com_object.Size = value

    # Lower case aliases for Size
    @property
    def size(self):
        return self.Size

    @size.setter
    def size(self, value):
        self.Size = value

    @property
    def Strikethrough(self):
        return ViewFont(self.com_object.Strikethrough)

    @Strikethrough.setter
    def Strikethrough(self, value):
        self.com_object.Strikethrough = value

    # Lower case aliases for Strikethrough
    @property
    def strikethrough(self):
        return self.Strikethrough

    @strikethrough.setter
    def strikethrough(self, value):
        self.Strikethrough = value

    @property
    def Underline(self):
        return ViewFont(self.com_object.Underline)

    @Underline.setter
    def Underline(self, value):
        self.com_object.Underline = value

    # Lower case aliases for Underline
    @property
    def underline(self):
        return self.Underline

    @underline.setter
    def underline(self, value):
        self.Underline = value


class Views:

    def __init__(self, views=None):
        self.com_object= views

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    # Lower case aliases for Count
    @property
    def count(self):
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    # Lower case aliases for Parent
    @property
    def parent(self):
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    # Lower case aliases for Session
    @property
    def session(self):
        return self.Session

    def Add(self, Name=None, ViewType=None, SaveOption=None):
        arguments = com_arguments([Name, ViewType, SaveOption])
        return View(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, ViewType=None, SaveOption=None):
        arguments = [Name, ViewType, SaveOption]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return View(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([Index])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

