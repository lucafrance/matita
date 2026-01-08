from . import com_arguments, unwrap
from .office import *

import win32com.client

class Account:

    def __init__(self, account=None):
        self.com_object= account

    @property
    def AccountType(self):
        return OlAccountType(self.com_object.AccountType)

    @property
    def accounttype(self):
        """Lower case alias for AccountType"""
        return self.AccountType

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoDiscoverConnectionMode(self):
        return OlAutoDiscoverConnectionMode(self.com_object.AutoDiscoverConnectionMode)

    @property
    def autodiscoverconnectionmode(self):
        """Lower case alias for AutoDiscoverConnectionMode"""
        return self.AutoDiscoverConnectionMode

    @property
    def AutoDiscoverXml(self):
        return self.com_object.AutoDiscoverXml

    @property
    def autodiscoverxml(self):
        """Lower case alias for AutoDiscoverXml"""
        return self.AutoDiscoverXml

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentUser(self):
        return Recipient(self.com_object.CurrentUser)

    @property
    def currentuser(self):
        """Lower case alias for CurrentUser"""
        return self.CurrentUser

    @property
    def DeliveryStore(self):
        return Store(self.com_object.DeliveryStore)

    @property
    def deliverystore(self):
        """Lower case alias for DeliveryStore"""
        return self.DeliveryStore

    @property
    def DisplayName(self):
        return Account(self.com_object.DisplayName)

    @property
    def displayname(self):
        """Lower case alias for DisplayName"""
        return self.DisplayName

    @property
    def ExchangeConnectionMode(self):
        return OlExchangeConnectionMode(self.com_object.ExchangeConnectionMode)

    @property
    def exchangeconnectionmode(self):
        """Lower case alias for ExchangeConnectionMode"""
        return self.ExchangeConnectionMode

    @property
    def ExchangeMailboxServerName(self):
        return self.com_object.ExchangeMailboxServerName

    @property
    def exchangemailboxservername(self):
        """Lower case alias for ExchangeMailboxServerName"""
        return self.ExchangeMailboxServerName

    @property
    def ExchangeMailboxServerVersion(self):
        return self.com_object.ExchangeMailboxServerVersion

    @property
    def exchangemailboxserverversion(self):
        """Lower case alias for ExchangeMailboxServerVersion"""
        return self.ExchangeMailboxServerVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def SmtpAddress(self):
        return Account(self.com_object.SmtpAddress)

    @property
    def smtpaddress(self):
        """Lower case alias for SmtpAddress"""
        return self.SmtpAddress

    @property
    def UserName(self):
        return Account(self.com_object.UserName)

    @property
    def username(self):
        """Lower case alias for UserName"""
        return self.UserName

    def GetAddressEntryFromID(self, ID=None):
        arguments = com_arguments([unwrap(a) for a in [ID]])
        return self.com_object.GetAddressEntryFromID(*arguments)

    # Lower case alias for GetAddressEntryFromID
    def getaddressentryfromid(self, ID=None):
        arguments = [ID]
        return self.GetAddressEntryFromID(*arguments)

    def GetRecipientFromID(self, EntryID=None):
        arguments = com_arguments([unwrap(a) for a in [EntryID]])
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

    @property
    def account(self):
        """Lower case alias for Account"""
        return self.Account

    @account.setter
    def account(self, value):
        """Lower case alias for Account.setter"""
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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SelectedAccount(self):
        return Account(self.com_object.SelectedAccount)

    @property
    def selectedaccount(self):
        """Lower case alias for SelectedAccount"""
        return self.SelectedAccount

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def copylike(self):
        """Lower case alias for CopyLike"""
        return self.CopyLike

    @copylike.setter
    def copylike(self, value):
        """Lower case alias for CopyLike.setter"""
        self.CopyLike = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def MessageClass(self):
        return Action(self.com_object.MessageClass)

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Prefix(self):
        return self.com_object.Prefix

    @Prefix.setter
    def Prefix(self, value):
        self.com_object.Prefix = value

    @property
    def prefix(self):
        """Lower case alias for Prefix"""
        return self.Prefix

    @prefix.setter
    def prefix(self, value):
        """Lower case alias for Prefix.setter"""
        self.Prefix = value

    @property
    def ReplyStyle(self):
        return OlActionReplyStyle(self.com_object.ReplyStyle)

    @ReplyStyle.setter
    def ReplyStyle(self, value):
        self.com_object.ReplyStyle = value

    @property
    def replystyle(self):
        """Lower case alias for ReplyStyle"""
        return self.ReplyStyle

    @replystyle.setter
    def replystyle(self, value):
        """Lower case alias for ReplyStyle.setter"""
        self.ReplyStyle = value

    @property
    def ResponseStyle(self):
        return OlActionResponseStyle(self.com_object.ResponseStyle)

    @ResponseStyle.setter
    def ResponseStyle(self, value):
        self.com_object.ResponseStyle = value

    @property
    def responsestyle(self):
        """Lower case alias for ResponseStyle"""
        return self.ResponseStyle

    @responsestyle.setter
    def responsestyle(self, value):
        """Lower case alias for ResponseStyle.setter"""
        self.ResponseStyle = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ShowOn(self):
        return OlActionShowOn(self.com_object.ShowOn)

    @ShowOn.setter
    def ShowOn(self, value):
        self.com_object.ShowOn = value

    @property
    def showon(self):
        """Lower case alias for ShowOn"""
        return self.ShowOn

    @showon.setter
    def showon(self, value):
        """Lower case alias for ShowOn.setter"""
        self.ShowOn = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Execute(self):
        return self.com_object.Execute()

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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self):
        return Action(self.com_object.Add())

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Action(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Type=None, Name=None, Address=None):
        arguments = com_arguments([unwrap(a) for a in [Type, Name, Address]])
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
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return AddressEntry(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Sort(self, Property=None, Order=None):
        arguments = com_arguments([unwrap(a) for a in [Property, Order]])
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

    @property
    def address(self):
        """Lower case alias for Address"""
        return self.Address

    @address.setter
    def address(self, value):
        """Lower case alias for Address.setter"""
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.com_object.AddressEntryUserType)

    @property
    def addressentryusertype(self):
        """Lower case alias for AddressEntryUserType"""
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

    @property
    def displaytype(self):
        """Lower case alias for DisplayType"""
        return self.DisplayType

    @property
    def ID(self):
        return self.com_object.ID

    @property
    def id(self):
        """Lower case alias for ID"""
        return self.ID

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Type(self):
        return self.com_object.Type

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @type.setter
    def type(self, value):
        """Lower case alias for Type.setter"""
        self.Type = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([unwrap(a) for a in [HWnd]])
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
        arguments = com_arguments([unwrap(a) for a in [Start, MinPerChar, CompleteFormat]])
        return self.com_object.GetFreeBusy(*arguments)

    # Lower case alias for GetFreeBusy
    def getfreebusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = [Start, MinPerChar, CompleteFormat]
        return self.GetFreeBusy(*arguments)

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([unwrap(a) for a in [MakePermanent, Refresh]])
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

    @property
    def addressentries(self):
        """Lower case alias for AddressEntries"""
        return self.AddressEntries

    @property
    def AddressListType(self):
        return OlAddressListType(self.com_object.AddressListType)

    @property
    def addresslisttype(self):
        """Lower case alias for AddressListType"""
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

    @property
    def id(self):
        """Lower case alias for ID"""
        return self.ID

    @property
    def Index(self):
        return self.com_object.Index

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def IsInitialAddressList(self):
        return AddressList(self.com_object.IsInitialAddressList)

    @property
    def isinitialaddresslist(self):
        """Lower case alias for IsInitialAddressList"""
        return self.IsInitialAddressList

    @property
    def IsReadOnly(self):
        return AddressList(self.com_object.IsReadOnly)

    @property
    def isreadonly(self):
        """Lower case alias for IsReadOnly"""
        return self.IsReadOnly

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def ResolutionOrder(self):
        return AddressList(self.com_object.ResolutionOrder)

    @property
    def resolutionorder(self):
        """Lower case alias for ResolutionOrder"""
        return self.ResolutionOrder

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def address(self):
        """Lower case alias for Address"""
        return self.Address

    @address.setter
    def address(self, value):
        """Lower case alias for Address.setter"""
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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class Application:

    def __init__(self, application=None):
        if application is None:
            self.com_object = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        else:
            self.com_object = application

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Assistance(self):
        return self.com_object.Assistance

    @property
    def assistance(self):
        """Lower case alias for Assistance"""
        return self.Assistance

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def COMAddIns(self):
        return self.com_object.COMAddIns

    @property
    def comaddins(self):
        """Lower case alias for COMAddIns"""
        return self.COMAddIns

    @property
    def DefaultProfileName(self):
        return self.com_object.DefaultProfileName

    @property
    def defaultprofilename(self):
        """Lower case alias for DefaultProfileName"""
        return self.DefaultProfileName

    @property
    def Explorers(self):
        return Explorers(self.com_object.Explorers)

    @property
    def explorers(self):
        """Lower case alias for Explorers"""
        return self.Explorers

    @property
    def Inspectors(self):
        return Inspectors(self.com_object.Inspectors)

    @property
    def inspectors(self):
        """Lower case alias for Inspectors"""
        return self.Inspectors

    @property
    def IsTrusted(self):
        return self.com_object.IsTrusted

    @property
    def istrusted(self):
        """Lower case alias for IsTrusted"""
        return self.IsTrusted

    @property
    def LanguageSettings(self):
        return self.com_object.LanguageSettings

    @property
    def languagesettings(self):
        """Lower case alias for LanguageSettings"""
        return self.LanguageSettings

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PickerDialog(self):
        return self.com_object.PickerDialog

    @property
    def pickerdialog(self):
        """Lower case alias for PickerDialog"""
        return self.PickerDialog

    @property
    def ProductCode(self):
        return self.com_object.ProductCode

    @property
    def productcode(self):
        """Lower case alias for ProductCode"""
        return self.ProductCode

    @property
    def Reminders(self):
        return Reminders(self.com_object.Reminders)

    @property
    def reminders(self):
        """Lower case alias for Reminders"""
        return self.Reminders

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def TimeZones(self):
        return TimeZones(self.com_object.TimeZones)

    @property
    def timezones(self):
        """Lower case alias for TimeZones"""
        return self.TimeZones

    @property
    def Version(self):
        return self.com_object.Version

    @Version.setter
    def Version(self, value):
        self.com_object.Version = value

    @property
    def version(self):
        """Lower case alias for Version"""
        return self.Version

    @version.setter
    def version(self, value):
        """Lower case alias for Version.setter"""
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
        arguments = com_arguments([unwrap(a) for a in [Scope, Filter, SearchSubFolders, Tag]])
        return Search(self.com_object.AdvancedSearch(*arguments))

    # Lower case alias for AdvancedSearch
    def advancedsearch(self, Scope=None, Filter=None, SearchSubFolders=None, Tag=None):
        arguments = [Scope, Filter, SearchSubFolders, Tag]
        return self.AdvancedSearch(*arguments)

    def CopyFile(self, FilePath=None, DestFolderPath=None):
        arguments = com_arguments([unwrap(a) for a in [FilePath, DestFolderPath]])
        return self.com_object.CopyFile(*arguments)

    # Lower case alias for CopyFile
    def copyfile(self, FilePath=None, DestFolderPath=None):
        arguments = [FilePath, DestFolderPath]
        return self.CopyFile(*arguments)

    def CreateItem(self, ItemType=None):
        arguments = com_arguments([unwrap(a) for a in [ItemType]])
        return self.com_object.CreateItem(*arguments)

    # Lower case alias for CreateItem
    def createitem(self, ItemType=None):
        arguments = [ItemType]
        return self.CreateItem(*arguments)

    def CreateItemFromTemplate(self, TemplatePath=None, InFolder=None):
        arguments = com_arguments([unwrap(a) for a in [TemplatePath, InFolder]])
        return self.com_object.CreateItemFromTemplate(*arguments)

    # Lower case alias for CreateItemFromTemplate
    def createitemfromtemplate(self, TemplatePath=None, InFolder=None):
        arguments = [TemplatePath, InFolder]
        return self.CreateItemFromTemplate(*arguments)

    def CreateObject(self, ObjectName=None):
        arguments = com_arguments([unwrap(a) for a in [ObjectName]])
        return self.com_object.CreateObject(*arguments)

    # Lower case alias for CreateObject
    def createobject(self, ObjectName=None):
        arguments = [ObjectName]
        return self.CreateObject(*arguments)

    def GetNamespace(self, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Type]])
        return NameSpace(self.com_object.GetNamespace(*arguments))

    # Lower case alias for GetNamespace
    def getnamespace(self, Type=None):
        arguments = [Type]
        return self.GetNamespace(*arguments)

    def GetObjectReference(self, Item=None, ReferenceType=None):
        arguments = com_arguments([unwrap(a) for a in [Item, ReferenceType]])
        return self.com_object.GetObjectReference(*arguments)

    # Lower case alias for GetObjectReference
    def getobjectreference(self, Item=None, ReferenceType=None):
        arguments = [Item, ReferenceType]
        return self.GetObjectReference(*arguments)

    def IsSearchSynchronous(self, LookInFolders=None):
        arguments = com_arguments([unwrap(a) for a in [LookInFolders]])
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
        arguments = com_arguments([unwrap(a) for a in [RegionName]])
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def AllDayEvent(self):
        return self.com_object.AllDayEvent

    @AllDayEvent.setter
    def AllDayEvent(self, value):
        self.com_object.AllDayEvent = value

    @property
    def alldayevent(self):
        """Lower case alias for AllDayEvent"""
        return self.AllDayEvent

    @alldayevent.setter
    def alldayevent(self, value):
        """Lower case alias for AllDayEvent.setter"""
        self.AllDayEvent = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def BusyStatus(self):
        return OlBusyStatus(self.com_object.BusyStatus)

    @BusyStatus.setter
    def BusyStatus(self, value):
        self.com_object.BusyStatus = value

    @property
    def busystatus(self):
        """Lower case alias for BusyStatus"""
        return self.BusyStatus

    @busystatus.setter
    def busystatus(self, value):
        """Lower case alias for BusyStatus.setter"""
        self.BusyStatus = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def Duration(self):
        return AppointmentItem(self.com_object.Duration)

    @Duration.setter
    def Duration(self, value):
        self.com_object.Duration = value

    @property
    def duration(self):
        """Lower case alias for Duration"""
        return self.Duration

    @duration.setter
    def duration(self, value):
        """Lower case alias for Duration.setter"""
        self.Duration = value

    @property
    def End(self):
        return AppointmentItem(self.com_object.End)

    @End.setter
    def End(self, value):
        self.com_object.End = value

    @property
    def end(self):
        """Lower case alias for End"""
        return self.End

    @end.setter
    def end(self, value):
        """Lower case alias for End.setter"""
        self.End = value

    @property
    def EndInEndTimeZone(self):
        return AppointmentItem.EndTimeZone(self.com_object.EndInEndTimeZone)

    @EndInEndTimeZone.setter
    def EndInEndTimeZone(self, value):
        self.com_object.EndInEndTimeZone = value

    @property
    def endinendtimezone(self):
        """Lower case alias for EndInEndTimeZone"""
        return self.EndInEndTimeZone

    @endinendtimezone.setter
    def endinendtimezone(self, value):
        """Lower case alias for EndInEndTimeZone.setter"""
        self.EndInEndTimeZone = value

    @property
    def EndTimeZone(self):
        return TimeZone(self.com_object.EndTimeZone)

    @EndTimeZone.setter
    def EndTimeZone(self, value):
        self.com_object.EndTimeZone = value

    @property
    def endtimezone(self):
        """Lower case alias for EndTimeZone"""
        return self.EndTimeZone

    @endtimezone.setter
    def endtimezone(self, value):
        """Lower case alias for EndTimeZone.setter"""
        self.EndTimeZone = value

    @property
    def EndUTC(self):
        return self.com_object.EndUTC

    @EndUTC.setter
    def EndUTC(self, value):
        self.com_object.EndUTC = value

    @property
    def endutc(self):
        """Lower case alias for EndUTC"""
        return self.EndUTC

    @endutc.setter
    def endutc(self, value):
        """Lower case alias for EndUTC.setter"""
        self.EndUTC = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def ForceUpdateToAllAttendees(self):
        return self.com_object.ForceUpdateToAllAttendees

    @ForceUpdateToAllAttendees.setter
    def ForceUpdateToAllAttendees(self, value):
        self.com_object.ForceUpdateToAllAttendees = value

    @property
    def forceupdatetoallattendees(self):
        """Lower case alias for ForceUpdateToAllAttendees"""
        return self.ForceUpdateToAllAttendees

    @forceupdatetoallattendees.setter
    def forceupdatetoallattendees(self, value):
        """Lower case alias for ForceUpdateToAllAttendees.setter"""
        self.ForceUpdateToAllAttendees = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def GlobalAppointmentID(self):
        return AppointmentItem(self.com_object.GlobalAppointmentID)

    @property
    def globalappointmentid(self):
        """Lower case alias for GlobalAppointmentID"""
        return self.GlobalAppointmentID

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    @property
    def internetcodepage(self):
        """Lower case alias for InternetCodepage"""
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        """Lower case alias for InternetCodepage.setter"""
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def IsRecurring(self):
        return self.com_object.IsRecurring

    @property
    def isrecurring(self):
        """Lower case alias for IsRecurring"""
        return self.IsRecurring

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def Location(self):
        return self.com_object.Location

    @Location.setter
    def Location(self, value):
        self.com_object.Location = value

    @property
    def location(self):
        """Lower case alias for Location"""
        return self.Location

    @location.setter
    def location(self, value):
        """Lower case alias for Location.setter"""
        self.Location = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MeetingStatus(self):
        return OlMeetingStatus(self.com_object.MeetingStatus)

    @MeetingStatus.setter
    def MeetingStatus(self, value):
        self.com_object.MeetingStatus = value

    @property
    def meetingstatus(self):
        """Lower case alias for MeetingStatus"""
        return self.MeetingStatus

    @meetingstatus.setter
    def meetingstatus(self, value):
        """Lower case alias for MeetingStatus.setter"""
        self.MeetingStatus = value

    @property
    def MeetingWorkspaceURL(self):
        return self.com_object.MeetingWorkspaceURL

    @property
    def meetingworkspaceurl(self):
        """Lower case alias for MeetingWorkspaceURL"""
        return self.MeetingWorkspaceURL

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OptionalAttendees(self):
        return self.com_object.OptionalAttendees

    @OptionalAttendees.setter
    def OptionalAttendees(self, value):
        self.com_object.OptionalAttendees = value

    @property
    def optionalattendees(self):
        """Lower case alias for OptionalAttendees"""
        return self.OptionalAttendees

    @optionalattendees.setter
    def optionalattendees(self, value):
        """Lower case alias for OptionalAttendees.setter"""
        self.OptionalAttendees = value

    @property
    def Organizer(self):
        return self.com_object.Organizer

    @property
    def organizer(self):
        """Lower case alias for Organizer"""
        return self.Organizer

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @property
    def RecurrenceState(self):
        return OlRecurrenceState(self.com_object.RecurrenceState)

    @property
    def recurrencestate(self):
        """Lower case alias for RecurrenceState"""
        return self.RecurrenceState

    @property
    def ReminderMinutesBeforeStart(self):
        return self.com_object.ReminderMinutesBeforeStart

    @ReminderMinutesBeforeStart.setter
    def ReminderMinutesBeforeStart(self, value):
        self.com_object.ReminderMinutesBeforeStart = value

    @property
    def reminderminutesbeforestart(self):
        """Lower case alias for ReminderMinutesBeforeStart"""
        return self.ReminderMinutesBeforeStart

    @reminderminutesbeforestart.setter
    def reminderminutesbeforestart(self, value):
        """Lower case alias for ReminderMinutesBeforeStart.setter"""
        self.ReminderMinutesBeforeStart = value

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    @property
    def reminderoverridedefault(self):
        """Lower case alias for ReminderOverrideDefault"""
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        """Lower case alias for ReminderOverrideDefault.setter"""
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    @property
    def reminderplaysound(self):
        """Lower case alias for ReminderPlaySound"""
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        """Lower case alias for ReminderPlaySound.setter"""
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    @property
    def reminderset(self):
        """Lower case alias for ReminderSet"""
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        """Lower case alias for ReminderSet.setter"""
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    @property
    def remindersoundfile(self):
        """Lower case alias for ReminderSoundFile"""
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        """Lower case alias for ReminderSoundFile.setter"""
        self.ReminderSoundFile = value

    @property
    def ReplyTime(self):
        return self.com_object.ReplyTime

    @ReplyTime.setter
    def ReplyTime(self, value):
        self.com_object.ReplyTime = value

    @property
    def replytime(self):
        """Lower case alias for ReplyTime"""
        return self.ReplyTime

    @replytime.setter
    def replytime(self, value):
        """Lower case alias for ReplyTime.setter"""
        self.ReplyTime = value

    @property
    def RequiredAttendees(self):
        return self.com_object.RequiredAttendees

    @RequiredAttendees.setter
    def RequiredAttendees(self, value):
        self.com_object.RequiredAttendees = value

    @property
    def requiredattendees(self):
        """Lower case alias for RequiredAttendees"""
        return self.RequiredAttendees

    @requiredattendees.setter
    def requiredattendees(self, value):
        """Lower case alias for RequiredAttendees.setter"""
        self.RequiredAttendees = value

    @property
    def Resources(self):
        return self.com_object.Resources

    @Resources.setter
    def Resources(self, value):
        self.com_object.Resources = value

    @property
    def resources(self):
        """Lower case alias for Resources"""
        return self.Resources

    @resources.setter
    def resources(self, value):
        """Lower case alias for Resources.setter"""
        self.Resources = value

    @property
    def ResponseRequested(self):
        return self.com_object.ResponseRequested

    @ResponseRequested.setter
    def ResponseRequested(self, value):
        self.com_object.ResponseRequested = value

    @property
    def responserequested(self):
        """Lower case alias for ResponseRequested"""
        return self.ResponseRequested

    @responserequested.setter
    def responserequested(self, value):
        """Lower case alias for ResponseRequested.setter"""
        self.ResponseRequested = value

    @property
    def ResponseStatus(self):
        return OlResponseStatus(self.com_object.ResponseStatus)

    @property
    def responsestatus(self):
        """Lower case alias for ResponseStatus"""
        return self.ResponseStatus

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    @property
    def sendusingaccount(self):
        """Lower case alias for SendUsingAccount"""
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        """Lower case alias for SendUsingAccount.setter"""
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Start(self):
        return self.com_object.Start

    @Start.setter
    def Start(self, value):
        self.com_object.Start = value

    @property
    def start(self):
        """Lower case alias for Start"""
        return self.Start

    @start.setter
    def start(self, value):
        """Lower case alias for Start.setter"""
        self.Start = value

    @property
    def StartInStartTimeZone(self):
        return AppointmentItem.StartTimeZone(self.com_object.StartInStartTimeZone)

    @StartInStartTimeZone.setter
    def StartInStartTimeZone(self, value):
        self.com_object.StartInStartTimeZone = value

    @property
    def startinstarttimezone(self):
        """Lower case alias for StartInStartTimeZone"""
        return self.StartInStartTimeZone

    @startinstarttimezone.setter
    def startinstarttimezone(self, value):
        """Lower case alias for StartInStartTimeZone.setter"""
        self.StartInStartTimeZone = value

    @property
    def StartTimeZone(self):
        return TimeZone(self.com_object.StartTimeZone)

    @StartTimeZone.setter
    def StartTimeZone(self, value):
        self.com_object.StartTimeZone = value

    @property
    def starttimezone(self):
        """Lower case alias for StartTimeZone"""
        return self.StartTimeZone

    @starttimezone.setter
    def starttimezone(self, value):
        """Lower case alias for StartTimeZone.setter"""
        self.StartTimeZone = value

    @property
    def StartUTC(self):
        return self.com_object.StartUTC

    @StartUTC.setter
    def StartUTC(self, value):
        self.com_object.StartUTC = value

    @property
    def startutc(self):
        """Lower case alias for StartUTC"""
        return self.StartUTC

    @startutc.setter
    def startutc(self, value):
        """Lower case alias for StartUTC.setter"""
        self.StartUTC = value

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def ClearRecurrencePattern(self):
        self.com_object.ClearRecurrencePattern()

    # Lower case alias for ClearRecurrencePattern
    def clearrecurrencepattern(self):
        return self.ClearRecurrencePattern()

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [DestinationFolder, CopyOptions]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Response, fNoUI, fAdditionalTextDialog]])
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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def actiontype(self):
        """Lower case alias for ActionType"""
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

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def blocklevel(self):
        """Lower case alias for BlockLevel"""
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

    @property
    def displayname(self):
        """Lower case alias for DisplayName"""
        return self.DisplayName

    @displayname.setter
    def displayname(self, value):
        """Lower case alias for DisplayName.setter"""
        self.DisplayName = value

    @property
    def FileName(self):
        return self.com_object.FileName

    @property
    def filename(self):
        """Lower case alias for FileName"""
        return self.FileName

    @property
    def Index(self):
        return self.com_object.Index

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PathName(self):
        return self.com_object.PathName

    @property
    def pathname(self):
        """Lower case alias for PathName"""
        return self.PathName

    @property
    def Position(self):
        return self.com_object.Position

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Type(self):
        return OlAttachmentType(self.com_object.Type)

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def GetTemporaryFilePath(self):
        return self.com_object.GetTemporaryFilePath()

    # Lower case alias for GetTemporaryFilePath
    def gettemporaryfilepath(self):
        return self.GetTemporaryFilePath()

    def SaveAsFile(self, Path=None):
        arguments = com_arguments([unwrap(a) for a in [Path]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Source=None, Type=None, Position=None, DisplayName=None):
        arguments = com_arguments([unwrap(a) for a in [Source, Type, Position, DisplayName]])
        return Attachment(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Source=None, Type=None, Position=None, DisplayName=None):
        arguments = [Source, Type, Position, DisplayName]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Attachment(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Location(self):
        return OlSelectionLocation(self.com_object.Location)

    @property
    def location(self):
        """Lower case alias for Location"""
        return self.Location

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def GetSelection(self, SelectionContents=None):
        arguments = com_arguments([unwrap(a) for a in [SelectionContents]])
        return Selection(self.com_object.GetSelection(*arguments))

    # Lower case alias for GetSelection
    def getselection(self, SelectionContents=None):
        arguments = [SelectionContents]
        return self.GetSelection(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for Filter.setter"""
        self.Filter = value

    @property
    def Font(self):
        return ViewFont(self.com_object.Font)

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Standard(self):
        return AutoFormatRule(self.com_object.Standard)

    @property
    def standard(self):
        """Lower case alias for Standard"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Name]])
        return AutoFormatRule(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None):
        arguments = [Name]
        return self.Add(*arguments)

    def Insert(self, Name=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Index]])
        return AutoFormatRule(self.com_object.Insert(*arguments))

    # Lower case alias for Insert
    def insert(self, Name=None, Index=None):
        arguments = [Name, Index]
        return self.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return AutoFormatRule(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def cardsize(self):
        """Lower case alias for CardSize"""
        return self.CardSize

    @cardsize.setter
    def cardsize(self, value):
        """Lower case alias for CardSize.setter"""
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

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for Filter.setter"""
        self.Filter = value

    @property
    def HeadingsFont(self):
        return ViewFont(self.com_object.HeadingsFont)

    @property
    def headingsfont(self):
        """Lower case alias for HeadingsFont"""
        return self.HeadingsFont

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @language.setter
    def language(self, value):
        """Lower case alias for Language.setter"""
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    @property
    def lockuserchanges(self):
        """Lower case alias for LockUserChanges"""
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        """Lower case alias for LockUserChanges.setter"""
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    @property
    def saveoption(self):
        """Lower case alias for SaveOption"""
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def SortFields(self):
        return OrderFields(self.com_object.SortFields)

    @property
    def sortfields(self):
        """Lower case alias for SortFields"""
        return self.SortFields

    @property
    def Standard(self):
        return BusinessCardView(self.com_object.Standard)

    @property
    def standard(self):
        """Lower case alias for Standard"""
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    @property
    def xml(self):
        """Lower case alias for XML"""
        return self.XML

    @xml.setter
    def xml(self, value):
        """Lower case alias for XML.setter"""
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SaveOption]])
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
        arguments = com_arguments([unwrap(a) for a in [Date]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    @property
    def navigationgroups(self):
        """Lower case alias for NavigationGroups"""
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.com_object.NavigationModuleType)

    @property
    def navigationmoduletype(self):
        """Lower case alias for NavigationModuleType"""
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return CalendarModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return CalendarModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
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

    @property
    def calendardetail(self):
        """Lower case alias for CalendarDetail"""
        return self.CalendarDetail

    @calendardetail.setter
    def calendardetail(self, value):
        """Lower case alias for CalendarDetail.setter"""
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

    @property
    def enddate(self):
        """Lower case alias for EndDate"""
        return self.EndDate

    @enddate.setter
    def enddate(self, value):
        """Lower case alias for EndDate.setter"""
        self.EndDate = value

    @property
    def Folder(self):
        return Folder(self.com_object.Folder)

    @property
    def folder(self):
        """Lower case alias for Folder"""
        return self.Folder

    @property
    def IncludeAttachments(self):
        return self.com_object.IncludeAttachments

    @IncludeAttachments.setter
    def IncludeAttachments(self, value):
        self.com_object.IncludeAttachments = value

    @property
    def includeattachments(self):
        """Lower case alias for IncludeAttachments"""
        return self.IncludeAttachments

    @includeattachments.setter
    def includeattachments(self, value):
        """Lower case alias for IncludeAttachments.setter"""
        self.IncludeAttachments = value

    @property
    def IncludePrivateDetails(self):
        return self.com_object.IncludePrivateDetails

    @IncludePrivateDetails.setter
    def IncludePrivateDetails(self, value):
        self.com_object.IncludePrivateDetails = value

    @property
    def includeprivatedetails(self):
        """Lower case alias for IncludePrivateDetails"""
        return self.IncludePrivateDetails

    @includeprivatedetails.setter
    def includeprivatedetails(self, value):
        """Lower case alias for IncludePrivateDetails.setter"""
        self.IncludePrivateDetails = value

    @property
    def IncludeWholeCalendar(self):
        return self.com_object.IncludeWholeCalendar

    @IncludeWholeCalendar.setter
    def IncludeWholeCalendar(self, value):
        self.com_object.IncludeWholeCalendar = value

    @property
    def includewholecalendar(self):
        """Lower case alias for IncludeWholeCalendar"""
        return self.IncludeWholeCalendar

    @includewholecalendar.setter
    def includewholecalendar(self, value):
        """Lower case alias for IncludeWholeCalendar.setter"""
        self.IncludeWholeCalendar = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RestrictToWorkingHours(self):
        return self.com_object.RestrictToWorkingHours

    @RestrictToWorkingHours.setter
    def RestrictToWorkingHours(self, value):
        self.com_object.RestrictToWorkingHours = value

    @property
    def restricttoworkinghours(self):
        """Lower case alias for RestrictToWorkingHours"""
        return self.RestrictToWorkingHours

    @restricttoworkinghours.setter
    def restricttoworkinghours(self, value):
        """Lower case alias for RestrictToWorkingHours.setter"""
        self.RestrictToWorkingHours = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def StartDate(self):
        return CalendarSharing(self.com_object.StartDate)

    @StartDate.setter
    def StartDate(self, value):
        self.com_object.StartDate = value

    @property
    def startdate(self):
        """Lower case alias for StartDate"""
        return self.StartDate

    @startdate.setter
    def startdate(self, value):
        """Lower case alias for StartDate.setter"""
        self.StartDate = value

    def ForwardAsICal(self, MailFormat=None):
        arguments = com_arguments([unwrap(a) for a in [MailFormat]])
        return MailItem(self.com_object.ForwardAsICal(*arguments))

    # Lower case alias for ForwardAsICal
    def forwardasical(self, MailFormat=None):
        arguments = [MailFormat]
        return self.ForwardAsICal(*arguments)

    def SaveAsICal(self, Path=None):
        arguments = com_arguments([unwrap(a) for a in [Path]])
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

    @property
    def autoformatrules(self):
        """Lower case alias for AutoFormatRules"""
        return self.AutoFormatRules

    @property
    def BoldDatesWithItems(self):
        return CalendarView(self.com_object.BoldDatesWithItems)

    @BoldDatesWithItems.setter
    def BoldDatesWithItems(self, value):
        self.com_object.BoldDatesWithItems = value

    @property
    def bolddateswithitems(self):
        """Lower case alias for BoldDatesWithItems"""
        return self.BoldDatesWithItems

    @bolddateswithitems.setter
    def bolddateswithitems(self, value):
        """Lower case alias for BoldDatesWithItems.setter"""
        self.BoldDatesWithItems = value

    @property
    def BoldSubjects(self):
        return CalendarView(self.com_object.BoldSubjects)

    @BoldSubjects.setter
    def BoldSubjects(self, value):
        self.com_object.BoldSubjects = value

    @property
    def boldsubjects(self):
        """Lower case alias for BoldSubjects"""
        return self.BoldSubjects

    @boldsubjects.setter
    def boldsubjects(self, value):
        """Lower case alias for BoldSubjects.setter"""
        self.BoldSubjects = value

    @property
    def CalendarViewMode(self):
        return OlCalendarViewMode(self.com_object.CalendarViewMode)

    @CalendarViewMode.setter
    def CalendarViewMode(self, value):
        self.com_object.CalendarViewMode = value

    @property
    def calendarviewmode(self):
        """Lower case alias for CalendarViewMode"""
        return self.CalendarViewMode

    @calendarviewmode.setter
    def calendarviewmode(self, value):
        """Lower case alias for CalendarViewMode.setter"""
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

    @property
    def daysinmultidaymode(self):
        """Lower case alias for DaysInMultiDayMode"""
        return self.DaysInMultiDayMode

    @daysinmultidaymode.setter
    def daysinmultidaymode(self, value):
        """Lower case alias for DaysInMultiDayMode.setter"""
        self.DaysInMultiDayMode = value

    @property
    def DayWeekTimeScale(self):
        return OlDayWeekTimeScale(self.com_object.DayWeekTimeScale)

    @DayWeekTimeScale.setter
    def DayWeekTimeScale(self, value):
        self.com_object.DayWeekTimeScale = value

    @property
    def dayweektimescale(self):
        """Lower case alias for DayWeekTimeScale"""
        return self.DayWeekTimeScale

    @dayweektimescale.setter
    def dayweektimescale(self, value):
        """Lower case alias for DayWeekTimeScale.setter"""
        self.DayWeekTimeScale = value

    @property
    def DisplayedDates(self):
        return CalendarView(self.com_object.DisplayedDates)

    @property
    def displayeddates(self):
        """Lower case alias for DisplayedDates"""
        return self.DisplayedDates

    @property
    def EndField(self):
        return CalendarView(self.com_object.EndField)

    @EndField.setter
    def EndField(self, value):
        self.com_object.EndField = value

    @property
    def endfield(self):
        """Lower case alias for EndField"""
        return self.EndField

    @endfield.setter
    def endfield(self, value):
        """Lower case alias for EndField.setter"""
        self.EndField = value

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for Filter.setter"""
        self.Filter = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @language.setter
    def language(self, value):
        """Lower case alias for Language.setter"""
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    @property
    def lockuserchanges(self):
        """Lower case alias for LockUserChanges"""
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        """Lower case alias for LockUserChanges.setter"""
        self.LockUserChanges = value

    @property
    def MonthShowEndTime(self):
        return CalendarView(self.com_object.MonthShowEndTime)

    @MonthShowEndTime.setter
    def MonthShowEndTime(self, value):
        self.com_object.MonthShowEndTime = value

    @property
    def monthshowendtime(self):
        """Lower case alias for MonthShowEndTime"""
        return self.MonthShowEndTime

    @monthshowendtime.setter
    def monthshowendtime(self, value):
        """Lower case alias for MonthShowEndTime.setter"""
        self.MonthShowEndTime = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    @property
    def saveoption(self):
        """Lower case alias for SaveOption"""
        return self.SaveOption

    @property
    def SelectedEndTime(self):
        return CalendarView(self.com_object.SelectedEndTime)

    @property
    def selectedendtime(self):
        """Lower case alias for SelectedEndTime"""
        return self.SelectedEndTime

    @property
    def SelectedStartTime(self):
        return CalendarView(self.com_object.SelectedStartTime)

    @property
    def selectedstarttime(self):
        """Lower case alias for SelectedStartTime"""
        return self.SelectedStartTime

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Standard(self):
        return CalendarView(self.com_object.Standard)

    @property
    def standard(self):
        """Lower case alias for Standard"""
        return self.Standard

    @property
    def StartField(self):
        return CalendarView(self.com_object.StartField)

    @StartField.setter
    def StartField(self, value):
        self.com_object.StartField = value

    @property
    def startfield(self):
        """Lower case alias for StartField"""
        return self.StartField

    @startfield.setter
    def startfield(self, value):
        """Lower case alias for StartField.setter"""
        self.StartField = value

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    @property
    def xml(self):
        """Lower case alias for XML"""
        return self.XML

    @xml.setter
    def xml(self, value):
        """Lower case alias for XML.setter"""
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SaveOption]])
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
        arguments = com_arguments([unwrap(a) for a in [Date]])
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

    @property
    def allowincellediting(self):
        """Lower case alias for AllowInCellEditing"""
        return self.AllowInCellEditing

    @allowincellediting.setter
    def allowincellediting(self, value):
        """Lower case alias for AllowInCellEditing.setter"""
        self.AllowInCellEditing = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.com_object.AutoFormatRules)

    @property
    def autoformatrules(self):
        """Lower case alias for AutoFormatRules"""
        return self.AutoFormatRules

    @property
    def BodyFont(self):
        return ViewFont(self.com_object.BodyFont)

    @property
    def bodyfont(self):
        """Lower case alias for BodyFont"""
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

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for Filter.setter"""
        self.Filter = value

    @property
    def HeadingsFont(self):
        return ViewFont(self.com_object.HeadingsFont)

    @property
    def headingsfont(self):
        """Lower case alias for HeadingsFont"""
        return self.HeadingsFont

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @language.setter
    def language(self, value):
        """Lower case alias for Language.setter"""
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    @property
    def lockuserchanges(self):
        """Lower case alias for LockUserChanges"""
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        """Lower case alias for LockUserChanges.setter"""
        self.LockUserChanges = value

    @property
    def MultiLineFieldHeight(self):
        return CardView(self.com_object.MultiLineFieldHeight)

    @MultiLineFieldHeight.setter
    def MultiLineFieldHeight(self, value):
        self.com_object.MultiLineFieldHeight = value

    @property
    def multilinefieldheight(self):
        """Lower case alias for MultiLineFieldHeight"""
        return self.MultiLineFieldHeight

    @multilinefieldheight.setter
    def multilinefieldheight(self, value):
        """Lower case alias for MultiLineFieldHeight.setter"""
        self.MultiLineFieldHeight = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    @property
    def saveoption(self):
        """Lower case alias for SaveOption"""
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ShowEmptyFields(self):
        return CardView(self.com_object.ShowEmptyFields)

    @ShowEmptyFields.setter
    def ShowEmptyFields(self, value):
        self.com_object.ShowEmptyFields = value

    @property
    def showemptyfields(self):
        """Lower case alias for ShowEmptyFields"""
        return self.ShowEmptyFields

    @showemptyfields.setter
    def showemptyfields(self, value):
        """Lower case alias for ShowEmptyFields.setter"""
        self.ShowEmptyFields = value

    @property
    def SortFields(self):
        return OrderFields(self.com_object.SortFields)

    @property
    def sortfields(self):
        """Lower case alias for SortFields"""
        return self.SortFields

    @property
    def Standard(self):
        return CardView(self.com_object.Standard)

    @property
    def standard(self):
        """Lower case alias for Standard"""
        return self.Standard

    @property
    def ViewFields(self):
        return ViewFields(self.com_object.ViewFields)

    @property
    def viewfields(self):
        """Lower case alias for ViewFields"""
        return self.ViewFields

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
        return self.ViewType

    @property
    def Width(self):
        return CardView(self.com_object.Width)

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    @property
    def width(self):
        """Lower case alias for Width"""
        return self.Width

    @width.setter
    def width(self, value):
        """Lower case alias for Width.setter"""
        self.Width = value

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    @property
    def xml(self):
        """Lower case alias for XML"""
        return self.XML

    @xml.setter
    def xml(self, value):
        """Lower case alias for XML.setter"""
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SaveOption]])
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
        arguments = com_arguments([unwrap(a) for a in [Date]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None, Color=None, ShortcutKey=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Color, ShortcutKey]])
        return Category(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Color=None, ShortcutKey=None):
        arguments = [Name, Color, ShortcutKey]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Category(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def categorybordercolor(self):
        """Lower case alias for CategoryBorderColor"""
        return self.CategoryBorderColor

    @property
    def CategoryGradientBottomColor(self):
        return Category(self.com_object.CategoryGradientBottomColor)

    @property
    def categorygradientbottomcolor(self):
        """Lower case alias for CategoryGradientBottomColor"""
        return self.CategoryGradientBottomColor

    @property
    def CategoryGradientTopColor(self):
        return Category(self.com_object.CategoryGradientTopColor)

    @property
    def categorygradienttopcolor(self):
        """Lower case alias for CategoryGradientTopColor"""
        return self.CategoryGradientTopColor

    @property
    def CategoryID(self):
        return Category(self.com_object.CategoryID)

    @property
    def categoryid(self):
        """Lower case alias for CategoryID"""
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

    @property
    def color(self):
        """Lower case alias for Color"""
        return self.Color

    @color.setter
    def color(self, value):
        """Lower case alias for Color.setter"""
        self.Color = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ShortcutKey(self):
        return OlCategoryShortcutKey(self.com_object.ShortcutKey)

    @ShortcutKey.setter
    def ShortcutKey(self, value):
        self.com_object.ShortcutKey = value

    @property
    def shortcutkey(self):
        """Lower case alias for ShortcutKey"""
        return self.ShortcutKey

    @shortcutkey.setter
    def shortcutkey(self, value):
        """Lower case alias for ShortcutKey.setter"""
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

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ConditionType(self):
        return OlRuleConditionType(self.com_object.ConditionType)

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class checkbox:

    def __init__(self, checkbox=None):
        self.com_object= checkbox


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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return Column(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def align(self):
        """Lower case alias for Align"""
        return self.Align

    @align.setter
    def align(self, value):
        """Lower case alias for Align.setter"""
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

    @property
    def fieldformat(self):
        """Lower case alias for FieldFormat"""
        return self.FieldFormat

    @fieldformat.setter
    def fieldformat(self, value):
        """Lower case alias for FieldFormat.setter"""
        self.FieldFormat = value

    @property
    def FieldType(self):
        return OlUserPropertyType(self.com_object.FieldType)

    @property
    def fieldtype(self):
        """Lower case alias for FieldType"""
        return self.FieldType

    @property
    def Label(self):
        return ColumnFormat(self.com_object.Label)

    @Label.setter
    def Label(self, value):
        self.com_object.Label = value

    @property
    def label(self):
        """Lower case alias for Label"""
        return self.Label

    @label.setter
    def label(self, value):
        """Lower case alias for Label.setter"""
        self.Label = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Width(self):
        return self.com_object.Width

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    @property
    def width(self):
        """Lower case alias for Width"""
        return self.Width

    @width.setter
    def width(self, value):
        """Lower case alias for Width.setter"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return Columns(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Name]])
        return Column(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None):
        arguments = [Name]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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


class combobox:

    def __init__(self, combobox=None):
        self.com_object= combobox


class commandbutton:

    def __init__(self, commandbutton=None):
        self.com_object= commandbutton


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

    @property
    def item(self):
        """Lower case alias for Item"""
        return self.Item

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Type(self):
        return OlObjectClass(self.com_object.Type)

    @property
    def type(self):
        """Lower case alias for Type"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def account(self):
        """Lower case alias for Account"""
        return self.Account

    @account.setter
    def account(self, value):
        """Lower case alias for Account.setter"""
        self.Account = value

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Anniversary(self):
        return self.com_object.Anniversary

    @Anniversary.setter
    def Anniversary(self, value):
        self.com_object.Anniversary = value

    @property
    def anniversary(self):
        """Lower case alias for Anniversary"""
        return self.Anniversary

    @anniversary.setter
    def anniversary(self, value):
        """Lower case alias for Anniversary.setter"""
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

    @property
    def assistantname(self):
        """Lower case alias for AssistantName"""
        return self.AssistantName

    @assistantname.setter
    def assistantname(self, value):
        """Lower case alias for AssistantName.setter"""
        self.AssistantName = value

    @property
    def AssistantTelephoneNumber(self):
        return self.com_object.AssistantTelephoneNumber

    @AssistantTelephoneNumber.setter
    def AssistantTelephoneNumber(self, value):
        self.com_object.AssistantTelephoneNumber = value

    @property
    def assistanttelephonenumber(self):
        """Lower case alias for AssistantTelephoneNumber"""
        return self.AssistantTelephoneNumber

    @assistanttelephonenumber.setter
    def assistanttelephonenumber(self, value):
        """Lower case alias for AssistantTelephoneNumber.setter"""
        self.AssistantTelephoneNumber = value

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Birthday(self):
        return self.com_object.Birthday

    @Birthday.setter
    def Birthday(self, value):
        self.com_object.Birthday = value

    @property
    def birthday(self):
        """Lower case alias for Birthday"""
        return self.Birthday

    @birthday.setter
    def birthday(self, value):
        """Lower case alias for Birthday.setter"""
        self.Birthday = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Business2TelephoneNumber(self):
        return self.com_object.Business2TelephoneNumber

    @Business2TelephoneNumber.setter
    def Business2TelephoneNumber(self, value):
        self.com_object.Business2TelephoneNumber = value

    @property
    def business2telephonenumber(self):
        """Lower case alias for Business2TelephoneNumber"""
        return self.Business2TelephoneNumber

    @business2telephonenumber.setter
    def business2telephonenumber(self, value):
        """Lower case alias for Business2TelephoneNumber.setter"""
        self.Business2TelephoneNumber = value

    @property
    def BusinessAddress(self):
        return self.com_object.BusinessAddress

    @BusinessAddress.setter
    def BusinessAddress(self, value):
        self.com_object.BusinessAddress = value

    @property
    def businessaddress(self):
        """Lower case alias for BusinessAddress"""
        return self.BusinessAddress

    @businessaddress.setter
    def businessaddress(self, value):
        """Lower case alias for BusinessAddress.setter"""
        self.BusinessAddress = value

    @property
    def BusinessAddressCity(self):
        return self.com_object.BusinessAddressCity

    @BusinessAddressCity.setter
    def BusinessAddressCity(self, value):
        self.com_object.BusinessAddressCity = value

    @property
    def businessaddresscity(self):
        """Lower case alias for BusinessAddressCity"""
        return self.BusinessAddressCity

    @businessaddresscity.setter
    def businessaddresscity(self, value):
        """Lower case alias for BusinessAddressCity.setter"""
        self.BusinessAddressCity = value

    @property
    def BusinessAddressCountry(self):
        return self.com_object.BusinessAddressCountry

    @BusinessAddressCountry.setter
    def BusinessAddressCountry(self, value):
        self.com_object.BusinessAddressCountry = value

    @property
    def businessaddresscountry(self):
        """Lower case alias for BusinessAddressCountry"""
        return self.BusinessAddressCountry

    @businessaddresscountry.setter
    def businessaddresscountry(self, value):
        """Lower case alias for BusinessAddressCountry.setter"""
        self.BusinessAddressCountry = value

    @property
    def BusinessAddressPostalCode(self):
        return self.com_object.BusinessAddressPostalCode

    @BusinessAddressPostalCode.setter
    def BusinessAddressPostalCode(self, value):
        self.com_object.BusinessAddressPostalCode = value

    @property
    def businessaddresspostalcode(self):
        """Lower case alias for BusinessAddressPostalCode"""
        return self.BusinessAddressPostalCode

    @businessaddresspostalcode.setter
    def businessaddresspostalcode(self, value):
        """Lower case alias for BusinessAddressPostalCode.setter"""
        self.BusinessAddressPostalCode = value

    @property
    def BusinessAddressPostOfficeBox(self):
        return self.com_object.BusinessAddressPostOfficeBox

    @BusinessAddressPostOfficeBox.setter
    def BusinessAddressPostOfficeBox(self, value):
        self.com_object.BusinessAddressPostOfficeBox = value

    @property
    def businessaddresspostofficebox(self):
        """Lower case alias for BusinessAddressPostOfficeBox"""
        return self.BusinessAddressPostOfficeBox

    @businessaddresspostofficebox.setter
    def businessaddresspostofficebox(self, value):
        """Lower case alias for BusinessAddressPostOfficeBox.setter"""
        self.BusinessAddressPostOfficeBox = value

    @property
    def BusinessAddressState(self):
        return self.com_object.BusinessAddressState

    @BusinessAddressState.setter
    def BusinessAddressState(self, value):
        self.com_object.BusinessAddressState = value

    @property
    def businessaddressstate(self):
        """Lower case alias for BusinessAddressState"""
        return self.BusinessAddressState

    @businessaddressstate.setter
    def businessaddressstate(self, value):
        """Lower case alias for BusinessAddressState.setter"""
        self.BusinessAddressState = value

    @property
    def BusinessAddressStreet(self):
        return self.com_object.BusinessAddressStreet

    @BusinessAddressStreet.setter
    def BusinessAddressStreet(self, value):
        self.com_object.BusinessAddressStreet = value

    @property
    def businessaddressstreet(self):
        """Lower case alias for BusinessAddressStreet"""
        return self.BusinessAddressStreet

    @businessaddressstreet.setter
    def businessaddressstreet(self, value):
        """Lower case alias for BusinessAddressStreet.setter"""
        self.BusinessAddressStreet = value

    @property
    def BusinessCardLayoutXml(self):
        return self.com_object.BusinessCardLayoutXml

    @BusinessCardLayoutXml.setter
    def BusinessCardLayoutXml(self, value):
        self.com_object.BusinessCardLayoutXml = value

    @property
    def businesscardlayoutxml(self):
        """Lower case alias for BusinessCardLayoutXml"""
        return self.BusinessCardLayoutXml

    @businesscardlayoutxml.setter
    def businesscardlayoutxml(self, value):
        """Lower case alias for BusinessCardLayoutXml.setter"""
        self.BusinessCardLayoutXml = value

    @property
    def BusinessCardType(self):
        return OlBusinessCardType(self.com_object.BusinessCardType)

    @property
    def businesscardtype(self):
        """Lower case alias for BusinessCardType"""
        return self.BusinessCardType

    @property
    def BusinessFaxNumber(self):
        return self.com_object.BusinessFaxNumber

    @BusinessFaxNumber.setter
    def BusinessFaxNumber(self, value):
        self.com_object.BusinessFaxNumber = value

    @property
    def businessfaxnumber(self):
        """Lower case alias for BusinessFaxNumber"""
        return self.BusinessFaxNumber

    @businessfaxnumber.setter
    def businessfaxnumber(self, value):
        """Lower case alias for BusinessFaxNumber.setter"""
        self.BusinessFaxNumber = value

    @property
    def BusinessHomePage(self):
        return self.com_object.BusinessHomePage

    @BusinessHomePage.setter
    def BusinessHomePage(self, value):
        self.com_object.BusinessHomePage = value

    @property
    def businesshomepage(self):
        """Lower case alias for BusinessHomePage"""
        return self.BusinessHomePage

    @businesshomepage.setter
    def businesshomepage(self, value):
        """Lower case alias for BusinessHomePage.setter"""
        self.BusinessHomePage = value

    @property
    def BusinessTelephoneNumber(self):
        return self.com_object.BusinessTelephoneNumber

    @BusinessTelephoneNumber.setter
    def BusinessTelephoneNumber(self, value):
        self.com_object.BusinessTelephoneNumber = value

    @property
    def businesstelephonenumber(self):
        """Lower case alias for BusinessTelephoneNumber"""
        return self.BusinessTelephoneNumber

    @businesstelephonenumber.setter
    def businesstelephonenumber(self, value):
        """Lower case alias for BusinessTelephoneNumber.setter"""
        self.BusinessTelephoneNumber = value

    @property
    def CallbackTelephoneNumber(self):
        return self.com_object.CallbackTelephoneNumber

    @CallbackTelephoneNumber.setter
    def CallbackTelephoneNumber(self, value):
        self.com_object.CallbackTelephoneNumber = value

    @property
    def callbacktelephonenumber(self):
        """Lower case alias for CallbackTelephoneNumber"""
        return self.CallbackTelephoneNumber

    @callbacktelephonenumber.setter
    def callbacktelephonenumber(self, value):
        """Lower case alias for CallbackTelephoneNumber.setter"""
        self.CallbackTelephoneNumber = value

    @property
    def CarTelephoneNumber(self):
        return self.com_object.CarTelephoneNumber

    @CarTelephoneNumber.setter
    def CarTelephoneNumber(self, value):
        self.com_object.CarTelephoneNumber = value

    @property
    def cartelephonenumber(self):
        """Lower case alias for CarTelephoneNumber"""
        return self.CarTelephoneNumber

    @cartelephonenumber.setter
    def cartelephonenumber(self, value):
        """Lower case alias for CarTelephoneNumber.setter"""
        self.CarTelephoneNumber = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
        self.Categories = value

    @property
    def Children(self):
        return self.com_object.Children

    @Children.setter
    def Children(self, value):
        self.com_object.Children = value

    @property
    def children(self):
        """Lower case alias for Children"""
        return self.Children

    @children.setter
    def children(self, value):
        """Lower case alias for Children.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def CompanyAndFullName(self):
        return self.com_object.CompanyAndFullName

    @property
    def companyandfullname(self):
        """Lower case alias for CompanyAndFullName"""
        return self.CompanyAndFullName

    @property
    def CompanyLastFirstNoSpace(self):
        return self.com_object.CompanyLastFirstNoSpace

    @property
    def companylastfirstnospace(self):
        """Lower case alias for CompanyLastFirstNoSpace"""
        return self.CompanyLastFirstNoSpace

    @property
    def CompanyLastFirstSpaceOnly(self):
        return self.com_object.CompanyLastFirstSpaceOnly

    @property
    def companylastfirstspaceonly(self):
        """Lower case alias for CompanyLastFirstSpaceOnly"""
        return self.CompanyLastFirstSpaceOnly

    @property
    def CompanyMainTelephoneNumber(self):
        return self.com_object.CompanyMainTelephoneNumber

    @CompanyMainTelephoneNumber.setter
    def CompanyMainTelephoneNumber(self, value):
        self.com_object.CompanyMainTelephoneNumber = value

    @property
    def companymaintelephonenumber(self):
        """Lower case alias for CompanyMainTelephoneNumber"""
        return self.CompanyMainTelephoneNumber

    @companymaintelephonenumber.setter
    def companymaintelephonenumber(self, value):
        """Lower case alias for CompanyMainTelephoneNumber.setter"""
        self.CompanyMainTelephoneNumber = value

    @property
    def CompanyName(self):
        return self.com_object.CompanyName

    @CompanyName.setter
    def CompanyName(self, value):
        self.com_object.CompanyName = value

    @property
    def companyname(self):
        """Lower case alias for CompanyName"""
        return self.CompanyName

    @companyname.setter
    def companyname(self, value):
        """Lower case alias for CompanyName.setter"""
        self.CompanyName = value

    @property
    def ComputerNetworkName(self):
        return self.com_object.ComputerNetworkName

    @ComputerNetworkName.setter
    def ComputerNetworkName(self, value):
        self.com_object.ComputerNetworkName = value

    @property
    def computernetworkname(self):
        """Lower case alias for ComputerNetworkName"""
        return self.ComputerNetworkName

    @computernetworkname.setter
    def computernetworkname(self, value):
        """Lower case alias for ComputerNetworkName.setter"""
        self.ComputerNetworkName = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def CustomerID(self):
        return self.com_object.CustomerID

    @CustomerID.setter
    def CustomerID(self, value):
        self.com_object.CustomerID = value

    @property
    def customerid(self):
        """Lower case alias for CustomerID"""
        return self.CustomerID

    @customerid.setter
    def customerid(self, value):
        """Lower case alias for CustomerID.setter"""
        self.CustomerID = value

    @property
    def Department(self):
        return self.com_object.Department

    @Department.setter
    def Department(self, value):
        self.com_object.Department = value

    @property
    def department(self):
        """Lower case alias for Department"""
        return self.Department

    @department.setter
    def department(self, value):
        """Lower case alias for Department.setter"""
        self.Department = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def Email1Address(self):
        return self.com_object.Email1Address

    @Email1Address.setter
    def Email1Address(self, value):
        self.com_object.Email1Address = value

    @property
    def email1address(self):
        """Lower case alias for Email1Address"""
        return self.Email1Address

    @email1address.setter
    def email1address(self, value):
        """Lower case alias for Email1Address.setter"""
        self.Email1Address = value

    @property
    def Email1AddressType(self):
        return self.com_object.Email1AddressType

    @Email1AddressType.setter
    def Email1AddressType(self, value):
        self.com_object.Email1AddressType = value

    @property
    def email1addresstype(self):
        """Lower case alias for Email1AddressType"""
        return self.Email1AddressType

    @email1addresstype.setter
    def email1addresstype(self, value):
        """Lower case alias for Email1AddressType.setter"""
        self.Email1AddressType = value

    @property
    def Email1DisplayName(self):
        return self.com_object.Email1DisplayName

    @Email1DisplayName.setter
    def Email1DisplayName(self, value):
        self.com_object.Email1DisplayName = value

    @property
    def email1displayname(self):
        """Lower case alias for Email1DisplayName"""
        return self.Email1DisplayName

    @email1displayname.setter
    def email1displayname(self, value):
        """Lower case alias for Email1DisplayName.setter"""
        self.Email1DisplayName = value

    @property
    def Email1EntryID(self):
        return self.com_object.Email1EntryID

    @property
    def email1entryid(self):
        """Lower case alias for Email1EntryID"""
        return self.Email1EntryID

    @property
    def Email2Address(self):
        return self.com_object.Email2Address

    @Email2Address.setter
    def Email2Address(self, value):
        self.com_object.Email2Address = value

    @property
    def email2address(self):
        """Lower case alias for Email2Address"""
        return self.Email2Address

    @email2address.setter
    def email2address(self, value):
        """Lower case alias for Email2Address.setter"""
        self.Email2Address = value

    @property
    def Email2AddressType(self):
        return self.com_object.Email2AddressType

    @Email2AddressType.setter
    def Email2AddressType(self, value):
        self.com_object.Email2AddressType = value

    @property
    def email2addresstype(self):
        """Lower case alias for Email2AddressType"""
        return self.Email2AddressType

    @email2addresstype.setter
    def email2addresstype(self, value):
        """Lower case alias for Email2AddressType.setter"""
        self.Email2AddressType = value

    @property
    def Email2DisplayName(self):
        return self.com_object.Email2DisplayName

    @Email2DisplayName.setter
    def Email2DisplayName(self, value):
        self.com_object.Email2DisplayName = value

    @property
    def email2displayname(self):
        """Lower case alias for Email2DisplayName"""
        return self.Email2DisplayName

    @email2displayname.setter
    def email2displayname(self, value):
        """Lower case alias for Email2DisplayName.setter"""
        self.Email2DisplayName = value

    @property
    def Email2EntryID(self):
        return self.com_object.Email2EntryID

    @property
    def email2entryid(self):
        """Lower case alias for Email2EntryID"""
        return self.Email2EntryID

    @property
    def Email3Address(self):
        return self.com_object.Email3Address

    @Email3Address.setter
    def Email3Address(self, value):
        self.com_object.Email3Address = value

    @property
    def email3address(self):
        """Lower case alias for Email3Address"""
        return self.Email3Address

    @email3address.setter
    def email3address(self, value):
        """Lower case alias for Email3Address.setter"""
        self.Email3Address = value

    @property
    def Email3AddressType(self):
        return self.com_object.Email3AddressType

    @Email3AddressType.setter
    def Email3AddressType(self, value):
        self.com_object.Email3AddressType = value

    @property
    def email3addresstype(self):
        """Lower case alias for Email3AddressType"""
        return self.Email3AddressType

    @email3addresstype.setter
    def email3addresstype(self, value):
        """Lower case alias for Email3AddressType.setter"""
        self.Email3AddressType = value

    @property
    def Email3DisplayName(self):
        return self.com_object.Email3DisplayName

    @Email3DisplayName.setter
    def Email3DisplayName(self, value):
        self.com_object.Email3DisplayName = value

    @property
    def email3displayname(self):
        """Lower case alias for Email3DisplayName"""
        return self.Email3DisplayName

    @email3displayname.setter
    def email3displayname(self, value):
        """Lower case alias for Email3DisplayName.setter"""
        self.Email3DisplayName = value

    @property
    def Email3EntryID(self):
        return self.com_object.Email3EntryID

    @property
    def email3entryid(self):
        """Lower case alias for Email3EntryID"""
        return self.Email3EntryID

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FileAs(self):
        return self.com_object.FileAs

    @FileAs.setter
    def FileAs(self, value):
        self.com_object.FileAs = value

    @property
    def fileas(self):
        """Lower case alias for FileAs"""
        return self.FileAs

    @fileas.setter
    def fileas(self, value):
        """Lower case alias for FileAs.setter"""
        self.FileAs = value

    @property
    def FirstName(self):
        return self.com_object.FirstName

    @FirstName.setter
    def FirstName(self, value):
        self.com_object.FirstName = value

    @property
    def firstname(self):
        """Lower case alias for FirstName"""
        return self.FirstName

    @firstname.setter
    def firstname(self, value):
        """Lower case alias for FirstName.setter"""
        self.FirstName = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def FTPSite(self):
        return self.com_object.FTPSite

    @FTPSite.setter
    def FTPSite(self, value):
        self.com_object.FTPSite = value

    @property
    def ftpsite(self):
        """Lower case alias for FTPSite"""
        return self.FTPSite

    @ftpsite.setter
    def ftpsite(self, value):
        """Lower case alias for FTPSite.setter"""
        self.FTPSite = value

    @property
    def FullName(self):
        return self.com_object.FullName

    @FullName.setter
    def FullName(self, value):
        self.com_object.FullName = value

    @property
    def fullname(self):
        """Lower case alias for FullName"""
        return self.FullName

    @fullname.setter
    def fullname(self, value):
        """Lower case alias for FullName.setter"""
        self.FullName = value

    @property
    def FullNameAndCompany(self):
        return self.com_object.FullNameAndCompany

    @property
    def fullnameandcompany(self):
        """Lower case alias for FullNameAndCompany"""
        return self.FullNameAndCompany

    @property
    def Gender(self):
        return OlGender(self.com_object.Gender)

    @Gender.setter
    def Gender(self, value):
        self.com_object.Gender = value

    @property
    def gender(self):
        """Lower case alias for Gender"""
        return self.Gender

    @gender.setter
    def gender(self, value):
        """Lower case alias for Gender.setter"""
        self.Gender = value

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def GovernmentIDNumber(self):
        return self.com_object.GovernmentIDNumber

    @GovernmentIDNumber.setter
    def GovernmentIDNumber(self, value):
        self.com_object.GovernmentIDNumber = value

    @property
    def governmentidnumber(self):
        """Lower case alias for GovernmentIDNumber"""
        return self.GovernmentIDNumber

    @governmentidnumber.setter
    def governmentidnumber(self, value):
        """Lower case alias for GovernmentIDNumber.setter"""
        self.GovernmentIDNumber = value

    @property
    def HasPicture(self):
        return self.com_object.HasPicture

    @property
    def haspicture(self):
        """Lower case alias for HasPicture"""
        return self.HasPicture

    @property
    def Hobby(self):
        return self.com_object.Hobby

    @Hobby.setter
    def Hobby(self, value):
        self.com_object.Hobby = value

    @property
    def hobby(self):
        """Lower case alias for Hobby"""
        return self.Hobby

    @hobby.setter
    def hobby(self, value):
        """Lower case alias for Hobby.setter"""
        self.Hobby = value

    @property
    def Home2TelephoneNumber(self):
        return self.com_object.Home2TelephoneNumber

    @Home2TelephoneNumber.setter
    def Home2TelephoneNumber(self, value):
        self.com_object.Home2TelephoneNumber = value

    @property
    def home2telephonenumber(self):
        """Lower case alias for Home2TelephoneNumber"""
        return self.Home2TelephoneNumber

    @home2telephonenumber.setter
    def home2telephonenumber(self, value):
        """Lower case alias for Home2TelephoneNumber.setter"""
        self.Home2TelephoneNumber = value

    @property
    def HomeAddress(self):
        return self.com_object.HomeAddress

    @HomeAddress.setter
    def HomeAddress(self, value):
        self.com_object.HomeAddress = value

    @property
    def homeaddress(self):
        """Lower case alias for HomeAddress"""
        return self.HomeAddress

    @homeaddress.setter
    def homeaddress(self, value):
        """Lower case alias for HomeAddress.setter"""
        self.HomeAddress = value

    @property
    def HomeAddressCity(self):
        return self.com_object.HomeAddressCity

    @HomeAddressCity.setter
    def HomeAddressCity(self, value):
        self.com_object.HomeAddressCity = value

    @property
    def homeaddresscity(self):
        """Lower case alias for HomeAddressCity"""
        return self.HomeAddressCity

    @homeaddresscity.setter
    def homeaddresscity(self, value):
        """Lower case alias for HomeAddressCity.setter"""
        self.HomeAddressCity = value

    @property
    def HomeAddressCountry(self):
        return self.com_object.HomeAddressCountry

    @HomeAddressCountry.setter
    def HomeAddressCountry(self, value):
        self.com_object.HomeAddressCountry = value

    @property
    def homeaddresscountry(self):
        """Lower case alias for HomeAddressCountry"""
        return self.HomeAddressCountry

    @homeaddresscountry.setter
    def homeaddresscountry(self, value):
        """Lower case alias for HomeAddressCountry.setter"""
        self.HomeAddressCountry = value

    @property
    def HomeAddressPostalCode(self):
        return self.com_object.HomeAddressPostalCode

    @HomeAddressPostalCode.setter
    def HomeAddressPostalCode(self, value):
        self.com_object.HomeAddressPostalCode = value

    @property
    def homeaddresspostalcode(self):
        """Lower case alias for HomeAddressPostalCode"""
        return self.HomeAddressPostalCode

    @homeaddresspostalcode.setter
    def homeaddresspostalcode(self, value):
        """Lower case alias for HomeAddressPostalCode.setter"""
        self.HomeAddressPostalCode = value

    @property
    def HomeAddressPostOfficeBox(self):
        return self.com_object.HomeAddressPostOfficeBox

    @HomeAddressPostOfficeBox.setter
    def HomeAddressPostOfficeBox(self, value):
        self.com_object.HomeAddressPostOfficeBox = value

    @property
    def homeaddresspostofficebox(self):
        """Lower case alias for HomeAddressPostOfficeBox"""
        return self.HomeAddressPostOfficeBox

    @homeaddresspostofficebox.setter
    def homeaddresspostofficebox(self, value):
        """Lower case alias for HomeAddressPostOfficeBox.setter"""
        self.HomeAddressPostOfficeBox = value

    @property
    def HomeAddressState(self):
        return self.com_object.HomeAddressState

    @HomeAddressState.setter
    def HomeAddressState(self, value):
        self.com_object.HomeAddressState = value

    @property
    def homeaddressstate(self):
        """Lower case alias for HomeAddressState"""
        return self.HomeAddressState

    @homeaddressstate.setter
    def homeaddressstate(self, value):
        """Lower case alias for HomeAddressState.setter"""
        self.HomeAddressState = value

    @property
    def HomeAddressStreet(self):
        return self.com_object.HomeAddressStreet

    @HomeAddressStreet.setter
    def HomeAddressStreet(self, value):
        self.com_object.HomeAddressStreet = value

    @property
    def homeaddressstreet(self):
        """Lower case alias for HomeAddressStreet"""
        return self.HomeAddressStreet

    @homeaddressstreet.setter
    def homeaddressstreet(self, value):
        """Lower case alias for HomeAddressStreet.setter"""
        self.HomeAddressStreet = value

    @property
    def HomeFaxNumber(self):
        return self.com_object.HomeFaxNumber

    @HomeFaxNumber.setter
    def HomeFaxNumber(self, value):
        self.com_object.HomeFaxNumber = value

    @property
    def homefaxnumber(self):
        """Lower case alias for HomeFaxNumber"""
        return self.HomeFaxNumber

    @homefaxnumber.setter
    def homefaxnumber(self, value):
        """Lower case alias for HomeFaxNumber.setter"""
        self.HomeFaxNumber = value

    @property
    def HomeTelephoneNumber(self):
        return self.com_object.HomeTelephoneNumber

    @HomeTelephoneNumber.setter
    def HomeTelephoneNumber(self, value):
        self.com_object.HomeTelephoneNumber = value

    @property
    def hometelephonenumber(self):
        """Lower case alias for HomeTelephoneNumber"""
        return self.HomeTelephoneNumber

    @hometelephonenumber.setter
    def hometelephonenumber(self, value):
        """Lower case alias for HomeTelephoneNumber.setter"""
        self.HomeTelephoneNumber = value

    @property
    def IMAddress(self):
        return self.com_object.IMAddress

    @IMAddress.setter
    def IMAddress(self, value):
        self.com_object.IMAddress = value

    @property
    def imaddress(self):
        """Lower case alias for IMAddress"""
        return self.IMAddress

    @imaddress.setter
    def imaddress(self, value):
        """Lower case alias for IMAddress.setter"""
        self.IMAddress = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def Initials(self):
        return self.com_object.Initials

    @Initials.setter
    def Initials(self, value):
        self.com_object.Initials = value

    @property
    def initials(self):
        """Lower case alias for Initials"""
        return self.Initials

    @initials.setter
    def initials(self, value):
        """Lower case alias for Initials.setter"""
        self.Initials = value

    @property
    def InternetFreeBusyAddress(self):
        return self.com_object.InternetFreeBusyAddress

    @InternetFreeBusyAddress.setter
    def InternetFreeBusyAddress(self, value):
        self.com_object.InternetFreeBusyAddress = value

    @property
    def internetfreebusyaddress(self):
        """Lower case alias for InternetFreeBusyAddress"""
        return self.InternetFreeBusyAddress

    @internetfreebusyaddress.setter
    def internetfreebusyaddress(self, value):
        """Lower case alias for InternetFreeBusyAddress.setter"""
        self.InternetFreeBusyAddress = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ISDNNumber(self):
        return self.com_object.ISDNNumber

    @ISDNNumber.setter
    def ISDNNumber(self, value):
        self.com_object.ISDNNumber = value

    @property
    def isdnnumber(self):
        """Lower case alias for ISDNNumber"""
        return self.ISDNNumber

    @isdnnumber.setter
    def isdnnumber(self, value):
        """Lower case alias for ISDNNumber.setter"""
        self.ISDNNumber = value

    @property
    def IsMarkedAsTask(self):
        return ContactItem(self.com_object.IsMarkedAsTask)

    @property
    def ismarkedastask(self):
        """Lower case alias for IsMarkedAsTask"""
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def JobTitle(self):
        return self.com_object.JobTitle

    @JobTitle.setter
    def JobTitle(self, value):
        self.com_object.JobTitle = value

    @property
    def jobtitle(self):
        """Lower case alias for JobTitle"""
        return self.JobTitle

    @jobtitle.setter
    def jobtitle(self, value):
        """Lower case alias for JobTitle.setter"""
        self.JobTitle = value

    @property
    def Journal(self):
        return self.com_object.Journal

    @Journal.setter
    def Journal(self, value):
        self.com_object.Journal = value

    @property
    def journal(self):
        """Lower case alias for Journal"""
        return self.Journal

    @journal.setter
    def journal(self, value):
        """Lower case alias for Journal.setter"""
        self.Journal = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @language.setter
    def language(self, value):
        """Lower case alias for Language.setter"""
        self.Language = value

    @property
    def LastFirstAndSuffix(self):
        return self.com_object.LastFirstAndSuffix

    @property
    def lastfirstandsuffix(self):
        """Lower case alias for LastFirstAndSuffix"""
        return self.LastFirstAndSuffix

    @property
    def LastFirstNoSpace(self):
        return self.com_object.LastFirstNoSpace

    @property
    def lastfirstnospace(self):
        """Lower case alias for LastFirstNoSpace"""
        return self.LastFirstNoSpace

    @property
    def LastFirstNoSpaceAndSuffix(self):
        return self.com_object.LastFirstNoSpaceAndSuffix

    @property
    def lastfirstnospaceandsuffix(self):
        """Lower case alias for LastFirstNoSpaceAndSuffix"""
        return self.LastFirstNoSpaceAndSuffix

    @property
    def LastFirstNoSpaceCompany(self):
        return self.com_object.LastFirstNoSpaceCompany

    @property
    def lastfirstnospacecompany(self):
        """Lower case alias for LastFirstNoSpaceCompany"""
        return self.LastFirstNoSpaceCompany

    @property
    def LastFirstSpaceOnly(self):
        return self.com_object.LastFirstSpaceOnly

    @property
    def lastfirstspaceonly(self):
        """Lower case alias for LastFirstSpaceOnly"""
        return self.LastFirstSpaceOnly

    @property
    def LastFirstSpaceOnlyCompany(self):
        return self.com_object.LastFirstSpaceOnlyCompany

    @property
    def lastfirstspaceonlycompany(self):
        """Lower case alias for LastFirstSpaceOnlyCompany"""
        return self.LastFirstSpaceOnlyCompany

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def LastName(self):
        return self.com_object.LastName

    @LastName.setter
    def LastName(self, value):
        self.com_object.LastName = value

    @property
    def lastname(self):
        """Lower case alias for LastName"""
        return self.LastName

    @lastname.setter
    def lastname(self, value):
        """Lower case alias for LastName.setter"""
        self.LastName = value

    @property
    def LastNameAndFirstName(self):
        return self.com_object.LastNameAndFirstName

    @property
    def lastnameandfirstname(self):
        """Lower case alias for LastNameAndFirstName"""
        return self.LastNameAndFirstName

    @property
    def MailingAddress(self):
        return self.com_object.MailingAddress

    @MailingAddress.setter
    def MailingAddress(self, value):
        self.com_object.MailingAddress = value

    @property
    def mailingaddress(self):
        """Lower case alias for MailingAddress"""
        return self.MailingAddress

    @mailingaddress.setter
    def mailingaddress(self, value):
        """Lower case alias for MailingAddress.setter"""
        self.MailingAddress = value

    @property
    def MailingAddressCity(self):
        return self.com_object.MailingAddressCity

    @MailingAddressCity.setter
    def MailingAddressCity(self, value):
        self.com_object.MailingAddressCity = value

    @property
    def mailingaddresscity(self):
        """Lower case alias for MailingAddressCity"""
        return self.MailingAddressCity

    @mailingaddresscity.setter
    def mailingaddresscity(self, value):
        """Lower case alias for MailingAddressCity.setter"""
        self.MailingAddressCity = value

    @property
    def MailingAddressCountry(self):
        return self.com_object.MailingAddressCountry

    @MailingAddressCountry.setter
    def MailingAddressCountry(self, value):
        self.com_object.MailingAddressCountry = value

    @property
    def mailingaddresscountry(self):
        """Lower case alias for MailingAddressCountry"""
        return self.MailingAddressCountry

    @mailingaddresscountry.setter
    def mailingaddresscountry(self, value):
        """Lower case alias for MailingAddressCountry.setter"""
        self.MailingAddressCountry = value

    @property
    def MailingAddressPostalCode(self):
        return self.com_object.MailingAddressPostalCode

    @MailingAddressPostalCode.setter
    def MailingAddressPostalCode(self, value):
        self.com_object.MailingAddressPostalCode = value

    @property
    def mailingaddresspostalcode(self):
        """Lower case alias for MailingAddressPostalCode"""
        return self.MailingAddressPostalCode

    @mailingaddresspostalcode.setter
    def mailingaddresspostalcode(self, value):
        """Lower case alias for MailingAddressPostalCode.setter"""
        self.MailingAddressPostalCode = value

    @property
    def MailingAddressPostOfficeBox(self):
        return self.com_object.MailingAddressPostOfficeBox

    @MailingAddressPostOfficeBox.setter
    def MailingAddressPostOfficeBox(self, value):
        self.com_object.MailingAddressPostOfficeBox = value

    @property
    def mailingaddresspostofficebox(self):
        """Lower case alias for MailingAddressPostOfficeBox"""
        return self.MailingAddressPostOfficeBox

    @mailingaddresspostofficebox.setter
    def mailingaddresspostofficebox(self, value):
        """Lower case alias for MailingAddressPostOfficeBox.setter"""
        self.MailingAddressPostOfficeBox = value

    @property
    def MailingAddressState(self):
        return self.com_object.MailingAddressState

    @MailingAddressState.setter
    def MailingAddressState(self, value):
        self.com_object.MailingAddressState = value

    @property
    def mailingaddressstate(self):
        """Lower case alias for MailingAddressState"""
        return self.MailingAddressState

    @mailingaddressstate.setter
    def mailingaddressstate(self, value):
        """Lower case alias for MailingAddressState.setter"""
        self.MailingAddressState = value

    @property
    def MailingAddressStreet(self):
        return self.com_object.MailingAddressStreet

    @MailingAddressStreet.setter
    def MailingAddressStreet(self, value):
        self.com_object.MailingAddressStreet = value

    @property
    def mailingaddressstreet(self):
        """Lower case alias for MailingAddressStreet"""
        return self.MailingAddressStreet

    @mailingaddressstreet.setter
    def mailingaddressstreet(self, value):
        """Lower case alias for MailingAddressStreet.setter"""
        self.MailingAddressStreet = value

    @property
    def ManagerName(self):
        return self.com_object.ManagerName

    @ManagerName.setter
    def ManagerName(self, value):
        self.com_object.ManagerName = value

    @property
    def managername(self):
        """Lower case alias for ManagerName"""
        return self.ManagerName

    @managername.setter
    def managername(self, value):
        """Lower case alias for ManagerName.setter"""
        self.ManagerName = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def MiddleName(self):
        return self.com_object.MiddleName

    @MiddleName.setter
    def MiddleName(self, value):
        self.com_object.MiddleName = value

    @property
    def middlename(self):
        """Lower case alias for MiddleName"""
        return self.MiddleName

    @middlename.setter
    def middlename(self, value):
        """Lower case alias for MiddleName.setter"""
        self.MiddleName = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def MobileTelephoneNumber(self):
        return self.com_object.MobileTelephoneNumber

    @MobileTelephoneNumber.setter
    def MobileTelephoneNumber(self, value):
        self.com_object.MobileTelephoneNumber = value

    @property
    def mobiletelephonenumber(self):
        """Lower case alias for MobileTelephoneNumber"""
        return self.MobileTelephoneNumber

    @mobiletelephonenumber.setter
    def mobiletelephonenumber(self, value):
        """Lower case alias for MobileTelephoneNumber.setter"""
        self.MobileTelephoneNumber = value

    @property
    def NetMeetingAlias(self):
        return self.com_object.NetMeetingAlias

    @NetMeetingAlias.setter
    def NetMeetingAlias(self, value):
        self.com_object.NetMeetingAlias = value

    @property
    def netmeetingalias(self):
        """Lower case alias for NetMeetingAlias"""
        return self.NetMeetingAlias

    @netmeetingalias.setter
    def netmeetingalias(self, value):
        """Lower case alias for NetMeetingAlias.setter"""
        self.NetMeetingAlias = value

    @property
    def NetMeetingServer(self):
        return self.com_object.NetMeetingServer

    @NetMeetingServer.setter
    def NetMeetingServer(self, value):
        self.com_object.NetMeetingServer = value

    @property
    def netmeetingserver(self):
        """Lower case alias for NetMeetingServer"""
        return self.NetMeetingServer

    @netmeetingserver.setter
    def netmeetingserver(self, value):
        """Lower case alias for NetMeetingServer.setter"""
        self.NetMeetingServer = value

    @property
    def NickName(self):
        return self.com_object.NickName

    @NickName.setter
    def NickName(self, value):
        self.com_object.NickName = value

    @property
    def nickname(self):
        """Lower case alias for NickName"""
        return self.NickName

    @nickname.setter
    def nickname(self, value):
        """Lower case alias for NickName.setter"""
        self.NickName = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OfficeLocation(self):
        return self.com_object.OfficeLocation

    @OfficeLocation.setter
    def OfficeLocation(self, value):
        self.com_object.OfficeLocation = value

    @property
    def officelocation(self):
        """Lower case alias for OfficeLocation"""
        return self.OfficeLocation

    @officelocation.setter
    def officelocation(self, value):
        """Lower case alias for OfficeLocation.setter"""
        self.OfficeLocation = value

    @property
    def OrganizationalIDNumber(self):
        return self.com_object.OrganizationalIDNumber

    @OrganizationalIDNumber.setter
    def OrganizationalIDNumber(self, value):
        self.com_object.OrganizationalIDNumber = value

    @property
    def organizationalidnumber(self):
        """Lower case alias for OrganizationalIDNumber"""
        return self.OrganizationalIDNumber

    @organizationalidnumber.setter
    def organizationalidnumber(self, value):
        """Lower case alias for OrganizationalIDNumber.setter"""
        self.OrganizationalIDNumber = value

    @property
    def OtherAddress(self):
        return self.com_object.OtherAddress

    @OtherAddress.setter
    def OtherAddress(self, value):
        self.com_object.OtherAddress = value

    @property
    def otheraddress(self):
        """Lower case alias for OtherAddress"""
        return self.OtherAddress

    @otheraddress.setter
    def otheraddress(self, value):
        """Lower case alias for OtherAddress.setter"""
        self.OtherAddress = value

    @property
    def OtherAddressCity(self):
        return self.com_object.OtherAddressCity

    @OtherAddressCity.setter
    def OtherAddressCity(self, value):
        self.com_object.OtherAddressCity = value

    @property
    def otheraddresscity(self):
        """Lower case alias for OtherAddressCity"""
        return self.OtherAddressCity

    @otheraddresscity.setter
    def otheraddresscity(self, value):
        """Lower case alias for OtherAddressCity.setter"""
        self.OtherAddressCity = value

    @property
    def OtherAddressCountry(self):
        return self.com_object.OtherAddressCountry

    @OtherAddressCountry.setter
    def OtherAddressCountry(self, value):
        self.com_object.OtherAddressCountry = value

    @property
    def otheraddresscountry(self):
        """Lower case alias for OtherAddressCountry"""
        return self.OtherAddressCountry

    @otheraddresscountry.setter
    def otheraddresscountry(self, value):
        """Lower case alias for OtherAddressCountry.setter"""
        self.OtherAddressCountry = value

    @property
    def OtherAddressPostalCode(self):
        return self.com_object.OtherAddressPostalCode

    @OtherAddressPostalCode.setter
    def OtherAddressPostalCode(self, value):
        self.com_object.OtherAddressPostalCode = value

    @property
    def otheraddresspostalcode(self):
        """Lower case alias for OtherAddressPostalCode"""
        return self.OtherAddressPostalCode

    @otheraddresspostalcode.setter
    def otheraddresspostalcode(self, value):
        """Lower case alias for OtherAddressPostalCode.setter"""
        self.OtherAddressPostalCode = value

    @property
    def OtherAddressPostOfficeBox(self):
        return self.com_object.OtherAddressPostOfficeBox

    @OtherAddressPostOfficeBox.setter
    def OtherAddressPostOfficeBox(self, value):
        self.com_object.OtherAddressPostOfficeBox = value

    @property
    def otheraddresspostofficebox(self):
        """Lower case alias for OtherAddressPostOfficeBox"""
        return self.OtherAddressPostOfficeBox

    @otheraddresspostofficebox.setter
    def otheraddresspostofficebox(self, value):
        """Lower case alias for OtherAddressPostOfficeBox.setter"""
        self.OtherAddressPostOfficeBox = value

    @property
    def OtherAddressState(self):
        return self.com_object.OtherAddressState

    @OtherAddressState.setter
    def OtherAddressState(self, value):
        self.com_object.OtherAddressState = value

    @property
    def otheraddressstate(self):
        """Lower case alias for OtherAddressState"""
        return self.OtherAddressState

    @otheraddressstate.setter
    def otheraddressstate(self, value):
        """Lower case alias for OtherAddressState.setter"""
        self.OtherAddressState = value

    @property
    def OtherAddressStreet(self):
        return self.com_object.OtherAddressStreet

    @OtherAddressStreet.setter
    def OtherAddressStreet(self, value):
        self.com_object.OtherAddressStreet = value

    @property
    def otheraddressstreet(self):
        """Lower case alias for OtherAddressStreet"""
        return self.OtherAddressStreet

    @otheraddressstreet.setter
    def otheraddressstreet(self, value):
        """Lower case alias for OtherAddressStreet.setter"""
        self.OtherAddressStreet = value

    @property
    def OtherFaxNumber(self):
        return self.com_object.OtherFaxNumber

    @OtherFaxNumber.setter
    def OtherFaxNumber(self, value):
        self.com_object.OtherFaxNumber = value

    @property
    def otherfaxnumber(self):
        """Lower case alias for OtherFaxNumber"""
        return self.OtherFaxNumber

    @otherfaxnumber.setter
    def otherfaxnumber(self, value):
        """Lower case alias for OtherFaxNumber.setter"""
        self.OtherFaxNumber = value

    @property
    def OtherTelephoneNumber(self):
        return self.com_object.OtherTelephoneNumber

    @OtherTelephoneNumber.setter
    def OtherTelephoneNumber(self, value):
        self.com_object.OtherTelephoneNumber = value

    @property
    def othertelephonenumber(self):
        """Lower case alias for OtherTelephoneNumber"""
        return self.OtherTelephoneNumber

    @othertelephonenumber.setter
    def othertelephonenumber(self, value):
        """Lower case alias for OtherTelephoneNumber.setter"""
        self.OtherTelephoneNumber = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def PagerNumber(self):
        return self.com_object.PagerNumber

    @PagerNumber.setter
    def PagerNumber(self, value):
        self.com_object.PagerNumber = value

    @property
    def pagernumber(self):
        """Lower case alias for PagerNumber"""
        return self.PagerNumber

    @pagernumber.setter
    def pagernumber(self, value):
        """Lower case alias for PagerNumber.setter"""
        self.PagerNumber = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PersonalHomePage(self):
        return self.com_object.PersonalHomePage

    @PersonalHomePage.setter
    def PersonalHomePage(self, value):
        self.com_object.PersonalHomePage = value

    @property
    def personalhomepage(self):
        """Lower case alias for PersonalHomePage"""
        return self.PersonalHomePage

    @personalhomepage.setter
    def personalhomepage(self, value):
        """Lower case alias for PersonalHomePage.setter"""
        self.PersonalHomePage = value

    @property
    def PrimaryTelephoneNumber(self):
        return self.com_object.PrimaryTelephoneNumber

    @PrimaryTelephoneNumber.setter
    def PrimaryTelephoneNumber(self, value):
        self.com_object.PrimaryTelephoneNumber = value

    @property
    def primarytelephonenumber(self):
        """Lower case alias for PrimaryTelephoneNumber"""
        return self.PrimaryTelephoneNumber

    @primarytelephonenumber.setter
    def primarytelephonenumber(self, value):
        """Lower case alias for PrimaryTelephoneNumber.setter"""
        self.PrimaryTelephoneNumber = value

    @property
    def Profession(self):
        return self.com_object.Profession

    @Profession.setter
    def Profession(self, value):
        self.com_object.Profession = value

    @property
    def profession(self):
        """Lower case alias for Profession"""
        return self.Profession

    @profession.setter
    def profession(self, value):
        """Lower case alias for Profession.setter"""
        self.Profession = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def RadioTelephoneNumber(self):
        return self.com_object.RadioTelephoneNumber

    @RadioTelephoneNumber.setter
    def RadioTelephoneNumber(self, value):
        self.com_object.RadioTelephoneNumber = value

    @property
    def radiotelephonenumber(self):
        """Lower case alias for RadioTelephoneNumber"""
        return self.RadioTelephoneNumber

    @radiotelephonenumber.setter
    def radiotelephonenumber(self, value):
        """Lower case alias for RadioTelephoneNumber.setter"""
        self.RadioTelephoneNumber = value

    @property
    def ReferredBy(self):
        return self.com_object.ReferredBy

    @ReferredBy.setter
    def ReferredBy(self, value):
        self.com_object.ReferredBy = value

    @property
    def referredby(self):
        """Lower case alias for ReferredBy"""
        return self.ReferredBy

    @referredby.setter
    def referredby(self, value):
        """Lower case alias for ReferredBy.setter"""
        self.ReferredBy = value

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    @property
    def reminderoverridedefault(self):
        """Lower case alias for ReminderOverrideDefault"""
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        """Lower case alias for ReminderOverrideDefault.setter"""
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    @property
    def reminderplaysound(self):
        """Lower case alias for ReminderPlaySound"""
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        """Lower case alias for ReminderPlaySound.setter"""
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    @property
    def reminderset(self):
        """Lower case alias for ReminderSet"""
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        """Lower case alias for ReminderSet.setter"""
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    @property
    def remindersoundfile(self):
        """Lower case alias for ReminderSoundFile"""
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        """Lower case alias for ReminderSoundFile.setter"""
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    @property
    def remindertime(self):
        """Lower case alias for ReminderTime"""
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        """Lower case alias for ReminderTime.setter"""
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def SelectedMailingAddress(self):
        return OlMailingAddress(self.com_object.SelectedMailingAddress)

    @SelectedMailingAddress.setter
    def SelectedMailingAddress(self, value):
        self.com_object.SelectedMailingAddress = value

    @property
    def selectedmailingaddress(self):
        """Lower case alias for SelectedMailingAddress"""
        return self.SelectedMailingAddress

    @selectedmailingaddress.setter
    def selectedmailingaddress(self, value):
        """Lower case alias for SelectedMailingAddress.setter"""
        self.SelectedMailingAddress = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Spouse(self):
        return self.com_object.Spouse

    @Spouse.setter
    def Spouse(self, value):
        self.com_object.Spouse = value

    @property
    def spouse(self):
        """Lower case alias for Spouse"""
        return self.Spouse

    @spouse.setter
    def spouse(self, value):
        """Lower case alias for Spouse.setter"""
        self.Spouse = value

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def Suffix(self):
        return self.com_object.Suffix

    @Suffix.setter
    def Suffix(self, value):
        self.com_object.Suffix = value

    @property
    def suffix(self):
        """Lower case alias for Suffix"""
        return self.Suffix

    @suffix.setter
    def suffix(self, value):
        """Lower case alias for Suffix.setter"""
        self.Suffix = value

    @property
    def TaskCompletedDate(self):
        return ContactItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    @property
    def taskcompleteddate(self):
        """Lower case alias for TaskCompletedDate"""
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        """Lower case alias for TaskCompletedDate.setter"""
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return ContactItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    @property
    def taskduedate(self):
        """Lower case alias for TaskDueDate"""
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        """Lower case alias for TaskDueDate.setter"""
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return ContactItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    @property
    def taskstartdate(self):
        """Lower case alias for TaskStartDate"""
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        """Lower case alias for TaskStartDate.setter"""
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return ContactItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    @property
    def tasksubject(self):
        """Lower case alias for TaskSubject"""
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        """Lower case alias for TaskSubject.setter"""
        self.TaskSubject = value

    @property
    def TelexNumber(self):
        return self.com_object.TelexNumber

    @TelexNumber.setter
    def TelexNumber(self, value):
        self.com_object.TelexNumber = value

    @property
    def telexnumber(self):
        """Lower case alias for TelexNumber"""
        return self.TelexNumber

    @telexnumber.setter
    def telexnumber(self, value):
        """Lower case alias for TelexNumber.setter"""
        self.TelexNumber = value

    @property
    def Title(self):
        return self.com_object.Title

    @Title.setter
    def Title(self, value):
        self.com_object.Title = value

    @property
    def title(self):
        """Lower case alias for Title"""
        return self.Title

    @title.setter
    def title(self, value):
        """Lower case alias for Title.setter"""
        self.Title = value

    @property
    def ToDoTaskOrdinal(self):
        return ContactItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    @property
    def todotaskordinal(self):
        """Lower case alias for ToDoTaskOrdinal"""
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        """Lower case alias for ToDoTaskOrdinal.setter"""
        self.ToDoTaskOrdinal = value

    @property
    def TTYTDDTelephoneNumber(self):
        return self.com_object.TTYTDDTelephoneNumber

    @TTYTDDTelephoneNumber.setter
    def TTYTDDTelephoneNumber(self, value):
        self.com_object.TTYTDDTelephoneNumber = value

    @property
    def ttytddtelephonenumber(self):
        """Lower case alias for TTYTDDTelephoneNumber"""
        return self.TTYTDDTelephoneNumber

    @ttytddtelephonenumber.setter
    def ttytddtelephonenumber(self, value):
        """Lower case alias for TTYTDDTelephoneNumber.setter"""
        self.TTYTDDTelephoneNumber = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def User1(self):
        return self.com_object.User1

    @User1.setter
    def User1(self, value):
        self.com_object.User1 = value

    @property
    def user1(self):
        """Lower case alias for User1"""
        return self.User1

    @user1.setter
    def user1(self, value):
        """Lower case alias for User1.setter"""
        self.User1 = value

    @property
    def User2(self):
        return self.com_object.User2

    @User2.setter
    def User2(self, value):
        self.com_object.User2 = value

    @property
    def user2(self):
        """Lower case alias for User2"""
        return self.User2

    @user2.setter
    def user2(self, value):
        """Lower case alias for User2.setter"""
        self.User2 = value

    @property
    def User3(self):
        return self.com_object.User3

    @User3.setter
    def User3(self, value):
        self.com_object.User3 = value

    @property
    def user3(self):
        """Lower case alias for User3"""
        return self.User3

    @user3.setter
    def user3(self, value):
        """Lower case alias for User3.setter"""
        self.User3 = value

    @property
    def User4(self):
        return self.com_object.User4

    @User4.setter
    def User4(self, value):
        self.com_object.User4 = value

    @property
    def user4(self):
        """Lower case alias for User4"""
        return self.User4

    @user4.setter
    def user4(self, value):
        """Lower case alias for User4.setter"""
        self.User4 = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    @property
    def WebPage(self):
        return self.com_object.WebPage

    @WebPage.setter
    def WebPage(self, value):
        self.com_object.WebPage = value

    @property
    def webpage(self):
        """Lower case alias for WebPage"""
        return self.WebPage

    @webpage.setter
    def webpage(self, value):
        """Lower case alias for WebPage.setter"""
        self.WebPage = value

    @property
    def YomiCompanyName(self):
        return self.com_object.YomiCompanyName

    @YomiCompanyName.setter
    def YomiCompanyName(self, value):
        self.com_object.YomiCompanyName = value

    @property
    def yomicompanyname(self):
        """Lower case alias for YomiCompanyName"""
        return self.YomiCompanyName

    @yomicompanyname.setter
    def yomicompanyname(self, value):
        """Lower case alias for YomiCompanyName.setter"""
        self.YomiCompanyName = value

    @property
    def YomiFirstName(self):
        return self.com_object.YomiFirstName

    @YomiFirstName.setter
    def YomiFirstName(self, value):
        self.com_object.YomiFirstName = value

    @property
    def yomifirstname(self):
        """Lower case alias for YomiFirstName"""
        return self.YomiFirstName

    @yomifirstname.setter
    def yomifirstname(self, value):
        """Lower case alias for YomiFirstName.setter"""
        self.YomiFirstName = value

    @property
    def YomiLastName(self):
        return self.com_object.YomiLastName

    @YomiLastName.setter
    def YomiLastName(self, value):
        self.com_object.YomiLastName = value

    @property
    def yomilastname(self):
        """Lower case alias for YomiLastName"""
        return self.YomiLastName

    @yomilastname.setter
    def yomilastname(self, value):
        """Lower case alias for YomiLastName.setter"""
        self.YomiLastName = value

    def AddBusinessCardLogoPicture(self, Path=None):
        arguments = com_arguments([unwrap(a) for a in [Path]])
        self.com_object.AddBusinessCardLogoPicture(*arguments)

    # Lower case alias for AddBusinessCardLogoPicture
    def addbusinesscardlogopicture(self, Path=None):
        arguments = [Path]
        return self.AddBusinessCardLogoPicture(*arguments)

    def AddPicture(self, Path=None):
        arguments = com_arguments([unwrap(a) for a in [Path]])
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
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [MarkInterval]])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
        self.com_object.SaveAs(*arguments)

    # Lower case alias for SaveAs
    def saveas(self, Path=None, Type=None):
        arguments = [Path, Type]
        return self.SaveAs(*arguments)

    def SaveBusinessCardImage(self, Path=None):
        arguments = com_arguments([unwrap(a) for a in [Path]])
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

    def showcheckaddressdialog(self, MailingAddress=None):
        arguments = com_arguments([unwrap(a) for a in [MailingAddress]])
        self.com_object.showcheckaddressdialog(*arguments)

    # Lower case alias for showcheckaddressdialog
    def showcheckaddressdialog(self, MailingAddress=None):
        arguments = [MailingAddress]
        return self.showcheckaddressdialog(*arguments)

    def showcheckfullnamedialog(self):
        self.com_object.showcheckfullnamedialog()

    # Lower case alias for showcheckfullnamedialog
    def showcheckfullnamedialog(self):
        return self.showcheckfullnamedialog()

    def ShowCheckPhoneDialog(self, PhoneNumber=None):
        arguments = com_arguments([unwrap(a) for a in [PhoneNumber]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    @property
    def navigationgroups(self):
        """Lower case alias for NavigationGroups"""
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    @property
    def navigationmoduletype(self):
        """Lower case alias for NavigationModuleType"""
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return ContactsModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return ContactsModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
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

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def Parent(self):
        return Conversation(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def ClearAlwaysAssignCategories(self, Store=None):
        arguments = com_arguments([unwrap(a) for a in [Store]])
        self.com_object.ClearAlwaysAssignCategories(*arguments)

    # Lower case alias for ClearAlwaysAssignCategories
    def clearalwaysassigncategories(self, Store=None):
        arguments = [Store]
        return self.ClearAlwaysAssignCategories(*arguments)

    def GetAlwaysAssignCategories(self, Store=None):
        arguments = com_arguments([unwrap(a) for a in [Store]])
        return self.com_object.GetAlwaysAssignCategories(*arguments)

    # Lower case alias for GetAlwaysAssignCategories
    def getalwaysassigncategories(self, Store=None):
        arguments = [Store]
        return self.GetAlwaysAssignCategories(*arguments)

    def GetAlwaysDelete(self, Store=None):
        arguments = com_arguments([unwrap(a) for a in [Store]])
        return OlAlwaysDeleteConversation(self.com_object.GetAlwaysDelete(*arguments))

    # Lower case alias for GetAlwaysDelete
    def getalwaysdelete(self, Store=None):
        arguments = [Store]
        return self.GetAlwaysDelete(*arguments)

    def GetAlwaysMoveToFolder(self, Store=None):
        arguments = com_arguments([unwrap(a) for a in [Store]])
        return Folder(self.com_object.GetAlwaysMoveToFolder(*arguments))

    # Lower case alias for GetAlwaysMoveToFolder
    def getalwaysmovetofolder(self, Store=None):
        arguments = [Store]
        return self.GetAlwaysMoveToFolder(*arguments)

    def GetChildren(self, Item=None):
        arguments = com_arguments([unwrap(a) for a in [Item]])
        return SimpleItems(self.com_object.GetChildren(*arguments))

    # Lower case alias for GetChildren
    def getchildren(self, Item=None):
        arguments = [Item]
        return self.GetChildren(*arguments)

    def GetParent(self, Item=None):
        arguments = com_arguments([unwrap(a) for a in [Item]])
        return self.com_object.GetParent(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Categories, Store]])
        self.com_object.SetAlwaysAssignCategories(*arguments)

    # Lower case alias for SetAlwaysAssignCategories
    def setalwaysassigncategories(self, Categories=None, Store=None):
        arguments = [Categories, Store]
        return self.SetAlwaysAssignCategories(*arguments)

    def SetAlwaysDelete(self, AlwaysDelete=None, Store=None):
        arguments = com_arguments([unwrap(a) for a in [AlwaysDelete, Store]])
        self.com_object.SetAlwaysDelete(*arguments)

    # Lower case alias for SetAlwaysDelete
    def setalwaysdelete(self, AlwaysDelete=None, Store=None):
        arguments = [AlwaysDelete, Store]
        return self.SetAlwaysDelete(*arguments)

    def SetAlwaysMoveToFolder(self, MoveToFolder=None, Store=None):
        arguments = com_arguments([unwrap(a) for a in [MoveToFolder, Store]])
        self.com_object.SetAlwaysMoveToFolder(*arguments)

    # Lower case alias for SetAlwaysMoveToFolder
    def setalwaysmovetofolder(self, MoveToFolder=None, Store=None):
        arguments = [MoveToFolder, Store]
        return self.SetAlwaysMoveToFolder(*arguments)

    def StopAlwaysDelete(self, Store=None):
        arguments = com_arguments([unwrap(a) for a in [Store]])
        self.com_object.StopAlwaysDelete(*arguments)

    # Lower case alias for StopAlwaysDelete
    def stopalwaysdelete(self, Store=None):
        arguments = [Store]
        return self.StopAlwaysDelete(*arguments)

    def StopAlwaysMoveToFolder(self, Store=None):
        arguments = com_arguments([unwrap(a) for a in [Store]])
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

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DLName(self):
        return self.com_object.DLName

    @DLName.setter
    def DLName(self, value):
        self.com_object.DLName = value

    @property
    def dlname(self):
        """Lower case alias for DLName"""
        return self.DLName

    @dlname.setter
    def dlname(self, value):
        """Lower case alias for DLName.setter"""
        self.DLName = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return DistListItem(self.com_object.IsMarkedAsTask)

    @property
    def ismarkedastask(self):
        """Lower case alias for IsMarkedAsTask"""
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MemberCount(self):
        return self.com_object.MemberCount

    @property
    def membercount(self):
        """Lower case alias for MemberCount"""
        return self.MemberCount

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    @property
    def reminderoverridedefault(self):
        """Lower case alias for ReminderOverrideDefault"""
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        """Lower case alias for ReminderOverrideDefault.setter"""
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    @property
    def reminderplaysound(self):
        """Lower case alias for ReminderPlaySound"""
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        """Lower case alias for ReminderPlaySound.setter"""
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    @property
    def reminderset(self):
        """Lower case alias for ReminderSet"""
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        """Lower case alias for ReminderSet.setter"""
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    @property
    def remindersoundfile(self):
        """Lower case alias for ReminderSoundFile"""
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        """Lower case alias for ReminderSoundFile.setter"""
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    @property
    def remindertime(self):
        """Lower case alias for ReminderTime"""
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        """Lower case alias for ReminderTime.setter"""
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def TaskCompletedDate(self):
        return DistListItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    @property
    def taskcompleteddate(self):
        """Lower case alias for TaskCompletedDate"""
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        """Lower case alias for TaskCompletedDate.setter"""
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return DistListItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    @property
    def taskduedate(self):
        """Lower case alias for TaskDueDate"""
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        """Lower case alias for TaskDueDate.setter"""
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return DistListItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    @property
    def taskstartdate(self):
        """Lower case alias for TaskStartDate"""
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        """Lower case alias for TaskStartDate.setter"""
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return DistListItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    @property
    def tasksubject(self):
        """Lower case alias for TaskSubject"""
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        """Lower case alias for TaskSubject.setter"""
        self.TaskSubject = value

    @property
    def ToDoTaskOrdinal(self):
        return DistListItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    @property
    def todotaskordinal(self):
        """Lower case alias for ToDoTaskOrdinal"""
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        """Lower case alias for ToDoTaskOrdinal.setter"""
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def AddMember(self, Recipient=None):
        arguments = com_arguments([unwrap(a) for a in [Recipient]])
        self.com_object.AddMember(*arguments)

    # Lower case alias for AddMember
    def addmember(self, Recipient=None):
        arguments = [Recipient]
        return self.AddMember(*arguments)

    def AddMembers(self, Recipients=None):
        arguments = com_arguments([unwrap(a) for a in [Recipients]])
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
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Recipient(self.com_object.GetMember(*arguments))

    # Lower case alias for GetMember
    def getmember(self, Index=None):
        arguments = [Index]
        return self.GetMember(*arguments)

    def MarkAsTask(self, MarkInterval=None):
        arguments = com_arguments([unwrap(a) for a in [MarkInterval]])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Recipient]])
        self.com_object.RemoveMember(*arguments)

    # Lower case alias for RemoveMember
    def removemember(self, Recipient=None):
        arguments = [Recipient]
        return self.RemoveMember(*arguments)

    def RemoveMembers(self, Recipients=None):
        arguments = com_arguments([unwrap(a) for a in [Recipients]])
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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return self.com_object.GetInspector

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return self.com_object.MarkForDownload

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def appointmentitem(self):
        """Lower case alias for AppointmentItem"""
        return self.AppointmentItem

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Deleted(self):
        return AppointmentItem(self.com_object.Deleted)

    @property
    def deleted(self):
        """Lower case alias for Deleted"""
        return self.Deleted

    @property
    def OriginalDate(self):
        return AppointmentItem(self.com_object.OriginalDate)

    @property
    def originaldate(self):
        """Lower case alias for OriginalDate"""
        return self.OriginalDate

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def address(self):
        """Lower case alias for Address"""
        return self.Address

    @address.setter
    def address(self, value):
        """Lower case alias for Address.setter"""
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.com_object.AddressEntryUserType)

    @property
    def addressentryusertype(self):
        """Lower case alias for AddressEntryUserType"""
        return self.AddressEntryUserType

    @property
    def Alias(self):
        return ExchangeDistributionList(self.com_object.Alias)

    @property
    def alias(self):
        """Lower case alias for Alias"""
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

    @property
    def comments(self):
        """Lower case alias for Comments"""
        return self.Comments

    @comments.setter
    def comments(self, value):
        """Lower case alias for Comments.setter"""
        self.Comments = value

    @property
    def DisplayType(self):
        return OlDisplayType(self.com_object.DisplayType)

    @property
    def displaytype(self):
        """Lower case alias for DisplayType"""
        return self.DisplayType

    @property
    def ID(self):
        return ExchangeDistributionList(self.com_object.ID)

    @property
    def id(self):
        """Lower case alias for ID"""
        return self.ID

    @property
    def Name(self):
        return ExchangeDistributionList(self.com_object.Name)

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return ExchangeDistributionList(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PrimarySmtpAddress(self):
        return ExchangeDistributionList(self.com_object.PrimarySmtpAddress)

    @property
    def primarysmtpaddress(self):
        """Lower case alias for PrimarySmtpAddress"""
        return self.PrimarySmtpAddress

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Type(self):
        return ExchangeDistributionList(self.com_object.Type)

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @type.setter
    def type(self, value):
        """Lower case alias for Type.setter"""
        self.Type = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([unwrap(a) for a in [HWnd]])
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
        arguments = com_arguments([unwrap(a) for a in [Start, MinPerChar, CompleteFormat]])
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

    def getunifiedgroup(self):
        return self.com_object.getunifiedgroup()

    # Lower case alias for getunifiedgroup
    def getunifiedgroup(self):
        return self.getunifiedgroup()

    def getunifiedgroupfromstore(self):
        return self.com_object.getunifiedgroupfromstore()

    # Lower case alias for getunifiedgroupfromstore
    def getunifiedgroupfromstore(self):
        return self.getunifiedgroupfromstore()

    def isunifiedgroup(self):
        return self.com_object.isunifiedgroup()

    # Lower case alias for isunifiedgroup
    def isunifiedgroup(self):
        return self.isunifiedgroup()

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([unwrap(a) for a in [MakePermanent, Refresh]])
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

    @property
    def address(self):
        """Lower case alias for Address"""
        return self.Address

    @address.setter
    def address(self, value):
        """Lower case alias for Address.setter"""
        self.Address = value

    @property
    def AddressEntryUserType(self):
        return OlAddressEntryUserType(self.com_object.AddressEntryUserType)

    @property
    def addressentryusertype(self):
        """Lower case alias for AddressEntryUserType"""
        return self.AddressEntryUserType

    @property
    def Alias(self):
        return ExchangeUser(self.com_object.Alias)

    @property
    def alias(self):
        """Lower case alias for Alias"""
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

    @property
    def assistantname(self):
        """Lower case alias for AssistantName"""
        return self.AssistantName

    @assistantname.setter
    def assistantname(self, value):
        """Lower case alias for AssistantName.setter"""
        self.AssistantName = value

    @property
    def BusinessTelephoneNumber(self):
        return ExchangeUser(self.com_object.BusinessTelephoneNumber)

    @BusinessTelephoneNumber.setter
    def BusinessTelephoneNumber(self, value):
        self.com_object.BusinessTelephoneNumber = value

    @property
    def businesstelephonenumber(self):
        """Lower case alias for BusinessTelephoneNumber"""
        return self.BusinessTelephoneNumber

    @businesstelephonenumber.setter
    def businesstelephonenumber(self, value):
        """Lower case alias for BusinessTelephoneNumber.setter"""
        self.BusinessTelephoneNumber = value

    @property
    def City(self):
        return ExchangeUser(self.com_object.City)

    @City.setter
    def City(self, value):
        self.com_object.City = value

    @property
    def city(self):
        """Lower case alias for City"""
        return self.City

    @city.setter
    def city(self, value):
        """Lower case alias for City.setter"""
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

    @property
    def comments(self):
        """Lower case alias for Comments"""
        return self.Comments

    @comments.setter
    def comments(self, value):
        """Lower case alias for Comments.setter"""
        self.Comments = value

    @property
    def CompanyName(self):
        return ExchangeUser(self.com_object.CompanyName)

    @CompanyName.setter
    def CompanyName(self, value):
        self.com_object.CompanyName = value

    @property
    def companyname(self):
        """Lower case alias for CompanyName"""
        return self.CompanyName

    @companyname.setter
    def companyname(self, value):
        """Lower case alias for CompanyName.setter"""
        self.CompanyName = value

    @property
    def Department(self):
        return ExchangeUser(self.com_object.Department)

    @Department.setter
    def Department(self, value):
        self.com_object.Department = value

    @property
    def department(self):
        """Lower case alias for Department"""
        return self.Department

    @department.setter
    def department(self, value):
        """Lower case alias for Department.setter"""
        self.Department = value

    @property
    def DisplayType(self):
        return OlDisplayType(self.com_object.DisplayType)

    @property
    def displaytype(self):
        """Lower case alias for DisplayType"""
        return self.DisplayType

    @property
    def FirstName(self):
        return ExchangeUser(self.com_object.FirstName)

    @FirstName.setter
    def FirstName(self, value):
        self.com_object.FirstName = value

    @property
    def firstname(self):
        """Lower case alias for FirstName"""
        return self.FirstName

    @firstname.setter
    def firstname(self, value):
        """Lower case alias for FirstName.setter"""
        self.FirstName = value

    @property
    def ID(self):
        return ExchangeUser(self.com_object.ID)

    @property
    def id(self):
        """Lower case alias for ID"""
        return self.ID

    @property
    def JobTitle(self):
        return ExchangeUser(self.com_object.JobTitle)

    @JobTitle.setter
    def JobTitle(self, value):
        self.com_object.JobTitle = value

    @property
    def jobtitle(self):
        """Lower case alias for JobTitle"""
        return self.JobTitle

    @jobtitle.setter
    def jobtitle(self, value):
        """Lower case alias for JobTitle.setter"""
        self.JobTitle = value

    @property
    def LastName(self):
        return ExchangeUser(self.com_object.LastName)

    @LastName.setter
    def LastName(self, value):
        self.com_object.LastName = value

    @property
    def lastname(self):
        """Lower case alias for LastName"""
        return self.LastName

    @lastname.setter
    def lastname(self, value):
        """Lower case alias for LastName.setter"""
        self.LastName = value

    @property
    def MobileTelephoneNumber(self):
        return ExchangeUser(self.com_object.MobileTelephoneNumber)

    @MobileTelephoneNumber.setter
    def MobileTelephoneNumber(self, value):
        self.com_object.MobileTelephoneNumber = value

    @property
    def mobiletelephonenumber(self):
        """Lower case alias for MobileTelephoneNumber"""
        return self.MobileTelephoneNumber

    @mobiletelephonenumber.setter
    def mobiletelephonenumber(self, value):
        """Lower case alias for MobileTelephoneNumber.setter"""
        self.MobileTelephoneNumber = value

    @property
    def Name(self):
        return ExchangeUser(self.com_object.Name)

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def OfficeLocation(self):
        return ExchangeUser(self.com_object.OfficeLocation)

    @OfficeLocation.setter
    def OfficeLocation(self, value):
        self.com_object.OfficeLocation = value

    @property
    def officelocation(self):
        """Lower case alias for OfficeLocation"""
        return self.OfficeLocation

    @officelocation.setter
    def officelocation(self, value):
        """Lower case alias for OfficeLocation.setter"""
        self.OfficeLocation = value

    @property
    def Parent(self):
        return ExchangeUser(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PostalCode(self):
        return ExchangeUser(self.com_object.PostalCode)

    @PostalCode.setter
    def PostalCode(self, value):
        self.com_object.PostalCode = value

    @property
    def postalcode(self):
        """Lower case alias for PostalCode"""
        return self.PostalCode

    @postalcode.setter
    def postalcode(self, value):
        """Lower case alias for PostalCode.setter"""
        self.PostalCode = value

    @property
    def PrimarySmtpAddress(self):
        return ExchangeUser(self.com_object.PrimarySmtpAddress)

    @property
    def primarysmtpaddress(self):
        """Lower case alias for PrimarySmtpAddress"""
        return self.PrimarySmtpAddress

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def StateOrProvince(self):
        return ExchangeUser(self.com_object.StateOrProvince)

    @StateOrProvince.setter
    def StateOrProvince(self, value):
        self.com_object.StateOrProvince = value

    @property
    def stateorprovince(self):
        """Lower case alias for StateOrProvince"""
        return self.StateOrProvince

    @stateorprovince.setter
    def stateorprovince(self, value):
        """Lower case alias for StateOrProvince.setter"""
        self.StateOrProvince = value

    @property
    def StreetAddress(self):
        return ExchangeUser(self.com_object.StreetAddress)

    @StreetAddress.setter
    def StreetAddress(self, value):
        self.com_object.StreetAddress = value

    @property
    def streetaddress(self):
        """Lower case alias for StreetAddress"""
        return self.StreetAddress

    @streetaddress.setter
    def streetaddress(self, value):
        """Lower case alias for StreetAddress.setter"""
        self.StreetAddress = value

    @property
    def Type(self):
        return ExchangeUser(self.com_object.Type)

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @type.setter
    def type(self, value):
        """Lower case alias for Type.setter"""
        self.Type = value

    @property
    def YomiCompanyName(self):
        return ExchangeUser(self.com_object.YomiCompanyName)

    @YomiCompanyName.setter
    def YomiCompanyName(self, value):
        self.com_object.YomiCompanyName = value

    @property
    def yomicompanyname(self):
        """Lower case alias for YomiCompanyName"""
        return self.YomiCompanyName

    @yomicompanyname.setter
    def yomicompanyname(self, value):
        """Lower case alias for YomiCompanyName.setter"""
        self.YomiCompanyName = value

    @property
    def YomiDepartment(self):
        return ExchangeUser(self.com_object.YomiDepartment)

    @YomiDepartment.setter
    def YomiDepartment(self, value):
        self.com_object.YomiDepartment = value

    @property
    def yomidepartment(self):
        """Lower case alias for YomiDepartment"""
        return self.YomiDepartment

    @yomidepartment.setter
    def yomidepartment(self, value):
        """Lower case alias for YomiDepartment.setter"""
        self.YomiDepartment = value

    @property
    def YomiDisplayName(self):
        return ExchangeUser(self.com_object.YomiDisplayName)

    @YomiDisplayName.setter
    def YomiDisplayName(self, value):
        self.com_object.YomiDisplayName = value

    @property
    def yomidisplayname(self):
        """Lower case alias for YomiDisplayName"""
        return self.YomiDisplayName

    @yomidisplayname.setter
    def yomidisplayname(self, value):
        """Lower case alias for YomiDisplayName.setter"""
        self.YomiDisplayName = value

    @property
    def YomiFirstName(self):
        return ExchangeUser(self.com_object.YomiFirstName)

    @YomiFirstName.setter
    def YomiFirstName(self, value):
        self.com_object.YomiFirstName = value

    @property
    def yomifirstname(self):
        """Lower case alias for YomiFirstName"""
        return self.YomiFirstName

    @yomifirstname.setter
    def yomifirstname(self, value):
        """Lower case alias for YomiFirstName.setter"""
        self.YomiFirstName = value

    @property
    def YomiLastName(self):
        return ExchangeUser(self.com_object.YomiLastName)

    @YomiLastName.setter
    def YomiLastName(self, value):
        self.com_object.YomiLastName = value

    @property
    def yomilastname(self):
        """Lower case alias for YomiLastName"""
        return self.YomiLastName

    @yomilastname.setter
    def yomilastname(self, value):
        """Lower case alias for YomiLastName.setter"""
        self.YomiLastName = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def Details(self, HWnd=None):
        arguments = com_arguments([unwrap(a) for a in [HWnd]])
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
        arguments = com_arguments([unwrap(a) for a in [Start, MinPerChar, CompleteFormat]])
        return self.com_object.GetFreeBusy(*arguments)

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
        return self.com_object.GetPicture()

    # Lower case alias for GetPicture
    def getpicture(self):
        return self.GetPicture()

    def getunifiedgroup(self):
        return self.com_object.getunifiedgroup()

    # Lower case alias for getunifiedgroup
    def getunifiedgroup(self):
        return self.getunifiedgroup()

    def getunifiedgroupfromstore(self):
        return self.com_object.getunifiedgroupfromstore()

    # Lower case alias for getunifiedgroupfromstore
    def getunifiedgroupfromstore(self):
        return self.getunifiedgroupfromstore()

    def isunifiedgroup(self):
        return self.com_object.isunifiedgroup()

    # Lower case alias for isunifiedgroup
    def isunifiedgroup(self):
        return self.isunifiedgroup()

    def Update(self, MakePermanent=None, Refresh=None):
        arguments = com_arguments([unwrap(a) for a in [MakePermanent, Refresh]])
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

    @property
    def accountselector(self):
        """Lower case alias for AccountSelector"""
        return self.AccountSelector

    @property
    def activeinlineresponse(self):
        return self.com_object.activeinlineresponse

    @property
    def activeinlineresponse(self):
        """Lower case alias for activeinlineresponse"""
        return self.activeinlineresponse

    @property
    def activeinlineresponsewordeditor(self):
        return self.com_object.activeinlineresponsewordeditor

    @property
    def activeinlineresponsewordeditor(self):
        """Lower case alias for activeinlineresponsewordeditor"""
        return self.activeinlineresponsewordeditor

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AttachmentSelection(self):
        return AttachmentSelection(self.com_object.AttachmentSelection)

    @property
    def attachmentselection(self):
        """Lower case alias for AttachmentSelection"""
        return self.AttachmentSelection

    @property
    def Caption(self):
        return self.com_object.Caption

    @property
    def caption(self):
        """Lower case alias for Caption"""
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

    @property
    def currentfolder(self):
        """Lower case alias for CurrentFolder"""
        return self.CurrentFolder

    @currentfolder.setter
    def currentfolder(self, value):
        """Lower case alias for CurrentFolder.setter"""
        self.CurrentFolder = value

    @property
    def CurrentView(self):
        return self.com_object.CurrentView

    @CurrentView.setter
    def CurrentView(self, value):
        self.com_object.CurrentView = value

    @property
    def currentview(self):
        """Lower case alias for CurrentView"""
        return self.CurrentView

    @currentview.setter
    def currentview(self, value):
        """Lower case alias for CurrentView.setter"""
        self.CurrentView = value

    @property
    def displaymode(self):
        return self.com_object.displaymode

    @property
    def displaymode(self):
        """Lower case alias for displaymode"""
        return self.displaymode

    @property
    def Height(self):
        return self.com_object.Height

    @Height.setter
    def Height(self, value):
        self.com_object.Height = value

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @height.setter
    def height(self, value):
        """Lower case alias for Height.setter"""
        self.Height = value

    @property
    def HTMLDocument(self):
        return self.com_object.HTMLDocument

    @property
    def htmldocument(self):
        """Lower case alias for HTMLDocument"""
        return self.HTMLDocument

    @property
    def Left(self):
        return self.com_object.Left

    @Left.setter
    def Left(self, value):
        self.com_object.Left = value

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

    @left.setter
    def left(self, value):
        """Lower case alias for Left.setter"""
        self.Left = value

    @property
    def NavigationPane(self):
        return NavigationPane(self.com_object.NavigationPane)

    @property
    def navigationpane(self):
        """Lower case alias for NavigationPane"""
        return self.NavigationPane

    @property
    def Panes(self):
        return Panes(self.com_object.Panes)

    @property
    def panes(self):
        """Lower case alias for Panes"""
        return self.Panes

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def previewpane(self):
        return self.com_object.previewpane

    @property
    def previewpane(self):
        """Lower case alias for previewpane"""
        return self.previewpane

    @property
    def Selection(self):
        return Selection(self.com_object.Selection)

    @property
    def selection(self):
        """Lower case alias for Selection"""
        return self.Selection

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Top(self):
        return self.com_object.Top

    @Top.setter
    def Top(self, value):
        self.com_object.Top = value

    @property
    def top(self):
        """Lower case alias for Top"""
        return self.Top

    @top.setter
    def top(self, value):
        """Lower case alias for Top.setter"""
        self.Top = value

    @property
    def Width(self):
        return self.com_object.Width

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    @property
    def width(self):
        """Lower case alias for Width"""
        return self.Width

    @width.setter
    def width(self, value):
        """Lower case alias for Width.setter"""
        self.Width = value

    @property
    def WindowState(self):
        return OlWindowState(self.com_object.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.com_object.WindowState = value

    @property
    def windowstate(self):
        """Lower case alias for WindowState"""
        return self.WindowState

    @windowstate.setter
    def windowstate(self, value):
        """Lower case alias for WindowState.setter"""
        self.WindowState = value

    def Activate(self):
        self.com_object.Activate()

    # Lower case alias for Activate
    def activate(self):
        return self.Activate()

    def AddToSelection(self, Item=None):
        arguments = com_arguments([unwrap(a) for a in [Item]])
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
        arguments = com_arguments([unwrap(a) for a in [Item]])
        return self.com_object.IsItemSelectableInView(*arguments)

    # Lower case alias for IsItemSelectableInView
    def isitemselectableinview(self, Item=None):
        arguments = [Item]
        return self.IsItemSelectableInView(*arguments)

    def IsPaneVisible(self, Pane=None):
        arguments = com_arguments([unwrap(a) for a in [Pane]])
        return self.com_object.IsPaneVisible(*arguments)

    # Lower case alias for IsPaneVisible
    def ispanevisible(self, Pane=None):
        arguments = [Pane]
        return self.IsPaneVisible(*arguments)

    def RemoveFromSelection(self, Item=None):
        arguments = com_arguments([unwrap(a) for a in [Item]])
        self.com_object.RemoveFromSelection(*arguments)

    # Lower case alias for RemoveFromSelection
    def removefromselection(self, Item=None):
        arguments = [Item]
        return self.RemoveFromSelection(*arguments)

    def Search(self, Query=None, SearchScope=None):
        arguments = com_arguments([unwrap(a) for a in [Query, SearchScope]])
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
        arguments = com_arguments([unwrap(a) for a in [Pane, Visible]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Folder=None, DisplayMode=None):
        arguments = com_arguments([unwrap(a) for a in [Folder, DisplayMode]])
        return Explorer(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Folder=None, DisplayMode=None):
        arguments = [Folder, DisplayMode]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def addressbookname(self):
        """Lower case alias for AddressBookName"""
        return self.AddressBookName

    @addressbookname.setter
    def addressbookname(self, value):
        """Lower case alias for AddressBookName.setter"""
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

    @property
    def currentview(self):
        """Lower case alias for CurrentView"""
        return self.CurrentView

    @property
    def CustomViewsOnly(self):
        return self.com_object.CustomViewsOnly

    @CustomViewsOnly.setter
    def CustomViewsOnly(self, value):
        self.com_object.CustomViewsOnly = value

    @property
    def customviewsonly(self):
        """Lower case alias for CustomViewsOnly"""
        return self.CustomViewsOnly

    @customviewsonly.setter
    def customviewsonly(self, value):
        """Lower case alias for CustomViewsOnly.setter"""
        self.CustomViewsOnly = value

    @property
    def DefaultItemType(self):
        return OlItemType(self.com_object.DefaultItemType)

    @property
    def defaultitemtype(self):
        """Lower case alias for DefaultItemType"""
        return self.DefaultItemType

    @property
    def DefaultMessageClass(self):
        return self.com_object.DefaultMessageClass

    @property
    def defaultmessageclass(self):
        """Lower case alias for DefaultMessageClass"""
        return self.DefaultMessageClass

    @property
    def Description(self):
        return self.com_object.Description

    @Description.setter
    def Description(self, value):
        self.com_object.Description = value

    @property
    def description(self):
        """Lower case alias for Description"""
        return self.Description

    @description.setter
    def description(self, value):
        """Lower case alias for Description.setter"""
        self.Description = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FolderPath(self):
        return self.com_object.FolderPath

    @property
    def folderpath(self):
        """Lower case alias for FolderPath"""
        return self.FolderPath

    @property
    def Folders(self):
        return Folders(self.com_object.Folders)

    @property
    def folders(self):
        """Lower case alias for Folders"""
        return self.Folders

    @property
    def InAppFolderSyncObject(self):
        return self.com_object.InAppFolderSyncObject

    @InAppFolderSyncObject.setter
    def InAppFolderSyncObject(self, value):
        self.com_object.InAppFolderSyncObject = value

    @property
    def inappfoldersyncobject(self):
        """Lower case alias for InAppFolderSyncObject"""
        return self.InAppFolderSyncObject

    @inappfoldersyncobject.setter
    def inappfoldersyncobject(self, value):
        """Lower case alias for InAppFolderSyncObject.setter"""
        self.InAppFolderSyncObject = value

    @property
    def IsSharePointFolder(self):
        return self.com_object.IsSharePointFolder

    @property
    def issharepointfolder(self):
        """Lower case alias for IsSharePointFolder"""
        return self.IsSharePointFolder

    @property
    def Items(self):
        return Items(self.com_object.Items)

    @property
    def items(self):
        """Lower case alias for Items"""
        return self.Items

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ShowAsOutlookAB(self):
        return self.com_object.ShowAsOutlookAB

    @ShowAsOutlookAB.setter
    def ShowAsOutlookAB(self, value):
        self.com_object.ShowAsOutlookAB = value

    @property
    def showasoutlookab(self):
        """Lower case alias for ShowAsOutlookAB"""
        return self.ShowAsOutlookAB

    @showasoutlookab.setter
    def showasoutlookab(self, value):
        """Lower case alias for ShowAsOutlookAB.setter"""
        self.ShowAsOutlookAB = value

    @property
    def ShowItemCount(self):
        return self.com_object.ShowItemCount

    @ShowItemCount.setter
    def ShowItemCount(self, value):
        self.com_object.ShowItemCount = value

    @property
    def showitemcount(self):
        """Lower case alias for ShowItemCount"""
        return self.ShowItemCount

    @showitemcount.setter
    def showitemcount(self, value):
        """Lower case alias for ShowItemCount.setter"""
        self.ShowItemCount = value

    @property
    def Store(self):
        return Store(self.com_object.Store)

    @property
    def store(self):
        """Lower case alias for Store"""
        return self.Store

    @property
    def StoreID(self):
        return self.com_object.StoreID

    @property
    def storeid(self):
        """Lower case alias for StoreID"""
        return self.StoreID

    @property
    def UnReadItemCount(self):
        return self.com_object.UnReadItemCount

    @property
    def unreaditemcount(self):
        """Lower case alias for UnReadItemCount"""
        return self.UnReadItemCount

    @property
    def UserDefinedProperties(self):
        return UserDefinedProperties(self.com_object.UserDefinedProperties)

    @property
    def userdefinedproperties(self):
        """Lower case alias for UserDefinedProperties"""
        return self.UserDefinedProperties

    @property
    def Views(self):
        return Views(self.com_object.Views)

    @property
    def views(self):
        """Lower case alias for Views"""
        return self.Views

    @property
    def WebViewOn(self):
        return self.com_object.WebViewOn

    @WebViewOn.setter
    def WebViewOn(self, value):
        self.com_object.WebViewOn = value

    @property
    def webviewon(self):
        """Lower case alias for WebViewOn"""
        return self.WebViewOn

    @webviewon.setter
    def webviewon(self, value):
        """Lower case alias for WebViewOn.setter"""
        self.WebViewOn = value

    @property
    def WebViewURL(self):
        return self.com_object.WebViewURL

    @WebViewURL.setter
    def WebViewURL(self, value):
        self.com_object.WebViewURL = value

    @property
    def webviewurl(self):
        """Lower case alias for WebViewURL"""
        return self.WebViewURL

    @webviewurl.setter
    def webviewurl(self, value):
        """Lower case alias for WebViewURL.setter"""
        self.WebViewURL = value

    def AddToPFFavorites(self):
        self.com_object.AddToPFFavorites()

    # Lower case alias for AddToPFFavorites
    def addtopffavorites(self):
        return self.AddToPFFavorites()

    def CopyTo(self, DestinationFolder=None):
        arguments = com_arguments([unwrap(a) for a in [DestinationFolder]])
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
        return self.com_object.GetCustomIcon()

    # Lower case alias for GetCustomIcon
    def getcustomicon(self):
        return self.GetCustomIcon()

    def GetExplorer(self, DisplayMode=None):
        arguments = com_arguments([unwrap(a) for a in [DisplayMode]])
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
        arguments = com_arguments([unwrap(a) for a in [StorageIdentifier, StorageIdentifierType]])
        return StorageItem(self.com_object.GetStorage(*arguments))

    # Lower case alias for GetStorage
    def getstorage(self, StorageIdentifier=None, StorageIdentifierType=None):
        arguments = [StorageIdentifier, StorageIdentifierType]
        return self.GetStorage(*arguments)

    def GetTable(self, Filter=None, TableContents=None):
        arguments = com_arguments([unwrap(a) for a in [Filter, TableContents]])
        return Folder(self.com_object.GetTable(*arguments))

    # Lower case alias for GetTable
    def gettable(self, Filter=None, TableContents=None):
        arguments = [Filter, TableContents]
        return self.GetTable(*arguments)

    def MoveTo(self, DestinationFolder=None):
        arguments = com_arguments([unwrap(a) for a in [DestinationFolder]])
        self.com_object.MoveTo(*arguments)

    # Lower case alias for MoveTo
    def moveto(self, DestinationFolder=None):
        arguments = [DestinationFolder]
        return self.MoveTo(*arguments)

    def SetCustomIcon(self, Picture=None):
        arguments = com_arguments([unwrap(a) for a in [Picture]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Type]])
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
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Folder(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)


class font:

    def __init__(self, font=None):
        self.com_object= font


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

    @property
    def category(self):
        """Lower case alias for Category"""
        return self.Category

    @category.setter
    def category(self, value):
        """Lower case alias for Category.setter"""
        self.Category = value

    @property
    def CategorySub(self):
        return self.com_object.CategorySub

    @CategorySub.setter
    def CategorySub(self, value):
        self.com_object.CategorySub = value

    @property
    def categorysub(self):
        """Lower case alias for CategorySub"""
        return self.CategorySub

    @categorysub.setter
    def categorysub(self, value):
        """Lower case alias for CategorySub.setter"""
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

    @property
    def comment(self):
        """Lower case alias for Comment"""
        return self.Comment

    @comment.setter
    def comment(self, value):
        """Lower case alias for Comment.setter"""
        self.Comment = value

    @property
    def ContactName(self):
        return FormDescription(self.com_object.ContactName)

    @ContactName.setter
    def ContactName(self, value):
        self.com_object.ContactName = value

    @property
    def contactname(self):
        """Lower case alias for ContactName"""
        return self.ContactName

    @contactname.setter
    def contactname(self, value):
        """Lower case alias for ContactName.setter"""
        self.ContactName = value

    @property
    def DisplayName(self):
        return self.com_object.DisplayName

    @DisplayName.setter
    def DisplayName(self, value):
        self.com_object.DisplayName = value

    @property
    def displayname(self):
        """Lower case alias for DisplayName"""
        return self.DisplayName

    @displayname.setter
    def displayname(self, value):
        """Lower case alias for DisplayName.setter"""
        self.DisplayName = value

    @property
    def Hidden(self):
        return self.com_object.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.com_object.Hidden = value

    @property
    def hidden(self):
        """Lower case alias for Hidden"""
        return self.Hidden

    @hidden.setter
    def hidden(self, value):
        """Lower case alias for Hidden.setter"""
        self.Hidden = value

    @property
    def Icon(self):
        return self.com_object.Icon

    @Icon.setter
    def Icon(self, value):
        self.com_object.Icon = value

    @property
    def icon(self):
        """Lower case alias for Icon"""
        return self.Icon

    @icon.setter
    def icon(self, value):
        """Lower case alias for Icon.setter"""
        self.Icon = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    @property
    def locked(self):
        """Lower case alias for Locked"""
        return self.Locked

    @locked.setter
    def locked(self, value):
        """Lower case alias for Locked.setter"""
        self.Locked = value

    @property
    def MessageClass(self):
        return FormDescription(self.com_object.MessageClass)

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @property
    def MiniIcon(self):
        return self.com_object.MiniIcon

    @MiniIcon.setter
    def MiniIcon(self, value):
        self.com_object.MiniIcon = value

    @property
    def miniicon(self):
        """Lower case alias for MiniIcon"""
        return self.MiniIcon

    @miniicon.setter
    def miniicon(self, value):
        """Lower case alias for MiniIcon.setter"""
        self.MiniIcon = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Number(self):
        return self.com_object.Number

    @Number.setter
    def Number(self, value):
        self.com_object.Number = value

    @property
    def number(self):
        """Lower case alias for Number"""
        return self.Number

    @number.setter
    def number(self, value):
        """Lower case alias for Number.setter"""
        self.Number = value

    @property
    def OneOff(self):
        return self.com_object.OneOff

    @OneOff.setter
    def OneOff(self, value):
        self.com_object.OneOff = value

    @property
    def oneoff(self):
        """Lower case alias for OneOff"""
        return self.OneOff

    @oneoff.setter
    def oneoff(self, value):
        """Lower case alias for OneOff.setter"""
        self.OneOff = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def ScriptText(self):
        return self.com_object.ScriptText

    @property
    def scripttext(self):
        """Lower case alias for ScriptText"""
        return self.ScriptText

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Template(self):
        return self.com_object.Template

    @Template.setter
    def Template(self, value):
        self.com_object.Template = value

    @property
    def template(self):
        """Lower case alias for Template"""
        return self.Template

    @template.setter
    def template(self, value):
        """Lower case alias for Template.setter"""
        self.Template = value

    @property
    def UseWordMail(self):
        return self.com_object.UseWordMail

    @UseWordMail.setter
    def UseWordMail(self, value):
        self.com_object.UseWordMail = value

    @property
    def usewordmail(self):
        """Lower case alias for UseWordMail"""
        return self.UseWordMail

    @usewordmail.setter
    def usewordmail(self, value):
        """Lower case alias for UseWordMail.setter"""
        self.UseWordMail = value

    @property
    def Version(self):
        return self.com_object.Version

    @Version.setter
    def Version(self, value):
        self.com_object.Version = value

    @property
    def version(self):
        """Lower case alias for Version"""
        return self.Version

    @version.setter
    def version(self, value):
        """Lower case alias for Version.setter"""
        self.Version = value

    def PublishForm(self, Registry=None, Folder=None):
        arguments = com_arguments([unwrap(a) for a in [Registry, Folder]])
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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def FormName(self):
        return self.com_object.FormName

    @FormName.setter
    def FormName(self, value):
        self.com_object.FormName = value

    @property
    def formname(self):
        """Lower case alias for FormName"""
        return self.FormName

    @formname.setter
    def formname(self, value):
        """Lower case alias for FormName.setter"""
        self.FormName = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def detail(self):
        """Lower case alias for Detail"""
        return self.Detail

    @detail.setter
    def detail(self, value):
        """Lower case alias for Detail.setter"""
        self.Detail = value

    @property
    def DisplayName(self):
        return self.com_object.DisplayName

    @property
    def displayname(self):
        """Lower case alias for DisplayName"""
        return self.DisplayName

    @property
    def EnableAutoLayout(self):
        return self.com_object.EnableAutoLayout

    @EnableAutoLayout.setter
    def EnableAutoLayout(self, value):
        self.com_object.EnableAutoLayout = value

    @property
    def enableautolayout(self):
        """Lower case alias for EnableAutoLayout"""
        return self.EnableAutoLayout

    @enableautolayout.setter
    def enableautolayout(self, value):
        """Lower case alias for EnableAutoLayout.setter"""
        self.EnableAutoLayout = value

    @property
    def Form(self):
        return self.com_object.Form

    @property
    def form(self):
        """Lower case alias for Form"""
        return self.Form

    @property
    def FormRegionMode(self):
        return OlFormRegionMode(self.com_object.FormRegionMode)

    @property
    def formregionmode(self):
        """Lower case alias for FormRegionMode"""
        return self.FormRegionMode

    @property
    def Inspector(self):
        return Inspector(self.com_object.Inspector)

    @property
    def inspector(self):
        """Lower case alias for Inspector"""
        return self.Inspector

    @property
    def InternalName(self):
        return self.com_object.InternalName

    @property
    def internalname(self):
        """Lower case alias for InternalName"""
        return self.InternalName

    @property
    def IsExpanded(self):
        return self.com_object.IsExpanded

    @property
    def isexpanded(self):
        """Lower case alias for IsExpanded"""
        return self.IsExpanded

    @property
    def Item(self):
        return self.com_object.Item

    @property
    def item(self):
        """Lower case alias for Item"""
        return self.Item

    @property
    def Language(self):
        return self.com_object.Language

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def SuppressControlReplacement(self):
        return self.com_object.SuppressControlReplacement

    @SuppressControlReplacement.setter
    def SuppressControlReplacement(self, value):
        self.com_object.SuppressControlReplacement = value

    @property
    def suppresscontrolreplacement(self):
        """Lower case alias for SuppressControlReplacement"""
        return self.SuppressControlReplacement

    @suppresscontrolreplacement.setter
    def suppresscontrolreplacement(self, value):
        """Lower case alias for SuppressControlReplacement.setter"""
        self.SuppressControlReplacement = value

    @property
    def Visible(self):
        return self.com_object.Visible

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
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
        arguments = com_arguments([unwrap(a) for a in [Control, PropertyName]])
        self.com_object.SetControlItemProperty(*arguments)

    # Lower case alias for SetControlItemProperty
    def setcontrolitemproperty(self, Control=None, PropertyName=None):
        arguments = [Control, PropertyName]
        return self.SetControlItemProperty(*arguments)


class formregionstartup:

    def __init__(self, formregionstartup=None):
        self.com_object= formregionstartup

    def BeforeFormRegionShow(self, FormRegion=None):
        arguments = com_arguments([unwrap(a) for a in [FormRegion]])
        self.com_object.BeforeFormRegionShow(*arguments)

    # Lower case alias for BeforeFormRegionShow
    def beforeformregionshow(self, FormRegion=None):
        arguments = [FormRegion]
        return self.BeforeFormRegionShow(*arguments)

    def GetFormRegionIcon(self, FormRegionName=None, LCID=None, Icon=None):
        arguments = com_arguments([unwrap(a) for a in [FormRegionName, LCID, Icon]])
        return self.com_object.GetFormRegionIcon(*arguments)

    # Lower case alias for GetFormRegionIcon
    def getformregionicon(self, FormRegionName=None, LCID=None, Icon=None):
        arguments = [FormRegionName, LCID, Icon]
        return self.GetFormRegionIcon(*arguments)

    def GetFormRegionManifest(self, FormRegionName=None, LCID=None):
        arguments = com_arguments([unwrap(a) for a in [FormRegionName, LCID]])
        return self.com_object.GetFormRegionManifest(*arguments)

    # Lower case alias for GetFormRegionManifest
    def getformregionmanifest(self, FormRegionName=None, LCID=None):
        arguments = [FormRegionName, LCID]
        return self.GetFormRegionManifest(*arguments)

    def GetFormRegionStorage(self, FormRegionName=None, Item=None, LCID=None, FormRegionMode=None, FormRegionSize=None):
        arguments = com_arguments([unwrap(a) for a in [FormRegionName, Item, LCID, FormRegionMode, FormRegionSize]])
        return self.com_object.GetFormRegionStorage(*arguments)

    # Lower case alias for GetFormRegionStorage
    def getformregionstorage(self, FormRegionName=None, Item=None, LCID=None, FormRegionMode=None, FormRegionSize=None):
        arguments = [FormRegionName, Item, LCID, FormRegionMode, FormRegionSize]
        return self.GetFormRegionStorage(*arguments)


class frame:

    def __init__(self, frame=None):
        self.com_object= frame


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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def FromRssFeed(self):
        return self.com_object.FromRssFeed

    @FromRssFeed.setter
    def FromRssFeed(self, value):
        self.com_object.FromRssFeed = value

    @property
    def fromrssfeed(self):
        """Lower case alias for FromRssFeed"""
        return self.FromRssFeed

    @fromrssfeed.setter
    def fromrssfeed(self, value):
        """Lower case alias for FromRssFeed.setter"""
        self.FromRssFeed = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for Filter.setter"""
        self.Filter = value

    @property
    def IconPlacement(self):
        return OlIconViewPlacement(self.com_object.IconPlacement)

    @IconPlacement.setter
    def IconPlacement(self, value):
        self.com_object.IconPlacement = value

    @property
    def iconplacement(self):
        """Lower case alias for IconPlacement"""
        return self.IconPlacement

    @iconplacement.setter
    def iconplacement(self, value):
        """Lower case alias for IconPlacement.setter"""
        self.IconPlacement = value

    @property
    def IconViewType(self):
        return OlIconViewType(self.com_object.IconViewType)

    @IconViewType.setter
    def IconViewType(self, value):
        self.com_object.IconViewType = value

    @property
    def iconviewtype(self):
        """Lower case alias for IconViewType"""
        return self.IconViewType

    @iconviewtype.setter
    def iconviewtype(self, value):
        """Lower case alias for IconViewType.setter"""
        self.IconViewType = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @language.setter
    def language(self, value):
        """Lower case alias for Language.setter"""
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    @property
    def lockuserchanges(self):
        """Lower case alias for LockUserChanges"""
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        """Lower case alias for LockUserChanges.setter"""
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    @property
    def saveoption(self):
        """Lower case alias for SaveOption"""
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def SortFields(self):
        return OrderFields(self.com_object.SortFields)

    @property
    def sortfields(self):
        """Lower case alias for SortFields"""
        return self.SortFields

    @property
    def Standard(self):
        return IconView(self.com_object.Standard)

    @property
    def standard(self):
        """Lower case alias for Standard"""
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    @property
    def xml(self):
        """Lower case alias for XML"""
        return self.XML

    @xml.setter
    def xml(self, value):
        """Lower case alias for XML.setter"""
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SaveOption]])
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
        arguments = com_arguments([unwrap(a) for a in [Date]])
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


class image:

    def __init__(self, image=None):
        self.com_object= image


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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def attachmentselection(self):
        """Lower case alias for AttachmentSelection"""
        return self.AttachmentSelection

    @property
    def Caption(self):
        return self.com_object.Caption

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentItem(self):
        return self.com_object.CurrentItem

    @property
    def currentitem(self):
        """Lower case alias for CurrentItem"""
        return self.CurrentItem

    @property
    def EditorType(self):
        return OlEditorType(self.com_object.EditorType)

    @property
    def editortype(self):
        """Lower case alias for EditorType"""
        return self.EditorType

    @property
    def Height(self):
        return self.com_object.Height

    @Height.setter
    def Height(self, value):
        self.com_object.Height = value

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @height.setter
    def height(self, value):
        """Lower case alias for Height.setter"""
        self.Height = value

    @property
    def Left(self):
        return self.com_object.Left

    @Left.setter
    def Left(self, value):
        self.com_object.Left = value

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

    @left.setter
    def left(self, value):
        """Lower case alias for Left.setter"""
        self.Left = value

    @property
    def ModifiedFormPages(self):
        return Pages(self.com_object.ModifiedFormPages)

    @property
    def modifiedformpages(self):
        """Lower case alias for ModifiedFormPages"""
        return self.ModifiedFormPages

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Top(self):
        return self.com_object.Top

    @Top.setter
    def Top(self, value):
        self.com_object.Top = value

    @property
    def top(self):
        """Lower case alias for Top"""
        return self.Top

    @top.setter
    def top(self, value):
        """Lower case alias for Top.setter"""
        self.Top = value

    @property
    def Width(self):
        return self.com_object.Width

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    @property
    def width(self):
        """Lower case alias for Width"""
        return self.Width

    @width.setter
    def width(self, value):
        """Lower case alias for Width.setter"""
        self.Width = value

    @property
    def WindowState(self):
        return OlWindowState(self.com_object.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.com_object.WindowState = value

    @property
    def windowstate(self):
        """Lower case alias for WindowState"""
        return self.WindowState

    @windowstate.setter
    def windowstate(self, value):
        """Lower case alias for WindowState.setter"""
        self.WindowState = value

    @property
    def WordEditor(self):
        return self.com_object.WordEditor

    @property
    def wordeditor(self):
        """Lower case alias for WordEditor"""
        return self.WordEditor

    def Activate(self):
        self.com_object.Activate()

    # Lower case alias for Activate
    def activate(self):
        return self.Activate()

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
        self.com_object.Close(*arguments)

    # Lower case alias for Close
    def close(self, SaveMode=None):
        arguments = [SaveMode]
        return self.Close(*arguments)

    def Display(self, Modal=None):
        arguments = com_arguments([unwrap(a) for a in [Modal]])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def HideFormPage(self, PageName=None):
        arguments = com_arguments([unwrap(a) for a in [PageName]])
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
        return self.com_object.NewFormRegion()

    # Lower case alias for NewFormRegion
    def newformregion(self):
        return self.NewFormRegion()

    def OpenFormRegion(self, Path=None):
        arguments = com_arguments([unwrap(a) for a in [Path]])
        return self.com_object.OpenFormRegion(*arguments)

    # Lower case alias for OpenFormRegion
    def openformregion(self, Path=None):
        arguments = [Path]
        return self.OpenFormRegion(*arguments)

    def SaveFormRegion(self, Page=None, FileName=None):
        arguments = com_arguments([unwrap(a) for a in [Page, FileName]])
        self.com_object.SaveFormRegion(*arguments)

    # Lower case alias for SaveFormRegion
    def saveformregion(self, Page=None, FileName=None):
        arguments = [Page, FileName]
        return self.SaveFormRegion(*arguments)

    def SetControlItemProperty(self, Control=None, PropertyName=None):
        arguments = com_arguments([unwrap(a) for a in [Control, PropertyName]])
        self.com_object.SetControlItemProperty(*arguments)

    # Lower case alias for SetControlItemProperty
    def setcontrolitemproperty(self, Control=None, PropertyName=None):
        arguments = [Control, PropertyName]
        return self.SetControlItemProperty(*arguments)

    def SetCurrentFormPage(self, PageName=None):
        arguments = com_arguments([unwrap(a) for a in [PageName]])
        self.com_object.SetCurrentFormPage(*arguments)

    # Lower case alias for SetCurrentFormPage
    def setcurrentformpage(self, PageName=None):
        arguments = [PageName]
        return self.SetCurrentFormPage(*arguments)

    def SetSchedulingStartTime(self, Start=None):
        arguments = com_arguments([unwrap(a) for a in [Start]])
        self.com_object.SetSchedulingStartTime(*arguments)

    # Lower case alias for SetSchedulingStartTime
    def setschedulingstarttime(self, Start=None):
        arguments = [Start]
        return self.SetSchedulingStartTime(*arguments)

    def ShowFormPage(self, PageName=None):
        arguments = com_arguments([unwrap(a) for a in [PageName]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self):
        return Inspector(self.com_object.Add())

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Type, AddToFolderFields, DisplayFormat]])
        self.com_object.Add(*arguments)

    # Lower case alias for Add
    def add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = [Name, Type, AddToFolderFields, DisplayFormat]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ItemProperty(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def isuserproperty(self):
        """Lower case alias for IsUserProperty"""
        return self.IsUserProperty

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.com_object.Type)

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def IncludeRecurrences(self):
        return Items(self.com_object.IncludeRecurrences)

    @IncludeRecurrences.setter
    def IncludeRecurrences(self, value):
        self.com_object.IncludeRecurrences = value

    @property
    def includerecurrences(self):
        """Lower case alias for IncludeRecurrences"""
        return self.IncludeRecurrences

    @includerecurrences.setter
    def includerecurrences(self, value):
        """Lower case alias for IncludeRecurrences.setter"""
        self.IncludeRecurrences = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self):
        return self.com_object.Add()

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Find(self, Filter=None):
        arguments = com_arguments([unwrap(a) for a in [Filter]])
        return self.com_object.Find(*arguments)

    # Lower case alias for Find
    def find(self, Filter=None):
        arguments = [Filter]
        return self.Find(*arguments)

    def FindNext(self):
        return self.com_object.FindNext()

    # Lower case alias for FindNext
    def findnext(self):
        return self.FindNext()

    def GetFirst(self):
        return self.com_object.GetFirst()

    # Lower case alias for GetFirst
    def getfirst(self):
        return self.GetFirst()

    def GetLast(self):
        return self.com_object.GetLast()

    # Lower case alias for GetLast
    def getlast(self):
        return self.GetLast()

    def GetNext(self):
        return self.com_object.GetNext()

    # Lower case alias for GetNext
    def getnext(self):
        return self.GetNext()

    def GetPrevious(self):
        return self.com_object.GetPrevious()

    # Lower case alias for GetPrevious
    def getprevious(self):
        return self.GetPrevious()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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
        arguments = com_arguments([unwrap(a) for a in [Filter]])
        return Items(self.com_object.Restrict(*arguments))

    # Lower case alias for Restrict
    def restrict(self, Filter=None):
        arguments = [Filter]
        return self.Restrict(*arguments)

    def SetColumns(self, Columns=None):
        arguments = com_arguments([unwrap(a) for a in [Columns]])
        self.com_object.SetColumns(*arguments)

    # Lower case alias for SetColumns
    def setcolumns(self, Columns=None):
        arguments = [Columns]
        return self.SetColumns(*arguments)

    def Sort(self, Property=None, Descending=None):
        arguments = com_arguments([unwrap(a) for a in [Property, Descending]])
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return Conflicts(self.com_object.Conflicts)

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ContactNames(self):
        return self.com_object.ContactNames

    @ContactNames.setter
    def ContactNames(self, value):
        self.com_object.ContactNames = value

    @property
    def contactnames(self):
        """Lower case alias for ContactNames"""
        return self.ContactNames

    @contactnames.setter
    def contactnames(self, value):
        """Lower case alias for ContactNames.setter"""
        self.ContactNames = value

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DocPosted(self):
        return self.com_object.DocPosted

    @DocPosted.setter
    def DocPosted(self, value):
        self.com_object.DocPosted = value

    @property
    def docposted(self):
        """Lower case alias for DocPosted"""
        return self.DocPosted

    @docposted.setter
    def docposted(self, value):
        """Lower case alias for DocPosted.setter"""
        self.DocPosted = value

    @property
    def DocPrinted(self):
        return self.com_object.DocPrinted

    @DocPrinted.setter
    def DocPrinted(self, value):
        self.com_object.DocPrinted = value

    @property
    def docprinted(self):
        """Lower case alias for DocPrinted"""
        return self.DocPrinted

    @docprinted.setter
    def docprinted(self, value):
        """Lower case alias for DocPrinted.setter"""
        self.DocPrinted = value

    @property
    def DocRouted(self):
        return self.com_object.DocRouted

    @DocRouted.setter
    def DocRouted(self, value):
        self.com_object.DocRouted = value

    @property
    def docrouted(self):
        """Lower case alias for DocRouted"""
        return self.DocRouted

    @docrouted.setter
    def docrouted(self, value):
        """Lower case alias for DocRouted.setter"""
        self.DocRouted = value

    @property
    def DocSaved(self):
        return self.com_object.DocSaved

    @DocSaved.setter
    def DocSaved(self, value):
        self.com_object.DocSaved = value

    @property
    def docsaved(self):
        """Lower case alias for DocSaved"""
        return self.DocSaved

    @docsaved.setter
    def docsaved(self, value):
        """Lower case alias for DocSaved.setter"""
        self.DocSaved = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def Duration(self):
        return JournalItem(self.com_object.Duration)

    @Duration.setter
    def Duration(self, value):
        self.com_object.Duration = value

    @property
    def duration(self):
        """Lower case alias for Duration"""
        return self.Duration

    @duration.setter
    def duration(self, value):
        """Lower case alias for Duration.setter"""
        self.Duration = value

    @property
    def End(self):
        return self.com_object.End

    @End.setter
    def End(self, value):
        self.com_object.End = value

    @property
    def end(self):
        """Lower case alias for End"""
        return self.End

    @end.setter
    def end(self, value):
        """Lower case alias for End.setter"""
        self.End = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Start(self):
        return self.com_object.Start

    @Start.setter
    def Start(self, value):
        self.com_object.Start = value

    @property
    def start(self):
        """Lower case alias for Start"""
        return self.Start

    @start.setter
    def start(self, value):
        """Lower case alias for Start.setter"""
        self.Start = value

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def Type(self):
        return self.com_object.Type

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @type.setter
    def type(self, value):
        """Lower case alias for Type.setter"""
        self.Type = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    @property
    def navigationgroups(self):
        """Lower case alias for NavigationGroups"""
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    @property
    def navigationmoduletype(self):
        """Lower case alias for NavigationModuleType"""
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return JournalModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return JournalModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
        self.Visible = value


class label:

    def __init__(self, label=None):
        self.com_object= label


class listbox:

    def __init__(self, listbox=None):
        self.com_object= listbox


class MailItem:

    def __init__(self, mailitem=None):
        self.com_object= mailitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def AlternateRecipientAllowed(self):
        return self.com_object.AlternateRecipientAllowed

    @AlternateRecipientAllowed.setter
    def AlternateRecipientAllowed(self, value):
        self.com_object.AlternateRecipientAllowed = value

    @property
    def alternaterecipientallowed(self):
        """Lower case alias for AlternateRecipientAllowed"""
        return self.AlternateRecipientAllowed

    @alternaterecipientallowed.setter
    def alternaterecipientallowed(self, value):
        """Lower case alias for AlternateRecipientAllowed.setter"""
        self.AlternateRecipientAllowed = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.com_object.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.com_object.AutoForwarded = value

    @property
    def autoforwarded(self):
        """Lower case alias for AutoForwarded"""
        return self.AutoForwarded

    @autoforwarded.setter
    def autoforwarded(self, value):
        """Lower case alias for AutoForwarded.setter"""
        self.AutoForwarded = value

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BCC(self):
        return MailItem(self.com_object.BCC)

    @BCC.setter
    def BCC(self, value):
        self.com_object.BCC = value

    @property
    def bcc(self):
        """Lower case alias for BCC"""
        return self.BCC

    @bcc.setter
    def bcc(self, value):
        """Lower case alias for BCC.setter"""
        self.BCC = value

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.com_object.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.com_object.BodyFormat = value

    @property
    def bodyformat(self):
        """Lower case alias for BodyFormat"""
        return self.BodyFormat

    @bodyformat.setter
    def bodyformat(self, value):
        """Lower case alias for BodyFormat.setter"""
        self.BodyFormat = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
        self.Categories = value

    @property
    def CC(self):
        return MailItem(self.com_object.CC)

    @CC.setter
    def CC(self, value):
        self.com_object.CC = value

    @property
    def cc(self):
        """Lower case alias for CC"""
        return self.CC

    @cc.setter
    def cc(self, value):
        """Lower case alias for CC.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return self.com_object.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.com_object.DeferredDeliveryTime = value

    @property
    def deferreddeliverytime(self):
        """Lower case alias for DeferredDeliveryTime"""
        return self.DeferredDeliveryTime

    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        """Lower case alias for DeferredDeliveryTime.setter"""
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.com_object.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.com_object.DeleteAfterSubmit = value

    @property
    def deleteaftersubmit(self):
        """Lower case alias for DeleteAfterSubmit"""
        return self.DeleteAfterSubmit

    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        """Lower case alias for DeleteAfterSubmit.setter"""
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.com_object.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.com_object.ExpiryTime = value

    @property
    def expirytime(self):
        """Lower case alias for ExpiryTime"""
        return self.ExpiryTime

    @expirytime.setter
    def expirytime(self, value):
        """Lower case alias for ExpiryTime.setter"""
        self.ExpiryTime = value

    @property
    def FlagRequest(self):
        return self.com_object.FlagRequest

    @FlagRequest.setter
    def FlagRequest(self, value):
        self.com_object.FlagRequest = value

    @property
    def flagrequest(self):
        """Lower case alias for FlagRequest"""
        return self.FlagRequest

    @flagrequest.setter
    def flagrequest(self, value):
        """Lower case alias for FlagRequest.setter"""
        self.FlagRequest = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def HTMLBody(self):
        return self.com_object.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.com_object.HTMLBody = value

    @property
    def htmlbody(self):
        """Lower case alias for HTMLBody"""
        return self.HTMLBody

    @htmlbody.setter
    def htmlbody(self, value):
        """Lower case alias for HTMLBody.setter"""
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    @property
    def internetcodepage(self):
        """Lower case alias for InternetCodepage"""
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        """Lower case alias for InternetCodepage.setter"""
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return MailItem(self.com_object.IsMarkedAsTask)

    @property
    def ismarkedastask(self):
        """Lower case alias for IsMarkedAsTask"""
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return self.com_object.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.com_object.OriginatorDeliveryReportRequested = value

    @property
    def originatordeliveryreportrequested(self):
        """Lower case alias for OriginatorDeliveryReportRequested"""
        return self.OriginatorDeliveryReportRequested

    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        """Lower case alias for OriginatorDeliveryReportRequested.setter"""
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Permission(self):
        return self.com_object.Permission

    @Permission.setter
    def Permission(self, value):
        self.com_object.Permission = value

    @property
    def permission(self):
        """Lower case alias for Permission"""
        return self.Permission

    @permission.setter
    def permission(self, value):
        """Lower case alias for Permission.setter"""
        self.Permission = value

    @property
    def PermissionService(self):
        return self.com_object.PermissionService

    @PermissionService.setter
    def PermissionService(self, value):
        self.com_object.PermissionService = value

    @property
    def permissionservice(self):
        """Lower case alias for PermissionService"""
        return self.PermissionService

    @permissionservice.setter
    def permissionservice(self, value):
        """Lower case alias for PermissionService.setter"""
        self.PermissionService = value

    @property
    def PermissionTemplateGuid(self):
        return MailItem(self.com_object.PermissionTemplateGuid)

    @PermissionTemplateGuid.setter
    def PermissionTemplateGuid(self, value):
        self.com_object.PermissionTemplateGuid = value

    @property
    def permissiontemplateguid(self):
        """Lower case alias for PermissionTemplateGuid"""
        return self.PermissionTemplateGuid

    @permissiontemplateguid.setter
    def permissiontemplateguid(self, value):
        """Lower case alias for PermissionTemplateGuid.setter"""
        self.PermissionTemplateGuid = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def ReadReceiptRequested(self):
        return self.com_object.ReadReceiptRequested

    @property
    def readreceiptrequested(self):
        """Lower case alias for ReadReceiptRequested"""
        return self.ReadReceiptRequested

    @property
    def ReceivedByEntryID(self):
        return self.com_object.ReceivedByEntryID

    @property
    def receivedbyentryid(self):
        """Lower case alias for ReceivedByEntryID"""
        return self.ReceivedByEntryID

    @property
    def ReceivedByName(self):
        return self.com_object.ReceivedByName

    @property
    def receivedbyname(self):
        """Lower case alias for ReceivedByName"""
        return self.ReceivedByName

    @property
    def ReceivedOnBehalfOfEntryID(self):
        return self.com_object.ReceivedOnBehalfOfEntryID

    @property
    def receivedonbehalfofentryid(self):
        """Lower case alias for ReceivedOnBehalfOfEntryID"""
        return self.ReceivedOnBehalfOfEntryID

    @property
    def ReceivedOnBehalfOfName(self):
        return self.com_object.ReceivedOnBehalfOfName

    @property
    def receivedonbehalfofname(self):
        """Lower case alias for ReceivedOnBehalfOfName"""
        return self.ReceivedOnBehalfOfName

    @property
    def ReceivedTime(self):
        return self.com_object.ReceivedTime

    @property
    def receivedtime(self):
        """Lower case alias for ReceivedTime"""
        return self.ReceivedTime

    @property
    def RecipientReassignmentProhibited(self):
        return self.com_object.RecipientReassignmentProhibited

    @RecipientReassignmentProhibited.setter
    def RecipientReassignmentProhibited(self, value):
        self.com_object.RecipientReassignmentProhibited = value

    @property
    def recipientreassignmentprohibited(self):
        """Lower case alias for RecipientReassignmentProhibited"""
        return self.RecipientReassignmentProhibited

    @recipientreassignmentprohibited.setter
    def recipientreassignmentprohibited(self, value):
        """Lower case alias for RecipientReassignmentProhibited.setter"""
        self.RecipientReassignmentProhibited = value

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    @property
    def reminderoverridedefault(self):
        """Lower case alias for ReminderOverrideDefault"""
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        """Lower case alias for ReminderOverrideDefault.setter"""
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    @property
    def reminderplaysound(self):
        """Lower case alias for ReminderPlaySound"""
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        """Lower case alias for ReminderPlaySound.setter"""
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    @property
    def reminderset(self):
        """Lower case alias for ReminderSet"""
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        """Lower case alias for ReminderSet.setter"""
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    @property
    def remindersoundfile(self):
        """Lower case alias for ReminderSoundFile"""
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        """Lower case alias for ReminderSoundFile.setter"""
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    @property
    def remindertime(self):
        """Lower case alias for ReminderTime"""
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        """Lower case alias for ReminderTime.setter"""
        self.ReminderTime = value

    @property
    def RemoteStatus(self):
        return OlRemoteStatus(self.com_object.RemoteStatus)

    @RemoteStatus.setter
    def RemoteStatus(self, value):
        self.com_object.RemoteStatus = value

    @property
    def remotestatus(self):
        """Lower case alias for RemoteStatus"""
        return self.RemoteStatus

    @remotestatus.setter
    def remotestatus(self, value):
        """Lower case alias for RemoteStatus.setter"""
        self.RemoteStatus = value

    @property
    def ReplyRecipientNames(self):
        return self.com_object.ReplyRecipientNames

    @property
    def replyrecipientnames(self):
        """Lower case alias for ReplyRecipientNames"""
        return self.ReplyRecipientNames

    @property
    def ReplyRecipients(self):
        return Recipients(self.com_object.ReplyRecipients)

    @property
    def replyrecipients(self):
        """Lower case alias for ReplyRecipients"""
        return self.ReplyRecipients

    @property
    def RetentionExpirationDate(self):
        return MailItem(self.com_object.RetentionExpirationDate)

    @property
    def retentionexpirationdate(self):
        """Lower case alias for RetentionExpirationDate"""
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.com_object.RetentionPolicyName

    @property
    def retentionpolicyname(self):
        """Lower case alias for RetentionPolicyName"""
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return Folder(self.com_object.SaveSentMessageFolder)

    @SaveSentMessageFolder.setter
    def SaveSentMessageFolder(self, value):
        self.com_object.SaveSentMessageFolder = value

    @property
    def savesentmessagefolder(self):
        """Lower case alias for SaveSentMessageFolder"""
        return self.SaveSentMessageFolder

    @savesentmessagefolder.setter
    def savesentmessagefolder(self, value):
        """Lower case alias for SaveSentMessageFolder.setter"""
        self.SaveSentMessageFolder = value

    @property
    def Sender(self):
        return self.com_object.Sender

    @Sender.setter
    def Sender(self, value):
        self.com_object.Sender = value

    @property
    def sender(self):
        """Lower case alias for Sender"""
        return self.Sender

    @sender.setter
    def sender(self, value):
        """Lower case alias for Sender.setter"""
        self.Sender = value

    @property
    def SenderEmailAddress(self):
        return self.com_object.SenderEmailAddress

    @property
    def senderemailaddress(self):
        """Lower case alias for SenderEmailAddress"""
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.com_object.SenderEmailType

    @property
    def senderemailtype(self):
        """Lower case alias for SenderEmailType"""
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.com_object.SenderName

    @property
    def sendername(self):
        """Lower case alias for SenderName"""
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    @property
    def sendusingaccount(self):
        """Lower case alias for SendUsingAccount"""
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        """Lower case alias for SendUsingAccount.setter"""
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Sent(self):
        return self.com_object.Sent

    @property
    def sent(self):
        """Lower case alias for Sent"""
        return self.Sent

    @property
    def SentOn(self):
        return self.com_object.SentOn

    @property
    def senton(self):
        """Lower case alias for SentOn"""
        return self.SentOn

    @property
    def SentOnBehalfOfName(self):
        return self.com_object.SentOnBehalfOfName

    @SentOnBehalfOfName.setter
    def SentOnBehalfOfName(self, value):
        self.com_object.SentOnBehalfOfName = value

    @property
    def sentonbehalfofname(self):
        """Lower case alias for SentOnBehalfOfName"""
        return self.SentOnBehalfOfName

    @sentonbehalfofname.setter
    def sentonbehalfofname(self, value):
        """Lower case alias for SentOnBehalfOfName.setter"""
        self.SentOnBehalfOfName = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def Submitted(self):
        return self.com_object.Submitted

    @property
    def submitted(self):
        """Lower case alias for Submitted"""
        return self.Submitted

    @property
    def TaskCompletedDate(self):
        return MailItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    @property
    def taskcompleteddate(self):
        """Lower case alias for TaskCompletedDate"""
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        """Lower case alias for TaskCompletedDate.setter"""
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return MailItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    @property
    def taskduedate(self):
        """Lower case alias for TaskDueDate"""
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        """Lower case alias for TaskDueDate.setter"""
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return MailItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    @property
    def taskstartdate(self):
        """Lower case alias for TaskStartDate"""
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        """Lower case alias for TaskStartDate.setter"""
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return MailItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    @property
    def tasksubject(self):
        """Lower case alias for TaskSubject"""
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        """Lower case alias for TaskSubject.setter"""
        self.TaskSubject = value

    @property
    def To(self):
        return self.com_object.To

    @To.setter
    def To(self, value):
        self.com_object.To = value

    @property
    def to(self):
        """Lower case alias for To"""
        return self.To

    @to.setter
    def to(self, value):
        """Lower case alias for To.setter"""
        self.To = value

    @property
    def ToDoTaskOrdinal(self):
        return MailItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    @property
    def todotaskordinal(self):
        """Lower case alias for ToDoTaskOrdinal"""
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        """Lower case alias for ToDoTaskOrdinal.setter"""
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    @property
    def VotingOptions(self):
        return self.com_object.VotingOptions

    @VotingOptions.setter
    def VotingOptions(self, value):
        self.com_object.VotingOptions = value

    @property
    def votingoptions(self):
        """Lower case alias for VotingOptions"""
        return self.VotingOptions

    @votingoptions.setter
    def votingoptions(self, value):
        """Lower case alias for VotingOptions.setter"""
        self.VotingOptions = value

    @property
    def VotingResponse(self):
        return self.com_object.VotingResponse

    @VotingResponse.setter
    def VotingResponse(self, value):
        self.com_object.VotingResponse = value

    @property
    def votingresponse(self):
        """Lower case alias for VotingResponse"""
        return self.VotingResponse

    @votingresponse.setter
    def votingresponse(self, value):
        """Lower case alias for VotingResponse.setter"""
        self.VotingResponse = value

    def AddBusinessCard(self, contact=None):
        arguments = com_arguments([unwrap(a) for a in [contact]])
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
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [MarkInterval]])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    @property
    def navigationgroups(self):
        """Lower case alias for NavigationGroups"""
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    @property
    def navigationmoduletype(self):
        """Lower case alias for NavigationModuleType"""
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return MailModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return MailModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
        self.Visible = value


class MarkAsTaskRuleAction:

    def __init__(self, markastaskruleaction=None):
        self.com_object= markastaskruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    @property
    def actiontype(self):
        """Lower case alias for ActionType"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def FlagTo(self):
        return self.com_object.FlagTo

    @FlagTo.setter
    def FlagTo(self, value):
        self.com_object.FlagTo = value

    @property
    def flagto(self):
        """Lower case alias for FlagTo"""
        return self.FlagTo

    @flagto.setter
    def flagto(self, value):
        """Lower case alias for FlagTo.setter"""
        self.FlagTo = value

    @property
    def MarkInterval(self):
        return OlMarkInterval(self.com_object.MarkInterval)

    @MarkInterval.setter
    def MarkInterval(self, value):
        self.com_object.MarkInterval = value

    @property
    def markinterval(self):
        """Lower case alias for MarkInterval"""
        return self.MarkInterval

    @markinterval.setter
    def markinterval(self, value):
        """Lower case alias for MarkInterval.setter"""
        self.MarkInterval = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class MeetingItem:

    def __init__(self, meetingitem=None):
        self.com_object= meetingitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.com_object.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.com_object.AutoForwarded = value

    @property
    def autoforwarded(self):
        """Lower case alias for AutoForwarded"""
        return self.AutoForwarded

    @autoforwarded.setter
    def autoforwarded(self, value):
        """Lower case alias for AutoForwarded.setter"""
        self.AutoForwarded = value

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return self.com_object.DeferredDeliveryTime

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.com_object.DeferredDeliveryTime = value

    @property
    def deferreddeliverytime(self):
        """Lower case alias for DeferredDeliveryTime"""
        return self.DeferredDeliveryTime

    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        """Lower case alias for DeferredDeliveryTime.setter"""
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.com_object.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.com_object.DeleteAfterSubmit = value

    @property
    def deleteaftersubmit(self):
        """Lower case alias for DeleteAfterSubmit"""
        return self.DeleteAfterSubmit

    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        """Lower case alias for DeleteAfterSubmit.setter"""
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.com_object.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.com_object.ExpiryTime = value

    @property
    def expirytime(self):
        """Lower case alias for ExpiryTime"""
        return self.ExpiryTime

    @expirytime.setter
    def expirytime(self, value):
        """Lower case alias for ExpiryTime.setter"""
        self.ExpiryTime = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def IsLatestVersion(self):
        return MeetingItem(self.com_object.IsLatestVersion)

    @property
    def islatestversion(self):
        """Lower case alias for IsLatestVersion"""
        return self.IsLatestVersion

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MeetingWorkspaceURL(self):
        return self.com_object.MeetingWorkspaceURL

    @property
    def meetingworkspaceurl(self):
        """Lower case alias for MeetingWorkspaceURL"""
        return self.MeetingWorkspaceURL

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return self.com_object.OriginatorDeliveryReportRequested

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.com_object.OriginatorDeliveryReportRequested = value

    @property
    def originatordeliveryreportrequested(self):
        """Lower case alias for OriginatorDeliveryReportRequested"""
        return self.OriginatorDeliveryReportRequested

    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        """Lower case alias for OriginatorDeliveryReportRequested.setter"""
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def ReceivedTime(self):
        return self.com_object.ReceivedTime

    @ReceivedTime.setter
    def ReceivedTime(self, value):
        self.com_object.ReceivedTime = value

    @property
    def receivedtime(self):
        """Lower case alias for ReceivedTime"""
        return self.ReceivedTime

    @receivedtime.setter
    def receivedtime(self, value):
        """Lower case alias for ReceivedTime.setter"""
        self.ReceivedTime = value

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    @property
    def reminderset(self):
        """Lower case alias for ReminderSet"""
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        """Lower case alias for ReminderSet.setter"""
        self.ReminderSet = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    @property
    def remindertime(self):
        """Lower case alias for ReminderTime"""
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        """Lower case alias for ReminderTime.setter"""
        self.ReminderTime = value

    @property
    def ReplyRecipients(self):
        return Recipients(self.com_object.ReplyRecipients)

    @property
    def replyrecipients(self):
        """Lower case alias for ReplyRecipients"""
        return self.ReplyRecipients

    @property
    def RetentionExpirationDate(self):
        return MeetingItem(self.com_object.RetentionExpirationDate)

    @property
    def retentionexpirationdate(self):
        """Lower case alias for RetentionExpirationDate"""
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.com_object.RetentionPolicyName

    @property
    def retentionpolicyname(self):
        """Lower case alias for RetentionPolicyName"""
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return self.com_object.SaveSentMessageFolder

    @property
    def savesentmessagefolder(self):
        """Lower case alias for SaveSentMessageFolder"""
        return self.SaveSentMessageFolder

    @property
    def SenderEmailAddress(self):
        return self.com_object.SenderEmailAddress

    @property
    def senderemailaddress(self):
        """Lower case alias for SenderEmailAddress"""
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.com_object.SenderEmailType

    @property
    def senderemailtype(self):
        """Lower case alias for SenderEmailType"""
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.com_object.SenderName

    @property
    def sendername(self):
        """Lower case alias for SenderName"""
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    @property
    def sendusingaccount(self):
        """Lower case alias for SendUsingAccount"""
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        """Lower case alias for SendUsingAccount.setter"""
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Sent(self):
        return self.com_object.Sent

    @property
    def sent(self):
        """Lower case alias for Sent"""
        return self.Sent

    @property
    def SentOn(self):
        return self.com_object.SentOn

    @property
    def senton(self):
        """Lower case alias for SentOn"""
        return self.SentOn

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def Submitted(self):
        return self.com_object.Submitted

    @property
    def submitted(self):
        """Lower case alias for Submitted"""
        return self.Submitted

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [AddToCalendar]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def actiontype(self):
        """Lower case alias for ActionType"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Folder(self):
        return Folder(self.com_object.Folder)

    @Folder.setter
    def Folder(self, value):
        self.com_object.Folder = value

    @property
    def folder(self):
        """Lower case alias for Folder"""
        return self.Folder

    @folder.setter
    def folder(self, value):
        """Lower case alias for Folder.setter"""
        self.Folder = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class multipage:

    def __init__(self, multipage=None):
        self.com_object= multipage


class NameSpace:

    def __init__(self, namespace=None):
        self.com_object= namespace

    @property
    def Accounts(self):
        return Accounts(self.com_object.Accounts)

    @property
    def accounts(self):
        """Lower case alias for Accounts"""
        return self.Accounts

    @property
    def AddressLists(self):
        return AddressLists(self.com_object.AddressLists)

    @property
    def addresslists(self):
        """Lower case alias for AddressLists"""
        return self.AddressLists

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoDiscoverConnectionMode(self):
        return OlAutoDiscoverConnectionMode(self.com_object.AutoDiscoverConnectionMode)

    @property
    def autodiscoverconnectionmode(self):
        """Lower case alias for AutoDiscoverConnectionMode"""
        return self.AutoDiscoverConnectionMode

    @property
    def AutoDiscoverXml(self):
        return self.com_object.AutoDiscoverXml

    @property
    def autodiscoverxml(self):
        """Lower case alias for AutoDiscoverXml"""
        return self.AutoDiscoverXml

    @property
    def Categories(self):
        return Categories(self.com_object.Categories)

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CurrentProfileName(self):
        return self.com_object.CurrentProfileName

    @property
    def currentprofilename(self):
        """Lower case alias for CurrentProfileName"""
        return self.CurrentProfileName

    @property
    def CurrentUser(self):
        return Recipient(self.com_object.CurrentUser)

    @property
    def currentuser(self):
        """Lower case alias for CurrentUser"""
        return self.CurrentUser

    @property
    def DefaultStore(self):
        return Store(self.com_object.DefaultStore)

    @property
    def defaultstore(self):
        """Lower case alias for DefaultStore"""
        return self.DefaultStore

    @property
    def ExchangeConnectionMode(self):
        return OlExchangeConnectionMode(self.com_object.ExchangeConnectionMode)

    @property
    def exchangeconnectionmode(self):
        """Lower case alias for ExchangeConnectionMode"""
        return self.ExchangeConnectionMode

    @property
    def ExchangeMailboxServerName(self):
        return self.com_object.ExchangeMailboxServerName

    @property
    def exchangemailboxservername(self):
        """Lower case alias for ExchangeMailboxServerName"""
        return self.ExchangeMailboxServerName

    @property
    def ExchangeMailboxServerVersion(self):
        return self.com_object.ExchangeMailboxServerVersion

    @property
    def exchangemailboxserverversion(self):
        """Lower case alias for ExchangeMailboxServerVersion"""
        return self.ExchangeMailboxServerVersion

    @property
    def Folders(self):
        return Folders(self.com_object.Folders)

    @property
    def folders(self):
        """Lower case alias for Folders"""
        return self.Folders

    @property
    def Offline(self):
        return self.com_object.Offline

    @property
    def offline(self):
        """Lower case alias for Offline"""
        return self.Offline

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Stores(self):
        return Stores(self.com_object.Stores)

    @property
    def stores(self):
        """Lower case alias for Stores"""
        return self.Stores

    @property
    def SyncObjects(self):
        return SyncObjects(self.com_object.SyncObjects)

    @property
    def syncobjects(self):
        """Lower case alias for SyncObjects"""
        return self.SyncObjects

    @property
    def Type(self):
        return self.com_object.Type

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    def AddStore(self, Store=None):
        arguments = com_arguments([unwrap(a) for a in [Store]])
        self.com_object.AddStore(*arguments)

    # Lower case alias for AddStore
    def addstore(self, Store=None):
        arguments = [Store]
        return self.AddStore(*arguments)

    def AddStoreEx(self, Store=None, Type=None):
        arguments = com_arguments([unwrap(a) for a in [Store, Type]])
        self.com_object.AddStoreEx(*arguments)

    # Lower case alias for AddStoreEx
    def addstoreex(self, Store=None, Type=None):
        arguments = [Store, Type]
        return self.AddStoreEx(*arguments)

    def CompareEntryIDs(self, FirstEntryID=None, SecondEntryID=None):
        arguments = com_arguments([unwrap(a) for a in [FirstEntryID, SecondEntryID]])
        return self.com_object.CompareEntryIDs(*arguments)

    # Lower case alias for CompareEntryIDs
    def compareentryids(self, FirstEntryID=None, SecondEntryID=None):
        arguments = [FirstEntryID, SecondEntryID]
        return self.CompareEntryIDs(*arguments)

    def CreateContactCard(self, Address=None):
        arguments = com_arguments([unwrap(a) for a in [Address]])
        return self.com_object.CreateContactCard(*arguments)

    # Lower case alias for CreateContactCard
    def createcontactcard(self, Address=None):
        arguments = [Address]
        return self.CreateContactCard(*arguments)

    def CreateRecipient(self, RecipientName=None):
        arguments = com_arguments([unwrap(a) for a in [RecipientName]])
        return Recipient(self.com_object.CreateRecipient(*arguments))

    # Lower case alias for CreateRecipient
    def createrecipient(self, RecipientName=None):
        arguments = [RecipientName]
        return self.CreateRecipient(*arguments)

    def CreateSharingItem(self, Context=None, Provider=None):
        arguments = com_arguments([unwrap(a) for a in [Context, Provider]])
        return SharingItem(self.com_object.CreateSharingItem(*arguments))

    # Lower case alias for CreateSharingItem
    def createsharingitem(self, Context=None, Provider=None):
        arguments = [Context, Provider]
        return self.CreateSharingItem(*arguments)

    def Dial(self, ContactItem=None):
        arguments = com_arguments([unwrap(a) for a in [ContactItem]])
        self.com_object.Dial(*arguments)

    # Lower case alias for Dial
    def dial(self, ContactItem=None):
        arguments = [ContactItem]
        return self.Dial(*arguments)

    def GetAddressEntryFromID(self, ID=None):
        arguments = com_arguments([unwrap(a) for a in [ID]])
        return self.com_object.GetAddressEntryFromID(*arguments)

    # Lower case alias for GetAddressEntryFromID
    def getaddressentryfromid(self, ID=None):
        arguments = [ID]
        return self.GetAddressEntryFromID(*arguments)

    def GetDefaultFolder(self, FolderType=None):
        arguments = com_arguments([unwrap(a) for a in [FolderType]])
        return Folder(self.com_object.GetDefaultFolder(*arguments))

    # Lower case alias for GetDefaultFolder
    def getdefaultfolder(self, FolderType=None):
        arguments = [FolderType]
        return self.GetDefaultFolder(*arguments)

    def GetFolderFromID(self, EntryIDFolder=None, EntryIDStore=None):
        arguments = com_arguments([unwrap(a) for a in [EntryIDFolder, EntryIDStore]])
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
        arguments = com_arguments([unwrap(a) for a in [EntryIDItem, EntryIDStore]])
        return self.com_object.GetItemFromID(*arguments)

    # Lower case alias for GetItemFromID
    def getitemfromid(self, EntryIDItem=None, EntryIDStore=None):
        arguments = [EntryIDItem, EntryIDStore]
        return self.GetItemFromID(*arguments)

    def GetRecipientFromID(self, EntryID=None):
        arguments = com_arguments([unwrap(a) for a in [EntryID]])
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
        arguments = com_arguments([unwrap(a) for a in [Recipient, FolderType]])
        return Folder(self.com_object.GetSharedDefaultFolder(*arguments))

    # Lower case alias for GetSharedDefaultFolder
    def getshareddefaultfolder(self, Recipient=None, FolderType=None):
        arguments = [Recipient, FolderType]
        return self.GetSharedDefaultFolder(*arguments)

    def GetStoreFromID(self, ID=None):
        arguments = com_arguments([unwrap(a) for a in [ID]])
        return self.com_object.GetStoreFromID(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Profile, Password, ShowDialog, NewSession]])
        self.com_object.Logon(*arguments)

    # Lower case alias for Logon
    def logon(self, Profile=None, Password=None, ShowDialog=None, NewSession=None):
        arguments = [Profile, Password, ShowDialog, NewSession]
        return self.Logon(*arguments)

    def OpenSharedFolder(self, Path=None, Name=None, DownloadAttachments=None, UseTTL=None):
        arguments = com_arguments([unwrap(a) for a in [Path, Name, DownloadAttachments, UseTTL]])
        return Folder(self.com_object.OpenSharedFolder(*arguments))

    # Lower case alias for OpenSharedFolder
    def opensharedfolder(self, Path=None, Name=None, DownloadAttachments=None, UseTTL=None):
        arguments = [Path, Name, DownloadAttachments, UseTTL]
        return self.OpenSharedFolder(*arguments)

    def OpenSharedItem(self, Path=None):
        arguments = com_arguments([unwrap(a) for a in [Path]])
        return self.com_object.OpenSharedItem(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Folder]])
        self.com_object.RemoveStore(*arguments)

    # Lower case alias for RemoveStore
    def removestore(self, Folder=None):
        arguments = [Folder]
        return self.RemoveStore(*arguments)

    def SendAndReceive(self, showProgressDialog=None):
        arguments = com_arguments([unwrap(a) for a in [showProgressDialog]])
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

    @property
    def displayname(self):
        """Lower case alias for DisplayName"""
        return self.DisplayName

    @property
    def Folder(self):
        return Folder(self.com_object.Folder)

    @property
    def folder(self):
        """Lower case alias for Folder"""
        return self.Folder

    @property
    def IsRemovable(self):
        return NavigationFolder(self.com_object.IsRemovable)

    @property
    def isremovable(self):
        """Lower case alias for IsRemovable"""
        return self.IsRemovable

    @property
    def IsSelected(self):
        return NavigationFolder(self.com_object.IsSelected)

    @IsSelected.setter
    def IsSelected(self, value):
        self.com_object.IsSelected = value

    @property
    def isselected(self):
        """Lower case alias for IsSelected"""
        return self.IsSelected

    @isselected.setter
    def isselected(self, value):
        """Lower case alias for IsSelected.setter"""
        self.IsSelected = value

    @property
    def IsSideBySide(self):
        return NavigationFolder(self.com_object.IsSideBySide)

    @IsSideBySide.setter
    def IsSideBySide(self, value):
        self.com_object.IsSideBySide = value

    @property
    def issidebyside(self):
        """Lower case alias for IsSideBySide"""
        return self.IsSideBySide

    @issidebyside.setter
    def issidebyside(self, value):
        """Lower case alias for IsSideBySide.setter"""
        self.IsSideBySide = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return NavigationFolder(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Folder=None):
        arguments = com_arguments([unwrap(a) for a in [Folder]])
        return NavigationFolder(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Folder=None):
        arguments = [Folder]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return NavigationFolder(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, RemovableFolder=None):
        arguments = com_arguments([unwrap(a) for a in [RemovableFolder]])
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

    @property
    def grouptype(self):
        """Lower case alias for GroupType"""
        return self.GroupType

    @property
    def Name(self):
        return NavigationGroup(self.com_object.Name)

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def NavigationFolders(self):
        return NavigationFolders(self.com_object.NavigationFolders)

    @property
    def navigationfolders(self):
        """Lower case alias for NavigationFolders"""
        return self.NavigationFolders

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return NavigationGroup(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Create(self, GroupDisplayName=None):
        arguments = com_arguments([unwrap(a) for a in [GroupDisplayName]])
        return NavigationGroup(self.com_object.Create(*arguments))

    # Lower case alias for Create
    def create(self, GroupDisplayName=None):
        arguments = [GroupDisplayName]
        return self.Create(*arguments)

    def Delete(self, Group=None):
        arguments = com_arguments([unwrap(a) for a in [Group]])
        self.com_object.Delete(*arguments)

    # Lower case alias for Delete
    def delete(self, Group=None):
        arguments = [Group]
        return self.Delete(*arguments)

    def GetDefaultNavigationGroup(self, DefaultFolderGroup=None):
        arguments = com_arguments([unwrap(a) for a in [DefaultFolderGroup]])
        return NavigationGroup(self.com_object.GetDefaultNavigationGroup(*arguments))

    # Lower case alias for GetDefaultNavigationGroup
    def getdefaultnavigationgroup(self, DefaultFolderGroup=None):
        arguments = [DefaultFolderGroup]
        return self.GetDefaultNavigationGroup(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    @property
    def navigationmoduletype(self):
        """Lower case alias for NavigationModuleType"""
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return NavigationModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return NavigationModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def GetNavigationModule(self, ModuleType=None):
        arguments = com_arguments([unwrap(a) for a in [ModuleType]])
        return NavigationModule(self.com_object.GetNavigationModule(*arguments))

    # Lower case alias for GetNavigationModule
    def getnavigationmodule(self, ModuleType=None):
        arguments = [ModuleType]
        return self.GetNavigationModule(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def currentmodule(self):
        """Lower case alias for CurrentModule"""
        return self.CurrentModule

    @currentmodule.setter
    def currentmodule(self, value):
        """Lower case alias for CurrentModule.setter"""
        self.CurrentModule = value

    @property
    def DisplayedModuleCount(self):
        return NavigationModule(self.com_object.DisplayedModuleCount)

    @DisplayedModuleCount.setter
    def DisplayedModuleCount(self, value):
        self.com_object.DisplayedModuleCount = value

    @property
    def displayedmodulecount(self):
        """Lower case alias for DisplayedModuleCount"""
        return self.DisplayedModuleCount

    @displayedmodulecount.setter
    def displayedmodulecount(self, value):
        """Lower case alias for DisplayedModuleCount.setter"""
        self.DisplayedModuleCount = value

    @property
    def IsCollapsed(self):
        return self.com_object.IsCollapsed

    @IsCollapsed.setter
    def IsCollapsed(self, value):
        self.com_object.IsCollapsed = value

    @property
    def iscollapsed(self):
        """Lower case alias for IsCollapsed"""
        return self.IsCollapsed

    @iscollapsed.setter
    def iscollapsed(self, value):
        """Lower case alias for IsCollapsed.setter"""
        self.IsCollapsed = value

    @property
    def Modules(self):
        return NavigationModules(self.com_object.Modules)

    @property
    def modules(self):
        """Lower case alias for Modules"""
        return self.Modules

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class NewItemAlertRuleAction:

    def __init__(self, newitemalertruleaction=None):
        self.com_object= newitemalertruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    @property
    def actiontype(self):
        """Lower case alias for ActionType"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    @property
    def text(self):
        """Lower case alias for Text"""
        return self.Text

    @text.setter
    def text(self, value):
        """Lower case alias for Text.setter"""
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

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
        self.Categories = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Height(self):
        return self.com_object.Height

    @Height.setter
    def Height(self, value):
        self.com_object.Height = value

    @property
    def height(self):
        """Lower case alias for Height"""
        return self.Height

    @height.setter
    def height(self, value):
        """Lower case alias for Height.setter"""
        self.Height = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def Left(self):
        return self.com_object.Left

    @Left.setter
    def Left(self, value):
        self.com_object.Left = value

    @property
    def left(self):
        """Lower case alias for Left"""
        return self.Left

    @left.setter
    def left(self, value):
        """Lower case alias for Left.setter"""
        self.Left = value

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def Top(self):
        return self.com_object.Top

    @Top.setter
    def Top(self, value):
        self.com_object.Top = value

    @property
    def top(self):
        """Lower case alias for Top"""
        return self.Top

    @top.setter
    def top(self, value):
        """Lower case alias for Top.setter"""
        self.Top = value

    @property
    def Width(self):
        return self.com_object.Width

    @Width.setter
    def Width(self, value):
        self.com_object.Width = value

    @property
    def width(self):
        """Lower case alias for Width"""
        return self.Width

    @width.setter
    def width(self, value):
        """Lower case alias for Width.setter"""
        self.Width = value

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    @property
    def navigationgroups(self):
        """Lower case alias for NavigationGroups"""
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return olNavigationModuleType(self.com_object.NavigationModuleType)

    @property
    def navigationmoduletype(self):
        """Lower case alias for NavigationModuleType"""
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return NotesModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return NotesModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
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

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
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

    @property
    def autosize(self):
        """Lower case alias for AutoSize"""
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        """Lower case alias for AutoSize.setter"""
        self.AutoSize = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    @property
    def backstyle(self):
        """Lower case alias for BackStyle"""
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        """Lower case alias for BackStyle.setter"""
        self.BackStyle = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
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

    @property
    def accelerator(self):
        """Lower case alias for Accelerator"""
        return self.Accelerator

    @accelerator.setter
    def accelerator(self, value):
        """Lower case alias for Accelerator.setter"""
        self.Accelerator = value

    @property
    def Alignment(self):
        return olAlignment(self.com_object.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.com_object.Alignment = value

    @property
    def alignment(self):
        """Lower case alias for Alignment"""
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        """Lower case alias for Alignment.setter"""
        self.Alignment = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    @property
    def backstyle(self):
        """Lower case alias for BackStyle"""
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        """Lower case alias for BackStyle.setter"""
        self.BackStyle = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @caption.setter
    def caption(self, value):
        """Lower case alias for Caption.setter"""
        self.Caption = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def TripleState(self):
        return self.com_object.TripleState

    @TripleState.setter
    def TripleState(self, value):
        self.com_object.TripleState = value

    @property
    def triplestate(self):
        """Lower case alias for TripleState"""
        return self.TripleState

    @triplestate.setter
    def triplestate(self, value):
        """Lower case alias for TripleState.setter"""
        self.TripleState = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
        self.Value = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    @property
    def wordwrap(self):
        """Lower case alias for WordWrap"""
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        """Lower case alias for WordWrap.setter"""
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

    @property
    def autosize(self):
        """Lower case alias for AutoSize"""
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        """Lower case alias for AutoSize.setter"""
        self.AutoSize = value

    @property
    def AutoTab(self):
        return self.com_object.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.com_object.AutoTab = value

    @property
    def autotab(self):
        """Lower case alias for AutoTab"""
        return self.AutoTab

    @autotab.setter
    def autotab(self, value):
        """Lower case alias for AutoTab.setter"""
        self.AutoTab = value

    @property
    def AutoWordSelect(self):
        return self.com_object.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.com_object.AutoWordSelect = value

    @property
    def autowordselect(self):
        """Lower case alias for AutoWordSelect"""
        return self.AutoWordSelect

    @autowordselect.setter
    def autowordselect(self, value):
        """Lower case alias for AutoWordSelect.setter"""
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    @property
    def borderstyle(self):
        """Lower case alias for BorderStyle"""
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        """Lower case alias for BorderStyle.setter"""
        self.BorderStyle = value

    @property
    def DragBehavior(self):
        return self.com_object.DragBehavior

    @DragBehavior.setter
    def DragBehavior(self, value):
        self.com_object.DragBehavior = value

    @property
    def dragbehavior(self):
        """Lower case alias for DragBehavior"""
        return self.DragBehavior

    @dragbehavior.setter
    def dragbehavior(self, value):
        """Lower case alias for DragBehavior.setter"""
        self.DragBehavior = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.com_object.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.com_object.EnterFieldBehavior = value

    @property
    def enterfieldbehavior(self):
        """Lower case alias for EnterFieldBehavior"""
        return self.EnterFieldBehavior

    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        """Lower case alias for EnterFieldBehavior.setter"""
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.com_object.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.com_object.HideSelection = value

    @property
    def hideselection(self):
        """Lower case alias for HideSelection"""
        return self.HideSelection

    @hideselection.setter
    def hideselection(self, value):
        """Lower case alias for HideSelection.setter"""
        self.HideSelection = value

    @property
    def ListCount(self):
        return self.com_object.ListCount

    @property
    def listcount(self):
        """Lower case alias for ListCount"""
        return self.ListCount

    @property
    def ListIndex(self):
        return self.com_object.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.com_object.ListIndex = value

    @property
    def listindex(self):
        """Lower case alias for ListIndex"""
        return self.ListIndex

    @listindex.setter
    def listindex(self, value):
        """Lower case alias for ListIndex.setter"""
        self.ListIndex = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    @property
    def locked(self):
        """Lower case alias for Locked"""
        return self.Locked

    @locked.setter
    def locked(self, value):
        """Lower case alias for Locked.setter"""
        self.Locked = value

    @property
    def MaxLength(self):
        return self.com_object.MaxLength

    @MaxLength.setter
    def MaxLength(self, value):
        self.com_object.MaxLength = value

    @property
    def maxlength(self):
        """Lower case alias for MaxLength"""
        return self.MaxLength

    @maxlength.setter
    def maxlength(self, value):
        """Lower case alias for MaxLength.setter"""
        self.MaxLength = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def SelectionMargin(self):
        return self.com_object.SelectionMargin

    @SelectionMargin.setter
    def SelectionMargin(self, value):
        self.com_object.SelectionMargin = value

    @property
    def selectionmargin(self):
        """Lower case alias for SelectionMargin"""
        return self.SelectionMargin

    @selectionmargin.setter
    def selectionmargin(self, value):
        """Lower case alias for SelectionMargin.setter"""
        self.SelectionMargin = value

    @property
    def SelLength(self):
        return self.com_object.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.com_object.SelLength = value

    @property
    def sellength(self):
        """Lower case alias for SelLength"""
        return self.SelLength

    @sellength.setter
    def sellength(self, value):
        """Lower case alias for SelLength.setter"""
        self.SelLength = value

    @property
    def SelStart(self):
        return self.com_object.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.com_object.SelStart = value

    @property
    def selstart(self):
        """Lower case alias for SelStart"""
        return self.SelStart

    @selstart.setter
    def selstart(self, value):
        """Lower case alias for SelStart.setter"""
        self.SelStart = value

    @property
    def SelText(self):
        return self.com_object.SelText

    @property
    def seltext(self):
        """Lower case alias for SelText"""
        return self.SelText

    @property
    def Style(self):
        return OlComboBoxStyle(self.com_object.Style)

    @Style.setter
    def Style(self, value):
        self.com_object.Style = value

    @property
    def style(self):
        """Lower case alias for Style"""
        return self.Style

    @style.setter
    def style(self, value):
        """Lower case alias for Style.setter"""
        self.Style = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    @property
    def text(self):
        """Lower case alias for Text"""
        return self.Text

    @text.setter
    def text(self, value):
        """Lower case alias for Text.setter"""
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    @property
    def textalign(self):
        """Lower case alias for TextAlign"""
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        """Lower case alias for TextAlign.setter"""
        self.TextAlign = value

    @property
    def TopIndex(self):
        return self.com_object.TopIndex

    @TopIndex.setter
    def TopIndex(self, value):
        self.com_object.TopIndex = value

    @property
    def topindex(self):
        """Lower case alias for TopIndex"""
        return self.TopIndex

    @topindex.setter
    def topindex(self, value):
        """Lower case alias for TopIndex.setter"""
        self.TopIndex = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
        self.Value = value

    def AddItem(self, ItemText=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [ItemText, Index]])
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
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.GetItem(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.RemoveItem(*arguments)

    # Lower case alias for RemoveItem
    def removeitem(self, Index=None):
        arguments = [Index]
        return self.RemoveItem(*arguments)

    def SetItem(self, Index=None, Item=None):
        arguments = com_arguments([unwrap(a) for a in [Index, Item]])
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

    @property
    def accelerator(self):
        """Lower case alias for Accelerator"""
        return self.Accelerator

    @accelerator.setter
    def accelerator(self, value):
        """Lower case alias for Accelerator.setter"""
        self.Accelerator = value

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    @property
    def autosize(self):
        """Lower case alias for AutoSize"""
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        """Lower case alias for AutoSize.setter"""
        self.AutoSize = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @caption.setter
    def caption(self, value):
        """Lower case alias for Caption.setter"""
        self.Caption = value

    @property
    def DisplayDropArrow(self):
        return self.com_object.DisplayDropArrow

    @DisplayDropArrow.setter
    def DisplayDropArrow(self, value):
        self.com_object.DisplayDropArrow = value

    @property
    def displaydroparrow(self):
        """Lower case alias for DisplayDropArrow"""
        return self.DisplayDropArrow

    @displaydroparrow.setter
    def displaydroparrow(self, value):
        """Lower case alias for DisplayDropArrow.setter"""
        self.DisplayDropArrow = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def Picture(self):
        return self.com_object.Picture

    @Picture.setter
    def Picture(self, value):
        self.com_object.Picture = value

    @property
    def picture(self):
        """Lower case alias for Picture"""
        return self.Picture

    @picture.setter
    def picture(self, value):
        """Lower case alias for Picture.setter"""
        self.Picture = value

    @property
    def PictureAlignment(self):
        return OlPictureAlignment(self.com_object.PictureAlignment)

    @PictureAlignment.setter
    def PictureAlignment(self, value):
        self.com_object.PictureAlignment = value

    @property
    def picturealignment(self):
        """Lower case alias for PictureAlignment"""
        return self.PictureAlignment

    @picturealignment.setter
    def picturealignment(self, value):
        """Lower case alias for PictureAlignment.setter"""
        self.PictureAlignment = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    @property
    def textalign(self):
        """Lower case alias for TextAlign"""
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        """Lower case alias for TextAlign.setter"""
        self.TextAlign = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    @property
    def wordwrap(self):
        """Lower case alias for WordWrap"""
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        """Lower case alias for WordWrap.setter"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value


class olkcontrol:

    def __init__(self, olkcontrol=None):
        self.com_object= olkcontrol

    @property
    def ControlProperty(self):
        return self.com_object.ControlProperty

    @ControlProperty.setter
    def ControlProperty(self, value):
        self.com_object.ControlProperty = value

    @property
    def controlproperty(self):
        """Lower case alias for ControlProperty"""
        return self.ControlProperty

    @controlproperty.setter
    def controlproperty(self, value):
        """Lower case alias for ControlProperty.setter"""
        self.ControlProperty = value

    @property
    def EnableAutoLayout(self):
        return self.com_object.EnableAutoLayout

    @EnableAutoLayout.setter
    def EnableAutoLayout(self, value):
        self.com_object.EnableAutoLayout = value

    @property
    def enableautolayout(self):
        """Lower case alias for EnableAutoLayout"""
        return self.EnableAutoLayout

    @enableautolayout.setter
    def enableautolayout(self, value):
        """Lower case alias for EnableAutoLayout.setter"""
        self.EnableAutoLayout = value

    @property
    def Format(self):
        return self.com_object.Format

    @Format.setter
    def Format(self, value):
        self.com_object.Format = value

    @property
    def format(self):
        """Lower case alias for Format"""
        return self.Format

    @format.setter
    def format(self, value):
        """Lower case alias for Format.setter"""
        self.Format = value

    @property
    def HorizontalLayout(self):
        return OlHorizontalLayout(self.com_object.HorizontalLayout)

    @HorizontalLayout.setter
    def HorizontalLayout(self, value):
        self.com_object.HorizontalLayout = value

    @property
    def horizontallayout(self):
        """Lower case alias for HorizontalLayout"""
        return self.HorizontalLayout

    @horizontallayout.setter
    def horizontallayout(self, value):
        """Lower case alias for HorizontalLayout.setter"""
        self.HorizontalLayout = value

    @property
    def ItemProperty(self):
        return self.com_object.ItemProperty

    @ItemProperty.setter
    def ItemProperty(self, value):
        self.com_object.ItemProperty = value

    @property
    def itemproperty(self):
        """Lower case alias for ItemProperty"""
        return self.ItemProperty

    @itemproperty.setter
    def itemproperty(self, value):
        """Lower case alias for ItemProperty.setter"""
        self.ItemProperty = value

    @property
    def MinimumHeight(self):
        return self.com_object.MinimumHeight

    @MinimumHeight.setter
    def MinimumHeight(self, value):
        self.com_object.MinimumHeight = value

    @property
    def minimumheight(self):
        """Lower case alias for MinimumHeight"""
        return self.MinimumHeight

    @minimumheight.setter
    def minimumheight(self, value):
        """Lower case alias for MinimumHeight.setter"""
        self.MinimumHeight = value

    @property
    def MinimumWidth(self):
        return self.com_object.MinimumWidth

    @MinimumWidth.setter
    def MinimumWidth(self, value):
        self.com_object.MinimumWidth = value

    @property
    def minimumwidth(self):
        """Lower case alias for MinimumWidth"""
        return self.MinimumWidth

    @minimumwidth.setter
    def minimumwidth(self, value):
        """Lower case alias for MinimumWidth.setter"""
        self.MinimumWidth = value

    @property
    def PossibleValues(self):
        return self.com_object.PossibleValues

    @PossibleValues.setter
    def PossibleValues(self, value):
        self.com_object.PossibleValues = value

    @property
    def possiblevalues(self):
        """Lower case alias for PossibleValues"""
        return self.PossibleValues

    @possiblevalues.setter
    def possiblevalues(self, value):
        """Lower case alias for PossibleValues.setter"""
        self.PossibleValues = value

    @property
    def VerticalLayout(self):
        return OlVerticalLayout(self.com_object.VerticalLayout)

    @VerticalLayout.setter
    def VerticalLayout(self, value):
        self.com_object.VerticalLayout = value

    @property
    def verticallayout(self):
        """Lower case alias for VerticalLayout"""
        return self.VerticalLayout

    @verticallayout.setter
    def verticallayout(self, value):
        """Lower case alias for VerticalLayout.setter"""
        self.VerticalLayout = value


class OlkDateControl:

    def __init__(self, olkdatecontrol=None):
        self.com_object= olkdatecontrol

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    @property
    def autosize(self):
        """Lower case alias for AutoSize"""
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        """Lower case alias for AutoSize.setter"""
        self.AutoSize = value

    @property
    def AutoWordSelect(self):
        return self.com_object.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.com_object.AutoWordSelect = value

    @property
    def autowordselect(self):
        """Lower case alias for AutoWordSelect"""
        return self.AutoWordSelect

    @autowordselect.setter
    def autowordselect(self, value):
        """Lower case alias for AutoWordSelect.setter"""
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    @property
    def backstyle(self):
        """Lower case alias for BackStyle"""
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        """Lower case alias for BackStyle.setter"""
        self.BackStyle = value

    @property
    def Date(self):
        return self.com_object.Date

    @Date.setter
    def Date(self, value):
        self.com_object.Date = value

    @property
    def date(self):
        """Lower case alias for Date"""
        return self.Date

    @date.setter
    def date(self, value):
        """Lower case alias for Date.setter"""
        self.Date = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.com_object.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.com_object.EnterFieldBehavior = value

    @property
    def enterfieldbehavior(self):
        """Lower case alias for EnterFieldBehavior"""
        return self.EnterFieldBehavior

    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        """Lower case alias for EnterFieldBehavior.setter"""
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.com_object.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.com_object.HideSelection = value

    @property
    def hideselection(self):
        """Lower case alias for HideSelection"""
        return self.HideSelection

    @hideselection.setter
    def hideselection(self, value):
        """Lower case alias for HideSelection.setter"""
        self.HideSelection = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    @property
    def locked(self):
        """Lower case alias for Locked"""
        return self.Locked

    @locked.setter
    def locked(self, value):
        """Lower case alias for Locked.setter"""
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def ShowNoneButton(self):
        return self.com_object.ShowNoneButton

    @ShowNoneButton.setter
    def ShowNoneButton(self, value):
        self.com_object.ShowNoneButton = value

    @property
    def shownonebutton(self):
        """Lower case alias for ShowNoneButton"""
        return self.ShowNoneButton

    @shownonebutton.setter
    def shownonebutton(self, value):
        """Lower case alias for ShowNoneButton.setter"""
        self.ShowNoneButton = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    @property
    def text(self):
        """Lower case alias for Text"""
        return self.Text

    @text.setter
    def text(self, value):
        """Lower case alias for Text.setter"""
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    @property
    def textalign(self):
        """Lower case alias for TextAlign"""
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        """Lower case alias for TextAlign.setter"""
        self.TextAlign = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
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

    @property
    def alignment(self):
        """Lower case alias for Alignment"""
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        """Lower case alias for Alignment.setter"""
        self.Alignment = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @caption.setter
    def caption(self, value):
        """Lower case alias for Caption.setter"""
        self.Caption = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
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

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
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

    @property
    def accelerator(self):
        """Lower case alias for Accelerator"""
        return self.Accelerator

    @accelerator.setter
    def accelerator(self, value):
        """Lower case alias for Accelerator.setter"""
        self.Accelerator = value

    @property
    def AutoSize(self):
        return self.com_object.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.com_object.AutoSize = value

    @property
    def autosize(self):
        """Lower case alias for AutoSize"""
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        """Lower case alias for AutoSize.setter"""
        self.AutoSize = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    @property
    def backstyle(self):
        """Lower case alias for BackStyle"""
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        """Lower case alias for BackStyle.setter"""
        self.BackStyle = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    @property
    def borderstyle(self):
        """Lower case alias for BorderStyle"""
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        """Lower case alias for BorderStyle.setter"""
        self.BorderStyle = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @caption.setter
    def caption(self, value):
        """Lower case alias for Caption.setter"""
        self.Caption = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    @property
    def textalign(self):
        """Lower case alias for TextAlign"""
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        """Lower case alias for TextAlign.setter"""
        self.TextAlign = value

    @property
    def UseHeaderColor(self):
        return self.com_object.UseHeaderColor

    @UseHeaderColor.setter
    def UseHeaderColor(self, value):
        self.com_object.UseHeaderColor = value

    @property
    def useheadercolor(self):
        """Lower case alias for UseHeaderColor"""
        return self.UseHeaderColor

    @useheadercolor.setter
    def useheadercolor(self, value):
        """Lower case alias for UseHeaderColor.setter"""
        self.UseHeaderColor = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    @property
    def wordwrap(self):
        """Lower case alias for WordWrap"""
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        """Lower case alias for WordWrap.setter"""
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

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    @property
    def borderstyle(self):
        """Lower case alias for BorderStyle"""
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        """Lower case alias for BorderStyle.setter"""
        self.BorderStyle = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def ListCount(self):
        return self.com_object.ListCount

    @property
    def listcount(self):
        """Lower case alias for ListCount"""
        return self.ListCount

    @property
    def ListIndex(self):
        return self.com_object.ListIndex

    @ListIndex.setter
    def ListIndex(self, value):
        self.com_object.ListIndex = value

    @property
    def listindex(self):
        """Lower case alias for ListIndex"""
        return self.ListIndex

    @listindex.setter
    def listindex(self, value):
        """Lower case alias for ListIndex.setter"""
        self.ListIndex = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    @property
    def locked(self):
        """Lower case alias for Locked"""
        return self.Locked

    @locked.setter
    def locked(self, value):
        """Lower case alias for Locked.setter"""
        self.Locked = value

    @property
    def MatchEntry(self):
        return olMatchEntry(self.com_object.MatchEntry)

    @MatchEntry.setter
    def MatchEntry(self, value):
        self.com_object.MatchEntry = value

    @property
    def matchentry(self):
        """Lower case alias for MatchEntry"""
        return self.MatchEntry

    @matchentry.setter
    def matchentry(self, value):
        """Lower case alias for MatchEntry.setter"""
        self.MatchEntry = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def MultiSelect(self):
        return OlMultiSelect(self.com_object.MultiSelect)

    @MultiSelect.setter
    def MultiSelect(self, value):
        self.com_object.MultiSelect = value

    @property
    def multiselect(self):
        """Lower case alias for MultiSelect"""
        return self.MultiSelect

    @multiselect.setter
    def multiselect(self, value):
        """Lower case alias for MultiSelect.setter"""
        self.MultiSelect = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    @property
    def text(self):
        """Lower case alias for Text"""
        return self.Text

    @text.setter
    def text(self, value):
        """Lower case alias for Text.setter"""
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    @property
    def textalign(self):
        """Lower case alias for TextAlign"""
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        """Lower case alias for TextAlign.setter"""
        self.TextAlign = value

    @property
    def TopIndex(self):
        return self.com_object.TopIndex

    @TopIndex.setter
    def TopIndex(self, value):
        self.com_object.TopIndex = value

    @property
    def topindex(self):
        """Lower case alias for TopIndex"""
        return self.TopIndex

    @topindex.setter
    def topindex(self, value):
        """Lower case alias for TopIndex.setter"""
        self.TopIndex = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
        self.Value = value

    def AddItem(self, ItemText=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [ItemText, Index]])
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
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.GetItem(*arguments)

    # Lower case alias for GetItem
    def getitem(self, Index=None):
        arguments = [Index]
        return self.GetItem(*arguments)

    def GetSelected(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.GetSelected(*arguments)

    # Lower case alias for GetSelected
    def getselected(self, Index=None):
        arguments = [Index]
        return self.GetSelected(*arguments)

    def RemoveItem(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.RemoveItem(*arguments)

    # Lower case alias for RemoveItem
    def removeitem(self, Index=None):
        arguments = [Index]
        return self.RemoveItem(*arguments)

    def SetItem(self, Index=None, Item=None):
        arguments = com_arguments([unwrap(a) for a in [Index, Item]])
        self.com_object.SetItem(*arguments)

    # Lower case alias for SetItem
    def setitem(self, Index=None, Item=None):
        arguments = [Index, Item]
        return self.SetItem(*arguments)

    def SetSelected(self, Index=None, Selected=None):
        arguments = com_arguments([unwrap(a) for a in [Index, Selected]])
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

    @property
    def accelerator(self):
        """Lower case alias for Accelerator"""
        return self.Accelerator

    @accelerator.setter
    def accelerator(self, value):
        """Lower case alias for Accelerator.setter"""
        self.Accelerator = value

    @property
    def Alignment(self):
        return olAlignment(self.com_object.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.com_object.Alignment = value

    @property
    def alignment(self):
        """Lower case alias for Alignment"""
        return self.Alignment

    @alignment.setter
    def alignment(self, value):
        """Lower case alias for Alignment.setter"""
        self.Alignment = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    @property
    def backstyle(self):
        """Lower case alias for BackStyle"""
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        """Lower case alias for BackStyle.setter"""
        self.BackStyle = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @caption.setter
    def caption(self, value):
        """Lower case alias for Caption.setter"""
        self.Caption = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def GroupName(self):
        return self.com_object.GroupName

    @GroupName.setter
    def GroupName(self, value):
        self.com_object.GroupName = value

    @property
    def groupname(self):
        """Lower case alias for GroupName"""
        return self.GroupName

    @groupname.setter
    def groupname(self, value):
        """Lower case alias for GroupName.setter"""
        self.GroupName = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
        self.Value = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    @property
    def wordwrap(self):
        """Lower case alias for WordWrap"""
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        """Lower case alias for WordWrap.setter"""
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

    @property
    def page(self):
        """Lower case alias for Page"""
        return self.Page

    @page.setter
    def page(self, value):
        """Lower case alias for Page.setter"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def PreferredHeight(self):
        return self.com_object.PreferredHeight

    @property
    def preferredheight(self):
        """Lower case alias for PreferredHeight"""
        return self.PreferredHeight

    @property
    def PreferredWidth(self):
        return self.com_object.PreferredWidth

    @property
    def preferredwidth(self):
        """Lower case alias for PreferredWidth"""
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

    @property
    def autosize(self):
        """Lower case alias for AutoSize"""
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        """Lower case alias for AutoSize.setter"""
        self.AutoSize = value

    @property
    def AutoTab(self):
        return self.com_object.AutoTab

    @AutoTab.setter
    def AutoTab(self, value):
        self.com_object.AutoTab = value

    @property
    def autotab(self):
        """Lower case alias for AutoTab"""
        return self.AutoTab

    @autotab.setter
    def autotab(self, value):
        """Lower case alias for AutoTab.setter"""
        self.AutoTab = value

    @property
    def AutoWordSelect(self):
        return self.com_object.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.com_object.AutoWordSelect = value

    @property
    def autowordselect(self):
        """Lower case alias for AutoWordSelect"""
        return self.AutoWordSelect

    @autowordselect.setter
    def autowordselect(self, value):
        """Lower case alias for AutoWordSelect.setter"""
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    @property
    def borderstyle(self):
        """Lower case alias for BorderStyle"""
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        """Lower case alias for BorderStyle.setter"""
        self.BorderStyle = value

    @property
    def DragBehavior(self):
        return self.com_object.DragBehavior

    @DragBehavior.setter
    def DragBehavior(self, value):
        self.com_object.DragBehavior = value

    @property
    def dragbehavior(self):
        """Lower case alias for DragBehavior"""
        return self.DragBehavior

    @dragbehavior.setter
    def dragbehavior(self, value):
        """Lower case alias for DragBehavior.setter"""
        self.DragBehavior = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.com_object.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.com_object.EnterFieldBehavior = value

    @property
    def enterfieldbehavior(self):
        """Lower case alias for EnterFieldBehavior"""
        return self.EnterFieldBehavior

    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        """Lower case alias for EnterFieldBehavior.setter"""
        self.EnterFieldBehavior = value

    @property
    def EnterKeyBehavior(self):
        return self.com_object.EnterKeyBehavior

    @EnterKeyBehavior.setter
    def EnterKeyBehavior(self, value):
        self.com_object.EnterKeyBehavior = value

    @property
    def enterkeybehavior(self):
        """Lower case alias for EnterKeyBehavior"""
        return self.EnterKeyBehavior

    @enterkeybehavior.setter
    def enterkeybehavior(self, value):
        """Lower case alias for EnterKeyBehavior.setter"""
        self.EnterKeyBehavior = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.com_object.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.com_object.HideSelection = value

    @property
    def hideselection(self):
        """Lower case alias for HideSelection"""
        return self.HideSelection

    @hideselection.setter
    def hideselection(self, value):
        """Lower case alias for HideSelection.setter"""
        self.HideSelection = value

    @property
    def IntegralHeight(self):
        return self.com_object.IntegralHeight

    @IntegralHeight.setter
    def IntegralHeight(self, value):
        self.com_object.IntegralHeight = value

    @property
    def integralheight(self):
        """Lower case alias for IntegralHeight"""
        return self.IntegralHeight

    @integralheight.setter
    def integralheight(self, value):
        """Lower case alias for IntegralHeight.setter"""
        self.IntegralHeight = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    @property
    def locked(self):
        """Lower case alias for Locked"""
        return self.Locked

    @locked.setter
    def locked(self, value):
        """Lower case alias for Locked.setter"""
        self.Locked = value

    @property
    def MaxLength(self):
        return self.com_object.MaxLength

    @MaxLength.setter
    def MaxLength(self, value):
        self.com_object.MaxLength = value

    @property
    def maxlength(self):
        """Lower case alias for MaxLength"""
        return self.MaxLength

    @maxlength.setter
    def maxlength(self, value):
        """Lower case alias for MaxLength.setter"""
        self.MaxLength = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def Multiline(self):
        return self.com_object.Multiline

    @Multiline.setter
    def Multiline(self, value):
        self.com_object.Multiline = value

    @property
    def multiline(self):
        """Lower case alias for Multiline"""
        return self.Multiline

    @multiline.setter
    def multiline(self, value):
        """Lower case alias for Multiline.setter"""
        self.Multiline = value

    @property
    def PasswordChar(self):
        return self.com_object.PasswordChar

    @PasswordChar.setter
    def PasswordChar(self, value):
        self.com_object.PasswordChar = value

    @property
    def passwordchar(self):
        """Lower case alias for PasswordChar"""
        return self.PasswordChar

    @passwordchar.setter
    def passwordchar(self, value):
        """Lower case alias for PasswordChar.setter"""
        self.PasswordChar = value

    @property
    def Scrollbars(self):
        return olScrollBars(self.com_object.Scrollbars)

    @Scrollbars.setter
    def Scrollbars(self, value):
        self.com_object.Scrollbars = value

    @property
    def scrollbars(self):
        """Lower case alias for Scrollbars"""
        return self.Scrollbars

    @scrollbars.setter
    def scrollbars(self, value):
        """Lower case alias for Scrollbars.setter"""
        self.Scrollbars = value

    @property
    def SelectionMargin(self):
        return self.com_object.SelectionMargin

    @SelectionMargin.setter
    def SelectionMargin(self, value):
        self.com_object.SelectionMargin = value

    @property
    def selectionmargin(self):
        """Lower case alias for SelectionMargin"""
        return self.SelectionMargin

    @selectionmargin.setter
    def selectionmargin(self, value):
        """Lower case alias for SelectionMargin.setter"""
        self.SelectionMargin = value

    @property
    def SelLength(self):
        return self.com_object.SelLength

    @SelLength.setter
    def SelLength(self, value):
        self.com_object.SelLength = value

    @property
    def sellength(self):
        """Lower case alias for SelLength"""
        return self.SelLength

    @sellength.setter
    def sellength(self, value):
        """Lower case alias for SelLength.setter"""
        self.SelLength = value

    @property
    def SelStart(self):
        return self.com_object.SelStart

    @SelStart.setter
    def SelStart(self, value):
        self.com_object.SelStart = value

    @property
    def selstart(self):
        """Lower case alias for SelStart"""
        return self.SelStart

    @selstart.setter
    def selstart(self, value):
        """Lower case alias for SelStart.setter"""
        self.SelStart = value

    @property
    def SelText(self):
        return self.com_object.SelText

    @property
    def seltext(self):
        """Lower case alias for SelText"""
        return self.SelText

    @property
    def TabKeyBehavior(self):
        return self.com_object.TabKeyBehavior

    @TabKeyBehavior.setter
    def TabKeyBehavior(self, value):
        self.com_object.TabKeyBehavior = value

    @property
    def tabkeybehavior(self):
        """Lower case alias for TabKeyBehavior"""
        return self.TabKeyBehavior

    @tabkeybehavior.setter
    def tabkeybehavior(self, value):
        """Lower case alias for TabKeyBehavior.setter"""
        self.TabKeyBehavior = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    @property
    def text(self):
        """Lower case alias for Text"""
        return self.Text

    @text.setter
    def text(self, value):
        """Lower case alias for Text.setter"""
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    @property
    def textalign(self):
        """Lower case alias for TextAlign"""
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        """Lower case alias for TextAlign.setter"""
        self.TextAlign = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
        self.Value = value

    @property
    def WordWrap(self):
        return self.com_object.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.com_object.WordWrap = value

    @property
    def wordwrap(self):
        """Lower case alias for WordWrap"""
        return self.WordWrap

    @wordwrap.setter
    def wordwrap(self, value):
        """Lower case alias for WordWrap.setter"""
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

    @property
    def autosize(self):
        """Lower case alias for AutoSize"""
        return self.AutoSize

    @autosize.setter
    def autosize(self, value):
        """Lower case alias for AutoSize.setter"""
        self.AutoSize = value

    @property
    def AutoWordSelect(self):
        return self.com_object.AutoWordSelect

    @AutoWordSelect.setter
    def AutoWordSelect(self, value):
        self.com_object.AutoWordSelect = value

    @property
    def autowordselect(self):
        """Lower case alias for AutoWordSelect"""
        return self.AutoWordSelect

    @autowordselect.setter
    def autowordselect(self, value):
        """Lower case alias for AutoWordSelect.setter"""
        self.AutoWordSelect = value

    @property
    def BackColor(self):
        return self.com_object.BackColor

    @BackColor.setter
    def BackColor(self, value):
        self.com_object.BackColor = value

    @property
    def backcolor(self):
        """Lower case alias for BackColor"""
        return self.BackColor

    @backcolor.setter
    def backcolor(self, value):
        """Lower case alias for BackColor.setter"""
        self.BackColor = value

    @property
    def BackStyle(self):
        return olBackStyle(self.com_object.BackStyle)

    @BackStyle.setter
    def BackStyle(self, value):
        self.com_object.BackStyle = value

    @property
    def backstyle(self):
        """Lower case alias for BackStyle"""
        return self.BackStyle

    @backstyle.setter
    def backstyle(self, value):
        """Lower case alias for BackStyle.setter"""
        self.BackStyle = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def EnterFieldBehavior(self):
        return olEnterFieldBehavior(self.com_object.EnterFieldBehavior)

    @EnterFieldBehavior.setter
    def EnterFieldBehavior(self, value):
        self.com_object.EnterFieldBehavior = value

    @property
    def enterfieldbehavior(self):
        """Lower case alias for EnterFieldBehavior"""
        return self.EnterFieldBehavior

    @enterfieldbehavior.setter
    def enterfieldbehavior(self, value):
        """Lower case alias for EnterFieldBehavior.setter"""
        self.EnterFieldBehavior = value

    @property
    def Font(self):
        return self.com_object.Font

    @property
    def font(self):
        """Lower case alias for Font"""
        return self.Font

    @property
    def ForeColor(self):
        return self.com_object.ForeColor

    @ForeColor.setter
    def ForeColor(self, value):
        self.com_object.ForeColor = value

    @property
    def forecolor(self):
        """Lower case alias for ForeColor"""
        return self.ForeColor

    @forecolor.setter
    def forecolor(self, value):
        """Lower case alias for ForeColor.setter"""
        self.ForeColor = value

    @property
    def HideSelection(self):
        return self.com_object.HideSelection

    @HideSelection.setter
    def HideSelection(self, value):
        self.com_object.HideSelection = value

    @property
    def hideselection(self):
        """Lower case alias for HideSelection"""
        return self.HideSelection

    @hideselection.setter
    def hideselection(self, value):
        """Lower case alias for HideSelection.setter"""
        self.HideSelection = value

    @property
    def IntervalTime(self):
        return self.com_object.IntervalTime

    @IntervalTime.setter
    def IntervalTime(self, value):
        self.com_object.IntervalTime = value

    @property
    def intervaltime(self):
        """Lower case alias for IntervalTime"""
        return self.IntervalTime

    @intervaltime.setter
    def intervaltime(self, value):
        """Lower case alias for IntervalTime.setter"""
        self.IntervalTime = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    @property
    def locked(self):
        """Lower case alias for Locked"""
        return self.Locked

    @locked.setter
    def locked(self, value):
        """Lower case alias for Locked.setter"""
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def ReferenceTime(self):
        return self.com_object.ReferenceTime

    @ReferenceTime.setter
    def ReferenceTime(self, value):
        self.com_object.ReferenceTime = value

    @property
    def referencetime(self):
        """Lower case alias for ReferenceTime"""
        return self.ReferenceTime

    @referencetime.setter
    def referencetime(self, value):
        """Lower case alias for ReferenceTime.setter"""
        self.ReferenceTime = value

    @property
    def Style(self):
        return OlTimeStyle(self.com_object.Style)

    @Style.setter
    def Style(self, value):
        self.com_object.Style = value

    @property
    def style(self):
        """Lower case alias for Style"""
        return self.Style

    @style.setter
    def style(self, value):
        """Lower case alias for Style.setter"""
        self.Style = value

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    @property
    def text(self):
        """Lower case alias for Text"""
        return self.Text

    @text.setter
    def text(self, value):
        """Lower case alias for Text.setter"""
        self.Text = value

    @property
    def TextAlign(self):
        return OlTextAlign(self.com_object.TextAlign)

    @TextAlign.setter
    def TextAlign(self, value):
        self.com_object.TextAlign = value

    @property
    def textalign(self):
        """Lower case alias for TextAlign"""
        return self.TextAlign

    @textalign.setter
    def textalign(self, value):
        """Lower case alias for TextAlign.setter"""
        self.TextAlign = value

    @property
    def Time(self):
        return self.com_object.Time

    @Time.setter
    def Time(self, value):
        self.com_object.Time = value

    @property
    def time(self):
        """Lower case alias for Time"""
        return self.Time

    @time.setter
    def time(self, value):
        """Lower case alias for Time.setter"""
        self.Time = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
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

    @property
    def appointmenttimefield(self):
        """Lower case alias for AppointmentTimeField"""
        return self.AppointmentTimeField

    @appointmenttimefield.setter
    def appointmenttimefield(self, value):
        """Lower case alias for AppointmentTimeField.setter"""
        self.AppointmentTimeField = value

    @property
    def BorderStyle(self):
        return OlBorderStyle(self.com_object.BorderStyle)

    @BorderStyle.setter
    def BorderStyle(self, value):
        self.com_object.BorderStyle = value

    @property
    def borderstyle(self):
        """Lower case alias for BorderStyle"""
        return self.BorderStyle

    @borderstyle.setter
    def borderstyle(self, value):
        """Lower case alias for BorderStyle.setter"""
        self.BorderStyle = value

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Locked(self):
        return self.com_object.Locked

    @Locked.setter
    def Locked(self, value):
        self.com_object.Locked = value

    @property
    def locked(self):
        """Lower case alias for Locked"""
        return self.Locked

    @locked.setter
    def locked(self, value):
        """Lower case alias for Locked.setter"""
        self.Locked = value

    @property
    def MouseIcon(self):
        return self.com_object.MouseIcon

    @MouseIcon.setter
    def MouseIcon(self, value):
        self.com_object.MouseIcon = value

    @property
    def mouseicon(self):
        """Lower case alias for MouseIcon"""
        return self.MouseIcon

    @mouseicon.setter
    def mouseicon(self, value):
        """Lower case alias for MouseIcon.setter"""
        self.MouseIcon = value

    @property
    def MousePointer(self):
        return OlMousePointer(self.com_object.MousePointer)

    @MousePointer.setter
    def MousePointer(self, value):
        self.com_object.MousePointer = value

    @property
    def mousepointer(self):
        """Lower case alias for MousePointer"""
        return self.MousePointer

    @mousepointer.setter
    def mousepointer(self, value):
        """Lower case alias for MousePointer.setter"""
        self.MousePointer = value

    @property
    def SelectedTimeZoneIndex(self):
        return Application.TimeZones(self.com_object.SelectedTimeZoneIndex)

    @SelectedTimeZoneIndex.setter
    def SelectedTimeZoneIndex(self, value):
        self.com_object.SelectedTimeZoneIndex = value

    @property
    def selectedtimezoneindex(self):
        """Lower case alias for SelectedTimeZoneIndex"""
        return self.SelectedTimeZoneIndex

    @selectedtimezoneindex.setter
    def selectedtimezoneindex(self, value):
        """Lower case alias for SelectedTimeZoneIndex.setter"""
        self.SelectedTimeZoneIndex = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
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

# olunifiedgroupfoldertype enumeration
olGroupCalendarFolder = 1
olGroupMailFolder = 0

# olunifiedgrouptype enumeration
PrivateGroup = 1
PublicGroup = 2

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

class optionbutton:

    def __init__(self, optionbutton=None):
        self.com_object= optionbutton


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

    @property
    def isdescending(self):
        """Lower case alias for IsDescending"""
        return self.IsDescending

    @isdescending.setter
    def isdescending(self, value):
        """Lower case alias for IsDescending.setter"""
        self.IsDescending = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ViewXMLSchemaName(self):
        return OrderField(self.com_object.ViewXMLSchemaName)

    @property
    def viewxmlschemaname(self):
        """Lower case alias for ViewXMLSchemaName"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, PropertyName=None, IsDescending=None):
        arguments = com_arguments([unwrap(a) for a in [PropertyName, IsDescending]])
        return OrderField(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, PropertyName=None, IsDescending=None):
        arguments = [PropertyName, IsDescending]
        return self.Add(*arguments)

    def Insert(self, PropertyName=None, Index=None, IsDescending=None):
        arguments = com_arguments([unwrap(a) for a in [PropertyName, Index, IsDescending]])
        return OrderField(self.com_object.Insert(*arguments))

    # Lower case alias for Insert
    def insert(self, PropertyName=None, Index=None, IsDescending=None):
        arguments = [PropertyName, Index, IsDescending]
        return self.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return OrderField(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Shortcuts(self):
        return OutlookBarShortcuts(self.com_object.Shortcuts)

    @property
    def shortcuts(self):
        """Lower case alias for Shortcuts"""
        return self.Shortcuts

    @property
    def ViewType(self):
        return OlOutlookBarViewType(self.com_object.ViewType)

    @ViewType.setter
    def ViewType(self, value):
        self.com_object.ViewType = value

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
        return self.ViewType

    @viewtype.setter
    def viewtype(self, value):
        """Lower case alias for ViewType.setter"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Index]])
        return OutlookBarGroup(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Index=None):
        arguments = [Name, Index]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return OutlookBarGroup(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def contents(self):
        """Lower case alias for Contents"""
        return self.Contents

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return self.com_object.Visible

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Target(self):
        return self.com_object.Target

    @property
    def target(self):
        """Lower case alias for Target"""
        return self.Target

    def SetIcon(self, Icon=None):
        arguments = com_arguments([unwrap(a) for a in [Icon]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Target=None, Name=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Target, Name, Index]])
        return OutlookBarShortcut(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Target=None, Name=None, Index=None):
        arguments = [Target, Name, Index]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return OutlookBarShortcut(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def groups(self):
        """Lower case alias for Groups"""
        return self.Groups

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class page:

    def __init__(self, page=None):
        self.com_object= page


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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self):
        return Page(self.com_object.Add())

    # Lower case alias for Add
    def add(self):
        return self.Add()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class peopleview:

    def __init__(self, peopleview=None):
        self.com_object= peopleview

    @property
    def application(self):
        return self.com_object.application

    @property
    def class(self):
        return self.com_object.class

    @property
    def filter(self):
        return self.com_object.filter

    @filter.setter
    def filter(self, value):
        self.com_object.filter = value

    @property
    def filter(self):
        """Lower case alias for filter"""
        return self.filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for filter.setter"""
        self.filter = value

    @property
    def language(self):
        return self.com_object.language

    @language.setter
    def language(self, value):
        self.com_object.language = value

    @property
    def language(self):
        """Lower case alias for language"""
        return self.language

    @language.setter
    def language(self, value):
        """Lower case alias for language.setter"""
        self.language = value

    @property
    def lockuserchanges(self):
        return self.com_object.lockuserchanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        self.com_object.lockuserchanges = value

    @property
    def lockuserchanges(self):
        """Lower case alias for lockuserchanges"""
        return self.lockuserchanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        """Lower case alias for lockuserchanges.setter"""
        self.lockuserchanges = value

    @property
    def name(self):
        return self.com_object.name

    @name.setter
    def name(self, value):
        self.com_object.name = value

    @property
    def name(self):
        """Lower case alias for name"""
        return self.name

    @name.setter
    def name(self, value):
        """Lower case alias for name.setter"""
        self.name = value

    @property
    def parent(self):
        return self.com_object.parent

    @property
    def parent(self):
        """Lower case alias for parent"""
        return self.parent

    @property
    def saveoption(self):
        return self.com_object.saveoption

    @property
    def saveoption(self):
        """Lower case alias for saveoption"""
        return self.saveoption

    @property
    def session(self):
        return self.com_object.session

    @property
    def session(self):
        """Lower case alias for session"""
        return self.session

    @property
    def sortfields(self):
        return self.com_object.sortfields

    @property
    def sortfields(self):
        """Lower case alias for sortfields"""
        return self.sortfields

    @property
    def standard(self):
        return self.com_object.standard

    @property
    def standard(self):
        """Lower case alias for standard"""
        return self.standard

    @property
    def viewtype(self):
        return self.com_object.viewtype

    @property
    def viewtype(self):
        """Lower case alias for viewtype"""
        return self.viewtype

    @property
    def xml(self):
        return self.com_object.xml

    @xml.setter
    def xml(self, value):
        self.com_object.xml = value

    @property
    def xml(self):
        """Lower case alias for xml"""
        return self.xml

    @xml.setter
    def xml(self, value):
        """Lower case alias for xml.setter"""
        self.xml = value

    def apply(self):
        self.com_object.apply()

    # Lower case alias for apply
    def apply(self):
        return self.apply()

    def copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SaveOption]])
        self.com_object.copy(*arguments)

    # Lower case alias for copy
    def copy(self, Name=None, SaveOption=None):
        arguments = [Name, SaveOption]
        return self.copy(*arguments)

    def delete(self):
        self.com_object.delete()

    # Lower case alias for delete
    def delete(self):
        return self.delete()

    def gotodate(self, Date=None):
        arguments = com_arguments([unwrap(a) for a in [Date]])
        self.com_object.gotodate(*arguments)

    # Lower case alias for gotodate
    def gotodate(self, Date=None):
        arguments = [Date]
        return self.gotodate(*arguments)

    def reset(self):
        self.com_object.reset()

    # Lower case alias for reset
    def reset(self):
        return self.reset()

    def save(self):
        self.com_object.save()

    # Lower case alias for save
    def save(self):
        return self.save()


class PlaySoundRuleAction:

    def __init__(self, playsoundruleaction=None):
        self.com_object= playsoundruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    @property
    def actiontype(self):
        """Lower case alias for ActionType"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def FilePath(self):
        return self.com_object.FilePath

    @FilePath.setter
    def FilePath(self, value):
        self.com_object.FilePath = value

    @property
    def filepath(self):
        """Lower case alias for FilePath"""
        return self.FilePath

    @filepath.setter
    def filepath(self, value):
        """Lower case alias for FilePath.setter"""
        self.FilePath = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class PostItem:

    def __init__(self, postitem=None):
        self.com_object= postitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.com_object.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.com_object.BodyFormat = value

    @property
    def bodyformat(self):
        """Lower case alias for BodyFormat"""
        return self.BodyFormat

    @bodyformat.setter
    def bodyformat(self, value):
        """Lower case alias for BodyFormat.setter"""
        self.BodyFormat = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def ExpiryTime(self):
        return self.com_object.ExpiryTime

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.com_object.ExpiryTime = value

    @property
    def expirytime(self):
        """Lower case alias for ExpiryTime"""
        return self.ExpiryTime

    @expirytime.setter
    def expirytime(self, value):
        """Lower case alias for ExpiryTime.setter"""
        self.ExpiryTime = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def HTMLBody(self):
        return self.com_object.HTMLBody

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.com_object.HTMLBody = value

    @property
    def htmlbody(self):
        """Lower case alias for HTMLBody"""
        return self.HTMLBody

    @htmlbody.setter
    def htmlbody(self, value):
        """Lower case alias for HTMLBody.setter"""
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    @property
    def internetcodepage(self):
        """Lower case alias for InternetCodepage"""
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        """Lower case alias for InternetCodepage.setter"""
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return PostItem(self.com_object.IsMarkedAsTask)

    @property
    def ismarkedastask(self):
        """Lower case alias for IsMarkedAsTask"""
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def ReceivedTime(self):
        return self.com_object.ReceivedTime

    @property
    def receivedtime(self):
        """Lower case alias for ReceivedTime"""
        return self.ReceivedTime

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    @property
    def reminderoverridedefault(self):
        """Lower case alias for ReminderOverrideDefault"""
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        """Lower case alias for ReminderOverrideDefault.setter"""
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    @property
    def reminderplaysound(self):
        """Lower case alias for ReminderPlaySound"""
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        """Lower case alias for ReminderPlaySound.setter"""
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    @property
    def reminderset(self):
        """Lower case alias for ReminderSet"""
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        """Lower case alias for ReminderSet.setter"""
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    @property
    def remindersoundfile(self):
        """Lower case alias for ReminderSoundFile"""
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        """Lower case alias for ReminderSoundFile.setter"""
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    @property
    def remindertime(self):
        """Lower case alias for ReminderTime"""
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        """Lower case alias for ReminderTime.setter"""
        self.ReminderTime = value

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def SenderEmailAddress(self):
        return self.com_object.SenderEmailAddress

    @property
    def senderemailaddress(self):
        """Lower case alias for SenderEmailAddress"""
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return self.com_object.SenderEmailType

    @property
    def senderemailtype(self):
        """Lower case alias for SenderEmailType"""
        return self.SenderEmailType

    @property
    def SenderName(self):
        return self.com_object.SenderName

    @property
    def sendername(self):
        """Lower case alias for SenderName"""
        return self.SenderName

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def SentOn(self):
        return self.com_object.SentOn

    @property
    def senton(self):
        """Lower case alias for SentOn"""
        return self.SentOn

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def TaskCompletedDate(self):
        return PostItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    @property
    def taskcompleteddate(self):
        """Lower case alias for TaskCompletedDate"""
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        """Lower case alias for TaskCompletedDate.setter"""
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return PostItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    @property
    def taskduedate(self):
        """Lower case alias for TaskDueDate"""
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        """Lower case alias for TaskDueDate.setter"""
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return PostItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    @property
    def taskstartdate(self):
        """Lower case alias for TaskStartDate"""
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        """Lower case alias for TaskStartDate.setter"""
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return PostItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    @property
    def tasksubject(self):
        """Lower case alias for TaskSubject"""
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        """Lower case alias for TaskSubject.setter"""
        self.TaskSubject = value

    @property
    def ToDoTaskOrdinal(self):
        return PostItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    @property
    def todotaskordinal(self):
        """Lower case alias for ToDoTaskOrdinal"""
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        """Lower case alias for ToDoTaskOrdinal.setter"""
        self.ToDoTaskOrdinal = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
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
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [MarkInterval]])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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


class previewpane:

    def __init__(self, previewpane=None):
        self.com_object= previewpane

    @property
    def application(self):
        return self.com_object.application

    @property
    def class(self):
        return self.com_object.class

    @property
    def parent(self):
        return PreviewPane(self.com_object.parent)

    @property
    def parent(self):
        """Lower case alias for parent"""
        return self.parent

    @property
    def session(self):
        return self.com_object.session

    @property
    def session(self):
        """Lower case alias for session"""
        return self.session

    @property
    def wordeditor(self):
        return self.com_object.wordeditor

    @property
    def wordeditor(self):
        """Lower case alias for wordeditor"""
        return self.wordeditor


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

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def BinaryToString(self, Value=None):
        arguments = com_arguments([unwrap(a) for a in [Value]])
        return self.com_object.BinaryToString(*arguments)

    # Lower case alias for BinaryToString
    def binarytostring(self, Value=None):
        arguments = [Value]
        return self.BinaryToString(*arguments)

    def DeleteProperties(self, SchemaNames=None):
        arguments = com_arguments([unwrap(a) for a in [SchemaNames]])
        return self.com_object.DeleteProperties(*arguments)

    # Lower case alias for DeleteProperties
    def deleteproperties(self, SchemaNames=None):
        arguments = [SchemaNames]
        return self.DeleteProperties(*arguments)

    def DeleteProperty(self, SchemaName=None):
        arguments = com_arguments([unwrap(a) for a in [SchemaName]])
        self.com_object.DeleteProperty(*arguments)

    # Lower case alias for DeleteProperty
    def deleteproperty(self, SchemaName=None):
        arguments = [SchemaName]
        return self.DeleteProperty(*arguments)

    def GetProperties(self, SchemaNames=None):
        arguments = com_arguments([unwrap(a) for a in [SchemaNames]])
        return self.com_object.GetProperties(*arguments)

    # Lower case alias for GetProperties
    def getproperties(self, SchemaNames=None):
        arguments = [SchemaNames]
        return self.GetProperties(*arguments)

    def GetProperty(self, SchemaName=None):
        arguments = com_arguments([unwrap(a) for a in [SchemaName]])
        return self.com_object.GetProperty(*arguments)

    # Lower case alias for GetProperty
    def getproperty(self, SchemaName=None):
        arguments = [SchemaName]
        return self.GetProperty(*arguments)

    def LocalTimeToUTC(self, Value=None):
        arguments = com_arguments([unwrap(a) for a in [Value]])
        return self.com_object.LocalTimeToUTC(*arguments)

    # Lower case alias for LocalTimeToUTC
    def localtimetoutc(self, Value=None):
        arguments = [Value]
        return self.LocalTimeToUTC(*arguments)

    def SetProperties(self, SchemaNames=None, Values=None):
        arguments = com_arguments([unwrap(a) for a in [SchemaNames, Values]])
        return self.com_object.SetProperties(*arguments)

    # Lower case alias for SetProperties
    def setproperties(self, SchemaNames=None, Values=None):
        arguments = [SchemaNames, Values]
        return self.SetProperties(*arguments)

    def SetProperty(self, SchemaName=None, Value=None):
        arguments = com_arguments([unwrap(a) for a in [SchemaName, Value]])
        self.com_object.SetProperty(*arguments)

    # Lower case alias for SetProperty
    def setproperty(self, SchemaName=None, Value=None):
        arguments = [SchemaName, Value]
        return self.SetProperty(*arguments)

    def StringToBinary(self, Value=None):
        arguments = com_arguments([unwrap(a) for a in [Value]])
        return self.com_object.StringToBinary(*arguments)

    # Lower case alias for StringToBinary
    def stringtobinary(self, Value=None):
        arguments = [Value]
        return self.StringToBinary(*arguments)

    def UTCToLocalTime(self, Value=None):
        arguments = com_arguments([unwrap(a) for a in [Value]])
        return self.com_object.UTCToLocalTime(*arguments)

    # Lower case alias for UTCToLocalTime
    def utctolocaltime(self, Value=None):
        arguments = [Value]
        return self.UTCToLocalTime(*arguments)


class PropertyPage:

    def __init__(self, propertypage=None):
        self.com_object= propertypage

    def Dirty(self, Dirty=None):
        arguments = com_arguments([unwrap(a) for a in [Dirty]])
        if hasattr(self.com_object, "GetDirty"):
            return self.com_object.GetDirty(*arguments)
        else:
            return self.com_object.Dirty(*arguments)

    def dirty(self, Dirty=None):
        """Lower case alias for Dirty"""
        arguments = [Dirty]
        return self.Dirty(*arguments)

    def Apply(self):
        return self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def GetPageInfo(self, HelpFile=None, HelpContext=None):
        arguments = com_arguments([unwrap(a) for a in [HelpFile, HelpContext]])
        return self.com_object.GetPageInfo(*arguments)

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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Page=None, Title=None):
        arguments = com_arguments([unwrap(a) for a in [Page, Title]])
        self.com_object.Add(*arguments)

    # Lower case alias for Add
    def add(self, Page=None, Title=None):
        arguments = [Page, Title]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def address(self):
        """Lower case alias for Address"""
        return self.Address

    @property
    def AddressEntry(self):
        return AddressEntry(self.com_object.AddressEntry)

    @AddressEntry.setter
    def AddressEntry(self, value):
        self.com_object.AddressEntry = value

    @property
    def addressentry(self):
        """Lower case alias for AddressEntry"""
        return self.AddressEntry

    @addressentry.setter
    def addressentry(self, value):
        """Lower case alias for AddressEntry.setter"""
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

    @property
    def autoresponse(self):
        """Lower case alias for AutoResponse"""
        return self.AutoResponse

    @autoresponse.setter
    def autoresponse(self, value):
        """Lower case alias for AutoResponse.setter"""
        self.AutoResponse = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DisplayType(self):
        return OlDisplayType(self.com_object.DisplayType)

    @property
    def displaytype(self):
        """Lower case alias for DisplayType"""
        return self.DisplayType

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def Index(self):
        return self.com_object.Index

    @property
    def index(self):
        """Lower case alias for Index"""
        return self.Index

    @property
    def MeetingResponseStatus(self):
        return OlResponseStatus(self.com_object.MeetingResponseStatus)

    @property
    def meetingresponsestatus(self):
        """Lower case alias for MeetingResponseStatus"""
        return self.MeetingResponseStatus

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Resolved(self):
        return self.com_object.Resolved

    @property
    def resolved(self):
        """Lower case alias for Resolved"""
        return self.Resolved

    @property
    def Sendable(self):
        return Recipient(self.com_object.Sendable)

    @Sendable.setter
    def Sendable(self, value):
        self.com_object.Sendable = value

    @property
    def sendable(self):
        """Lower case alias for Sendable"""
        return self.Sendable

    @sendable.setter
    def sendable(self, value):
        """Lower case alias for Sendable.setter"""
        self.Sendable = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def TrackingStatus(self):
        return OlTrackingStatus(self.com_object.TrackingStatus)

    @TrackingStatus.setter
    def TrackingStatus(self, value):
        self.com_object.TrackingStatus = value

    @property
    def trackingstatus(self):
        """Lower case alias for TrackingStatus"""
        return self.TrackingStatus

    @trackingstatus.setter
    def trackingstatus(self, value):
        """Lower case alias for TrackingStatus.setter"""
        self.TrackingStatus = value

    @property
    def TrackingStatusTime(self):
        return self.com_object.TrackingStatusTime

    @TrackingStatusTime.setter
    def TrackingStatusTime(self, value):
        self.com_object.TrackingStatusTime = value

    @property
    def trackingstatustime(self):
        """Lower case alias for TrackingStatusTime"""
        return self.TrackingStatusTime

    @trackingstatustime.setter
    def trackingstatustime(self, value):
        """Lower case alias for TrackingStatusTime.setter"""
        self.TrackingStatusTime = value

    @property
    def Type(self):
        return self.com_object.Type

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @type.setter
    def type(self, value):
        """Lower case alias for Type.setter"""
        self.Type = value

    def Delete(self):
        self.com_object.Delete()

    # Lower case alias for Delete
    def delete(self):
        return self.Delete()

    def FreeBusy(self, Start=None, MinPerChar=None, CompleteFormat=None):
        arguments = com_arguments([unwrap(a) for a in [Start, MinPerChar, CompleteFormat]])
        return self.com_object.FreeBusy(*arguments)

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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Name]])
        return Recipient(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None):
        arguments = [Name]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Recipient(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def dayofmonth(self):
        """Lower case alias for DayOfMonth"""
        return self.DayOfMonth

    @dayofmonth.setter
    def dayofmonth(self, value):
        """Lower case alias for DayOfMonth.setter"""
        self.DayOfMonth = value

    @property
    def DayOfWeekMask(self):
        return OlDaysOfWeek(self.com_object.DayOfWeekMask)

    @DayOfWeekMask.setter
    def DayOfWeekMask(self, value):
        self.com_object.DayOfWeekMask = value

    @property
    def dayofweekmask(self):
        """Lower case alias for DayOfWeekMask"""
        return self.DayOfWeekMask

    @dayofweekmask.setter
    def dayofweekmask(self, value):
        """Lower case alias for DayOfWeekMask.setter"""
        self.DayOfWeekMask = value

    @property
    def Duration(self):
        return RecurrencePattern(self.com_object.Duration)

    @Duration.setter
    def Duration(self, value):
        self.com_object.Duration = value

    @property
    def duration(self):
        """Lower case alias for Duration"""
        return self.Duration

    @duration.setter
    def duration(self, value):
        """Lower case alias for Duration.setter"""
        self.Duration = value

    @property
    def EndTime(self):
        return self.com_object.EndTime

    @EndTime.setter
    def EndTime(self, value):
        self.com_object.EndTime = value

    @property
    def endtime(self):
        """Lower case alias for EndTime"""
        return self.EndTime

    @endtime.setter
    def endtime(self, value):
        """Lower case alias for EndTime.setter"""
        self.EndTime = value

    @property
    def Exceptions(self):
        return Exceptions(self.com_object.Exceptions)

    @property
    def exceptions(self):
        """Lower case alias for Exceptions"""
        return self.Exceptions

    @property
    def Instance(self):
        return self.com_object.Instance

    @Instance.setter
    def Instance(self, value):
        self.com_object.Instance = value

    @property
    def instance(self):
        """Lower case alias for Instance"""
        return self.Instance

    @instance.setter
    def instance(self, value):
        """Lower case alias for Instance.setter"""
        self.Instance = value

    @property
    def Interval(self):
        return self.com_object.Interval

    @Interval.setter
    def Interval(self, value):
        self.com_object.Interval = value

    @property
    def interval(self):
        """Lower case alias for Interval"""
        return self.Interval

    @interval.setter
    def interval(self, value):
        """Lower case alias for Interval.setter"""
        self.Interval = value

    @property
    def MonthOfYear(self):
        return self.com_object.MonthOfYear

    @MonthOfYear.setter
    def MonthOfYear(self, value):
        self.com_object.MonthOfYear = value

    @property
    def monthofyear(self):
        """Lower case alias for MonthOfYear"""
        return self.MonthOfYear

    @monthofyear.setter
    def monthofyear(self, value):
        """Lower case alias for MonthOfYear.setter"""
        self.MonthOfYear = value

    @property
    def NoEndDate(self):
        return self.com_object.NoEndDate

    @NoEndDate.setter
    def NoEndDate(self, value):
        self.com_object.NoEndDate = value

    @property
    def noenddate(self):
        """Lower case alias for NoEndDate"""
        return self.NoEndDate

    @noenddate.setter
    def noenddate(self, value):
        """Lower case alias for NoEndDate.setter"""
        self.NoEndDate = value

    @property
    def Occurrences(self):
        return self.com_object.Occurrences

    @Occurrences.setter
    def Occurrences(self, value):
        self.com_object.Occurrences = value

    @property
    def occurrences(self):
        """Lower case alias for Occurrences"""
        return self.Occurrences

    @occurrences.setter
    def occurrences(self, value):
        """Lower case alias for Occurrences.setter"""
        self.Occurrences = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PatternEndDate(self):
        return self.com_object.PatternEndDate

    @PatternEndDate.setter
    def PatternEndDate(self, value):
        self.com_object.PatternEndDate = value

    @property
    def patternenddate(self):
        """Lower case alias for PatternEndDate"""
        return self.PatternEndDate

    @patternenddate.setter
    def patternenddate(self, value):
        """Lower case alias for PatternEndDate.setter"""
        self.PatternEndDate = value

    @property
    def PatternStartDate(self):
        return self.com_object.PatternStartDate

    @PatternStartDate.setter
    def PatternStartDate(self, value):
        self.com_object.PatternStartDate = value

    @property
    def patternstartdate(self):
        """Lower case alias for PatternStartDate"""
        return self.PatternStartDate

    @patternstartdate.setter
    def patternstartdate(self, value):
        """Lower case alias for PatternStartDate.setter"""
        self.PatternStartDate = value

    @property
    def RecurrenceType(self):
        return OlRecurrenceType(self.com_object.RecurrenceType)

    @RecurrenceType.setter
    def RecurrenceType(self, value):
        self.com_object.RecurrenceType = value

    @property
    def recurrencetype(self):
        """Lower case alias for RecurrenceType"""
        return self.RecurrenceType

    @recurrencetype.setter
    def recurrencetype(self, value):
        """Lower case alias for RecurrenceType.setter"""
        self.RecurrenceType = value

    @property
    def Regenerate(self):
        return self.com_object.Regenerate

    @Regenerate.setter
    def Regenerate(self, value):
        self.com_object.Regenerate = value

    @property
    def regenerate(self):
        """Lower case alias for Regenerate"""
        return self.Regenerate

    @regenerate.setter
    def regenerate(self, value):
        """Lower case alias for Regenerate.setter"""
        self.Regenerate = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def StartTime(self):
        return self.com_object.StartTime

    @StartTime.setter
    def StartTime(self, value):
        self.com_object.StartTime = value

    @property
    def starttime(self):
        """Lower case alias for StartTime"""
        return self.StartTime

    @starttime.setter
    def starttime(self, value):
        """Lower case alias for StartTime.setter"""
        self.StartTime = value

    def GetOccurrence(self, StartDate=None):
        arguments = com_arguments([unwrap(a) for a in [StartDate]])
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

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def IsVisible(self):
        return self.com_object.IsVisible

    @property
    def isvisible(self):
        """Lower case alias for IsVisible"""
        return self.IsVisible

    @property
    def Item(self):
        return self.com_object.Item

    @property
    def item(self):
        """Lower case alias for Item"""
        return self.Item

    @property
    def NextReminderDate(self):
        return self.com_object.NextReminderDate

    @property
    def nextreminderdate(self):
        """Lower case alias for NextReminderDate"""
        return self.NextReminderDate

    @property
    def OriginalReminderDate(self):
        return self.com_object.OriginalReminderDate

    @property
    def originalreminderdate(self):
        """Lower case alias for OriginalReminderDate"""
        return self.OriginalReminderDate

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Dismiss(self):
        self.com_object.Dismiss()

    # Lower case alias for Dismiss
    def dismiss(self):
        return self.Dismiss()

    def Snooze(self, SnoozeTime=None):
        arguments = com_arguments([unwrap(a) for a in [SnoozeTime]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Reminder(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def HasAttachment(self):
        return self.com_object.HasAttachment

    @property
    def hasattachment(self):
        """Lower case alias for HasAttachment"""
        return self.HasAttachment

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def RemoteMessageClass(self):
        return self.com_object.RemoteMessageClass

    @property
    def remotemessageclass(self):
        """Lower case alias for RemoteMessageClass"""
        return self.RemoteMessageClass

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def TransferSize(self):
        return self.com_object.TransferSize

    @property
    def transfersize(self):
        """Lower case alias for TransferSize"""
        return self.TransferSize

    @property
    def TransferTime(self):
        return self.com_object.TransferTime

    @property
    def transfertime(self):
        """Lower case alias for TransferTime"""
        return self.TransferTime

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def RetentionExpirationDate(self):
        return ReportItem(self.com_object.RetentionExpirationDate)

    @property
    def retentionexpirationdate(self):
        """Lower case alias for RetentionExpirationDate"""
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.com_object.RetentionPolicyName

    @property
    def retentionpolicyname(self):
        """Lower case alias for RetentionPolicyName"""
        return self.RetentionPolicyName

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def DefaultItemType(self):
        return OlItemType(self.com_object.DefaultItemType)

    @DefaultItemType.setter
    def DefaultItemType(self, value):
        self.com_object.DefaultItemType = value

    @property
    def defaultitemtype(self):
        """Lower case alias for DefaultItemType"""
        return self.DefaultItemType

    @defaultitemtype.setter
    def defaultitemtype(self, value):
        """Lower case alias for DefaultItemType.setter"""
        self.DefaultItemType = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def GetFirst(self):
        return self.com_object.GetFirst()

    # Lower case alias for GetFirst
    def getfirst(self):
        return self.GetFirst()

    def GetLast(self):
        return self.com_object.GetLast()

    # Lower case alias for GetLast
    def getlast(self):
        return self.GetLast()

    def GetNext(self):
        return self.com_object.GetNext()

    # Lower case alias for GetNext
    def getnext(self):
        return self.GetNext()

    def GetPrevious(self):
        return self.com_object.GetPrevious()

    # Lower case alias for GetPrevious
    def getprevious(self):
        return self.GetPrevious()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Columns]])
        self.com_object.SetColumns(*arguments)

    # Lower case alias for SetColumns
    def setcolumns(self, Columns=None):
        arguments = [Columns]
        return self.SetColumns(*arguments)

    def Sort(self, Property=None, Descending=None):
        arguments = com_arguments([unwrap(a) for a in [Property, Descending]])
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

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def BinaryToString(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.BinaryToString(*arguments)

    # Lower case alias for BinaryToString
    def binarytostring(self, Index=None):
        arguments = [Index]
        return self.BinaryToString(*arguments)

    def GetValues(self):
        return self.com_object.GetValues()

    # Lower case alias for GetValues
    def getvalues(self):
        return self.GetValues()

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def LocalTimeToUTC(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.LocalTimeToUTC(*arguments)

    # Lower case alias for LocalTimeToUTC
    def localtimetoutc(self, Index=None):
        arguments = [Index]
        return self.LocalTimeToUTC(*arguments)

    def UTCToLocalTime(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.UTCToLocalTime(*arguments)

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

    @property
    def actions(self):
        """Lower case alias for Actions"""
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

    @property
    def conditions(self):
        """Lower case alias for Conditions"""
        return self.Conditions

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Exceptions(self):
        return RuleConditions(self.com_object.Exceptions)

    @property
    def exceptions(self):
        """Lower case alias for Exceptions"""
        return self.Exceptions

    @property
    def ExecutionOrder(self):
        return Rules(self.com_object.ExecutionOrder)

    @ExecutionOrder.setter
    def ExecutionOrder(self, value):
        self.com_object.ExecutionOrder = value

    @property
    def executionorder(self):
        """Lower case alias for ExecutionOrder"""
        return self.ExecutionOrder

    @executionorder.setter
    def executionorder(self, value):
        """Lower case alias for ExecutionOrder.setter"""
        self.ExecutionOrder = value

    @property
    def IsLocalRule(self):
        return self.com_object.IsLocalRule

    @property
    def islocalrule(self):
        """Lower case alias for IsLocalRule"""
        return self.IsLocalRule

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RuleType(self):
        return OlRuleType(self.com_object.RuleType)

    @property
    def ruletype(self):
        """Lower case alias for RuleType"""
        return self.RuleType

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Execute(self, ShowProgress=None, Folder=None, IncludeSubfolders=None, RuleExecuteOption=None):
        arguments = com_arguments([unwrap(a) for a in [ShowProgress, Folder, IncludeSubfolders, RuleExecuteOption]])
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

    @property
    def actiontype(self):
        """Lower case alias for ActionType"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def assigntocategory(self):
        """Lower case alias for AssignToCategory"""
        return self.AssignToCategory

    @property
    def CC(self):
        return SendRuleAction(self.com_object.CC)

    @property
    def cc(self):
        """Lower case alias for CC"""
        return self.CC

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ClearCategories(self):
        return RuleAction(self.com_object.ClearCategories)

    @property
    def clearcategories(self):
        """Lower case alias for ClearCategories"""
        return self.ClearCategories

    @property
    def CopyToFolder(self):
        return MoveOrCopyRuleAction(self.com_object.CopyToFolder)

    @property
    def copytofolder(self):
        """Lower case alias for CopyToFolder"""
        return self.CopyToFolder

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Delete(self):
        return RuleAction(self.com_object.Delete)

    @property
    def delete(self):
        """Lower case alias for Delete"""
        return self.Delete

    @property
    def DeletePermanently(self):
        return RuleAction(self.com_object.DeletePermanently)

    @property
    def deletepermanently(self):
        """Lower case alias for DeletePermanently"""
        return self.DeletePermanently

    @property
    def DesktopAlert(self):
        return RuleAction(self.com_object.DesktopAlert)

    @property
    def desktopalert(self):
        """Lower case alias for DesktopAlert"""
        return self.DesktopAlert

    @property
    def Forward(self):
        return SendRuleAction(self.com_object.Forward)

    @property
    def forward(self):
        """Lower case alias for Forward"""
        return self.Forward

    @property
    def ForwardAsAttachment(self):
        return SendRuleAction(self.com_object.ForwardAsAttachment)

    @property
    def forwardasattachment(self):
        """Lower case alias for ForwardAsAttachment"""
        return self.ForwardAsAttachment

    @property
    def MarkAsTask(self):
        return MarkAsTaskRuleAction(self.com_object.MarkAsTask)

    @property
    def markastask(self):
        """Lower case alias for MarkAsTask"""
        return self.MarkAsTask

    @property
    def MoveToFolder(self):
        return MoveOrCopyRuleAction(self.com_object.MoveToFolder)

    @property
    def movetofolder(self):
        """Lower case alias for MoveToFolder"""
        return self.MoveToFolder

    @property
    def NewItemAlert(self):
        return NewItemAlertRuleAction(self.com_object.NewItemAlert)

    @property
    def newitemalert(self):
        """Lower case alias for NewItemAlert"""
        return self.NewItemAlert

    @property
    def NotifyDelivery(self):
        return RuleAction(self.com_object.NotifyDelivery)

    @property
    def notifydelivery(self):
        """Lower case alias for NotifyDelivery"""
        return self.NotifyDelivery

    @property
    def NotifyRead(self):
        return RuleAction(self.com_object.NotifyRead)

    @property
    def notifyread(self):
        """Lower case alias for NotifyRead"""
        return self.NotifyRead

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PlaySound(self):
        return PlaySoundRuleAction(self.com_object.PlaySound)

    @property
    def playsound(self):
        """Lower case alias for PlaySound"""
        return self.PlaySound

    @property
    def Redirect(self):
        return SendRuleAction(self.com_object.Redirect)

    @property
    def redirect(self):
        """Lower case alias for Redirect"""
        return self.Redirect

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Stop(self):
        return RuleAction(self.com_object.Stop)

    @property
    def stop(self):
        """Lower case alias for Stop"""
        return self.Stop

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return RuleCondition(self.com_object.Enabled)

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class RuleConditions:

    def __init__(self, ruleconditions=None):
        self.com_object= ruleconditions

    @property
    def Account(self):
        return AccountRuleCondition(self.com_object.Account)

    @property
    def account(self):
        """Lower case alias for Account"""
        return self.Account

    @property
    def AnyCategory(self):
        return RuleCondition(self.com_object.AnyCategory)

    @property
    def anycategory(self):
        """Lower case alias for AnyCategory"""
        return self.AnyCategory

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Body(self):
        return TextRuleCondition(self.com_object.Body)

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @property
    def BodyOrSubject(self):
        return TextRuleCondition(self.com_object.BodyOrSubject)

    @property
    def bodyorsubject(self):
        """Lower case alias for BodyOrSubject"""
        return self.BodyOrSubject

    @property
    def Category(self):
        return CategoryRuleCondition(self.com_object.Category)

    @property
    def category(self):
        """Lower case alias for Category"""
        return self.Category

    @property
    def CC(self):
        return RuleCondition(self.com_object.CC)

    @property
    def cc(self):
        """Lower case alias for CC"""
        return self.CC

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def Count(self):
        return self.com_object.Count

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def FormName(self):
        return FormNameRuleCondition(self.com_object.FormName)

    @property
    def formname(self):
        """Lower case alias for FormName"""
        return self.FormName

    @property
    def From(self):
        return ToOrFromRuleCondition(self.com_object.From)

    @property
    def FromAnyRSSFeed(self):
        return RuleCondition(self.com_object.FromAnyRSSFeed)

    @property
    def fromanyrssfeed(self):
        """Lower case alias for FromAnyRSSFeed"""
        return self.FromAnyRSSFeed

    @property
    def FromRssFeed(self):
        return FromRssFeedRuleCondition(self.com_object.FromRssFeed)

    @property
    def fromrssfeed(self):
        """Lower case alias for FromRssFeed"""
        return self.FromRssFeed

    @property
    def HasAttachment(self):
        return RuleCondition(self.com_object.HasAttachment)

    @property
    def hasattachment(self):
        """Lower case alias for HasAttachment"""
        return self.HasAttachment

    @property
    def Importance(self):
        return ImportanceRuleCondition(self.com_object.Importance)

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @property
    def MeetingInviteOrUpdate(self):
        return RuleCondition(self.com_object.MeetingInviteOrUpdate)

    @property
    def meetinginviteorupdate(self):
        """Lower case alias for MeetingInviteOrUpdate"""
        return self.MeetingInviteOrUpdate

    @property
    def MessageHeader(self):
        return TextRuleCondition(self.com_object.MessageHeader)

    @property
    def messageheader(self):
        """Lower case alias for MessageHeader"""
        return self.MessageHeader

    @property
    def NotTo(self):
        return RuleCondition(self.com_object.NotTo)

    @property
    def notto(self):
        """Lower case alias for NotTo"""
        return self.NotTo

    @property
    def OnLocalMachine(self):
        return RuleCondition(self.com_object.OnLocalMachine)

    @property
    def onlocalmachine(self):
        """Lower case alias for OnLocalMachine"""
        return self.OnLocalMachine

    @property
    def OnlyToMe(self):
        return RuleCondition(self.com_object.OnlyToMe)

    @property
    def onlytome(self):
        """Lower case alias for OnlyToMe"""
        return self.OnlyToMe

    @property
    def OnOtherMachine(self):
        return RuleCondition(self.com_object.OnOtherMachine)

    @property
    def onothermachine(self):
        """Lower case alias for OnOtherMachine"""
        return self.OnOtherMachine

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RecipientAddress(self):
        return AddressRuleCondition(self.com_object.RecipientAddress)

    @property
    def recipientaddress(self):
        """Lower case alias for RecipientAddress"""
        return self.RecipientAddress

    @property
    def SenderAddress(self):
        return AddressRuleCondition(self.com_object.SenderAddress)

    @property
    def senderaddress(self):
        """Lower case alias for SenderAddress"""
        return self.SenderAddress

    @property
    def SenderInAddressList(self):
        return SenderInAddressListRuleCondition(self.com_object.SenderInAddressList)

    @property
    def senderinaddresslist(self):
        """Lower case alias for SenderInAddressList"""
        return self.SenderInAddressList

    @property
    def SentTo(self):
        return ToOrFromRuleCondition(self.com_object.SentTo)

    @property
    def sentto(self):
        """Lower case alias for SentTo"""
        return self.SentTo

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Subject(self):
        return TextRuleCondition(self.com_object.Subject)

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @property
    def ToMe(self):
        return RuleCondition(self.com_object.ToMe)

    @property
    def tome(self):
        """Lower case alias for ToMe"""
        return self.ToMe

    @property
    def ToOrCc(self):
        return RuleCondition(self.com_object.ToOrCc)

    @property
    def toorcc(self):
        """Lower case alias for ToOrCc"""
        return self.ToOrCc

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def IsRssRulesProcessingEnabled(self):
        return self.com_object.IsRssRulesProcessingEnabled

    @IsRssRulesProcessingEnabled.setter
    def IsRssRulesProcessingEnabled(self, value):
        self.com_object.IsRssRulesProcessingEnabled = value

    @property
    def isrssrulesprocessingenabled(self):
        """Lower case alias for IsRssRulesProcessingEnabled"""
        return self.IsRssRulesProcessingEnabled

    @isrssrulesprocessingenabled.setter
    def isrssrulesprocessingenabled(self, value):
        """Lower case alias for IsRssRulesProcessingEnabled.setter"""
        self.IsRssRulesProcessingEnabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Create(self, Name=None, RuleType=None):
        arguments = com_arguments([unwrap(a) for a in [Name, RuleType]])
        return Rule(self.com_object.Create(*arguments))

    # Lower case alias for Create
    def create(self, Name=None, RuleType=None):
        arguments = [Name, RuleType]
        return self.Create(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return Rule(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

    def Save(self, ShowProgress=None):
        arguments = com_arguments([unwrap(a) for a in [ShowProgress]])
        self.com_object.Save(*arguments)

    # Lower case alias for Save
    def save(self, ShowProgress=None):
        arguments = [ShowProgress]
        return self.Save(*arguments)


class scrollbar:

    def __init__(self, scrollbar=None):
        self.com_object= scrollbar


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

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @property
    def IsSynchronous(self):
        return self.com_object.IsSynchronous

    @property
    def issynchronous(self):
        """Lower case alias for IsSynchronous"""
        return self.IsSynchronous

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Results(self):
        return Results(self.com_object.Results)

    @property
    def results(self):
        """Lower case alias for Results"""
        return self.Results

    @property
    def Scope(self):
        return self.com_object.Scope

    @property
    def scope(self):
        """Lower case alias for Scope"""
        return self.Scope

    @property
    def SearchSubFolders(self):
        return self.com_object.SearchSubFolders

    @property
    def searchsubfolders(self):
        """Lower case alias for SearchSubFolders"""
        return self.SearchSubFolders

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Tag(self):
        return self.com_object.Tag

    @property
    def tag(self):
        """Lower case alias for Tag"""
        return self.Tag

    def GetTable(self):
        return Table(self.com_object.GetTable())

    # Lower case alias for GetTable
    def gettable(self):
        return self.GetTable()

    def Save(self, SchFldrName=None):
        arguments = com_arguments([unwrap(a) for a in [SchFldrName]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Location(self):
        return OlSelectionLocation(self.com_object.Location)

    @property
    def location(self):
        """Lower case alias for Location"""
        return self.Location

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def GetSelection(self, SelectionContents=None):
        arguments = com_arguments([unwrap(a) for a in [SelectionContents]])
        return Selection(self.com_object.GetSelection(*arguments))

    # Lower case alias for GetSelection
    def getselection(self, SelectionContents=None):
        arguments = [SelectionContents]
        return self.GetSelection(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

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

    @property
    def allowmultipleselection(self):
        """Lower case alias for AllowMultipleSelection"""
        return self.AllowMultipleSelection

    @allowmultipleselection.setter
    def allowmultipleselection(self, value):
        """Lower case alias for AllowMultipleSelection.setter"""
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

    @property
    def bcclabel(self):
        """Lower case alias for BccLabel"""
        return self.BccLabel

    @bcclabel.setter
    def bcclabel(self, value):
        """Lower case alias for BccLabel.setter"""
        self.BccLabel = value

    @property
    def Caption(self):
        return self.com_object.Caption

    @Caption.setter
    def Caption(self, value):
        self.com_object.Caption = value

    @property
    def caption(self):
        """Lower case alias for Caption"""
        return self.Caption

    @caption.setter
    def caption(self, value):
        """Lower case alias for Caption.setter"""
        self.Caption = value

    @property
    def CcLabel(self):
        return self.com_object.CcLabel

    @CcLabel.setter
    def CcLabel(self, value):
        self.com_object.CcLabel = value

    @property
    def cclabel(self):
        """Lower case alias for CcLabel"""
        return self.CcLabel

    @cclabel.setter
    def cclabel(self, value):
        """Lower case alias for CcLabel.setter"""
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

    @property
    def forceresolution(self):
        """Lower case alias for ForceResolution"""
        return self.ForceResolution

    @forceresolution.setter
    def forceresolution(self, value):
        """Lower case alias for ForceResolution.setter"""
        self.ForceResolution = value

    @property
    def InitialAddressList(self):
        return AddressList(self.com_object.InitialAddressList)

    @InitialAddressList.setter
    def InitialAddressList(self, value):
        self.com_object.InitialAddressList = value

    @property
    def initialaddresslist(self):
        """Lower case alias for InitialAddressList"""
        return self.InitialAddressList

    @initialaddresslist.setter
    def initialaddresslist(self, value):
        """Lower case alias for InitialAddressList.setter"""
        self.InitialAddressList = value

    @property
    def NumberOfRecipientSelectors(self):
        return OlRecipientSelectors(self.com_object.NumberOfRecipientSelectors)

    @NumberOfRecipientSelectors.setter
    def NumberOfRecipientSelectors(self, value):
        self.com_object.NumberOfRecipientSelectors = value

    @property
    def numberofrecipientselectors(self):
        """Lower case alias for NumberOfRecipientSelectors"""
        return self.NumberOfRecipientSelectors

    @numberofrecipientselectors.setter
    def numberofrecipientselectors(self, value):
        """Lower case alias for NumberOfRecipientSelectors.setter"""
        self.NumberOfRecipientSelectors = value

    @property
    def Parent(self):
        return SelectNamesDialog(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @Recipients.setter
    def Recipients(self, value):
        self.com_object.Recipients = value

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @recipients.setter
    def recipients(self, value):
        """Lower case alias for Recipients.setter"""
        self.Recipients = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ShowOnlyInitialAddressList(self):
        return AddressList(self.com_object.ShowOnlyInitialAddressList)

    @ShowOnlyInitialAddressList.setter
    def ShowOnlyInitialAddressList(self, value):
        self.com_object.ShowOnlyInitialAddressList = value

    @property
    def showonlyinitialaddresslist(self):
        """Lower case alias for ShowOnlyInitialAddressList"""
        return self.ShowOnlyInitialAddressList

    @showonlyinitialaddresslist.setter
    def showonlyinitialaddresslist(self, value):
        """Lower case alias for ShowOnlyInitialAddressList.setter"""
        self.ShowOnlyInitialAddressList = value

    @property
    def ToLabel(self):
        return self.com_object.ToLabel

    @ToLabel.setter
    def ToLabel(self, value):
        self.com_object.ToLabel = value

    @property
    def tolabel(self):
        """Lower case alias for ToLabel"""
        return self.ToLabel

    @tolabel.setter
    def tolabel(self, value):
        """Lower case alias for ToLabel.setter"""
        self.ToLabel = value

    def Display(self):
        return self.com_object.Display()

    # Lower case alias for Display
    def display(self):
        return self.Display()

    def SetDefaultDisplayMode(self, defaultMode=None):
        arguments = com_arguments([unwrap(a) for a in [defaultMode]])
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

    @property
    def addresslist(self):
        """Lower case alias for AddressList"""
        return self.AddressList

    @addresslist.setter
    def addresslist(self, value):
        """Lower case alias for AddressList.setter"""
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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class SendRuleAction:

    def __init__(self, sendruleaction=None):
        self.com_object= sendruleaction

    @property
    def ActionType(self):
        return OlRuleActionType(self.com_object.ActionType)

    @property
    def actiontype(self):
        """Lower case alias for ActionType"""
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

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session


class SharingItem:

    def __init__(self, sharingitem=None):
        self.com_object= sharingitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def AllowWriteAccess(self):
        return self.com_object.AllowWriteAccess

    @AllowWriteAccess.setter
    def AllowWriteAccess(self, value):
        self.com_object.AllowWriteAccess = value

    @property
    def allowwriteaccess(self):
        """Lower case alias for AllowWriteAccess"""
        return self.AllowWriteAccess

    @allowwriteaccess.setter
    def allowwriteaccess(self, value):
        """Lower case alias for AllowWriteAccess.setter"""
        self.AllowWriteAccess = value

    @property
    def AlternateRecipientAllowed(self):
        return self.com_object.AlternateRecipientAllowed

    @AlternateRecipientAllowed.setter
    def AlternateRecipientAllowed(self, value):
        self.com_object.AlternateRecipientAllowed = value

    @property
    def alternaterecipientallowed(self):
        """Lower case alias for AlternateRecipientAllowed"""
        return self.AlternateRecipientAllowed

    @alternaterecipientallowed.setter
    def alternaterecipientallowed(self, value):
        """Lower case alias for AlternateRecipientAllowed.setter"""
        self.AlternateRecipientAllowed = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoForwarded(self):
        return self.com_object.AutoForwarded

    @AutoForwarded.setter
    def AutoForwarded(self, value):
        self.com_object.AutoForwarded = value

    @property
    def autoforwarded(self):
        """Lower case alias for AutoForwarded"""
        return self.AutoForwarded

    @autoforwarded.setter
    def autoforwarded(self, value):
        """Lower case alias for AutoForwarded.setter"""
        self.AutoForwarded = value

    @property
    def BCC(self):
        return SharingItem(self.com_object.BCC)

    @BCC.setter
    def BCC(self, value):
        self.com_object.BCC = value

    @property
    def bcc(self):
        """Lower case alias for BCC"""
        return self.BCC

    @bcc.setter
    def bcc(self, value):
        """Lower case alias for BCC.setter"""
        self.BCC = value

    @property
    def BillingInformation(self):
        return SharingItem(self.com_object.BillingInformation)

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return SharingItem(self.com_object.Body)

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def BodyFormat(self):
        return OlBodyFormat(self.com_object.BodyFormat)

    @BodyFormat.setter
    def BodyFormat(self, value):
        self.com_object.BodyFormat = value

    @property
    def bodyformat(self):
        """Lower case alias for BodyFormat"""
        return self.BodyFormat

    @bodyformat.setter
    def bodyformat(self, value):
        """Lower case alias for BodyFormat.setter"""
        self.BodyFormat = value

    @property
    def Categories(self):
        return SharingItem(self.com_object.Categories)

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
        self.Categories = value

    @property
    def CC(self):
        return SharingItem(self.com_object.CC)

    @CC.setter
    def CC(self, value):
        self.com_object.CC = value

    @property
    def cc(self):
        """Lower case alias for CC"""
        return self.CC

    @cc.setter
    def cc(self, value):
        """Lower case alias for CC.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return SharingItem(self.com_object.ConversationIndex)

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return SharingItem(self.com_object.ConversationTopic)

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return SharingItem(self.com_object.CreationTime)

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DeferredDeliveryTime(self):
        return SharingItem(self.com_object.DeferredDeliveryTime)

    @DeferredDeliveryTime.setter
    def DeferredDeliveryTime(self, value):
        self.com_object.DeferredDeliveryTime = value

    @property
    def deferreddeliverytime(self):
        """Lower case alias for DeferredDeliveryTime"""
        return self.DeferredDeliveryTime

    @deferreddeliverytime.setter
    def deferreddeliverytime(self, value):
        """Lower case alias for DeferredDeliveryTime.setter"""
        self.DeferredDeliveryTime = value

    @property
    def DeleteAfterSubmit(self):
        return self.com_object.DeleteAfterSubmit

    @DeleteAfterSubmit.setter
    def DeleteAfterSubmit(self, value):
        self.com_object.DeleteAfterSubmit = value

    @property
    def deleteaftersubmit(self):
        """Lower case alias for DeleteAfterSubmit"""
        return self.DeleteAfterSubmit

    @deleteaftersubmit.setter
    def deleteaftersubmit(self, value):
        """Lower case alias for DeleteAfterSubmit.setter"""
        self.DeleteAfterSubmit = value

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return SharingItem(self.com_object.EntryID)

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def ExpiryTime(self):
        return SharingItem(self.com_object.ExpiryTime)

    @ExpiryTime.setter
    def ExpiryTime(self, value):
        self.com_object.ExpiryTime = value

    @property
    def expirytime(self):
        """Lower case alias for ExpiryTime"""
        return self.ExpiryTime

    @expirytime.setter
    def expirytime(self, value):
        """Lower case alias for ExpiryTime.setter"""
        self.ExpiryTime = value

    @property
    def FlagRequest(self):
        return SharingItem(self.com_object.FlagRequest)

    @FlagRequest.setter
    def FlagRequest(self, value):
        self.com_object.FlagRequest = value

    @property
    def flagrequest(self):
        """Lower case alias for FlagRequest"""
        return self.FlagRequest

    @flagrequest.setter
    def flagrequest(self, value):
        """Lower case alias for FlagRequest.setter"""
        self.FlagRequest = value

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def HTMLBody(self):
        return SharingItem(self.com_object.HTMLBody)

    @HTMLBody.setter
    def HTMLBody(self, value):
        self.com_object.HTMLBody = value

    @property
    def htmlbody(self):
        """Lower case alias for HTMLBody"""
        return self.HTMLBody

    @htmlbody.setter
    def htmlbody(self, value):
        """Lower case alias for HTMLBody.setter"""
        self.HTMLBody = value

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    @property
    def internetcodepage(self):
        """Lower case alias for InternetCodepage"""
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        """Lower case alias for InternetCodepage.setter"""
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return SharingItem(self.com_object.IsConflict)

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def IsMarkedAsTask(self):
        return SharingItem(self.com_object.IsMarkedAsTask)

    @property
    def ismarkedastask(self):
        """Lower case alias for IsMarkedAsTask"""
        return self.IsMarkedAsTask

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return SharingItem(self.com_object.LastModificationTime)

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return SharingItem(self.com_object.MessageClass)

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return SharingItem(self.com_object.Mileage)

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return SharingItem(self.com_object.NoAging)

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OriginatorDeliveryReportRequested(self):
        return SharingItem(self.com_object.OriginatorDeliveryReportRequested)

    @OriginatorDeliveryReportRequested.setter
    def OriginatorDeliveryReportRequested(self, value):
        self.com_object.OriginatorDeliveryReportRequested = value

    @property
    def originatordeliveryreportrequested(self):
        """Lower case alias for OriginatorDeliveryReportRequested"""
        return self.OriginatorDeliveryReportRequested

    @originatordeliveryreportrequested.setter
    def originatordeliveryreportrequested(self, value):
        """Lower case alias for OriginatorDeliveryReportRequested.setter"""
        self.OriginatorDeliveryReportRequested = value

    @property
    def OutlookInternalVersion(self):
        return SharingItem(self.com_object.OutlookInternalVersion)

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return SharingItem(self.com_object.OutlookVersion)

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return SharingItem(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Permission(self):
        return self.com_object.Permission

    @Permission.setter
    def Permission(self, value):
        self.com_object.Permission = value

    @property
    def permission(self):
        """Lower case alias for Permission"""
        return self.Permission

    @permission.setter
    def permission(self, value):
        """Lower case alias for Permission.setter"""
        self.Permission = value

    @property
    def PermissionService(self):
        return self.com_object.PermissionService

    @PermissionService.setter
    def PermissionService(self, value):
        self.com_object.PermissionService = value

    @property
    def permissionservice(self):
        """Lower case alias for PermissionService"""
        return self.PermissionService

    @permissionservice.setter
    def permissionservice(self, value):
        """Lower case alias for PermissionService.setter"""
        self.PermissionService = value

    @property
    def PermissionTemplateGuid(self):
        return SharingItem(self.com_object.PermissionTemplateGuid)

    @PermissionTemplateGuid.setter
    def PermissionTemplateGuid(self, value):
        self.com_object.PermissionTemplateGuid = value

    @property
    def permissiontemplateguid(self):
        """Lower case alias for PermissionTemplateGuid"""
        return self.PermissionTemplateGuid

    @permissiontemplateguid.setter
    def permissiontemplateguid(self, value):
        """Lower case alias for PermissionTemplateGuid.setter"""
        self.PermissionTemplateGuid = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def ReadReceiptRequested(self):
        return self.com_object.ReadReceiptRequested

    @property
    def readreceiptrequested(self):
        """Lower case alias for ReadReceiptRequested"""
        return self.ReadReceiptRequested

    @property
    def ReceivedByEntryID(self):
        return self.com_object.ReceivedByEntryID

    @property
    def receivedbyentryid(self):
        """Lower case alias for ReceivedByEntryID"""
        return self.ReceivedByEntryID

    @property
    def ReceivedByName(self):
        return SharingItem(self.com_object.ReceivedByName)

    @property
    def receivedbyname(self):
        """Lower case alias for ReceivedByName"""
        return self.ReceivedByName

    @property
    def ReceivedOnBehalfOfEntryID(self):
        return self.com_object.ReceivedOnBehalfOfEntryID

    @property
    def receivedonbehalfofentryid(self):
        """Lower case alias for ReceivedOnBehalfOfEntryID"""
        return self.ReceivedOnBehalfOfEntryID

    @property
    def ReceivedOnBehalfOfName(self):
        return SharingItem(self.com_object.ReceivedOnBehalfOfName)

    @property
    def receivedonbehalfofname(self):
        """Lower case alias for ReceivedOnBehalfOfName"""
        return self.ReceivedOnBehalfOfName

    @property
    def ReceivedTime(self):
        return SharingItem(self.com_object.ReceivedTime)

    @property
    def receivedtime(self):
        """Lower case alias for ReceivedTime"""
        return self.ReceivedTime

    @property
    def RecipientReassignmentProhibited(self):
        return SharingItem(self.com_object.RecipientReassignmentProhibited)

    @RecipientReassignmentProhibited.setter
    def RecipientReassignmentProhibited(self, value):
        self.com_object.RecipientReassignmentProhibited = value

    @property
    def recipientreassignmentprohibited(self):
        """Lower case alias for RecipientReassignmentProhibited"""
        return self.RecipientReassignmentProhibited

    @recipientreassignmentprohibited.setter
    def recipientreassignmentprohibited(self, value):
        """Lower case alias for RecipientReassignmentProhibited.setter"""
        self.RecipientReassignmentProhibited = value

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return SharingItem(self.com_object.ReminderOverrideDefault)

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    @property
    def reminderoverridedefault(self):
        """Lower case alias for ReminderOverrideDefault"""
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        """Lower case alias for ReminderOverrideDefault.setter"""
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return SharingItem(self.com_object.ReminderPlaySound)

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    @property
    def reminderplaysound(self):
        """Lower case alias for ReminderPlaySound"""
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        """Lower case alias for ReminderPlaySound.setter"""
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return SharingItem(self.com_object.ReminderSet)

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    @property
    def reminderset(self):
        """Lower case alias for ReminderSet"""
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        """Lower case alias for ReminderSet.setter"""
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    @property
    def remindersoundfile(self):
        """Lower case alias for ReminderSoundFile"""
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        """Lower case alias for ReminderSoundFile.setter"""
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return SharingItem(self.com_object.ReminderTime)

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    @property
    def remindertime(self):
        """Lower case alias for ReminderTime"""
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        """Lower case alias for ReminderTime.setter"""
        self.ReminderTime = value

    @property
    def RemoteID(self):
        return SharingItem(self.com_object.RemoteID)

    @property
    def remoteid(self):
        """Lower case alias for RemoteID"""
        return self.RemoteID

    @property
    def RemoteName(self):
        return SharingItem(self.com_object.RemoteName)

    @property
    def remotename(self):
        """Lower case alias for RemoteName"""
        return self.RemoteName

    @property
    def RemotePath(self):
        return SharingItem(self.com_object.RemotePath)

    @property
    def remotepath(self):
        """Lower case alias for RemotePath"""
        return self.RemotePath

    @property
    def RemoteStatus(self):
        return OlRemoteStatus(self.com_object.RemoteStatus)

    @RemoteStatus.setter
    def RemoteStatus(self, value):
        self.com_object.RemoteStatus = value

    @property
    def remotestatus(self):
        """Lower case alias for RemoteStatus"""
        return self.RemoteStatus

    @remotestatus.setter
    def remotestatus(self, value):
        """Lower case alias for RemoteStatus.setter"""
        self.RemoteStatus = value

    @property
    def ReplyRecipientNames(self):
        return SharingItem(self.com_object.ReplyRecipientNames)

    @property
    def replyrecipientnames(self):
        """Lower case alias for ReplyRecipientNames"""
        return self.ReplyRecipientNames

    @property
    def ReplyRecipients(self):
        return Recipients(self.com_object.ReplyRecipients)

    @property
    def replyrecipients(self):
        """Lower case alias for ReplyRecipients"""
        return self.ReplyRecipients

    @property
    def RequestedFolder(self):
        return OlDefaultFolders(self.com_object.RequestedFolder)

    @property
    def requestedfolder(self):
        """Lower case alias for RequestedFolder"""
        return self.RequestedFolder

    @property
    def RetentionExpirationDate(self):
        return SharingItem(self.com_object.RetentionExpirationDate)

    @property
    def retentionexpirationdate(self):
        """Lower case alias for RetentionExpirationDate"""
        return self.RetentionExpirationDate

    @property
    def RetentionPolicyName(self):
        return self.com_object.RetentionPolicyName

    @property
    def retentionpolicyname(self):
        """Lower case alias for RetentionPolicyName"""
        return self.RetentionPolicyName

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return SharingItem(self.com_object.Saved)

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def SaveSentMessageFolder(self):
        return Folder(self.com_object.SaveSentMessageFolder)

    @SaveSentMessageFolder.setter
    def SaveSentMessageFolder(self, value):
        self.com_object.SaveSentMessageFolder = value

    @property
    def savesentmessagefolder(self):
        """Lower case alias for SaveSentMessageFolder"""
        return self.SaveSentMessageFolder

    @savesentmessagefolder.setter
    def savesentmessagefolder(self, value):
        """Lower case alias for SaveSentMessageFolder.setter"""
        self.SaveSentMessageFolder = value

    @property
    def SenderEmailAddress(self):
        return SharingItem(self.com_object.SenderEmailAddress)

    @property
    def senderemailaddress(self):
        """Lower case alias for SenderEmailAddress"""
        return self.SenderEmailAddress

    @property
    def SenderEmailType(self):
        return SharingItem(self.com_object.SenderEmailType)

    @property
    def senderemailtype(self):
        """Lower case alias for SenderEmailType"""
        return self.SenderEmailType

    @property
    def SenderName(self):
        return SharingItem(self.com_object.SenderName)

    @property
    def sendername(self):
        """Lower case alias for SenderName"""
        return self.SenderName

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    @property
    def sendusingaccount(self):
        """Lower case alias for SendUsingAccount"""
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        """Lower case alias for SendUsingAccount.setter"""
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Sent(self):
        return SharingItem(self.com_object.Sent)

    @property
    def sent(self):
        """Lower case alias for Sent"""
        return self.Sent

    @property
    def SentOn(self):
        return SharingItem(self.com_object.SentOn)

    @property
    def senton(self):
        """Lower case alias for SentOn"""
        return self.SentOn

    @property
    def SentOnBehalfOfName(self):
        return SharingItem(self.com_object.SentOnBehalfOfName)

    @SentOnBehalfOfName.setter
    def SentOnBehalfOfName(self, value):
        self.com_object.SentOnBehalfOfName = value

    @property
    def sentonbehalfofname(self):
        """Lower case alias for SentOnBehalfOfName"""
        return self.SentOnBehalfOfName

    @sentonbehalfofname.setter
    def sentonbehalfofname(self, value):
        """Lower case alias for SentOnBehalfOfName.setter"""
        self.SentOnBehalfOfName = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def SharingProvider(self):
        return OlSharingProvider(self.com_object.SharingProvider)

    @property
    def sharingprovider(self):
        """Lower case alias for SharingProvider"""
        return self.SharingProvider

    @property
    def SharingProviderGuid(self):
        return SharingItem(self.com_object.SharingProviderGuid)

    @property
    def sharingproviderguid(self):
        """Lower case alias for SharingProviderGuid"""
        return self.SharingProviderGuid

    @property
    def Size(self):
        return SharingItem(self.com_object.Size)

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return SharingItem(self.com_object.Subject)

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def Submitted(self):
        return SharingItem(self.com_object.Submitted)

    @property
    def submitted(self):
        """Lower case alias for Submitted"""
        return self.Submitted

    @property
    def TaskCompletedDate(self):
        return SharingItem(self.com_object.TaskCompletedDate)

    @TaskCompletedDate.setter
    def TaskCompletedDate(self, value):
        self.com_object.TaskCompletedDate = value

    @property
    def taskcompleteddate(self):
        """Lower case alias for TaskCompletedDate"""
        return self.TaskCompletedDate

    @taskcompleteddate.setter
    def taskcompleteddate(self, value):
        """Lower case alias for TaskCompletedDate.setter"""
        self.TaskCompletedDate = value

    @property
    def TaskDueDate(self):
        return SharingItem(self.com_object.TaskDueDate)

    @TaskDueDate.setter
    def TaskDueDate(self, value):
        self.com_object.TaskDueDate = value

    @property
    def taskduedate(self):
        """Lower case alias for TaskDueDate"""
        return self.TaskDueDate

    @taskduedate.setter
    def taskduedate(self, value):
        """Lower case alias for TaskDueDate.setter"""
        self.TaskDueDate = value

    @property
    def TaskStartDate(self):
        return SharingItem(self.com_object.TaskStartDate)

    @TaskStartDate.setter
    def TaskStartDate(self, value):
        self.com_object.TaskStartDate = value

    @property
    def taskstartdate(self):
        """Lower case alias for TaskStartDate"""
        return self.TaskStartDate

    @taskstartdate.setter
    def taskstartdate(self, value):
        """Lower case alias for TaskStartDate.setter"""
        self.TaskStartDate = value

    @property
    def TaskSubject(self):
        return SharingItem(self.com_object.TaskSubject)

    @TaskSubject.setter
    def TaskSubject(self, value):
        self.com_object.TaskSubject = value

    @property
    def tasksubject(self):
        """Lower case alias for TaskSubject"""
        return self.TaskSubject

    @tasksubject.setter
    def tasksubject(self, value):
        """Lower case alias for TaskSubject.setter"""
        self.TaskSubject = value

    @property
    def To(self):
        return SharingItem(self.com_object.To)

    @To.setter
    def To(self, value):
        self.com_object.To = value

    @property
    def to(self):
        """Lower case alias for To"""
        return self.To

    @to.setter
    def to(self, value):
        """Lower case alias for To.setter"""
        self.To = value

    @property
    def ToDoTaskOrdinal(self):
        return SharingItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    @property
    def todotaskordinal(self):
        """Lower case alias for ToDoTaskOrdinal"""
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        """Lower case alias for ToDoTaskOrdinal.setter"""
        self.ToDoTaskOrdinal = value

    @property
    def Type(self):
        return OlSharingMsgType(self.com_object.Type)

    @Type.setter
    def Type(self, value):
        self.com_object.Type = value

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @type.setter
    def type(self, value):
        """Lower case alias for Type.setter"""
        self.Type = value

    @property
    def UnRead(self):
        return SharingItem(self.com_object.UnRead)

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def AddBusinessCard(self, contact=None):
        arguments = com_arguments([unwrap(a) for a in [contact]])
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
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [MarkInterval]])
        self.com_object.MarkAsTask(*arguments)

    # Lower case alias for MarkAsTask
    def markastask(self, MarkInterval=None):
        arguments = [MarkInterval]
        return self.MarkAsTask(*arguments)

    def Move(self, DestFldr=None):
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return SimpleItems(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return self.com_object.Item(*arguments)

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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.com_object.NavigationModuleType)

    @property
    def navigationmoduletype(self):
        """Lower case alias for NavigationModuleType"""
        return self.NavigationModuleType

    @property
    def Parent(self):
        return SolutionsModule(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return SolutionsModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return self.com_object.Visible

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
        self.Visible = value

    def AddSolution(self, Solution=None, Scope=None):
        arguments = com_arguments([unwrap(a) for a in [Solution, Scope]])
        self.com_object.AddSolution(*arguments)

    # Lower case alias for AddSolution
    def addsolution(self, Solution=None, Scope=None):
        arguments = [Solution, Scope]
        return self.AddSolution(*arguments)


class spinbutton:

    def __init__(self, spinbutton=None):
        self.com_object= spinbutton


class StorageItem:

    def __init__(self, storageitem=None):
        self.com_object= storageitem

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def CreationTime(self):
        return StorageItem(self.com_object.CreationTime)

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def Creator(self):
        return StorageItem(self.com_object.Creator)

    @Creator.setter
    def Creator(self, value):
        self.com_object.Creator = value

    @property
    def creator(self):
        """Lower case alias for Creator"""
        return self.Creator

    @creator.setter
    def creator(self, value):
        """Lower case alias for Creator.setter"""
        self.Creator = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return StorageItem(self.com_object.Size)

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
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

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DisplayName(self):
        return Store(self.com_object.DisplayName)

    @property
    def displayname(self):
        """Lower case alias for DisplayName"""
        return self.DisplayName

    @property
    def ExchangeStoreType(self):
        return OlExchangeStoreType(self.com_object.ExchangeStoreType)

    @property
    def exchangestoretype(self):
        """Lower case alias for ExchangeStoreType"""
        return self.ExchangeStoreType

    @property
    def FilePath(self):
        return self.com_object.FilePath

    @property
    def filepath(self):
        """Lower case alias for FilePath"""
        return self.FilePath

    @property
    def IsCachedExchange(self):
        return Store(self.com_object.IsCachedExchange)

    @property
    def iscachedexchange(self):
        """Lower case alias for IsCachedExchange"""
        return self.IsCachedExchange

    @property
    def IsConversationEnabled(self):
        return self.com_object.IsConversationEnabled

    @property
    def isconversationenabled(self):
        """Lower case alias for IsConversationEnabled"""
        return self.IsConversationEnabled

    @property
    def IsDataFileStore(self):
        return Store(self.com_object.IsDataFileStore)

    @property
    def isdatafilestore(self):
        """Lower case alias for IsDataFileStore"""
        return self.IsDataFileStore

    @property
    def IsInstantSearchEnabled(self):
        return self.com_object.IsInstantSearchEnabled

    @property
    def isinstantsearchenabled(self):
        """Lower case alias for IsInstantSearchEnabled"""
        return self.IsInstantSearchEnabled

    @property
    def IsOpen(self):
        return Store(self.com_object.IsOpen)

    @property
    def isopen(self):
        """Lower case alias for IsOpen"""
        return self.IsOpen

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def StoreID(self):
        return Store(self.com_object.StoreID)

    @property
    def storeid(self):
        """Lower case alias for StoreID"""
        return self.StoreID

    def createunifiedgroup(self, Name=None, Name=None, Alias=None, Description=None, FAutoSubscribeMembers=None, GroupType=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Name, Alias, Description, FAutoSubscribeMembers, GroupType]])
        return self.com_object.createunifiedgroup(*arguments)

    # Lower case alias for createunifiedgroup
    def createunifiedgroup(self, Name=None, Name=None, Alias=None, Description=None, FAutoSubscribeMembers=None, GroupType=None):
        arguments = [Name, Name, Alias, Description, FAutoSubscribeMembers, GroupType]
        return self.createunifiedgroup(*arguments)

    def deleteunifiedgroup(self):
        return self.com_object.deleteunifiedgroup()

    # Lower case alias for deleteunifiedgroup
    def deleteunifiedgroup(self):
        return self.deleteunifiedgroup()

    def GetDefaultFolder(self, FolderType=None):
        arguments = com_arguments([unwrap(a) for a in [FolderType]])
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
        arguments = com_arguments([unwrap(a) for a in [FolderType]])
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def appfolders(self):
        """Lower case alias for AppFolders"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return SyncObject(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class tab:

    def __init__(self, tab=None):
        self.com_object= tab


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

    @property
    def columns(self):
        """Lower case alias for Columns"""
        return self.Columns

    @property
    def EndOfTable(self):
        return Table(self.com_object.EndOfTable)

    @property
    def endoftable(self):
        """Lower case alias for EndOfTable"""
        return self.EndOfTable

    @property
    def Parent(self):
        return Table(self.com_object.Parent)

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def FindNextRow(self):
        return self.com_object.FindNextRow()

    # Lower case alias for FindNextRow
    def findnextrow(self):
        return self.FindNextRow()

    def FindRow(self, Filter=None):
        arguments = com_arguments([unwrap(a) for a in [Filter]])
        return self.com_object.FindRow(*arguments)

    # Lower case alias for FindRow
    def findrow(self, Filter=None):
        arguments = [Filter]
        return self.FindRow(*arguments)

    def GetArray(self, MaxRows=None):
        arguments = com_arguments([unwrap(a) for a in [MaxRows]])
        return self.com_object.GetArray(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Filter]])
        return Table(self.com_object.Restrict(*arguments))

    # Lower case alias for Restrict
    def restrict(self, Filter=None):
        arguments = [Filter]
        return self.Restrict(*arguments)

    def Sort(self, SortProperty=None, Descending=None):
        arguments = com_arguments([unwrap(a) for a in [SortProperty, Descending]])
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

    @property
    def allowincellediting(self):
        """Lower case alias for AllowInCellEditing"""
        return self.AllowInCellEditing

    @allowincellediting.setter
    def allowincellediting(self, value):
        """Lower case alias for AllowInCellEditing.setter"""
        self.AllowInCellEditing = value

    @property
    def AlwaysExpandConversation(self):
        return self.com_object.AlwaysExpandConversation

    @AlwaysExpandConversation.setter
    def AlwaysExpandConversation(self, value):
        self.com_object.AlwaysExpandConversation = value

    @property
    def alwaysexpandconversation(self):
        """Lower case alias for AlwaysExpandConversation"""
        return self.AlwaysExpandConversation

    @alwaysexpandconversation.setter
    def alwaysexpandconversation(self, value):
        """Lower case alias for AlwaysExpandConversation.setter"""
        self.AlwaysExpandConversation = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def AutoFormatRules(self):
        return AutoFormatRules(self.com_object.AutoFormatRules)

    @property
    def autoformatrules(self):
        """Lower case alias for AutoFormatRules"""
        return self.AutoFormatRules

    @property
    def AutomaticColumnSizing(self):
        return TableView(self.com_object.AutomaticColumnSizing)

    @AutomaticColumnSizing.setter
    def AutomaticColumnSizing(self, value):
        self.com_object.AutomaticColumnSizing = value

    @property
    def automaticcolumnsizing(self):
        """Lower case alias for AutomaticColumnSizing"""
        return self.AutomaticColumnSizing

    @automaticcolumnsizing.setter
    def automaticcolumnsizing(self, value):
        """Lower case alias for AutomaticColumnSizing.setter"""
        self.AutomaticColumnSizing = value

    @property
    def AutomaticGrouping(self):
        return TableView(self.com_object.AutomaticGrouping)

    @AutomaticGrouping.setter
    def AutomaticGrouping(self, value):
        self.com_object.AutomaticGrouping = value

    @property
    def automaticgrouping(self):
        """Lower case alias for AutomaticGrouping"""
        return self.AutomaticGrouping

    @automaticgrouping.setter
    def automaticgrouping(self, value):
        """Lower case alias for AutomaticGrouping.setter"""
        self.AutomaticGrouping = value

    @property
    def AutoPreview(self):
        return OlAutoPreview(self.com_object.AutoPreview)

    @AutoPreview.setter
    def AutoPreview(self, value):
        self.com_object.AutoPreview = value

    @property
    def autopreview(self):
        """Lower case alias for AutoPreview"""
        return self.AutoPreview

    @autopreview.setter
    def autopreview(self, value):
        """Lower case alias for AutoPreview.setter"""
        self.AutoPreview = value

    @property
    def AutoPreviewFont(self):
        return ViewFont(self.com_object.AutoPreviewFont)

    @property
    def autopreviewfont(self):
        """Lower case alias for AutoPreviewFont"""
        return self.AutoPreviewFont

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def ColumnFont(self):
        return ViewFont(self.com_object.ColumnFont)

    @property
    def columnfont(self):
        """Lower case alias for ColumnFont"""
        return self.ColumnFont

    @property
    def DefaultExpandCollapseSetting(self):
        return OlDefaultExpandCollapseSetting(self.com_object.DefaultExpandCollapseSetting)

    @DefaultExpandCollapseSetting.setter
    def DefaultExpandCollapseSetting(self, value):
        self.com_object.DefaultExpandCollapseSetting = value

    @property
    def defaultexpandcollapsesetting(self):
        """Lower case alias for DefaultExpandCollapseSetting"""
        return self.DefaultExpandCollapseSetting

    @defaultexpandcollapsesetting.setter
    def defaultexpandcollapsesetting(self, value):
        """Lower case alias for DefaultExpandCollapseSetting.setter"""
        self.DefaultExpandCollapseSetting = value

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for Filter.setter"""
        self.Filter = value

    @property
    def GridLineStyle(self):
        return OlGridLineStyle(self.com_object.GridLineStyle)

    @GridLineStyle.setter
    def GridLineStyle(self, value):
        self.com_object.GridLineStyle = value

    @property
    def gridlinestyle(self):
        """Lower case alias for GridLineStyle"""
        return self.GridLineStyle

    @gridlinestyle.setter
    def gridlinestyle(self, value):
        """Lower case alias for GridLineStyle.setter"""
        self.GridLineStyle = value

    @property
    def GroupByFields(self):
        return OrderFields(self.com_object.GroupByFields)

    @property
    def groupbyfields(self):
        """Lower case alias for GroupByFields"""
        return self.GroupByFields

    @property
    def HideReadingPaneHeaderInfo(self):
        return TableView(self.com_object.HideReadingPaneHeaderInfo)

    @HideReadingPaneHeaderInfo.setter
    def HideReadingPaneHeaderInfo(self, value):
        self.com_object.HideReadingPaneHeaderInfo = value

    @property
    def hidereadingpaneheaderinfo(self):
        """Lower case alias for HideReadingPaneHeaderInfo"""
        return self.HideReadingPaneHeaderInfo

    @hidereadingpaneheaderinfo.setter
    def hidereadingpaneheaderinfo(self, value):
        """Lower case alias for HideReadingPaneHeaderInfo.setter"""
        self.HideReadingPaneHeaderInfo = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @language.setter
    def language(self, value):
        """Lower case alias for Language.setter"""
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    @property
    def lockuserchanges(self):
        """Lower case alias for LockUserChanges"""
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        """Lower case alias for LockUserChanges.setter"""
        self.LockUserChanges = value

    @property
    def MaxLinesInMultiLineView(self):
        return TableView(self.com_object.MaxLinesInMultiLineView)

    @MaxLinesInMultiLineView.setter
    def MaxLinesInMultiLineView(self, value):
        self.com_object.MaxLinesInMultiLineView = value

    @property
    def maxlinesinmultilineview(self):
        """Lower case alias for MaxLinesInMultiLineView"""
        return self.MaxLinesInMultiLineView

    @maxlinesinmultilineview.setter
    def maxlinesinmultilineview(self, value):
        """Lower case alias for MaxLinesInMultiLineView.setter"""
        self.MaxLinesInMultiLineView = value

    @property
    def Multiline(self):
        return OlMultiLine(self.com_object.Multiline)

    @Multiline.setter
    def Multiline(self, value):
        self.com_object.Multiline = value

    @property
    def multiline(self):
        """Lower case alias for Multiline"""
        return self.Multiline

    @multiline.setter
    def multiline(self, value):
        """Lower case alias for Multiline.setter"""
        self.Multiline = value

    @property
    def MultiLineWidth(self):
        return TableView(self.com_object.MultiLineWidth)

    @MultiLineWidth.setter
    def MultiLineWidth(self, value):
        self.com_object.MultiLineWidth = value

    @property
    def multilinewidth(self):
        """Lower case alias for MultiLineWidth"""
        return self.MultiLineWidth

    @multilinewidth.setter
    def multilinewidth(self, value):
        """Lower case alias for MultiLineWidth.setter"""
        self.MultiLineWidth = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def RowFont(self):
        return ViewFont(self.com_object.RowFont)

    @property
    def rowfont(self):
        """Lower case alias for RowFont"""
        return self.RowFont

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    @property
    def saveoption(self):
        """Lower case alias for SaveOption"""
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ShowConversationByDate(self):
        return self.com_object.ShowConversationByDate

    @ShowConversationByDate.setter
    def ShowConversationByDate(self, value):
        self.com_object.ShowConversationByDate = value

    @property
    def showconversationbydate(self):
        """Lower case alias for ShowConversationByDate"""
        return self.ShowConversationByDate

    @showconversationbydate.setter
    def showconversationbydate(self, value):
        """Lower case alias for ShowConversationByDate.setter"""
        self.ShowConversationByDate = value

    @property
    def ShowConversationSendersAboveSubject(self):
        return self.com_object.ShowConversationSendersAboveSubject

    @ShowConversationSendersAboveSubject.setter
    def ShowConversationSendersAboveSubject(self, value):
        self.com_object.ShowConversationSendersAboveSubject = value

    @property
    def showconversationsendersabovesubject(self):
        """Lower case alias for ShowConversationSendersAboveSubject"""
        return self.ShowConversationSendersAboveSubject

    @showconversationsendersabovesubject.setter
    def showconversationsendersabovesubject(self, value):
        """Lower case alias for ShowConversationSendersAboveSubject.setter"""
        self.ShowConversationSendersAboveSubject = value

    @property
    def ShowFullConversations(self):
        return self.com_object.ShowFullConversations

    @ShowFullConversations.setter
    def ShowFullConversations(self, value):
        self.com_object.ShowFullConversations = value

    @property
    def showfullconversations(self):
        """Lower case alias for ShowFullConversations"""
        return self.ShowFullConversations

    @showfullconversations.setter
    def showfullconversations(self, value):
        """Lower case alias for ShowFullConversations.setter"""
        self.ShowFullConversations = value

    @property
    def ShowItemsInGroups(self):
        return TableView(self.com_object.ShowItemsInGroups)

    @ShowItemsInGroups.setter
    def ShowItemsInGroups(self, value):
        self.com_object.ShowItemsInGroups = value

    @property
    def showitemsingroups(self):
        """Lower case alias for ShowItemsInGroups"""
        return self.ShowItemsInGroups

    @showitemsingroups.setter
    def showitemsingroups(self, value):
        """Lower case alias for ShowItemsInGroups.setter"""
        self.ShowItemsInGroups = value

    @property
    def ShowNewItemRow(self):
        return TableView(self.com_object.ShowNewItemRow)

    @ShowNewItemRow.setter
    def ShowNewItemRow(self, value):
        self.com_object.ShowNewItemRow = value

    @property
    def shownewitemrow(self):
        """Lower case alias for ShowNewItemRow"""
        return self.ShowNewItemRow

    @shownewitemrow.setter
    def shownewitemrow(self, value):
        """Lower case alias for ShowNewItemRow.setter"""
        self.ShowNewItemRow = value

    @property
    def ShowReadingPane(self):
        return TableView(self.com_object.ShowReadingPane)

    @ShowReadingPane.setter
    def ShowReadingPane(self, value):
        self.com_object.ShowReadingPane = value

    @property
    def showreadingpane(self):
        """Lower case alias for ShowReadingPane"""
        return self.ShowReadingPane

    @showreadingpane.setter
    def showreadingpane(self, value):
        """Lower case alias for ShowReadingPane.setter"""
        self.ShowReadingPane = value

    @property
    def SortFields(self):
        return OrderFields(self.com_object.SortFields)

    @property
    def sortfields(self):
        """Lower case alias for SortFields"""
        return self.SortFields

    @property
    def Standard(self):
        return TableView(self.com_object.Standard)

    @property
    def standard(self):
        """Lower case alias for Standard"""
        return self.Standard

    @property
    def ViewFields(self):
        return ViewFields(self.com_object.ViewFields)

    @property
    def viewfields(self):
        """Lower case alias for ViewFields"""
        return self.ViewFields

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    @property
    def xml(self):
        """Lower case alias for XML"""
        return self.XML

    @xml.setter
    def xml(self, value):
        """Lower case alias for XML.setter"""
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SaveOption]])
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
        arguments = com_arguments([unwrap(a) for a in [Date]])
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


class tabs:

    def __init__(self, tabs=None):
        self.com_object= tabs


class tabstrip:

    def __init__(self, tabstrip=None):
        self.com_object= tabstrip


class TaskItem:

    def __init__(self, taskitem=None):
        self.com_object= taskitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def ActualWork(self):
        return self.com_object.ActualWork

    @ActualWork.setter
    def ActualWork(self, value):
        self.com_object.ActualWork = value

    @property
    def actualwork(self):
        """Lower case alias for ActualWork"""
        return self.ActualWork

    @actualwork.setter
    def actualwork(self, value):
        """Lower case alias for ActualWork.setter"""
        self.ActualWork = value

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def CardData(self):
        return self.com_object.CardData

    @CardData.setter
    def CardData(self, value):
        self.com_object.CardData = value

    @property
    def carddata(self):
        """Lower case alias for CardData"""
        return self.CardData

    @carddata.setter
    def carddata(self, value):
        """Lower case alias for CardData.setter"""
        self.CardData = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Complete(self):
        return self.com_object.Complete

    @Complete.setter
    def Complete(self, value):
        self.com_object.Complete = value

    @property
    def complete(self):
        """Lower case alias for Complete"""
        return self.Complete

    @complete.setter
    def complete(self, value):
        """Lower case alias for Complete.setter"""
        self.Complete = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ContactNames(self):
        return self.com_object.ContactNames

    @ContactNames.setter
    def ContactNames(self, value):
        self.com_object.ContactNames = value

    @property
    def contactnames(self):
        """Lower case alias for ContactNames"""
        return self.ContactNames

    @contactnames.setter
    def contactnames(self, value):
        """Lower case alias for ContactNames.setter"""
        self.ContactNames = value

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DateCompleted(self):
        return self.com_object.DateCompleted

    @DateCompleted.setter
    def DateCompleted(self, value):
        self.com_object.DateCompleted = value

    @property
    def datecompleted(self):
        """Lower case alias for DateCompleted"""
        return self.DateCompleted

    @datecompleted.setter
    def datecompleted(self, value):
        """Lower case alias for DateCompleted.setter"""
        self.DateCompleted = value

    @property
    def DelegationState(self):
        return OlTaskDelegationState(self.com_object.DelegationState)

    @property
    def delegationstate(self):
        """Lower case alias for DelegationState"""
        return self.DelegationState

    @property
    def Delegator(self):
        return self.com_object.Delegator

    @property
    def delegator(self):
        """Lower case alias for Delegator"""
        return self.Delegator

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def DueDate(self):
        return self.com_object.DueDate

    @DueDate.setter
    def DueDate(self, value):
        self.com_object.DueDate = value

    @property
    def duedate(self):
        """Lower case alias for DueDate"""
        return self.DueDate

    @duedate.setter
    def duedate(self, value):
        """Lower case alias for DueDate.setter"""
        self.DueDate = value

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def InternetCodepage(self):
        return self.com_object.InternetCodepage

    @InternetCodepage.setter
    def InternetCodepage(self, value):
        self.com_object.InternetCodepage = value

    @property
    def internetcodepage(self):
        """Lower case alias for InternetCodepage"""
        return self.InternetCodepage

    @internetcodepage.setter
    def internetcodepage(self, value):
        """Lower case alias for InternetCodepage.setter"""
        self.InternetCodepage = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def IsRecurring(self):
        return self.com_object.IsRecurring

    @property
    def isrecurring(self):
        """Lower case alias for IsRecurring"""
        return self.IsRecurring

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def Ordinal(self):
        return self.com_object.Ordinal

    @Ordinal.setter
    def Ordinal(self, value):
        self.com_object.Ordinal = value

    @property
    def ordinal(self):
        """Lower case alias for Ordinal"""
        return self.Ordinal

    @ordinal.setter
    def ordinal(self, value):
        """Lower case alias for Ordinal.setter"""
        self.Ordinal = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Owner(self):
        return self.com_object.Owner

    @Owner.setter
    def Owner(self, value):
        self.com_object.Owner = value

    @property
    def owner(self):
        """Lower case alias for Owner"""
        return self.Owner

    @owner.setter
    def owner(self, value):
        """Lower case alias for Owner.setter"""
        self.Owner = value

    @property
    def Ownership(self):
        return OlTaskOwnership(self.com_object.Ownership)

    @property
    def ownership(self):
        """Lower case alias for Ownership"""
        return self.Ownership

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PercentComplete(self):
        return self.com_object.PercentComplete

    @PercentComplete.setter
    def PercentComplete(self, value):
        self.com_object.PercentComplete = value

    @property
    def percentcomplete(self):
        """Lower case alias for PercentComplete"""
        return self.PercentComplete

    @percentcomplete.setter
    def percentcomplete(self, value):
        """Lower case alias for PercentComplete.setter"""
        self.PercentComplete = value

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @property
    def ReminderOverrideDefault(self):
        return self.com_object.ReminderOverrideDefault

    @ReminderOverrideDefault.setter
    def ReminderOverrideDefault(self, value):
        self.com_object.ReminderOverrideDefault = value

    @property
    def reminderoverridedefault(self):
        """Lower case alias for ReminderOverrideDefault"""
        return self.ReminderOverrideDefault

    @reminderoverridedefault.setter
    def reminderoverridedefault(self, value):
        """Lower case alias for ReminderOverrideDefault.setter"""
        self.ReminderOverrideDefault = value

    @property
    def ReminderPlaySound(self):
        return self.com_object.ReminderPlaySound

    @ReminderPlaySound.setter
    def ReminderPlaySound(self, value):
        self.com_object.ReminderPlaySound = value

    @property
    def reminderplaysound(self):
        """Lower case alias for ReminderPlaySound"""
        return self.ReminderPlaySound

    @reminderplaysound.setter
    def reminderplaysound(self, value):
        """Lower case alias for ReminderPlaySound.setter"""
        self.ReminderPlaySound = value

    @property
    def ReminderSet(self):
        return self.com_object.ReminderSet

    @ReminderSet.setter
    def ReminderSet(self, value):
        self.com_object.ReminderSet = value

    @property
    def reminderset(self):
        """Lower case alias for ReminderSet"""
        return self.ReminderSet

    @reminderset.setter
    def reminderset(self, value):
        """Lower case alias for ReminderSet.setter"""
        self.ReminderSet = value

    @property
    def ReminderSoundFile(self):
        return self.com_object.ReminderSoundFile

    @ReminderSoundFile.setter
    def ReminderSoundFile(self, value):
        self.com_object.ReminderSoundFile = value

    @property
    def remindersoundfile(self):
        """Lower case alias for ReminderSoundFile"""
        return self.ReminderSoundFile

    @remindersoundfile.setter
    def remindersoundfile(self, value):
        """Lower case alias for ReminderSoundFile.setter"""
        self.ReminderSoundFile = value

    @property
    def ReminderTime(self):
        return self.com_object.ReminderTime

    @ReminderTime.setter
    def ReminderTime(self, value):
        self.com_object.ReminderTime = value

    @property
    def remindertime(self):
        """Lower case alias for ReminderTime"""
        return self.ReminderTime

    @remindertime.setter
    def remindertime(self, value):
        """Lower case alias for ReminderTime.setter"""
        self.ReminderTime = value

    @property
    def ResponseState(self):
        return OlTaskResponse(self.com_object.ResponseState)

    @property
    def responsestate(self):
        """Lower case alias for ResponseState"""
        return self.ResponseState

    @property
    def Role(self):
        return self.com_object.Role

    @Role.setter
    def Role(self, value):
        self.com_object.Role = value

    @property
    def role(self):
        """Lower case alias for Role"""
        return self.Role

    @role.setter
    def role(self, value):
        """Lower case alias for Role.setter"""
        self.Role = value

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def SchedulePlusPriority(self):
        return self.com_object.SchedulePlusPriority

    @SchedulePlusPriority.setter
    def SchedulePlusPriority(self, value):
        self.com_object.SchedulePlusPriority = value

    @property
    def schedulepluspriority(self):
        """Lower case alias for SchedulePlusPriority"""
        return self.SchedulePlusPriority

    @schedulepluspriority.setter
    def schedulepluspriority(self, value):
        """Lower case alias for SchedulePlusPriority.setter"""
        self.SchedulePlusPriority = value

    @property
    def SendUsingAccount(self):
        return Account(self.com_object.SendUsingAccount)

    @SendUsingAccount.setter
    def SendUsingAccount(self, value):
        self.com_object.SendUsingAccount = value

    @property
    def sendusingaccount(self):
        """Lower case alias for SendUsingAccount"""
        return self.SendUsingAccount

    @sendusingaccount.setter
    def sendusingaccount(self, value):
        """Lower case alias for SendUsingAccount.setter"""
        self.SendUsingAccount = value

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def StartDate(self):
        return self.com_object.StartDate

    @StartDate.setter
    def StartDate(self, value):
        self.com_object.StartDate = value

    @property
    def startdate(self):
        """Lower case alias for StartDate"""
        return self.StartDate

    @startdate.setter
    def startdate(self, value):
        """Lower case alias for StartDate.setter"""
        self.StartDate = value

    @property
    def Status(self):
        return OlTaskStatus(self.com_object.Status)

    @Status.setter
    def Status(self, value):
        self.com_object.Status = value

    @property
    def status(self):
        """Lower case alias for Status"""
        return self.Status

    @status.setter
    def status(self, value):
        """Lower case alias for Status.setter"""
        self.Status = value

    @property
    def StatusOnCompletionRecipients(self):
        return self.com_object.StatusOnCompletionRecipients

    @StatusOnCompletionRecipients.setter
    def StatusOnCompletionRecipients(self, value):
        self.com_object.StatusOnCompletionRecipients = value

    @property
    def statusoncompletionrecipients(self):
        """Lower case alias for StatusOnCompletionRecipients"""
        return self.StatusOnCompletionRecipients

    @statusoncompletionrecipients.setter
    def statusoncompletionrecipients(self, value):
        """Lower case alias for StatusOnCompletionRecipients.setter"""
        self.StatusOnCompletionRecipients = value

    @property
    def StatusUpdateRecipients(self):
        return self.com_object.StatusUpdateRecipients

    @StatusUpdateRecipients.setter
    def StatusUpdateRecipients(self, value):
        self.com_object.StatusUpdateRecipients = value

    @property
    def statusupdaterecipients(self):
        """Lower case alias for StatusUpdateRecipients"""
        return self.StatusUpdateRecipients

    @statusupdaterecipients.setter
    def statusupdaterecipients(self, value):
        """Lower case alias for StatusUpdateRecipients.setter"""
        self.StatusUpdateRecipients = value

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def TeamTask(self):
        return self.com_object.TeamTask

    @TeamTask.setter
    def TeamTask(self, value):
        self.com_object.TeamTask = value

    @property
    def teamtask(self):
        """Lower case alias for TeamTask"""
        return self.TeamTask

    @teamtask.setter
    def teamtask(self, value):
        """Lower case alias for TeamTask.setter"""
        self.TeamTask = value

    @property
    def ToDoTaskOrdinal(self):
        return TaskItem(self.com_object.ToDoTaskOrdinal)

    @ToDoTaskOrdinal.setter
    def ToDoTaskOrdinal(self, value):
        self.com_object.ToDoTaskOrdinal = value

    @property
    def todotaskordinal(self):
        """Lower case alias for ToDoTaskOrdinal"""
        return self.ToDoTaskOrdinal

    @todotaskordinal.setter
    def todotaskordinal(self, value):
        """Lower case alias for ToDoTaskOrdinal.setter"""
        self.ToDoTaskOrdinal = value

    @property
    def TotalWork(self):
        return self.com_object.TotalWork

    @TotalWork.setter
    def TotalWork(self, value):
        self.com_object.TotalWork = value

    @property
    def totalwork(self):
        """Lower case alias for TotalWork"""
        return self.TotalWork

    @totalwork.setter
    def totalwork(self, value):
        """Lower case alias for TotalWork.setter"""
        self.TotalWork = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
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
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Response, fNoUI, fAdditionalTextDialog]])
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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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
        return self.com_object.StatusReport()

    # Lower case alias for StatusReport
    def statusreport(self):
        return self.StatusReport()


class TaskRequestAcceptItem:

    def __init__(self, taskrequestacceptitem=None):
        self.com_object= taskrequestacceptitem

    @property
    def Actions(self):
        return Actions(self.com_object.Actions)

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([unwrap(a) for a in [AddToTaskList]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([unwrap(a) for a in [AddToTaskList]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([unwrap(a) for a in [AddToTaskList]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def actions(self):
        """Lower case alias for Actions"""
        return self.Actions

    @property
    def Application(self):
        return Application(self.com_object.Application)

    @property
    def Attachments(self):
        return Attachments(self.com_object.Attachments)

    @property
    def attachments(self):
        """Lower case alias for Attachments"""
        return self.Attachments

    @property
    def AutoResolvedWinner(self):
        return self.com_object.AutoResolvedWinner

    @property
    def autoresolvedwinner(self):
        """Lower case alias for AutoResolvedWinner"""
        return self.AutoResolvedWinner

    @property
    def BillingInformation(self):
        return self.com_object.BillingInformation

    @BillingInformation.setter
    def BillingInformation(self, value):
        self.com_object.BillingInformation = value

    @property
    def billinginformation(self):
        """Lower case alias for BillingInformation"""
        return self.BillingInformation

    @billinginformation.setter
    def billinginformation(self, value):
        """Lower case alias for BillingInformation.setter"""
        self.BillingInformation = value

    @property
    def Body(self):
        return self.com_object.Body

    @Body.setter
    def Body(self, value):
        self.com_object.Body = value

    @property
    def body(self):
        """Lower case alias for Body"""
        return self.Body

    @body.setter
    def body(self, value):
        """Lower case alias for Body.setter"""
        self.Body = value

    @property
    def Categories(self):
        return self.com_object.Categories

    @Categories.setter
    def Categories(self, value):
        self.com_object.Categories = value

    @property
    def categories(self):
        """Lower case alias for Categories"""
        return self.Categories

    @categories.setter
    def categories(self, value):
        """Lower case alias for Categories.setter"""
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

    @property
    def companies(self):
        """Lower case alias for Companies"""
        return self.Companies

    @companies.setter
    def companies(self, value):
        """Lower case alias for Companies.setter"""
        self.Companies = value

    @property
    def Conflicts(self):
        return self.com_object.Conflicts

    @property
    def conflicts(self):
        """Lower case alias for Conflicts"""
        return self.Conflicts

    @property
    def ConversationID(self):
        return Conversation(self.com_object.ConversationID)

    @property
    def conversationid(self):
        """Lower case alias for ConversationID"""
        return self.ConversationID

    @property
    def ConversationIndex(self):
        return self.com_object.ConversationIndex

    @property
    def conversationindex(self):
        """Lower case alias for ConversationIndex"""
        return self.ConversationIndex

    @property
    def ConversationTopic(self):
        return self.com_object.ConversationTopic

    @property
    def conversationtopic(self):
        """Lower case alias for ConversationTopic"""
        return self.ConversationTopic

    @property
    def CreationTime(self):
        return self.com_object.CreationTime

    @property
    def creationtime(self):
        """Lower case alias for CreationTime"""
        return self.CreationTime

    @property
    def DownloadState(self):
        return OlDownloadState(self.com_object.DownloadState)

    @property
    def downloadstate(self):
        """Lower case alias for DownloadState"""
        return self.DownloadState

    @property
    def EntryID(self):
        return self.com_object.EntryID

    @property
    def entryid(self):
        """Lower case alias for EntryID"""
        return self.EntryID

    @property
    def FormDescription(self):
        return FormDescription(self.com_object.FormDescription)

    @property
    def formdescription(self):
        """Lower case alias for FormDescription"""
        return self.FormDescription

    @property
    def GetInspector(self):
        return Inspector(self.com_object.GetInspector)

    @property
    def getinspector(self):
        """Lower case alias for GetInspector"""
        return self.GetInspector

    @property
    def Importance(self):
        return OlImportance(self.com_object.Importance)

    @Importance.setter
    def Importance(self, value):
        self.com_object.Importance = value

    @property
    def importance(self):
        """Lower case alias for Importance"""
        return self.Importance

    @importance.setter
    def importance(self, value):
        """Lower case alias for Importance.setter"""
        self.Importance = value

    @property
    def IsConflict(self):
        return self.com_object.IsConflict

    @property
    def isconflict(self):
        """Lower case alias for IsConflict"""
        return self.IsConflict

    @property
    def ItemProperties(self):
        return ItemProperties(self.com_object.ItemProperties)

    @property
    def itemproperties(self):
        """Lower case alias for ItemProperties"""
        return self.ItemProperties

    @property
    def LastModificationTime(self):
        return self.com_object.LastModificationTime

    @property
    def lastmodificationtime(self):
        """Lower case alias for LastModificationTime"""
        return self.LastModificationTime

    @property
    def MarkForDownload(self):
        return OlRemoteStatus(self.com_object.MarkForDownload)

    @MarkForDownload.setter
    def MarkForDownload(self, value):
        self.com_object.MarkForDownload = value

    @property
    def markfordownload(self):
        """Lower case alias for MarkForDownload"""
        return self.MarkForDownload

    @markfordownload.setter
    def markfordownload(self, value):
        """Lower case alias for MarkForDownload.setter"""
        self.MarkForDownload = value

    @property
    def MessageClass(self):
        return self.com_object.MessageClass

    @MessageClass.setter
    def MessageClass(self, value):
        self.com_object.MessageClass = value

    @property
    def messageclass(self):
        """Lower case alias for MessageClass"""
        return self.MessageClass

    @messageclass.setter
    def messageclass(self, value):
        """Lower case alias for MessageClass.setter"""
        self.MessageClass = value

    @property
    def Mileage(self):
        return self.com_object.Mileage

    @Mileage.setter
    def Mileage(self, value):
        self.com_object.Mileage = value

    @property
    def mileage(self):
        """Lower case alias for Mileage"""
        return self.Mileage

    @mileage.setter
    def mileage(self, value):
        """Lower case alias for Mileage.setter"""
        self.Mileage = value

    @property
    def NoAging(self):
        return self.com_object.NoAging

    @NoAging.setter
    def NoAging(self, value):
        self.com_object.NoAging = value

    @property
    def noaging(self):
        """Lower case alias for NoAging"""
        return self.NoAging

    @noaging.setter
    def noaging(self, value):
        """Lower case alias for NoAging.setter"""
        self.NoAging = value

    @property
    def OutlookInternalVersion(self):
        return self.com_object.OutlookInternalVersion

    @property
    def outlookinternalversion(self):
        """Lower case alias for OutlookInternalVersion"""
        return self.OutlookInternalVersion

    @property
    def OutlookVersion(self):
        return self.com_object.OutlookVersion

    @property
    def outlookversion(self):
        """Lower case alias for OutlookVersion"""
        return self.OutlookVersion

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def PropertyAccessor(self):
        return PropertyAccessor(self.com_object.PropertyAccessor)

    @property
    def propertyaccessor(self):
        """Lower case alias for PropertyAccessor"""
        return self.PropertyAccessor

    @property
    def RTFBody(self):
        return self.com_object.RTFBody

    @RTFBody.setter
    def RTFBody(self, value):
        self.com_object.RTFBody = value

    @property
    def rtfbody(self):
        """Lower case alias for RTFBody"""
        return self.RTFBody

    @rtfbody.setter
    def rtfbody(self, value):
        """Lower case alias for RTFBody.setter"""
        self.RTFBody = value

    @property
    def Saved(self):
        return self.com_object.Saved

    @property
    def saved(self):
        """Lower case alias for Saved"""
        return self.Saved

    @property
    def Sensitivity(self):
        return OlSensitivity(self.com_object.Sensitivity)

    @Sensitivity.setter
    def Sensitivity(self, value):
        self.com_object.Sensitivity = value

    @property
    def sensitivity(self):
        """Lower case alias for Sensitivity"""
        return self.Sensitivity

    @sensitivity.setter
    def sensitivity(self, value):
        """Lower case alias for Sensitivity.setter"""
        self.Sensitivity = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @property
    def Subject(self):
        return self.com_object.Subject

    @Subject.setter
    def Subject(self, value):
        self.com_object.Subject = value

    @property
    def subject(self):
        """Lower case alias for Subject"""
        return self.Subject

    @subject.setter
    def subject(self, value):
        """Lower case alias for Subject.setter"""
        self.Subject = value

    @property
    def UnRead(self):
        return self.com_object.UnRead

    @UnRead.setter
    def UnRead(self, value):
        self.com_object.UnRead = value

    @property
    def unread(self):
        """Lower case alias for UnRead"""
        return self.UnRead

    @unread.setter
    def unread(self, value):
        """Lower case alias for UnRead.setter"""
        self.UnRead = value

    @property
    def UserProperties(self):
        return UserProperties(self.com_object.UserProperties)

    @property
    def userproperties(self):
        """Lower case alias for UserProperties"""
        return self.UserProperties

    def Close(self, SaveMode=None):
        arguments = com_arguments([unwrap(a) for a in [SaveMode]])
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
        arguments = com_arguments([unwrap(a) for a in [Modal]])
        self.com_object.Display(*arguments)

    # Lower case alias for Display
    def display(self, Modal=None):
        arguments = [Modal]
        return self.Display(*arguments)

    def GetAssociatedTask(self, AddToTaskList=None):
        arguments = com_arguments([unwrap(a) for a in [AddToTaskList]])
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
        arguments = com_arguments([unwrap(a) for a in [DestFldr]])
        return self.com_object.Move(*arguments)

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
        arguments = com_arguments([unwrap(a) for a in [Path, Type]])
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

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def NavigationGroups(self):
        return NavigationGroups(self.com_object.NavigationGroups)

    @property
    def navigationgroups(self):
        """Lower case alias for NavigationGroups"""
        return self.NavigationGroups

    @property
    def NavigationModuleType(self):
        return OlNavigationModuleType(self.com_object.NavigationModuleType)

    @property
    def navigationmoduletype(self):
        """Lower case alias for NavigationModuleType"""
        return self.NavigationModuleType

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Position(self):
        return TasksModule(self.com_object.Position)

    @Position.setter
    def Position(self, value):
        self.com_object.Position = value

    @property
    def position(self):
        """Lower case alias for Position"""
        return self.Position

    @position.setter
    def position(self, value):
        """Lower case alias for Position.setter"""
        self.Position = value

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Visible(self):
        return TasksModule(self.com_object.Visible)

    @Visible.setter
    def Visible(self, value):
        self.com_object.Visible = value

    @property
    def visible(self):
        """Lower case alias for Visible"""
        return self.Visible

    @visible.setter
    def visible(self, value):
        """Lower case alias for Visible.setter"""
        self.Visible = value


class textbox:

    def __init__(self, textbox=None):
        self.com_object= textbox


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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Text(self):
        return self.com_object.Text

    @Text.setter
    def Text(self, value):
        self.com_object.Text = value

    @property
    def text(self):
        """Lower case alias for Text"""
        return self.Text

    @text.setter
    def text(self, value):
        """Lower case alias for Text.setter"""
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

    @property
    def defaultexpandcollapsesetting(self):
        """Lower case alias for DefaultExpandCollapseSetting"""
        return self.DefaultExpandCollapseSetting

    @defaultexpandcollapsesetting.setter
    def defaultexpandcollapsesetting(self, value):
        """Lower case alias for DefaultExpandCollapseSetting.setter"""
        self.DefaultExpandCollapseSetting = value

    @property
    def EndField(self):
        return TimelineView(self.com_object.EndField)

    @EndField.setter
    def EndField(self, value):
        self.com_object.EndField = value

    @property
    def endfield(self):
        """Lower case alias for EndField"""
        return self.EndField

    @endfield.setter
    def endfield(self, value):
        """Lower case alias for EndField.setter"""
        self.EndField = value

    @property
    def Filter(self):
        return self.com_object.Filter

    @Filter.setter
    def Filter(self, value):
        self.com_object.Filter = value

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for Filter.setter"""
        self.Filter = value

    @property
    def GroupByFields(self):
        return OrderFields(self.com_object.GroupByFields)

    @property
    def groupbyfields(self):
        """Lower case alias for GroupByFields"""
        return self.GroupByFields

    @property
    def ItemFont(self):
        return ViewFont(self.com_object.ItemFont)

    @property
    def itemfont(self):
        """Lower case alias for ItemFont"""
        return self.ItemFont

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @language.setter
    def language(self, value):
        """Lower case alias for Language.setter"""
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    @property
    def lockuserchanges(self):
        """Lower case alias for LockUserChanges"""
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        """Lower case alias for LockUserChanges.setter"""
        self.LockUserChanges = value

    @property
    def LowerScaleFont(self):
        return ViewFont(self.com_object.LowerScaleFont)

    @property
    def lowerscalefont(self):
        """Lower case alias for LowerScaleFont"""
        return self.LowerScaleFont

    @property
    def MaxLabelWidth(self):
        return TimelineView(self.com_object.MaxLabelWidth)

    @MaxLabelWidth.setter
    def MaxLabelWidth(self, value):
        self.com_object.MaxLabelWidth = value

    @property
    def maxlabelwidth(self):
        """Lower case alias for MaxLabelWidth"""
        return self.MaxLabelWidth

    @maxlabelwidth.setter
    def maxlabelwidth(self, value):
        """Lower case alias for MaxLabelWidth.setter"""
        self.MaxLabelWidth = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    @property
    def saveoption(self):
        """Lower case alias for SaveOption"""
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ShowLabelWhenViewingByMonth(self):
        return TimelineView(self.com_object.ShowLabelWhenViewingByMonth)

    @ShowLabelWhenViewingByMonth.setter
    def ShowLabelWhenViewingByMonth(self, value):
        self.com_object.ShowLabelWhenViewingByMonth = value

    @property
    def showlabelwhenviewingbymonth(self):
        """Lower case alias for ShowLabelWhenViewingByMonth"""
        return self.ShowLabelWhenViewingByMonth

    @showlabelwhenviewingbymonth.setter
    def showlabelwhenviewingbymonth(self, value):
        """Lower case alias for ShowLabelWhenViewingByMonth.setter"""
        self.ShowLabelWhenViewingByMonth = value

    @property
    def ShowWeekNumbers(self):
        return TimelineView(self.com_object.ShowWeekNumbers)

    @ShowWeekNumbers.setter
    def ShowWeekNumbers(self, value):
        self.com_object.ShowWeekNumbers = value

    @property
    def showweeknumbers(self):
        """Lower case alias for ShowWeekNumbers"""
        return self.ShowWeekNumbers

    @showweeknumbers.setter
    def showweeknumbers(self, value):
        """Lower case alias for ShowWeekNumbers.setter"""
        self.ShowWeekNumbers = value

    @property
    def Standard(self):
        return TimelineView(self.com_object.Standard)

    @property
    def standard(self):
        """Lower case alias for Standard"""
        return self.Standard

    @property
    def StartField(self):
        return TimelineView(self.com_object.StartField)

    @StartField.setter
    def StartField(self, value):
        self.com_object.StartField = value

    @property
    def startfield(self):
        """Lower case alias for StartField"""
        return self.StartField

    @startfield.setter
    def startfield(self, value):
        """Lower case alias for StartField.setter"""
        self.StartField = value

    @property
    def TimelineViewMode(self):
        return OlTimelineViewMode(self.com_object.TimelineViewMode)

    @TimelineViewMode.setter
    def TimelineViewMode(self, value):
        self.com_object.TimelineViewMode = value

    @property
    def timelineviewmode(self):
        """Lower case alias for TimelineViewMode"""
        return self.TimelineViewMode

    @timelineviewmode.setter
    def timelineviewmode(self, value):
        """Lower case alias for TimelineViewMode.setter"""
        self.TimelineViewMode = value

    @property
    def UpperScaleFont(self):
        return ViewFont(self.com_object.UpperScaleFont)

    @property
    def upperscalefont(self):
        """Lower case alias for UpperScaleFont"""
        return self.UpperScaleFont

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    @property
    def xml(self):
        """Lower case alias for XML"""
        return self.XML

    @xml.setter
    def xml(self, value):
        """Lower case alias for XML.setter"""
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SaveOption]])
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
        arguments = com_arguments([unwrap(a) for a in [Date]])
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

    @property
    def bias(self):
        """Lower case alias for Bias"""
        return self.Bias

    @property
    def Class(self):
        return OlObjectClass(self.com_object.Class)

    @property
    def DaylightBias(self):
        return self.com_object.DaylightBias

    @property
    def daylightbias(self):
        """Lower case alias for DaylightBias"""
        return self.DaylightBias

    @property
    def DaylightDate(self):
        return self.com_object.DaylightDate

    @property
    def daylightdate(self):
        """Lower case alias for DaylightDate"""
        return self.DaylightDate

    @property
    def DaylightDesignation(self):
        return self.com_object.DaylightDesignation

    @property
    def daylightdesignation(self):
        """Lower case alias for DaylightDesignation"""
        return self.DaylightDesignation

    @property
    def ID(self):
        return self.com_object.ID

    @property
    def id(self):
        """Lower case alias for ID"""
        return self.ID

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def StandardBias(self):
        return self.com_object.StandardBias

    @property
    def standardbias(self):
        """Lower case alias for StandardBias"""
        return self.StandardBias

    @property
    def StandardDate(self):
        return self.com_object.StandardDate

    @property
    def standarddate(self):
        """Lower case alias for StandardDate"""
        return self.StandardDate

    @property
    def StandardDesignation(self):
        return self.com_object.StandardDesignation

    @property
    def standarddesignation(self):
        """Lower case alias for StandardDesignation"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def CurrentTimeZone(self):
        return TimeZone(self.com_object.CurrentTimeZone)

    @property
    def currenttimezone(self):
        """Lower case alias for CurrentTimeZone"""
        return self.CurrentTimeZone

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def ConvertTime(self, SourceDateTime=None, SourceTimeZone=None, DestinationTimeZone=None):
        arguments = com_arguments([unwrap(a) for a in [SourceDateTime, SourceTimeZone, DestinationTimeZone]])
        return self.com_object.ConvertTime(*arguments)

    # Lower case alias for ConvertTime
    def converttime(self, SourceDateTime=None, SourceTimeZone=None, DestinationTimeZone=None):
        arguments = [SourceDateTime, SourceTimeZone, DestinationTimeZone]
        return self.ConvertTime(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return TimeZone(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)


class togglebutton:

    def __init__(self, togglebutton=None):
        self.com_object= togglebutton


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

    @property
    def conditiontype(self):
        """Lower case alias for ConditionType"""
        return self.ConditionType

    @property
    def Enabled(self):
        return self.com_object.Enabled

    @Enabled.setter
    def Enabled(self, value):
        self.com_object.Enabled = value

    @property
    def enabled(self):
        """Lower case alias for Enabled"""
        return self.Enabled

    @enabled.setter
    def enabled(self, value):
        """Lower case alias for Enabled.setter"""
        self.Enabled = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Recipients(self):
        return Recipients(self.com_object.Recipients)

    @property
    def recipients(self):
        """Lower case alias for Recipients"""
        return self.Recipients

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None, Type=None, DisplayFormat=None, Formula=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Type, DisplayFormat, Formula]])
        return UserDefinedProperty(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Type=None, DisplayFormat=None, Formula=None):
        arguments = [Name, Type, DisplayFormat, Formula]
        return self.Add(*arguments)

    def Find(self, Name=None):
        arguments = com_arguments([unwrap(a) for a in [Name]])
        return self.com_object.Find(*arguments)

    # Lower case alias for Find
    def find(self, Name=None):
        arguments = [Name]
        return self.Find(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def displayformat(self):
        """Lower case alias for DisplayFormat"""
        return self.DisplayFormat

    @property
    def Formula(self):
        return self.com_object.Formula

    @property
    def formula(self):
        """Lower case alias for Formula"""
        return self.Formula

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.com_object.Type)

    @property
    def type(self):
        """Lower case alias for Type"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Type, AddToFolderFields, DisplayFormat]])
        return UserProperty(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, Type=None, AddToFolderFields=None, DisplayFormat=None):
        arguments = [Name, Type, AddToFolderFields, DisplayFormat]
        return self.Add(*arguments)

    def Find(self, Name=None, Custom=None):
        arguments = com_arguments([unwrap(a) for a in [Name, Custom]])
        return self.com_object.Find(*arguments)

    # Lower case alias for Find
    def find(self, Name=None, Custom=None):
        arguments = [Name, Custom]
        return self.Find(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return UserProperty(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def formula(self):
        """Lower case alias for Formula"""
        return self.Formula

    @formula.setter
    def formula(self, value):
        """Lower case alias for Formula.setter"""
        self.Formula = value

    @property
    def Name(self):
        return self.com_object.Name

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Type(self):
        return OlUserPropertyType(self.com_object.Type)

    @property
    def type(self):
        """Lower case alias for Type"""
        return self.Type

    @property
    def ValidationFormula(self):
        return self.com_object.ValidationFormula

    @ValidationFormula.setter
    def ValidationFormula(self, value):
        self.com_object.ValidationFormula = value

    @property
    def validationformula(self):
        """Lower case alias for ValidationFormula"""
        return self.ValidationFormula

    @validationformula.setter
    def validationformula(self, value):
        """Lower case alias for ValidationFormula.setter"""
        self.ValidationFormula = value

    @property
    def ValidationText(self):
        return self.com_object.ValidationText

    @ValidationText.setter
    def ValidationText(self, value):
        self.com_object.ValidationText = value

    @property
    def validationtext(self):
        """Lower case alias for ValidationText"""
        return self.ValidationText

    @validationtext.setter
    def validationtext(self, value):
        """Lower case alias for ValidationText.setter"""
        self.ValidationText = value

    @property
    def Value(self):
        return self.com_object.Value

    @Value.setter
    def Value(self, value):
        self.com_object.Value = value

    @property
    def value(self):
        """Lower case alias for Value"""
        return self.Value

    @value.setter
    def value(self, value):
        """Lower case alias for Value.setter"""
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

    @property
    def filter(self):
        """Lower case alias for Filter"""
        return self.Filter

    @filter.setter
    def filter(self, value):
        """Lower case alias for Filter.setter"""
        self.Filter = value

    @property
    def Language(self):
        return self.com_object.Language

    @Language.setter
    def Language(self, value):
        self.com_object.Language = value

    @property
    def language(self):
        """Lower case alias for Language"""
        return self.Language

    @language.setter
    def language(self, value):
        """Lower case alias for Language.setter"""
        self.Language = value

    @property
    def LockUserChanges(self):
        return self.com_object.LockUserChanges

    @LockUserChanges.setter
    def LockUserChanges(self, value):
        self.com_object.LockUserChanges = value

    @property
    def lockuserchanges(self):
        """Lower case alias for LockUserChanges"""
        return self.LockUserChanges

    @lockuserchanges.setter
    def lockuserchanges(self, value):
        """Lower case alias for LockUserChanges.setter"""
        self.LockUserChanges = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def SaveOption(self):
        return OlViewSaveOption(self.com_object.SaveOption)

    @property
    def saveoption(self):
        """Lower case alias for SaveOption"""
        return self.SaveOption

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Standard(self):
        return self.com_object.Standard

    @property
    def standard(self):
        """Lower case alias for Standard"""
        return self.Standard

    @property
    def ViewType(self):
        return OlViewType(self.com_object.ViewType)

    @property
    def viewtype(self):
        """Lower case alias for ViewType"""
        return self.ViewType

    @property
    def XML(self):
        return self.com_object.XML

    @XML.setter
    def XML(self, value):
        self.com_object.XML = value

    @property
    def xml(self):
        """Lower case alias for XML"""
        return self.XML

    @xml.setter
    def xml(self, value):
        """Lower case alias for XML.setter"""
        self.XML = value

    def Apply(self):
        self.com_object.Apply()

    # Lower case alias for Apply
    def apply(self):
        return self.Apply()

    def Copy(self, Name=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, SaveOption]])
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
        arguments = com_arguments([unwrap(a) for a in [Date]])
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

    @property
    def columnformat(self):
        """Lower case alias for ColumnFormat"""
        return self.ColumnFormat

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def ViewXMLSchemaName(self):
        return ViewField(self.com_object.ViewXMLSchemaName)

    @property
    def viewxmlschemaname(self):
        """Lower case alias for ViewXMLSchemaName"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, PropertyName=None):
        arguments = com_arguments([unwrap(a) for a in [PropertyName]])
        return ViewField(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, PropertyName=None):
        arguments = [PropertyName]
        return self.Add(*arguments)

    def Insert(self, PropertyName=None, Index=None):
        arguments = com_arguments([unwrap(a) for a in [PropertyName, Index]])
        return ViewField(self.com_object.Insert(*arguments))

    # Lower case alias for Insert
    def insert(self, PropertyName=None, Index=None):
        arguments = [PropertyName, Index]
        return self.Insert(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return ViewField(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
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

    @property
    def bold(self):
        """Lower case alias for Bold"""
        return self.Bold

    @bold.setter
    def bold(self, value):
        """Lower case alias for Bold.setter"""
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

    @property
    def color(self):
        """Lower case alias for Color"""
        return self.Color

    @color.setter
    def color(self, value):
        """Lower case alias for Color.setter"""
        self.Color = value

    @property
    def ExtendedColor(self):
        return OlCategoryColor(self.com_object.ExtendedColor)

    @ExtendedColor.setter
    def ExtendedColor(self, value):
        self.com_object.ExtendedColor = value

    @property
    def extendedcolor(self):
        """Lower case alias for ExtendedColor"""
        return self.ExtendedColor

    @extendedcolor.setter
    def extendedcolor(self, value):
        """Lower case alias for ExtendedColor.setter"""
        self.ExtendedColor = value

    @property
    def Italic(self):
        return ViewFont(self.com_object.Italic)

    @Italic.setter
    def Italic(self, value):
        self.com_object.Italic = value

    @property
    def italic(self):
        """Lower case alias for Italic"""
        return self.Italic

    @italic.setter
    def italic(self, value):
        """Lower case alias for Italic.setter"""
        self.Italic = value

    @property
    def Name(self):
        return self.com_object.Name

    @Name.setter
    def Name(self, value):
        self.com_object.Name = value

    @property
    def name(self):
        """Lower case alias for Name"""
        return self.Name

    @name.setter
    def name(self, value):
        """Lower case alias for Name.setter"""
        self.Name = value

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    @property
    def Size(self):
        return self.com_object.Size

    @Size.setter
    def Size(self, value):
        self.com_object.Size = value

    @property
    def size(self):
        """Lower case alias for Size"""
        return self.Size

    @size.setter
    def size(self, value):
        """Lower case alias for Size.setter"""
        self.Size = value

    @property
    def Strikethrough(self):
        return ViewFont(self.com_object.Strikethrough)

    @Strikethrough.setter
    def Strikethrough(self, value):
        self.com_object.Strikethrough = value

    @property
    def strikethrough(self):
        """Lower case alias for Strikethrough"""
        return self.Strikethrough

    @strikethrough.setter
    def strikethrough(self, value):
        """Lower case alias for Strikethrough.setter"""
        self.Strikethrough = value

    @property
    def Underline(self):
        return ViewFont(self.com_object.Underline)

    @Underline.setter
    def Underline(self, value):
        self.com_object.Underline = value

    @property
    def underline(self):
        """Lower case alias for Underline"""
        return self.Underline

    @underline.setter
    def underline(self, value):
        """Lower case alias for Underline.setter"""
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

    @property
    def count(self):
        """Lower case alias for Count"""
        return self.Count

    @property
    def Parent(self):
        return self.com_object.Parent

    @property
    def parent(self):
        """Lower case alias for Parent"""
        return self.Parent

    @property
    def Session(self):
        return NameSpace(self.com_object.Session)

    @property
    def session(self):
        """Lower case alias for Session"""
        return self.Session

    def Add(self, Name=None, ViewType=None, SaveOption=None):
        arguments = com_arguments([unwrap(a) for a in [Name, ViewType, SaveOption]])
        return View(self.com_object.Add(*arguments))

    # Lower case alias for Add
    def add(self, Name=None, ViewType=None, SaveOption=None):
        arguments = [Name, ViewType, SaveOption]
        return self.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        return View(self.com_object.Item(*arguments))

    # Lower case alias for Item
    def item(self, Index=None):
        arguments = [Index]
        return self.Item(*arguments)

    def Remove(self, Index=None):
        arguments = com_arguments([unwrap(a) for a in [Index]])
        self.com_object.Remove(*arguments)

    # Lower case alias for Remove
    def remove(self, Index=None):
        arguments = [Index]
        return self.Remove(*arguments)

