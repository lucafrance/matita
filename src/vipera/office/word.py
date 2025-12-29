import win32com.client

class AddIn:

    def __init__(self, addin=None):
        self.addin = addin

    @property
    def Application(self):
        return Application(self.addin.Application)

    @property
    def Autoload(self):
        return self.addin.Autoload

    @property
    def Compiled(self):
        return self.addin.Compiled

    @property
    def Creator(self):
        return self.addin.Creator

    @property
    def Index(self):
        return self.addin.Index

    @property
    def Installed(self):
        return self.addin.Installed

    @property
    def Name(self):
        return self.addin.Name

    @property
    def Parent(self):
        return self.addin.Parent

    @property
    def Path(self):
        return self.addin.Path

    def Delete(self):
        self.addin.Delete()

class Adjustments:

    def __init__(self, adjustments=None):
        self.adjustments = adjustments

    @property
    def Application(self):
        return Application(self.adjustments.Application)

    @property
    def Count(self):
        return Adjustments(self.adjustments.Count)

    @property
    def Creator(self):
        return self.adjustments.Creator

    @property
    def Item(self):
        return self.adjustments.Item

    @Item.setter
    def Item(self, value):
        self.adjustments.Item = value

    @property
    def Parent(self):
        return self.adjustments.Parent

class Application:

    def __init__(self, application=None):
        self.application = application

    def new(self):
        self.application = win32com.client.Dispatch("Word.Application")
        return self

    @property
    def ActiveDocument(self):
        return Document(self.application.ActiveDocument)

    @property
    def ActiveEncryptionSession(self):
        return self.application.ActiveEncryptionSession

    @property
    def ActivePrinter(self):
        return self.application.ActivePrinter

    @ActivePrinter.setter
    def ActivePrinter(self, value):
        self.application.ActivePrinter = value

    @property
    def ActiveProtectedViewWindow(self):
        return self.application.ActiveProtectedViewWindow

    @property
    def ActiveWindow(self):
        return Window(self.application.ActiveWindow)

    @property
    def AddIns(self):
        return self.application.AddIns

    @property
    def Application(self):
        return Application(self.application.Application)

    @property
    def ArbitraryXMLSupportAvailable(self):
        return self.application.ArbitraryXMLSupportAvailable

    @property
    def Assistance(self):
        return self.application.Assistance

    @property
    def AutoCaptions(self):
        return self.application.AutoCaptions

    @property
    def AutoCorrect(self):
        return AutoCorrect(self.application.AutoCorrect)

    @property
    def AutoCorrectEmail(self):
        return AutoCorrect(self.application.AutoCorrectEmail)

    @property
    def AutomationSecurity(self):
        return self.application.AutomationSecurity

    @AutomationSecurity.setter
    def AutomationSecurity(self, value):
        self.application.AutomationSecurity = value

    @property
    def BackgroundPrintingStatus(self):
        return self.application.BackgroundPrintingStatus

    @property
    def BackgroundSavingStatus(self):
        return self.application.BackgroundSavingStatus

    @property
    def Bibliography(self):
        return Bibliography(self.application.Bibliography)

    @property
    def BrowseExtraFileTypes(self):
        return self.application.BrowseExtraFileTypes

    @property
    def Browser(self):
        return Browser(self.application.Browser)

    @property
    def Build(self):
        return self.application.Build

    @property
    def CapsLock(self):
        return self.application.CapsLock

    @property
    def Caption(self):
        return self.application.Caption

    @Caption.setter
    def Caption(self, value):
        self.application.Caption = value

    @property
    def CaptionLabels(self):
        return self.application.CaptionLabels

    @property
    def CheckLanguage(self):
        return self.application.CheckLanguage

    @property
    def COMAddIns(self):
        return self.application.COMAddIns

    @property
    def CommandBars(self):
        return self.application.CommandBars

    @property
    def Creator(self):
        return self.application.Creator

    @property
    def CustomDictionaries(self):
        return self.application.CustomDictionaries

    @property
    def CustomizationContext(self):
        return Template(self.application.CustomizationContext)

    @CustomizationContext.setter
    def CustomizationContext(self, value):
        self.application.CustomizationContext = value

    @property
    def DefaultLegalBlackline(self):
        return self.application.DefaultLegalBlackline

    @property
    def DefaultSaveFormat(self):
        return self.application.DefaultSaveFormat

    @DefaultSaveFormat.setter
    def DefaultSaveFormat(self, value):
        self.application.DefaultSaveFormat = value

    @property
    def DefaultTableSeparator(self):
        return self.application.DefaultTableSeparator

    @DefaultTableSeparator.setter
    def DefaultTableSeparator(self, value):
        self.application.DefaultTableSeparator = value

    @property
    def Dialogs(self):
        return self.application.Dialogs

    @property
    def DisplayAlerts(self):
        return WdAlertLevel(self.application.DisplayAlerts)

    @DisplayAlerts.setter
    def DisplayAlerts(self, value):
        self.application.DisplayAlerts = value

    @property
    def DisplayAutoCompleteTips(self):
        return self.application.DisplayAutoCompleteTips

    @property
    def DisplayDocumentInformationPanel(self):
        return self.application.DisplayDocumentInformationPanel

    @DisplayDocumentInformationPanel.setter
    def DisplayDocumentInformationPanel(self, value):
        self.application.DisplayDocumentInformationPanel = value

    @property
    def DisplayRecentFiles(self):
        return self.application.DisplayRecentFiles

    @property
    def DisplayScreenTips(self):
        return self.application.DisplayScreenTips

    @property
    def DisplayScrollBars(self):
        return self.application.DisplayScrollBars

    @property
    def Documents(self):
        return self.application.Documents

    @property
    def DontResetInsertionPointProperties(self):
        return self.application.DontResetInsertionPointProperties

    @DontResetInsertionPointProperties.setter
    def DontResetInsertionPointProperties(self, value):
        self.application.DontResetInsertionPointProperties = value

    @property
    def EmailOptions(self):
        return EmailOptions(self.application.EmailOptions)

    @property
    def EmailTemplate(self):
        return self.application.EmailTemplate

    @EmailTemplate.setter
    def EmailTemplate(self, value):
        self.application.EmailTemplate = value

    @property
    def EnableCancelKey(self):
        return WdEnableCancelKey(self.application.EnableCancelKey)

    @EnableCancelKey.setter
    def EnableCancelKey(self, value):
        self.application.EnableCancelKey = value

    @property
    def FeatureInstall(self):
        return self.application.FeatureInstall

    @FeatureInstall.setter
    def FeatureInstall(self, value):
        self.application.FeatureInstall = value

    @property
    def FileConverters(self):
        return self.application.FileConverters

    def FileDialog(self, *args, FileDialogType=None):
        arguments = {"FileDialogType": FileDialogType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.FileDialog(*args, **arguments)

    @property
    def FileValidation(self):
        return self.application.FileValidation

    @FileValidation.setter
    def FileValidation(self, value):
        self.application.FileValidation = value

    def FindKey(self, *args, KeyCode=None, KeyCode2=None):
        arguments = {"KeyCode": KeyCode, "KeyCode2": KeyCode2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return KeyBinding(self.application.FindKey(*args, **arguments))

    @property
    def FocusInMailHeader(self):
        return self.application.FocusInMailHeader

    @property
    def FontNames(self):
        return FontNames(self.application.FontNames)

    @property
    def HangulHanjaDictionaries(self):
        return self.application.HangulHanjaDictionaries

    @property
    def Height(self):
        return self.application.Height

    @Height.setter
    def Height(self, value):
        self.application.Height = value

    def International(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.International(*args, **arguments)

    def IsObjectValid(self, *args, Object=None):
        arguments = {"Object": Object}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.IsObjectValid(*args, **arguments)

    @property
    def IsSandboxed(self):
        return self.application.IsSandboxed

    @property
    def KeyBindings(self):
        return self.application.KeyBindings

    def KeysBoundTo(self, *args, KeyCategory=None, Command=None, CommandParameter=None):
        arguments = {"KeyCategory": KeyCategory, "Command": Command, "CommandParameter": CommandParameter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.KeysBoundTo(*args, **arguments)

    @property
    def LandscapeFontNames(self):
        return FontNames(self.application.LandscapeFontNames)

    @property
    def Language(self):
        return self.application.Language

    @property
    def Languages(self):
        return self.application.Languages

    @property
    def LanguageSettings(self):
        return self.application.LanguageSettings

    @property
    def Left(self):
        return self.application.Left

    @Left.setter
    def Left(self, value):
        self.application.Left = value

    @property
    def ListGalleries(self):
        return self.application.ListGalleries

    @property
    def MacroContainer(self):
        return Template(self.application.MacroContainer)

    @property
    def MailingLabel(self):
        return MailingLabel(self.application.MailingLabel)

    @property
    def MailMessage(self):
        return MailMessage(self.application.MailMessage)

    @property
    def MailSystem(self):
        return WdMailSystem(self.application.MailSystem)

    @property
    def MAPIAvailable(self):
        return self.application.MAPIAvailable

    @property
    def MathCoprocessorAvailable(self):
        return self.application.MathCoprocessorAvailable

    @property
    def MouseAvailable(self):
        return self.application.MouseAvailable

    @property
    def Name(self):
        return self.application.Name

    @property
    def NewDocument(self):
        return self.application.NewDocument

    @property
    def NormalTemplate(self):
        return Template(self.application.NormalTemplate)

    @property
    def NumLock(self):
        return self.application.NumLock

    @property
    def OMathAutoCorrect(self):
        return OMathAutoCorrect(self.application.OMathAutoCorrect)

    @property
    def OpenAttachmentsInFullScreen(self):
        return self.application.OpenAttachmentsInFullScreen

    @OpenAttachmentsInFullScreen.setter
    def OpenAttachmentsInFullScreen(self, value):
        self.application.OpenAttachmentsInFullScreen = value

    @property
    def Options(self):
        return Options(self.application.Options)

    @property
    def Parent(self):
        return self.application.Parent

    @property
    def Path(self):
        return self.application.Path

    @property
    def PathSeparator(self):
        return self.application.PathSeparator

    @property
    def PickerDialog(self):
        return self.application.PickerDialog

    @property
    def PortraitFontNames(self):
        return FontNames(self.application.PortraitFontNames)

    @property
    def PrintPreview(self):
        return self.application.PrintPreview

    @property
    def ProtectedViewWindows(self):
        return self.application.ProtectedViewWindows

    @property
    def RecentFiles(self):
        return self.application.RecentFiles

    @property
    def RestrictLinkedStyles(self):
        return self.application.RestrictLinkedStyles

    @RestrictLinkedStyles.setter
    def RestrictLinkedStyles(self, value):
        self.application.RestrictLinkedStyles = value

    @property
    def ScreenUpdating(self):
        return self.application.ScreenUpdating

    @property
    def Selection(self):
        return Selection(self.application.Selection)

    @property
    def SensitivityLabelPolicy(self):
        return self.application.SensitivityLabelPolicy

    @property
    def ShowStartupDialog(self):
        return self.application.ShowStartupDialog

    @property
    def ShowStylePreviews(self):
        return self.application.ShowStylePreviews

    @ShowStylePreviews.setter
    def ShowStylePreviews(self, value):
        self.application.ShowStylePreviews = value

    @property
    def ShowVisualBasicEditor(self):
        return self.application.ShowVisualBasicEditor

    @property
    def SmartArtColors(self):
        return self.application.SmartArtColors

    @property
    def SmartArtLayouts(self):
        return self.application.SmartArtLayouts

    @property
    def SmartArtQuickStyles(self):
        return self.application.SmartArtQuickStyles

    @property
    def SpecialMode(self):
        return self.application.SpecialMode

    @property
    def StartupPath(self):
        return self.application.StartupPath

    @StartupPath.setter
    def StartupPath(self, value):
        self.application.StartupPath = value

    @property
    def StatusBar(self):
        return self.application.StatusBar

    def SynonymInfo(self, *args, Word=None, LanguageID=None):
        arguments = {"Word": Word, "LanguageID": LanguageID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return SynonymInfo(self.application.SynonymInfo(*args, **arguments))

    @property
    def System(self):
        return System(self.application.System)

    @property
    def TaskPanes(self):
        return TaskPanes(self.application.TaskPanes)

    @property
    def Tasks(self):
        return self.application.Tasks

    @property
    def Templates(self):
        return self.application.Templates

    @property
    def Top(self):
        return self.application.Top

    @Top.setter
    def Top(self, value):
        self.application.Top = value

    @property
    def UndoRecord(self):
        return self.application.UndoRecord

    @property
    def UsableHeight(self):
        return self.application.UsableHeight

    @property
    def UsableWidth(self):
        return self.application.UsableWidth

    @property
    def UserAddress(self):
        return self.application.UserAddress

    @UserAddress.setter
    def UserAddress(self, value):
        self.application.UserAddress = value

    @property
    def UserControl(self):
        return self.application.UserControl

    @property
    def UserInitials(self):
        return self.application.UserInitials

    @UserInitials.setter
    def UserInitials(self, value):
        self.application.UserInitials = value

    @property
    def UserName(self):
        return self.application.UserName

    @UserName.setter
    def UserName(self, value):
        self.application.UserName = value

    @property
    def VBE(self):
        return self.application.VBE

    @property
    def Version(self):
        return self.application.Version

    @property
    def Visible(self):
        return self.application.Visible

    @property
    def Width(self):
        return self.application.Width

    @Width.setter
    def Width(self, value):
        self.application.Width = value

    @property
    def Windows(self):
        return self.application.Windows

    @property
    def WindowState(self):
        return WdWindowState(self.application.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.application.WindowState = value

    @property
    def WordBasic(self):
        return self.application.WordBasic

    @property
    def XMLNamespaces(self):
        return self.application.XMLNamespaces

    def Activate(self):
        self.application.Activate()

    def AddAddress(self, *args, TagID=None, Value=None):
        arguments = {"TagID": TagID, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.AddAddress(*args, **arguments)

    def Application(self):
        self.application.Application()

    def AutomaticChange(self):
        self.application.AutomaticChange()

    def BuildKeyCode(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.BuildKeyCode(*args, **arguments)

    def CentimetersToPoints(self, *args, Centimeters=None):
        arguments = {"Centimeters": Centimeters}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.CentimetersToPoints(*args, **arguments)

    def ChangeFileOpenDirectory(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.ChangeFileOpenDirectory(*args, **arguments)

    def CheckGrammar(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CheckGrammar(*args, **arguments)

    def CheckSpelling(self, *args, Word=None, CustomDictionary=None, IgnoreUppercase=None, MainDictionary=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = {"Word": Word, "CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "MainDictionary": MainDictionary, "CustomDictionary2": CustomDictionary2, "CustomDictionary3": CustomDictionary3, "CustomDictionary4": CustomDictionary4, "CustomDictionary5": CustomDictionary5, "CustomDictionary6": CustomDictionary6, "CustomDictionary7": CustomDictionary7, "CustomDictionary8": CustomDictionary8, "CustomDictionary9": CustomDictionary9, "CustomDictionary10": CustomDictionary10}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CheckSpelling(*args, **arguments)

    def CleanString(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.CleanString(*args, **arguments)

    def CompareDocuments(self, *args, OriginalDocument=None, RevisedDocument=None, Destination=None, Granularity=None, CompareFormatting=None, CompareCaseChanges=None, CompareWhitespace=None, CompareTables=None, CompareHeaders=None, CompareFootnotes=None, CompareTextboxes=None, CompareFields=None, CompareComments=None, CompareMoves=None, RevisedAuthor=None, IgnoreAllComparisonWarnings=None):
        arguments = {"OriginalDocument": OriginalDocument, "RevisedDocument": RevisedDocument, "Destination": Destination, "Granularity": Granularity, "CompareFormatting": CompareFormatting, "CompareCaseChanges": CompareCaseChanges, "CompareWhitespace": CompareWhitespace, "CompareTables": CompareTables, "CompareHeaders": CompareHeaders, "CompareFootnotes": CompareFootnotes, "CompareTextboxes": CompareTextboxes, "CompareFields": CompareFields, "CompareComments": CompareComments, "CompareMoves": CompareMoves, "RevisedAuthor": RevisedAuthor, "IgnoreAllComparisonWarnings": IgnoreAllComparisonWarnings}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.CompareDocuments(*args, **arguments)

    def DDEInitiate(self, *args, App=None, Topic=None):
        arguments = {"App": App, "Topic": Topic}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDEInitiate(*args, **arguments)

    def DDEPoke(self, *args, Channel=None, Item=None, Data=None):
        arguments = {"Channel": Channel, "Item": Item, "Data": Data}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDEPoke(*args, **arguments)

    def DDERequest(self, *args, Channel=None, Item=None):
        arguments = {"Channel": Channel, "Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDERequest(*args, **arguments)

    def DDETerminate(self, *args, Channel=None):
        arguments = {"Channel": Channel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.DDETerminate(*args, **arguments)

    def DDETerminateAll(self):
        self.application.DDETerminateAll()

    def DefaultWebOptions(self):
        return self.application.DefaultWebOptions()

    def GetAddress(self, *args, Name=None, AddressProperties=None, UseAutoText=None, DisplaySelectDialog=None, SelectDialog=None, CheckNamesDialog=None, RecentAddressesChoice=None, UpdateRecentAddresses=None):
        arguments = {"Name": Name, "AddressProperties": AddressProperties, "UseAutoText": UseAutoText, "DisplaySelectDialog": DisplaySelectDialog, "SelectDialog": SelectDialog, "CheckNamesDialog": CheckNamesDialog, "RecentAddressesChoice": RecentAddressesChoice, "UpdateRecentAddresses": UpdateRecentAddresses}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.GetAddress(*args, **arguments)

    def GetDefaultTheme(self, *args, DocumentType=None):
        arguments = {"DocumentType": DocumentType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.GetDefaultTheme(*args, **arguments)

    def GetSpellingSuggestions(self, *args, Word=None, CustomDictionary=None, IgnoreUppercase=None, MainDictionary=None, SuggestionMode=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = {"Word": Word, "CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "MainDictionary": MainDictionary, "SuggestionMode": SuggestionMode, "CustomDictionary2": CustomDictionary2, "CustomDictionary3": CustomDictionary3, "CustomDictionary4": CustomDictionary4, "CustomDictionary5": CustomDictionary5, "CustomDictionary6": CustomDictionary6, "CustomDictionary7": CustomDictionary7, "CustomDictionary8": CustomDictionary8, "CustomDictionary9": CustomDictionary9, "CustomDictionary10": CustomDictionary10}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.GetSpellingSuggestions(*args, **arguments)

    def GoBack(self):
        self.application.GoBack()

    def GoForward(self):
        self.application.GoForward()

    def Help(self, *args, HelpType=None):
        arguments = {"HelpType": HelpType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Help(*args, **arguments)

    def HelpTool(self):
        self.application.HelpTool()

    def InchesToPoints(self, *args, Inches=None):
        arguments = {"Inches": Inches}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.InchesToPoints(*args, **arguments)

    def Keyboard(self, *args, LangId=None):
        arguments = {"LangId": LangId}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Keyboard(*args, **arguments)

    def KeyboardBidi(self):
        self.application.KeyboardBidi()

    def KeyboardLatin(self):
        self.application.KeyboardLatin()

    def KeyString(self, *args, KeyCode=None, KeyCode2=None):
        arguments = {"KeyCode": KeyCode, "KeyCode2": KeyCode2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.KeyString(*args, **arguments)

    def LinesToPoints(self, *args, Lines=None):
        arguments = {"Lines": Lines}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.LinesToPoints(*args, **arguments)

    def ListCommands(self, *args, ListAllCommands=None):
        arguments = {"ListAllCommands": ListAllCommands}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.ListCommands(*args, **arguments)

    def LoadMasterList(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.LoadMasterList(*args, **arguments)

    def LookupNameProperties(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.LookupNameProperties(*args, **arguments)

    def MergeDocuments(self, *args, OriginalDocument=None, RevisedDocument=None, Destination=None, Granularity=None, CompareFormatting=None, CompareCaseChanges=None, CompareWhitespace=None, CompareTables=None, CompareHeaders=None, CompareFootnotes=None, CompareTextboxes=None, CompareFields=None, CompareComments=None, OriginalAuthor=None, RevisedAuthor=None, FormatFrom=None):
        arguments = {"OriginalDocument": OriginalDocument, "RevisedDocument": RevisedDocument, "Destination": Destination, "Granularity": Granularity, "CompareFormatting": CompareFormatting, "CompareCaseChanges": CompareCaseChanges, "CompareWhitespace": CompareWhitespace, "CompareTables": CompareTables, "CompareHeaders": CompareHeaders, "CompareFootnotes": CompareFootnotes, "CompareTextboxes": CompareTextboxes, "CompareFields": CompareFields, "CompareComments": CompareComments, "OriginalAuthor": OriginalAuthor, "RevisedAuthor": RevisedAuthor, "FormatFrom": FormatFrom}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.MergeDocuments(*args, **arguments)

    def MillimetersToPoints(self, *args, Millimeters=None):
        arguments = {"Millimeters": Millimeters}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.MillimetersToPoints(*args, **arguments)

    def Move(self, *args, Left=None, Top=None):
        arguments = {"Left": Left, "Top": Top}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Move(*args, **arguments)

    def NewWindow(self):
        return self.application.NewWindow()

    def OnTime(self, *args, When=None, Name=None, Tolerance=None):
        arguments = {"When": When, "Name": Name, "Tolerance": Tolerance}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.OnTime(*args, **arguments)

    def OrganizerCopy(self, *args, Source=None, Destination=None, Name=None, Object=None):
        arguments = {"Source": Source, "Destination": Destination, "Name": Name, "Object": Object}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.OrganizerCopy(*args, **arguments)

    def OrganizerDelete(self, *args, Source=None, Name=None, Object=None):
        arguments = {"Source": Source, "Name": Name, "Object": Object}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.OrganizerDelete(*args, **arguments)

    def OrganizerRename(self, *args, Source=None, Name=None, NewName=None, Object=None):
        arguments = {"Source": Source, "Name": Name, "NewName": NewName, "Object": Object}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.OrganizerRename(*args, **arguments)

    def PicasToPoints(self, *args, Picas=None):
        arguments = {"Picas": Picas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PicasToPoints(*args, **arguments)

    def PixelsToPoints(self, *args, Pixels=None, fVertical=None):
        arguments = {"Pixels": Pixels, "fVertical": fVertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PixelsToPoints(*args, **arguments)

    def PointsToCentimeters(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PointsToCentimeters(*args, **arguments)

    def PointsToInches(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PointsToInches(*args, **arguments)

    def PointsToLines(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PointsToLines(*args, **arguments)

    def PointsToMillimeters(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PointsToMillimeters(*args, **arguments)

    def PointsToPicas(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PointsToPicas(*args, **arguments)

    def PointsToPixels(self, *args, Points=None, fVertical=None):
        arguments = {"Points": Points, "fVertical": fVertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.PointsToPixels(*args, **arguments)

    def PrintOut(self, *args, Background=None, Append=None, Range=None, OutputFileName=None, From=None, To=None, Item=None, Copies=None, Pages=None, PageType=None, PrintToFile=None, Collate=None, FileName=None, ActivePrinterMacGX=None, ManualDuplexPrint=None, PrintZoomColumn=None, PrintZoomRow=None, PrintZoomPaperWidth=None, PrintZoomPaperHeight=None):
        arguments = {"Background": Background, "Append": Append, "Range": Range, "OutputFileName": OutputFileName, "From": From, "To": To, "Item": Item, "Copies": Copies, "Pages": Pages, "PageType": PageType, "PrintToFile": PrintToFile, "Collate": Collate, "FileName": FileName, "ActivePrinterMacGX": ActivePrinterMacGX, "ManualDuplexPrint": ManualDuplexPrint, "PrintZoomColumn": PrintZoomColumn, "PrintZoomRow": PrintZoomRow, "PrintZoomPaperWidth": PrintZoomPaperWidth, "PrintZoomPaperHeight": PrintZoomPaperHeight}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.PrintOut(*args, **arguments)

    def ProductCode(self):
        return self.application.ProductCode()

    def PutFocusInMailHeader(self):
        self.application.PutFocusInMailHeader()

    def Quit(self, *args, SaveChanges=None, OriginalFormat=None, RouteDocument=None):
        arguments = {"SaveChanges": SaveChanges, "OriginalFormat": OriginalFormat, "RouteDocument": RouteDocument}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Quit(*args, **arguments)

    def Repeat(self, *args, Times=None):
        arguments = {"Times": Times}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.application.Repeat(*args, **arguments)

    def ResetIgnoreAll(self):
        self.application.ResetIgnoreAll()

    def Resize(self, *args, Width=None, Height=None):
        arguments = {"Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Resize(*args, **arguments)

    def Run(self, *args, MacroName=None, varg1=None, varg2=None, varg3=None, varg4=None, varg5=None, varg6=None, varg7=None, varg8=None, varg9=None, varg10=None, varg11=None, varg12=None, varg13=None, varg14=None, varg15=None, varg16=None, varg17=None, varg18=None, varg19=None, varg20=None, varg21=None, varg22=None, varg23=None, varg24=None, varg25=None, varg26=None, varg27=None, varg28=None, varg29=None, varg30=None):
        arguments = {"MacroName": MacroName, "varg1": varg1, "varg2": varg2, "varg3": varg3, "varg4": varg4, "varg5": varg5, "varg6": varg6, "varg7": varg7, "varg8": varg8, "varg9": varg9, "varg10": varg10, "varg11": varg11, "varg12": varg12, "varg13": varg13, "varg14": varg14, "varg15": varg15, "varg16": varg16, "varg17": varg17, "varg18": varg18, "varg19": varg19, "varg20": varg20, "varg21": varg21, "varg22": varg22, "varg23": varg23, "varg24": varg24, "varg25": varg25, "varg26": varg26, "varg27": varg27, "varg28": varg28, "varg29": varg29, "varg30": varg30}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.Run(*args, **arguments)

    def ScreenRefresh(self):
        self.application.ScreenRefresh()

    def SetDefaultTheme(self, *args, Name=None, DocumentType=None):
        arguments = {"Name": Name, "DocumentType": DocumentType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.SetDefaultTheme(*args, **arguments)

    def ShowClipboard(self):
        self.application.ShowClipboard()

    def ShowMe(self):
        self.application.ShowMe()

    def SubstituteFont(self, *args, UnavailableFont=None, SubstituteFont=None):
        arguments = {"UnavailableFont": UnavailableFont, "SubstituteFont": SubstituteFont}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.application.SubstituteFont(*args, **arguments)

    def ToggleKeyboard(self):
        self.application.ToggleKeyboard()

class AutoCaption:

    def __init__(self, autocaption=None):
        self.autocaption = autocaption

    @property
    def Application(self):
        return Application(self.autocaption.Application)

    @property
    def AutoInsert(self):
        return self.autocaption.AutoInsert

    @property
    def CaptionLabel(self):
        return self.autocaption.CaptionLabel

    @CaptionLabel.setter
    def CaptionLabel(self, value):
        self.autocaption.CaptionLabel = value

    @property
    def Creator(self):
        return self.autocaption.Creator

    @property
    def Index(self):
        return self.autocaption.Index

    @property
    def Name(self):
        return self.autocaption.Name

    @Name.setter
    def Name(self, value):
        self.autocaption.Name = value

    @property
    def Parent(self):
        return self.autocaption.Parent

class AutoCorrect:

    def __init__(self, autocorrect=None):
        self.autocorrect = autocorrect

    @property
    def Application(self):
        return Application(self.autocorrect.Application)

    @property
    def CorrectCapsLock(self):
        return self.autocorrect.CorrectCapsLock

    @property
    def CorrectDays(self):
        return self.autocorrect.CorrectDays

    @property
    def CorrectHangulAndAlphabet(self):
        return self.autocorrect.CorrectHangulAndAlphabet

    @property
    def CorrectInitialCaps(self):
        return self.autocorrect.CorrectInitialCaps

    @property
    def CorrectKeyboardSetting(self):
        return self.autocorrect.CorrectKeyboardSetting

    @property
    def CorrectSentenceCaps(self):
        return self.autocorrect.CorrectSentenceCaps

    @property
    def CorrectTableCells(self):
        return self.autocorrect.CorrectTableCells

    @property
    def Creator(self):
        return self.autocorrect.Creator

    @property
    def DisplayAutoCorrectOptions(self):
        return self.autocorrect.DisplayAutoCorrectOptions

    @property
    def Entries(self):
        return self.autocorrect.Entries

    @property
    def FirstLetterAutoAdd(self):
        return self.autocorrect.FirstLetterAutoAdd

    @property
    def FirstLetterExceptions(self):
        return self.autocorrect.FirstLetterExceptions

    @property
    def HangulAndAlphabetAutoAdd(self):
        return self.autocorrect.HangulAndAlphabetAutoAdd

    @property
    def HangulAndAlphabetExceptions(self):
        return self.autocorrect.HangulAndAlphabetExceptions

    @property
    def OtherCorrectionsAutoAdd(self):
        return self.autocorrect.OtherCorrectionsAutoAdd

    @property
    def OtherCorrectionsExceptions(self):
        return self.autocorrect.OtherCorrectionsExceptions

    @property
    def Parent(self):
        return self.autocorrect.Parent

    @property
    def ReplaceText(self):
        return self.autocorrect.ReplaceText

    @property
    def ReplaceTextFromSpellingChecker(self):
        return self.autocorrect.ReplaceTextFromSpellingChecker

    @property
    def TwoInitialCapsAutoAdd(self):
        return self.autocorrect.TwoInitialCapsAutoAdd

    @property
    def TwoInitialCapsExceptions(self):
        return self.autocorrect.TwoInitialCapsExceptions

class AutoCorrectEntry:

    def __init__(self, autocorrectentry=None):
        self.autocorrectentry = autocorrectentry

    @property
    def Application(self):
        return Application(self.autocorrectentry.Application)

    @property
    def Creator(self):
        return self.autocorrectentry.Creator

    @property
    def Index(self):
        return self.autocorrectentry.Index

    @property
    def Name(self):
        return self.autocorrectentry.Name

    @Name.setter
    def Name(self, value):
        self.autocorrectentry.Name = value

    @property
    def Parent(self):
        return self.autocorrectentry.Parent

    @property
    def RichText(self):
        return self.autocorrectentry.RichText

    @property
    def Value(self):
        return self.autocorrectentry.Value

    @Value.setter
    def Value(self, value):
        self.autocorrectentry.Value = value

    def Apply(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.autocorrectentry.Apply(*args, **arguments)

    def Delete(self):
        self.autocorrectentry.Delete()

class AutoTextEntry:

    def __init__(self, autotextentry=None):
        self.autotextentry = autotextentry

    @property
    def Application(self):
        return Application(self.autotextentry.Application)

    @property
    def Creator(self):
        return self.autotextentry.Creator

    @property
    def Index(self):
        return self.autotextentry.Index

    @property
    def Name(self):
        return self.autotextentry.Name

    @Name.setter
    def Name(self, value):
        self.autotextentry.Name = value

    @property
    def Parent(self):
        return self.autotextentry.Parent

    @property
    def StyleName(self):
        return self.autotextentry.StyleName

    @property
    def Value(self):
        return self.autotextentry.Value

    @Value.setter
    def Value(self, value):
        self.autotextentry.Value = value

    def Delete(self):
        self.autotextentry.Delete()

    def Insert(self, *args, Where=None, RichText=None):
        arguments = {"Where": Where, "RichText": RichText}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.autotextentry.Insert(*args, **arguments)

class Axes:

    def __init__(self, axes=None):
        self.axes = axes

    @property
    def Application(self):
        return self.axes.Application

    @property
    def Count(self):
        return self.axes.Count

    @property
    def Creator(self):
        return self.axes.Creator

    @property
    def Parent(self):
        return self.axes.Parent

    def Item(self, *args, Type=None, AxisGroup=None):
        arguments = {"Type": Type, "AxisGroup": AxisGroup}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.axes.Item(*args, **arguments)

class Axis:

    def __init__(self, axis=None):
        self.axis = axis

    @property
    def Application(self):
        return self.axis.Application

    @property
    def AxisBetweenCategories(self):
        return self.axis.AxisBetweenCategories

    @property
    def AxisGroup(self):
        return self.axis.AxisGroup

    @property
    def AxisTitle(self):
        return AxisTitle(self.axis.AxisTitle)

    @property
    def BaseUnit(self):
        return self.axis.BaseUnit

    @BaseUnit.setter
    def BaseUnit(self, value):
        self.axis.BaseUnit = value

    @property
    def BaseUnitIsAuto(self):
        return self.axis.BaseUnitIsAuto

    @property
    def Border(self):
        return ChartBorder(self.axis.Border)

    @property
    def CategoryNames(self):
        return self.axis.CategoryNames

    @CategoryNames.setter
    def CategoryNames(self, value):
        self.axis.CategoryNames = value

    @property
    def CategoryType(self):
        return self.axis.CategoryType

    @CategoryType.setter
    def CategoryType(self, value):
        self.axis.CategoryType = value

    @property
    def Creator(self):
        return self.axis.Creator

    @property
    def Crosses(self):
        return self.axis.Crosses

    @Crosses.setter
    def Crosses(self, value):
        self.axis.Crosses = value

    @property
    def CrossesAt(self):
        return self.axis.CrossesAt

    @CrossesAt.setter
    def CrossesAt(self, value):
        self.axis.CrossesAt = value

    @property
    def DisplayUnit(self):
        return self.axis.DisplayUnit

    @DisplayUnit.setter
    def DisplayUnit(self, value):
        self.axis.DisplayUnit = value

    @property
    def DisplayUnitCustom(self):
        return self.axis.DisplayUnitCustom

    @property
    def DisplayUnitLabel(self):
        return DisplayUnitLabel(self.axis.DisplayUnitLabel)

    @property
    def Format(self):
        return ChartFormat(self.axis.Format)

    @property
    def HasDisplayUnitLabel(self):
        return self.axis.HasDisplayUnitLabel

    @property
    def HasMajorGridlines(self):
        return self.axis.HasMajorGridlines

    @property
    def HasMinorGridlines(self):
        return self.axis.HasMinorGridlines

    @property
    def HasTitle(self):
        return self.axis.HasTitle

    @property
    def Height(self):
        return self.axis.Height

    @property
    def Left(self):
        return self.axis.Left

    @property
    def LogBase(self):
        return self.axis.LogBase

    @LogBase.setter
    def LogBase(self, value):
        self.axis.LogBase = value

    @property
    def MajorGridlines(self):
        return Gridlines(self.axis.MajorGridlines)

    @property
    def MajorTickMark(self):
        return self.axis.MajorTickMark

    @MajorTickMark.setter
    def MajorTickMark(self, value):
        self.axis.MajorTickMark = value

    @property
    def MajorUnit(self):
        return self.axis.MajorUnit

    @MajorUnit.setter
    def MajorUnit(self, value):
        self.axis.MajorUnit = value

    @property
    def MajorUnitIsAuto(self):
        return self.axis.MajorUnitIsAuto

    @property
    def MajorUnitScale(self):
        return self.axis.MajorUnitScale

    @MajorUnitScale.setter
    def MajorUnitScale(self, value):
        self.axis.MajorUnitScale = value

    @property
    def MaximumScale(self):
        return self.axis.MaximumScale

    @MaximumScale.setter
    def MaximumScale(self, value):
        self.axis.MaximumScale = value

    @property
    def MaximumScaleIsAuto(self):
        return self.axis.MaximumScaleIsAuto

    @property
    def MinimumScale(self):
        return self.axis.MinimumScale

    @MinimumScale.setter
    def MinimumScale(self, value):
        self.axis.MinimumScale = value

    @property
    def MinimumScaleIsAuto(self):
        return self.axis.MinimumScaleIsAuto

    @property
    def MinorGridlines(self):
        return Gridlines(self.axis.MinorGridlines)

    @property
    def MinorTickMark(self):
        return self.axis.MinorTickMark

    @MinorTickMark.setter
    def MinorTickMark(self, value):
        self.axis.MinorTickMark = value

    @property
    def MinorUnit(self):
        return self.axis.MinorUnit

    @MinorUnit.setter
    def MinorUnit(self, value):
        self.axis.MinorUnit = value

    @property
    def MinorUnitIsAuto(self):
        return self.axis.MinorUnitIsAuto

    @property
    def MinorUnitScale(self):
        return self.axis.MinorUnitScale

    @MinorUnitScale.setter
    def MinorUnitScale(self, value):
        self.axis.MinorUnitScale = value

    @property
    def Parent(self):
        return self.axis.Parent

    @property
    def ReversePlotOrder(self):
        return self.axis.ReversePlotOrder

    @property
    def ScaleType(self):
        return self.axis.ScaleType

    @ScaleType.setter
    def ScaleType(self, value):
        self.axis.ScaleType = value

    @property
    def TickLabelPosition(self):
        return self.axis.TickLabelPosition

    @property
    def TickLabels(self):
        return TickLabels(self.axis.TickLabels)

    @property
    def TickLabelSpacing(self):
        return self.axis.TickLabelSpacing

    @TickLabelSpacing.setter
    def TickLabelSpacing(self, value):
        self.axis.TickLabelSpacing = value

    @property
    def TickLabelSpacingIsAuto(self):
        return self.axis.TickLabelSpacingIsAuto

    @TickLabelSpacingIsAuto.setter
    def TickLabelSpacingIsAuto(self, value):
        self.axis.TickLabelSpacingIsAuto = value

    @property
    def TickMarkSpacing(self):
        return self.axis.TickMarkSpacing

    @TickMarkSpacing.setter
    def TickMarkSpacing(self, value):
        self.axis.TickMarkSpacing = value

    @property
    def Top(self):
        return self.axis.Top

    @property
    def Type(self):
        return self.axis.Type

    @property
    def Width(self):
        return self.axis.Width

    def Delete(self):
        self.axis.Delete()

    def Select(self):
        self.axis.Select()

class AxisTitle:

    def __init__(self, axistitle=None):
        self.axistitle = axistitle

    @property
    def Application(self):
        return self.axistitle.Application

    @property
    def Caption(self):
        return self.axistitle.Caption

    @Caption.setter
    def Caption(self, value):
        self.axistitle.Caption = value

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ChartCharacters(self.axistitle.Characters(*args, **arguments))

    @property
    def Creator(self):
        return self.axistitle.Creator

    @property
    def Format(self):
        return ChartFormat(self.axistitle.Format)

    @property
    def Formula(self):
        return self.axistitle.Formula

    @property
    def FormulaLocal(self):
        return self.axistitle.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.axistitle.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.axistitle.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.axistitle.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.axistitle.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.axistitle.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.axistitle.Height

    @property
    def HorizontalAlignment(self):
        return self.axistitle.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.axistitle.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.axistitle.IncludeInLayout

    @property
    def Left(self):
        return self.axistitle.Left

    @Left.setter
    def Left(self, value):
        self.axistitle.Left = value

    @property
    def Name(self):
        return self.axistitle.Name

    @property
    def Orientation(self):
        return self.axistitle.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.axistitle.Orientation = value

    @property
    def Parent(self):
        return self.axistitle.Parent

    @property
    def Position(self):
        return self.axistitle.Position

    @Position.setter
    def Position(self, value):
        self.axistitle.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.axistitle.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.axistitle.ReadingOrder = value

    @property
    def Shadow(self):
        return self.axistitle.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.axistitle.Shadow = value

    @property
    def Text(self):
        return self.axistitle.Text

    @Text.setter
    def Text(self, value):
        self.axistitle.Text = value

    @property
    def Top(self):
        return self.axistitle.Top

    @Top.setter
    def Top(self, value):
        self.axistitle.Top = value

    @property
    def VerticalAlignment(self):
        return self.axistitle.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.axistitle.VerticalAlignment = value

    @property
    def Width(self):
        return self.axistitle.Width

    def Delete(self):
        self.axistitle.Delete()

    def Select(self):
        self.axistitle.Select()

class Bibliography:

    def __init__(self, bibliography=None):
        self.bibliography = bibliography

    @property
    def Application(self):
        return Application(self.bibliography.Application)

    @property
    def BibliographyStyle(self):
        return self.bibliography.BibliographyStyle

    @BibliographyStyle.setter
    def BibliographyStyle(self, value):
        self.bibliography.BibliographyStyle = value

    @property
    def Creator(self):
        return self.bibliography.Creator

    @property
    def Parent(self):
        return self.bibliography.Parent

    @property
    def Sources(self):
        return Sources(self.bibliography.Sources)

    def GenerateUniqueTag(self):
        return self.bibliography.GenerateUniqueTag()

class Bookmark:

    def __init__(self, bookmark=None):
        self.bookmark = bookmark

    @property
    def Application(self):
        return Application(self.bookmark.Application)

    @property
    def Column(self):
        return self.bookmark.Column

    @property
    def Creator(self):
        return self.bookmark.Creator

    @property
    def Empty(self):
        return self.bookmark.Empty

    @property
    def End(self):
        return self.bookmark.End

    @End.setter
    def End(self, value):
        self.bookmark.End = value

    @property
    def Name(self):
        return self.bookmark.Name

    @property
    def Parent(self):
        return self.bookmark.Parent

    @property
    def Range(self):
        return Range(self.bookmark.Range)

    @property
    def Start(self):
        return self.bookmark.Start

    @Start.setter
    def Start(self, value):
        self.bookmark.Start = value

    @property
    def StoryType(self):
        return WdStoryType(self.bookmark.StoryType)

    def Copy(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.bookmark.Copy(*args, **arguments)

    def Delete(self):
        self.bookmark.Delete()

    def Select(self):
        self.bookmark.Select()

class Border:

    def __init__(self, border=None):
        self.border = border

    @property
    def Application(self):
        return Application(self.border.Application)

    @property
    def ArtStyle(self):
        return WdPageBorderArt(self.border.ArtStyle)

    @ArtStyle.setter
    def ArtStyle(self, value):
        self.border.ArtStyle = value

    @property
    def ArtWidth(self):
        return self.border.ArtWidth

    @ArtWidth.setter
    def ArtWidth(self, value):
        self.border.ArtWidth = value

    @property
    def Color(self):
        return Border(self.border.Color)

    @Color.setter
    def Color(self, value):
        self.border.Color = value

    @property
    def ColorIndex(self):
        return WdColorIndex(self.border.ColorIndex)

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.border.ColorIndex = value

    @property
    def Creator(self):
        return self.border.Creator

    @property
    def Inside(self):
        return self.border.Inside

    @property
    def LineStyle(self):
        return WdLineStyle(self.border.LineStyle)

    @LineStyle.setter
    def LineStyle(self, value):
        self.border.LineStyle = value

    @property
    def LineWidth(self):
        return self.border.LineWidth

    @LineWidth.setter
    def LineWidth(self, value):
        self.border.LineWidth = value

    @property
    def Parent(self):
        return self.border.Parent

    @property
    def Visible(self):
        return self.border.Visible

class Break:

    def __init__(self, break=None):
        self.break = break

    @property
    def Application(self):
        return Application(self.break.Application)

    @property
    def Creator(self):
        return self.break.Creator

    @property
    def PageIndex(self):
        return self.break.PageIndex

    @property
    def Parent(self):
        return self.break.Parent

    @property
    def Range(self):
        return Range(self.break.Range)

class Breaks:

    def __init__(self, breaks=None):
        self.breaks = breaks

    def __call__(self, item):
        return Break(self.breaks(item))

    @property
    def Application(self):
        return Application(self.breaks.Application)

    @property
    def Count(self):
        return Breaks(self.breaks.Count)

    @property
    def Creator(self):
        return self.breaks.Creator

    @property
    def Parent(self):
        return self.breaks.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.breaks.Item(*args, **arguments)

class Browser:

    def __init__(self, browser=None):
        self.browser = browser

    @property
    def Application(self):
        return Application(self.browser.Application)

    @property
    def Creator(self):
        return self.browser.Creator

    @property
    def Parent(self):
        return self.browser.Parent

    @property
    def Target(self):
        return self.browser.Target

    @Target.setter
    def Target(self, value):
        self.browser.Target = value

    def Next(self):
        self.browser.Next()

    def Previous(self):
        self.browser.Previous()

class BuildingBlock:

    def __init__(self, buildingblock=None):
        self.buildingblock = buildingblock

    @property
    def Application(self):
        return Application(self.buildingblock.Application)

    @property
    def Category(self):
        return Category(self.buildingblock.Category)

    @property
    def Creator(self):
        return self.buildingblock.Creator

    @property
    def Description(self):
        return self.buildingblock.Description

    @Description.setter
    def Description(self, value):
        self.buildingblock.Description = value

    @property
    def ID(self):
        return self.buildingblock.ID

    @property
    def Index(self):
        return self.buildingblock.Index

    @property
    def InsertOptions(self):
        return self.buildingblock.InsertOptions

    @InsertOptions.setter
    def InsertOptions(self, value):
        self.buildingblock.InsertOptions = value

    @property
    def Name(self):
        return self.buildingblock.Name

    @Name.setter
    def Name(self, value):
        self.buildingblock.Name = value

    @property
    def Parent(self):
        return self.buildingblock.Parent

    @property
    def Type(self):
        return BuildingBlockType(self.buildingblock.Type)

    @property
    def Value(self):
        return self.buildingblock.Value

    @Value.setter
    def Value(self, value):
        self.buildingblock.Value = value

    def Delete(self):
        self.buildingblock.Delete()

    def Insert(self, *args, Where=None, RichText=None):
        arguments = {"Where": Where, "RichText": RichText}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.buildingblock.Insert(*args, **arguments)

class BuildingBlockEntries:

    def __init__(self, buildingblockentries=None):
        self.buildingblockentries = buildingblockentries

    @property
    def Application(self):
        return Application(self.buildingblockentries.Application)

    @property
    def Count(self):
        return BuildingBlockEntries(self.buildingblockentries.Count)

    @property
    def Creator(self):
        return self.buildingblockentries.Creator

    @property
    def Parent(self):
        return self.buildingblockentries.Parent

    def Add(self, *args, Name=None, Type=None, Category=None, Range=None, Description=None, InsertOptions=None):
        arguments = {"Name": Name, "Type": Type, "Category": Category, "Range": Range, "Description": Description, "InsertOptions": InsertOptions}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.buildingblockentries.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.buildingblockentries.Item(*args, **arguments)

class BuildingBlocks:

    def __init__(self, buildingblocks=None):
        self.buildingblocks = buildingblocks

    @property
    def Application(self):
        return Application(self.buildingblocks.Application)

    @property
    def Count(self):
        return BuildingBlocks(self.buildingblocks.Count)

    @property
    def Creator(self):
        return self.buildingblocks.Creator

    @property
    def Parent(self):
        return self.buildingblocks.Parent

    def Add(self, *args, Name=None, Range=None, Description=None, InsertOptions=None):
        arguments = {"Name": Name, "Range": Range, "Description": Description, "InsertOptions": InsertOptions}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.buildingblocks.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.buildingblocks.Item(*args, **arguments)

class BuildingBlockType:

    def __init__(self, buildingblocktype=None):
        self.buildingblocktype = buildingblocktype

    @property
    def Application(self):
        return Application(self.buildingblocktype.Application)

    @property
    def Categories(self):
        return Categories(self.buildingblocktype.Categories)

    @property
    def Creator(self):
        return self.buildingblocktype.Creator

    @property
    def Index(self):
        return self.buildingblocktype.Index

    @property
    def Name(self):
        return self.buildingblocktype.Name

    @property
    def Parent(self):
        return self.buildingblocktype.Parent

class BuildingBlockTypes:

    def __init__(self, buildingblocktypes=None):
        self.buildingblocktypes = buildingblocktypes

    @property
    def Application(self):
        return Application(self.buildingblocktypes.Application)

    @property
    def Count(self):
        return BuildingBlockTypes(self.buildingblocktypes.Count)

    @property
    def Creator(self):
        return self.buildingblocktypes.Creator

    @property
    def Parent(self):
        return self.buildingblocktypes.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.buildingblocktypes.Item(*args, **arguments)

class CalloutFormat:

    def __init__(self, calloutformat=None):
        self.calloutformat = calloutformat

    @property
    def Accent(self):
        return self.calloutformat.Accent

    @property
    def Angle(self):
        return self.calloutformat.Angle

    @Angle.setter
    def Angle(self, value):
        self.calloutformat.Angle = value

    @property
    def Application(self):
        return Application(self.calloutformat.Application)

    @property
    def AutoLength(self):
        return self.calloutformat.AutoLength

    @property
    def Border(self):
        return self.calloutformat.Border

    @Border.setter
    def Border(self, value):
        self.calloutformat.Border = value

    @property
    def Creator(self):
        return self.calloutformat.Creator

    @property
    def Drop(self):
        return self.calloutformat.Drop

    @property
    def DropType(self):
        return self.calloutformat.DropType

    @property
    def Gap(self):
        return self.calloutformat.Gap

    @Gap.setter
    def Gap(self, value):
        self.calloutformat.Gap = value

    @property
    def Length(self):
        return self.calloutformat.Length

    @property
    def Parent(self):
        return self.calloutformat.Parent

    @property
    def Type(self):
        return self.calloutformat.Type

    @Type.setter
    def Type(self, value):
        self.calloutformat.Type = value

    def CustomDrop(self, *args, Drop=None):
        arguments = {"Drop": Drop}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.calloutformat.CustomDrop(*args, **arguments)

    def CustomLength(self, *args, Length=None):
        arguments = {"Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.calloutformat.CustomLength(*args, **arguments)

    def PresetDrop(self, *args, DropType=None):
        arguments = {"DropType": DropType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.calloutformat.PresetDrop(*args, **arguments)

class CanvasShapes:

    def __init__(self, canvasshapes=None):
        self.canvasshapes = canvasshapes

    @property
    def Application(self):
        return Application(self.canvasshapes.Application)

    @property
    def Count(self):
        return self.canvasshapes.Count

    @property
    def Creator(self):
        return self.canvasshapes.Creator

    @property
    def Parent(self):
        return self.canvasshapes.Parent

    def AddCallout(self, *args, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Type": Type, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddCallout(*args, **arguments)

    def AddConnector(self, *args, Type=None, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = {"Type": Type, "BeginX": BeginX, "BeginY": BeginY, "EndX": EndX, "EndY": EndY}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddConnector(*args, **arguments)

    def AddCurve(self, *args, SafeArrayOfPoints=None):
        arguments = {"SafeArrayOfPoints": SafeArrayOfPoints}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddCurve(*args, **arguments)

    def AddLabel(self, *args, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Orientation": Orientation, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddLabel(*args, **arguments)

    def AddLine(self, *args, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = {"BeginX": BeginX, "BeginY": BeginY, "EndX": EndX, "EndY": EndY}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddLine(*args, **arguments)

    def AddPicture(self, *args, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"FileName": FileName, "LinkToFile": LinkToFile, "SaveWithDocument": SaveWithDocument, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddPicture(*args, **arguments)

    def AddPolyline(self, *args, SafeArrayOfPoints=None):
        arguments = {"SafeArrayOfPoints": SafeArrayOfPoints}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddPolyline(*args, **arguments)

    def AddShape(self, *args, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Type": Type, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddShape(*args, **arguments)

    def AddTextbox(self, *args, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = {"Orientation": Orientation, "Left": Left, "Top": Top, "Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddTextbox(*args, **arguments)

    def AddTextEffect(self, *args, PresetTextEffect=None, Text=None, FontName=None, FontSize=None, FontBold=None, FontItalic=None, Left=None, Top=None):
        arguments = {"PresetTextEffect": PresetTextEffect, "Text": Text, "FontName": FontName, "FontSize": FontSize, "FontBold": FontBold, "FontItalic": FontItalic, "Left": Left, "Top": Top}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.AddTextEffect(*args, **arguments)

    def BuildFreeform(self, *args, EditingType=None, X1=None, Y1=None):
        arguments = {"EditingType": EditingType, "X1": X1, "Y1": Y1}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.canvasshapes.BuildFreeform(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.canvasshapes.Item(*args, **arguments)

    def Range(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.canvasshapes.Range(*args, **arguments)

    def SelectAll(self):
        self.canvasshapes.SelectAll()

class CaptionLabel:

    def __init__(self, captionlabel=None):
        self.captionlabel = captionlabel

    @property
    def Application(self):
        return Application(self.captionlabel.Application)

    @property
    def BuiltIn(self):
        return self.captionlabel.BuiltIn

    @property
    def ChapterStyleLevel(self):
        return self.captionlabel.ChapterStyleLevel

    @ChapterStyleLevel.setter
    def ChapterStyleLevel(self, value):
        self.captionlabel.ChapterStyleLevel = value

    @property
    def Creator(self):
        return self.captionlabel.Creator

    @property
    def ID(self):
        return WdCaptionLabelID(self.captionlabel.ID)

    @property
    def IncludeChapterNumber(self):
        return self.captionlabel.IncludeChapterNumber

    @property
    def Name(self):
        return self.captionlabel.Name

    @property
    def NumberStyle(self):
        return CaptionLabel(self.captionlabel.NumberStyle)

    @NumberStyle.setter
    def NumberStyle(self, value):
        self.captionlabel.NumberStyle = value

    @property
    def Parent(self):
        return self.captionlabel.Parent

    @property
    def Position(self):
        return WdCaptionPosition(self.captionlabel.Position)

    @Position.setter
    def Position(self, value):
        self.captionlabel.Position = value

    @property
    def Separator(self):
        return WdSeparatorType(self.captionlabel.Separator)

    @Separator.setter
    def Separator(self, value):
        self.captionlabel.Separator = value

    def Delete(self):
        self.captionlabel.Delete()

class Categories:

    def __init__(self, categories=None):
        self.categories = categories

    @property
    def Application(self):
        return Application(self.categories.Application)

    @property
    def Count(self):
        return Categories(self.categories.Count)

    @property
    def Creator(self):
        return self.categories.Creator

    @property
    def Parent(self):
        return self.categories.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.categories.Item(*args, **arguments)

class Category:

    def __init__(self, category=None):
        self.category = category

    @property
    def Application(self):
        return Application(self.category.Application)

    @property
    def BuildingBlocks(self):
        return BuildingBlocks(self.category.BuildingBlocks)

    @property
    def Creator(self):
        return self.category.Creator

    @property
    def Index(self):
        return self.category.Index

    @property
    def Name(self):
        return self.category.Name

    @property
    def Parent(self):
        return self.category.Parent

    @property
    def Type(self):
        return BuildingBlockType(self.category.Type)

class Cell:

    def __init__(self, cell=None):
        self.cell = cell

    @property
    def Application(self):
        return Application(self.cell.Application)

    @property
    def Borders(self):
        return self.cell.Borders

    @property
    def BottomPadding(self):
        return self.cell.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.cell.BottomPadding = value

    @property
    def Column(self):
        return Column(self.cell.Column)

    @property
    def ColumnIndex(self):
        return self.cell.ColumnIndex

    @property
    def Creator(self):
        return self.cell.Creator

    @property
    def FitText(self):
        return self.cell.FitText

    @property
    def Height(self):
        return self.cell.Height

    @Height.setter
    def Height(self, value):
        self.cell.Height = value

    @property
    def HeightRule(self):
        return WdRowHeightRule(self.cell.HeightRule)

    @HeightRule.setter
    def HeightRule(self, value):
        self.cell.HeightRule = value

    @property
    def Id(self):
        return self.cell.Id

    @Id.setter
    def Id(self, value):
        self.cell.Id = value

    @property
    def LeftPadding(self):
        return self.cell.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.cell.LeftPadding = value

    @property
    def NestingLevel(self):
        return self.cell.NestingLevel

    @property
    def Next(self):
        return Cell(self.cell.Next)

    @property
    def Parent(self):
        return self.cell.Parent

    @property
    def PreferredWidth(self):
        return self.cell.PreferredWidth

    @PreferredWidth.setter
    def PreferredWidth(self, value):
        self.cell.PreferredWidth = value

    @property
    def PreferredWidthType(self):
        return WdPreferredWidthType(self.cell.PreferredWidthType)

    @PreferredWidthType.setter
    def PreferredWidthType(self, value):
        self.cell.PreferredWidthType = value

    @property
    def Previous(self):
        return self.cell.Previous

    @property
    def Range(self):
        return Range(self.cell.Range)

    @property
    def RightPadding(self):
        return self.cell.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.cell.RightPadding = value

    @property
    def Row(self):
        return Row(self.cell.Row)

    @property
    def RowIndex(self):
        return self.cell.RowIndex

    @property
    def Shading(self):
        return Shading(self.cell.Shading)

    @property
    def Tables(self):
        return self.cell.Tables

    @property
    def TopPadding(self):
        return self.cell.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.cell.TopPadding = value

    @property
    def VerticalAlignment(self):
        return WdCellVerticalAlignment(self.cell.VerticalAlignment)

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.cell.VerticalAlignment = value

    @property
    def Width(self):
        return self.cell.Width

    @Width.setter
    def Width(self, value):
        self.cell.Width = value

    @property
    def WordWrap(self):
        return self.cell.WordWrap

    def AutoSum(self):
        self.cell.AutoSum()

    def Delete(self, *args, ShiftCells=None):
        arguments = {"ShiftCells": ShiftCells}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cell.Delete(*args, **arguments)

    def Formula(self, *args, Formula=None, NumFormat=None):
        arguments = {"Formula": Formula, "NumFormat": NumFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cell.Formula(*args, **arguments)

    def Merge(self, *args, MergeTo=None):
        arguments = {"MergeTo": MergeTo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cell.Merge(*args, **arguments)

    def Select(self):
        self.cell.Select()

    def SetHeight(self, *args, RowHeight=None, HeightRule=None):
        arguments = {"RowHeight": RowHeight, "HeightRule": HeightRule}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cell.SetHeight(*args, **arguments)

    def SetWidth(self, *args, ColumnWidth=None, RulerStyle=None):
        arguments = {"ColumnWidth": ColumnWidth, "RulerStyle": RulerStyle}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cell.SetWidth(*args, **arguments)

    def Split(self, *args, NumRows=None, NumColumns=None):
        arguments = {"NumRows": NumRows, "NumColumns": NumColumns}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.cell.Split(*args, **arguments)

class Chart:

    def __init__(self, chart=None):
        self.chart = chart

    @property
    def Application(self):
        return self.chart.Application

    @property
    def AutoScaling(self):
        return self.chart.AutoScaling

    @property
    def BackWall(self):
        return Walls(self.chart.BackWall)

    @property
    def BarShape(self):
        return self.chart.BarShape

    @BarShape.setter
    def BarShape(self, value):
        self.chart.BarShape = value

    @property
    def ChartArea(self):
        return ChartArea(self.chart.ChartArea)

    @property
    def ChartData(self):
        return ChartData(self.chart.ChartData)

    def ChartGroups(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.ChartGroups(*args, **arguments)

    @property
    def ChartStyle(self):
        return self.chart.ChartStyle

    @ChartStyle.setter
    def ChartStyle(self, value):
        self.chart.ChartStyle = value

    @property
    def ChartTitle(self):
        return ChartTitle(self.chart.ChartTitle)

    @property
    def ChartType(self):
        return self.chart.ChartType

    @ChartType.setter
    def ChartType(self, value):
        self.chart.ChartType = value

    @property
    def Creator(self):
        return self.chart.Creator

    @property
    def DataTable(self):
        return DataTable(self.chart.DataTable)

    @property
    def DepthPercent(self):
        return self.chart.DepthPercent

    @DepthPercent.setter
    def DepthPercent(self, value):
        self.chart.DepthPercent = value

    @property
    def DisplayBlanksAs(self):
        return self.chart.DisplayBlanksAs

    @DisplayBlanksAs.setter
    def DisplayBlanksAs(self, value):
        self.chart.DisplayBlanksAs = value

    @property
    def Elevation(self):
        return self.chart.Elevation

    @Elevation.setter
    def Elevation(self, value):
        self.chart.Elevation = value

    @property
    def Floor(self):
        return Floor(self.chart.Floor)

    @property
    def GapDepth(self):
        return self.chart.GapDepth

    @GapDepth.setter
    def GapDepth(self, value):
        self.chart.GapDepth = value

    @property
    def HasAxis(self):
        return self.chart.HasAxis

    @HasAxis.setter
    def HasAxis(self, value):
        self.chart.HasAxis = value

    @property
    def HasDataTable(self):
        return self.chart.HasDataTable

    @property
    def HasLegend(self):
        return self.chart.HasLegend

    @property
    def HasTitle(self):
        return self.chart.HasTitle

    @property
    def HeightPercent(self):
        return self.chart.HeightPercent

    @HeightPercent.setter
    def HeightPercent(self, value):
        self.chart.HeightPercent = value

    @property
    def Legend(self):
        return Legend(self.chart.Legend)

    @property
    def Parent(self):
        return self.chart.Parent

    @property
    def Perspective(self):
        return self.chart.Perspective

    @Perspective.setter
    def Perspective(self, value):
        self.chart.Perspective = value

    @property
    def PivotLayout(self):
        return self.chart.PivotLayout

    @property
    def PlotArea(self):
        return PlotArea(self.chart.PlotArea)

    @property
    def PlotBy(self):
        return self.chart.PlotBy

    @PlotBy.setter
    def PlotBy(self, value):
        self.chart.PlotBy = value

    @property
    def PlotVisibleOnly(self):
        return self.chart.PlotVisibleOnly

    @property
    def RightAngleAxes(self):
        return self.chart.RightAngleAxes

    @property
    def Rotation(self):
        return self.chart.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.chart.Rotation = value

    @property
    def Shapes(self):
        return self.chart.Shapes

    @property
    def ShowAllFieldButtons(self):
        return self.chart.ShowAllFieldButtons

    @ShowAllFieldButtons.setter
    def ShowAllFieldButtons(self, value):
        self.chart.ShowAllFieldButtons = value

    @property
    def ShowAxisFieldButtons(self):
        return self.chart.ShowAxisFieldButtons

    @ShowAxisFieldButtons.setter
    def ShowAxisFieldButtons(self, value):
        self.chart.ShowAxisFieldButtons = value

    @property
    def ShowDataLabelsOverMaximum(self):
        return self.chart.ShowDataLabelsOverMaximum

    @ShowDataLabelsOverMaximum.setter
    def ShowDataLabelsOverMaximum(self, value):
        self.chart.ShowDataLabelsOverMaximum = value

    @property
    def ShowLegendFieldButtons(self):
        return self.chart.ShowLegendFieldButtons

    @ShowLegendFieldButtons.setter
    def ShowLegendFieldButtons(self, value):
        self.chart.ShowLegendFieldButtons = value

    @property
    def ShowReportFilterFieldButtons(self):
        return self.chart.ShowReportFilterFieldButtons

    @ShowReportFilterFieldButtons.setter
    def ShowReportFilterFieldButtons(self, value):
        self.chart.ShowReportFilterFieldButtons = value

    @property
    def ShowValueFieldButtons(self):
        return self.chart.ShowValueFieldButtons

    @ShowValueFieldButtons.setter
    def ShowValueFieldButtons(self, value):
        self.chart.ShowValueFieldButtons = value

    @property
    def SideWall(self):
        return Walls(self.chart.SideWall)

    @property
    def Walls(self):
        return Walls(self.chart.Walls)

    def ApplyChartTemplate(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ApplyChartTemplate(*args, **arguments)

    def ApplyDataLabels(self, *args, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = {"Type": Type, "LegendKey": LegendKey, "AutoText": AutoText, "HasLeaderLines": HasLeaderLines, "ShowSeriesName": ShowSeriesName, "ShowCategoryName": ShowCategoryName, "ShowValue": ShowValue, "ShowPercentage": ShowPercentage, "ShowBubbleSize": ShowBubbleSize, "Separator": Separator}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ApplyDataLabels(*args, **arguments)

    def ApplyLayout(self, *args, Layout=None, ChartType=None):
        arguments = {"Layout": Layout, "ChartType": ChartType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ApplyLayout(*args, **arguments)

    def Axes(self, *args, Type=None, AxisGroup=None):
        arguments = {"Type": Type, "AxisGroup": AxisGroup}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.Axes(*args, **arguments)

    def ChartWizard(self, *args, Source=None, Gallery=None, Format=None, PlotBy=None, CategoryLabels=None, SeriesLabels=None, HasLegend=None, Title=None, CategoryTitle=None, ValueTitle=None, ExtraTitle=None):
        arguments = {"Source": Source, "Gallery": Gallery, "Format": Format, "PlotBy": PlotBy, "CategoryLabels": CategoryLabels, "SeriesLabels": SeriesLabels, "HasLegend": HasLegend, "Title": Title, "CategoryTitle": CategoryTitle, "ValueTitle": ValueTitle, "ExtraTitle": ExtraTitle}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.ChartWizard(*args, **arguments)

    def ClearToMatchStyle(self):
        self.chart.ClearToMatchStyle()

    def Copy(self, *args, Before=None, After=None):
        arguments = {"Before": Before, "After": After}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.Copy(*args, **arguments)

    def CopyPicture(self, *args, Appearance=None, Format=None, Size=None):
        arguments = {"Appearance": Appearance, "Format": Format, "Size": Size}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.CopyPicture(*args, **arguments)

    def Delete(self):
        self.chart.Delete()

    def Export(self, *args, FileName=None, FilterName=None, Interactive=None):
        arguments = {"FileName": FileName, "FilterName": FilterName, "Interactive": Interactive}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.chart.Export(*args, **arguments)

    def GetChartElement(self, *args, x=None, y=None, ElementID=None, Arg1=None, Arg2=None):
        arguments = {"x": x, "y": y, "ElementID": ElementID, "Arg1": Arg1, "Arg2": Arg2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.GetChartElement(*args, **arguments)

    def Paste(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.Paste(*args, **arguments)

    def Refresh(self):
        self.chart.Refresh()

    def SaveChartTemplate(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.SaveChartTemplate(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.Select(*args, **arguments)

    def SeriesCollection(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return SeriesCollection(self.chart.SeriesCollection(*args, **arguments))

    def SetBackgroundPicture(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.SetBackgroundPicture(*args, **arguments)

    def SetDefaultChart(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.SetDefaultChart(*args, **arguments)

    def SetElement(self, *args, Element=None):
        arguments = {"Element": Element}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.SetElement(*args, **arguments)

    def SetSourceData(self, *args, Source=None, PlotBy=None):
        arguments = {"Source": Source, "PlotBy": PlotBy}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chart.SetSourceData(*args, **arguments)

class ChartArea:

    def __init__(self, chartarea=None):
        self.chartarea = chartarea

    @property
    def Application(self):
        return self.chartarea.Application

    @property
    def Creator(self):
        return self.chartarea.Creator

    @property
    def Format(self):
        return ChartFormat(self.chartarea.Format)

    @property
    def Height(self):
        return self.chartarea.Height

    @Height.setter
    def Height(self, value):
        self.chartarea.Height = value

    @property
    def Left(self):
        return self.chartarea.Left

    @Left.setter
    def Left(self, value):
        self.chartarea.Left = value

    @property
    def Name(self):
        return self.chartarea.Name

    @property
    def Parent(self):
        return self.chartarea.Parent

    @property
    def Shadow(self):
        return self.chartarea.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.chartarea.Shadow = value

    @property
    def Top(self):
        return self.chartarea.Top

    @Top.setter
    def Top(self, value):
        self.chartarea.Top = value

    @property
    def Width(self):
        return self.chartarea.Width

    @Width.setter
    def Width(self, value):
        self.chartarea.Width = value

    def Clear(self):
        self.chartarea.Clear()

    def ClearContents(self):
        self.chartarea.ClearContents()

    def ClearFormats(self):
        self.chartarea.ClearFormats()

    def Copy(self):
        self.chartarea.Copy()

    def Select(self):
        self.chartarea.Select()

class ChartBorder:

    def __init__(self, chartborder=None):
        self.chartborder = chartborder

    @property
    def Application(self):
        return self.chartborder.Application

    @property
    def Color(self):
        return self.chartborder.Color

    @Color.setter
    def Color(self, value):
        self.chartborder.Color = value

    @property
    def ColorIndex(self):
        return self.chartborder.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.chartborder.ColorIndex = value

    @property
    def Creator(self):
        return self.chartborder.Creator

    @property
    def LineStyle(self):
        return self.chartborder.LineStyle

    @LineStyle.setter
    def LineStyle(self, value):
        self.chartborder.LineStyle = value

    @property
    def Parent(self):
        return self.chartborder.Parent

    @property
    def Weight(self):
        return self.chartborder.Weight

    @Weight.setter
    def Weight(self, value):
        self.chartborder.Weight = value

class ChartCharacters:

    def __init__(self, chartcharacters=None):
        self.chartcharacters = chartcharacters

    @property
    def Application(self):
        return self.chartcharacters.Application

    @property
    def Caption(self):
        return self.chartcharacters.Caption

    @property
    def Count(self):
        return self.chartcharacters.Count

    @property
    def Creator(self):
        return self.chartcharacters.Creator

    @property
    def Font(self):
        return ChartFont(self.chartcharacters.Font)

    @property
    def Parent(self):
        return self.chartcharacters.Parent

    @property
    def PhoneticCharacters(self):
        return self.chartcharacters.PhoneticCharacters

    @PhoneticCharacters.setter
    def PhoneticCharacters(self, value):
        self.chartcharacters.PhoneticCharacters = value

    @property
    def Text(self):
        return self.chartcharacters.Text

    @Text.setter
    def Text(self, value):
        self.chartcharacters.Text = value

    def Delete(self):
        self.chartcharacters.Delete()

    def Insert(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.chartcharacters.Insert(*args, **arguments)

class ChartColorFormat:

    def __init__(self, chartcolorformat=None):
        self.chartcolorformat = chartcolorformat

    @property
    def Application(self):
        return self.chartcolorformat.Application

    @property
    def Creator(self):
        return self.chartcolorformat.Creator

    @property
    def Parent(self):
        return self.chartcolorformat.Parent

    @property
    def RGB(self):
        return self.chartcolorformat.RGB

    @property
    def SchemeColor(self):
        return self.chartcolorformat.SchemeColor

    @SchemeColor.setter
    def SchemeColor(self, value):
        self.chartcolorformat.SchemeColor = value

    @property
    def Type(self):
        return self.chartcolorformat.Type

class ChartData:

    def __init__(self, chartdata=None):
        self.chartdata = chartdata

    @property
    def IsLinked(self):
        return self.chartdata.IsLinked

    @property
    def Workbook(self):
        return self.chartdata.Workbook

    def Activate(self):
        self.chartdata.Activate()

    def BreakLink(self):
        self.chartdata.BreakLink()

class ChartFont:

    def __init__(self, chartfont=None):
        self.chartfont = chartfont

    @property
    def Application(self):
        return self.chartfont.Application

    @property
    def Background(self):
        return self.chartfont.Background

    @Background.setter
    def Background(self, value):
        self.chartfont.Background = value

    @property
    def Bold(self):
        return self.chartfont.Bold

    @property
    def Color(self):
        return self.chartfont.Color

    @Color.setter
    def Color(self, value):
        self.chartfont.Color = value

    @property
    def ColorIndex(self):
        return self.chartfont.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.chartfont.ColorIndex = value

    @property
    def Creator(self):
        return self.chartfont.Creator

    @property
    def FontStyle(self):
        return self.chartfont.FontStyle

    @FontStyle.setter
    def FontStyle(self, value):
        self.chartfont.FontStyle = value

    @property
    def Italic(self):
        return self.chartfont.Italic

    @property
    def Name(self):
        return self.chartfont.Name

    @Name.setter
    def Name(self, value):
        self.chartfont.Name = value

    @property
    def Parent(self):
        return self.chartfont.Parent

    @property
    def Size(self):
        return self.chartfont.Size

    @Size.setter
    def Size(self, value):
        self.chartfont.Size = value

    @property
    def StrikeThrough(self):
        return self.chartfont.StrikeThrough

    @property
    def Subscript(self):
        return self.chartfont.Subscript

    @property
    def Superscript(self):
        return self.chartfont.Superscript

    @property
    def Underline(self):
        return self.chartfont.Underline

    @Underline.setter
    def Underline(self, value):
        self.chartfont.Underline = value

class ChartFormat:

    def __init__(self, chartformat=None):
        self.chartformat = chartformat

    @property
    def Application(self):
        return self.chartformat.Application

    @property
    def Creator(self):
        return self.chartformat.Creator

    @property
    def Fill(self):
        return FillFormat(self.chartformat.Fill)

    @property
    def Glow(self):
        return GlowFormat(self.chartformat.Glow)

    @property
    def Line(self):
        return LineFormat(self.chartformat.Line)

    @property
    def Parent(self):
        return self.chartformat.Parent

    @property
    def PictureFormat(self):
        return PictureFormat(self.chartformat.PictureFormat)

    @property
    def Shadow(self):
        return ShadowFormat(self.chartformat.Shadow)

    @property
    def SoftEdge(self):
        return SoftEdgeFormat(self.chartformat.SoftEdge)

    @property
    def TextFrame2(self):
        return self.chartformat.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.chartformat.ThreeD)

class ChartGroup:

    def __init__(self, chartgroup=None):
        self.chartgroup = chartgroup

    @property
    def Application(self):
        return self.chartgroup.Application

    @property
    def AxisGroup(self):
        return self.chartgroup.AxisGroup

    @property
    def BubbleScale(self):
        return self.chartgroup.BubbleScale

    @BubbleScale.setter
    def BubbleScale(self, value):
        self.chartgroup.BubbleScale = value

    @property
    def Creator(self):
        return self.chartgroup.Creator

    @property
    def DoughnutHoleSize(self):
        return self.chartgroup.DoughnutHoleSize

    @DoughnutHoleSize.setter
    def DoughnutHoleSize(self, value):
        self.chartgroup.DoughnutHoleSize = value

    @property
    def DownBars(self):
        return DownBars(self.chartgroup.DownBars)

    @property
    def DropLines(self):
        return DropLines(self.chartgroup.DropLines)

    @property
    def FirstSliceAngle(self):
        return self.chartgroup.FirstSliceAngle

    @FirstSliceAngle.setter
    def FirstSliceAngle(self, value):
        self.chartgroup.FirstSliceAngle = value

    @property
    def GapWidth(self):
        return self.chartgroup.GapWidth

    @property
    def Has3DShading(self):
        return self.chartgroup.Has3DShading

    @property
    def HasDropLines(self):
        return self.chartgroup.HasDropLines

    @property
    def HasHiLoLines(self):
        return self.chartgroup.HasHiLoLines

    @property
    def HasRadarAxisLabels(self):
        return self.chartgroup.HasRadarAxisLabels

    @property
    def HasSeriesLines(self):
        return self.chartgroup.HasSeriesLines

    @property
    def HasUpDownBars(self):
        return self.chartgroup.HasUpDownBars

    @property
    def HiLoLines(self):
        return HiLoLines(self.chartgroup.HiLoLines)

    @property
    def Index(self):
        return self.chartgroup.Index

    @property
    def Overlap(self):
        return self.chartgroup.Overlap

    @property
    def Parent(self):
        return self.chartgroup.Parent

    @property
    def RadarAxisLabels(self):
        return TickLabels(self.chartgroup.RadarAxisLabels)

    @property
    def SecondPlotSize(self):
        return self.chartgroup.SecondPlotSize

    @SecondPlotSize.setter
    def SecondPlotSize(self, value):
        self.chartgroup.SecondPlotSize = value

    @property
    def SeriesLines(self):
        return SeriesLines(self.chartgroup.SeriesLines)

    @property
    def ShowNegativeBubbles(self):
        return self.chartgroup.ShowNegativeBubbles

    @property
    def SizeRepresents(self):
        return self.chartgroup.SizeRepresents

    @SizeRepresents.setter
    def SizeRepresents(self, value):
        self.chartgroup.SizeRepresents = value

    @property
    def SplitType(self):
        return self.chartgroup.SplitType

    @SplitType.setter
    def SplitType(self, value):
        self.chartgroup.SplitType = value

    @property
    def SplitValue(self):
        return self.chartgroup.SplitValue

    @SplitValue.setter
    def SplitValue(self, value):
        self.chartgroup.SplitValue = value

    @property
    def UpBars(self):
        return UpBars(self.chartgroup.UpBars)

    @property
    def VaryByCategories(self):
        return self.chartgroup.VaryByCategories

    def SeriesCollection(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return SeriesCollection(self.chartgroup.SeriesCollection(*args, **arguments))

class ChartGroups:

    def __init__(self, chartgroups=None):
        self.chartgroups = chartgroups

    @property
    def Application(self):
        return self.chartgroups.Application

    @property
    def Count(self):
        return self.chartgroups.Count

    @property
    def Creator(self):
        return self.chartgroups.Creator

    @property
    def Parent(self):
        return self.chartgroups.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ChartGroup(self.chartgroups.Item(*args, **arguments))

class ChartTitle:

    def __init__(self, charttitle=None):
        self.charttitle = charttitle

    @property
    def Application(self):
        return self.charttitle.Application

    @property
    def Caption(self):
        return self.charttitle.Caption

    @Caption.setter
    def Caption(self, value):
        self.charttitle.Caption = value

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ChartCharacters(self.charttitle.Characters(*args, **arguments))

    @property
    def Creator(self):
        return self.charttitle.Creator

    @property
    def Format(self):
        return ChartFormat(self.charttitle.Format)

    @property
    def Formula(self):
        return self.charttitle.Formula

    @property
    def FormulaLocal(self):
        return self.charttitle.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.charttitle.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.charttitle.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.charttitle.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.charttitle.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.charttitle.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.charttitle.Height

    @property
    def HorizontalAlignment(self):
        return self.charttitle.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.charttitle.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.charttitle.IncludeInLayout

    @property
    def Left(self):
        return self.charttitle.Left

    @Left.setter
    def Left(self, value):
        self.charttitle.Left = value

    @property
    def Name(self):
        return self.charttitle.Name

    @property
    def Orientation(self):
        return self.charttitle.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.charttitle.Orientation = value

    @property
    def Parent(self):
        return self.charttitle.Parent

    @property
    def Position(self):
        return self.charttitle.Position

    @Position.setter
    def Position(self, value):
        self.charttitle.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.charttitle.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.charttitle.ReadingOrder = value

    @property
    def Shadow(self):
        return self.charttitle.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.charttitle.Shadow = value

    @property
    def Text(self):
        return self.charttitle.Text

    @Text.setter
    def Text(self, value):
        self.charttitle.Text = value

    @property
    def Top(self):
        return self.charttitle.Top

    @Top.setter
    def Top(self, value):
        self.charttitle.Top = value

    @property
    def VerticalAlignment(self):
        return self.charttitle.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.charttitle.VerticalAlignment = value

    @property
    def Width(self):
        return self.charttitle.Width

    def Delete(self):
        self.charttitle.Delete()

    def Select(self):
        self.charttitle.Select()

class CheckBox:

    def __init__(self, checkbox=None):
        self.checkbox = checkbox

    @property
    def Application(self):
        return Application(self.checkbox.Application)

    @property
    def AutoSize(self):
        return self.checkbox.AutoSize

    @property
    def Creator(self):
        return self.checkbox.Creator

    @property
    def Default(self):
        return self.checkbox.Default

    @Default.setter
    def Default(self, value):
        self.checkbox.Default = value

    @property
    def Parent(self):
        return self.checkbox.Parent

    @property
    def Size(self):
        return self.checkbox.Size

    @Size.setter
    def Size(self, value):
        self.checkbox.Size = value

    @property
    def Valid(self):
        return self.checkbox.Valid

    @property
    def Value(self):
        return self.checkbox.Value

class CoAuthLock:

    def __init__(self, coauthlock=None):
        self.coauthlock = coauthlock

    @property
    def Application(self):
        return self.coauthlock.Application

    @property
    def Creator(self):
        return self.coauthlock.Creator

    @property
    def HeaderFooter(self):
        return self.coauthlock.HeaderFooter

    @property
    def Parent(self):
        return CoAuthLock(self.coauthlock.Parent)

    @property
    def Range(self):
        return self.coauthlock.Range

    @property
    def Type(self):
        return self.coauthlock.Type

    def Unlock(self):
        return self.coauthlock.Unlock()

class CoAuthLocks:

    def __init__(self, coauthlocks=None):
        self.coauthlocks = coauthlocks

    def __call__(self, item):
        return CoAuthLock(self.coauthlocks(item))

    @property
    def Application(self):
        return self.coauthlocks.Application

    @property
    def Count(self):
        return CoAuthLocks(self.coauthlocks.Count)

    @property
    def Creator(self):
        return self.coauthlocks.Creator

    @property
    def Parent(self):
        return CoAuthLocks(self.coauthlocks.Parent)

    def Add(self, *args, Range=None, Type=None):
        arguments = {"Range": Range, "Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CoAuthLock(self.coauthlocks.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CoAuthLock(self.coauthlocks.Item(*args, **arguments))

    def RemoveEphemeralLocks(self):
        return self.coauthlocks.RemoveEphemeralLocks()

class CoAuthor:

    def __init__(self, coauthor=None):
        self.coauthor = coauthor

    @property
    def Application(self):
        return self.coauthor.Application

    @property
    def Creator(self):
        return self.coauthor.Creator

    @property
    def EmailAddress(self):
        return self.coauthor.EmailAddress

    @property
    def ID(self):
        return self.coauthor.ID

    @property
    def IsMe(self):
        return self.coauthor.IsMe

    @property
    def Locks(self):
        return self.coauthor.Locks

    @property
    def Name(self):
        return self.coauthor.Name

    @property
    def Parent(self):
        return self.coauthor.Parent

class CoAuthoring:

    def __init__(self, coauthoring=None):
        self.coauthoring = coauthoring

    @property
    def Application(self):
        return self.coauthoring.Application

    @property
    def Authors(self):
        return CoAuthors(self.coauthoring.Authors)

    @property
    def CanMerge(self):
        return self.coauthoring.CanMerge

    @property
    def CanShare(self):
        return self.coauthoring.CanShare

    @property
    def Conflicts(self):
        return Conflicts(self.coauthoring.Conflicts)

    @property
    def Creator(self):
        return self.coauthoring.Creator

    @property
    def Locks(self):
        return CoAuthLocks(self.coauthoring.Locks)

    @property
    def Me(self):
        return CoAuthor(self.coauthoring.Me)

    @property
    def Parent(self):
        return CoAuthoring(self.coauthoring.Parent)

    @property
    def PendingUpdates(self):
        return self.coauthoring.PendingUpdates

    @property
    def Updates(self):
        return self.coauthoring.Updates

class CoAuthors:

    def __init__(self, coauthors=None):
        self.coauthors = coauthors

    def __call__(self, item):
        return CoAuthor(self.coauthors(item))

    @property
    def Application(self):
        return self.coauthors.Application

    @property
    def Count(self):
        return self.coauthors.Count

    @property
    def Creator(self):
        return self.coauthors.Creator

    @property
    def Parent(self):
        return self.coauthors.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.coauthors.Item(*args, **arguments)

class CoAuthUpdate:

    def __init__(self, coauthupdate=None):
        self.coauthupdate = coauthupdate

    @property
    def Application(self):
        return self.coauthupdate.Application

    @property
    def Creator(self):
        return self.coauthupdate.Creator

    @property
    def Parent(self):
        return self.coauthupdate.Parent

    @property
    def Range(self):
        return self.coauthupdate.Range

class ColorFormat:

    def __init__(self, colorformat=None):
        self.colorformat = colorformat

    @property
    def Application(self):
        return Application(self.colorformat.Application)

    @property
    def Brightness(self):
        return self.colorformat.Brightness

    @property
    def Creator(self):
        return self.colorformat.Creator

    @property
    def ObjectThemeColor(self):
        return WdThemeColorIndex(self.colorformat.ObjectThemeColor)

    @ObjectThemeColor.setter
    def ObjectThemeColor(self, value):
        self.colorformat.ObjectThemeColor = value

    @property
    def Parent(self):
        return self.colorformat.Parent

    @property
    def RGB(self):
        return self.colorformat.RGB

    @RGB.setter
    def RGB(self, value):
        self.colorformat.RGB = value

    @property
    def TintAndShade(self):
        return self.colorformat.TintAndShade

    @property
    def Type(self):
        return self.colorformat.Type

    @Type.setter
    def Type(self, value):
        self.colorformat.Type = value

class Column:

    def __init__(self, column=None):
        self.column = column

    @property
    def Application(self):
        return Application(self.column.Application)

    @property
    def Borders(self):
        return self.column.Borders

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.column.Cells(*args, **arguments)

    @property
    def Creator(self):
        return self.column.Creator

    @property
    def Index(self):
        return self.column.Index

    @property
    def IsFirst(self):
        return self.column.IsFirst

    @property
    def IsLast(self):
        return self.column.IsLast

    @property
    def NestingLevel(self):
        return self.column.NestingLevel

    @property
    def Next(self):
        return self.column.Next

    @property
    def Parent(self):
        return self.column.Parent

    @property
    def PreferredWidth(self):
        return self.column.PreferredWidth

    @PreferredWidth.setter
    def PreferredWidth(self, value):
        self.column.PreferredWidth = value

    @property
    def PreferredWidthType(self):
        return WdPreferredWidthType(self.column.PreferredWidthType)

    @PreferredWidthType.setter
    def PreferredWidthType(self, value):
        self.column.PreferredWidthType = value

    @property
    def Previous(self):
        return self.column.Previous

    @property
    def Shading(self):
        return Shading(self.column.Shading)

    @property
    def Width(self):
        return self.column.Width

    @Width.setter
    def Width(self, value):
        self.column.Width = value

    def AutoFit(self):
        self.column.AutoFit()

    def Delete(self):
        self.column.Delete()

    def Select(self):
        self.column.Select()

    def SetWidth(self, *args, ColumnWidth=None, RulerStyle=None):
        arguments = {"ColumnWidth": ColumnWidth, "RulerStyle": RulerStyle}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.column.SetWidth(*args, **arguments)

    def Sort(self, *args, ExcludeHeader=None, SortFieldType=None, SortOrder=None, CaseSensitive=None, BidiSort=None, IgnoreThe=None, IgnoreKashida=None, IgnoreDiacritics=None, IgnoreHe=None, LanguageID=None):
        arguments = {"ExcludeHeader": ExcludeHeader, "SortFieldType": SortFieldType, "SortOrder": SortOrder, "CaseSensitive": CaseSensitive, "BidiSort": BidiSort, "IgnoreThe": IgnoreThe, "IgnoreKashida": IgnoreKashida, "IgnoreDiacritics": IgnoreDiacritics, "IgnoreHe": IgnoreHe, "LanguageID": LanguageID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.column.Sort(*args, **arguments)

class Comment:

    def __init__(self, comment=None):
        self.comment = comment

    @property
    def Application(self):
        return Application(self.comment.Application)

    @property
    def Creator(self):
        return self.comment.Creator

    @property
    def Date(self):
        return self.comment.Date

    @property
    def Index(self):
        return self.comment.Index

    @property
    def IsInk(self):
        return self.comment.IsInk

    @property
    def Parent(self):
        return self.comment.Parent

    @property
    def Range(self):
        return Range(self.comment.Range)

    @property
    def Reference(self):
        return Range(self.comment.Reference)

    @property
    def Scope(self):
        return Range(self.comment.Scope)

    def Edit(self):
        self.comment.Edit()

class ConditionalStyle:

    def __init__(self, conditionalstyle=None):
        self.conditionalstyle = conditionalstyle

    @property
    def Application(self):
        return Application(self.conditionalstyle.Application)

    @property
    def Borders(self):
        return self.conditionalstyle.Borders

    @property
    def BottomPadding(self):
        return self.conditionalstyle.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.conditionalstyle.BottomPadding = value

    @property
    def Creator(self):
        return self.conditionalstyle.Creator

    @property
    def Font(self):
        return Font(self.conditionalstyle.Font)

    @Font.setter
    def Font(self, value):
        self.conditionalstyle.Font = value

    @property
    def LeftPadding(self):
        return self.conditionalstyle.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.conditionalstyle.LeftPadding = value

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.conditionalstyle.ParagraphFormat)

    @ParagraphFormat.setter
    def ParagraphFormat(self, value):
        self.conditionalstyle.ParagraphFormat = value

    @property
    def Parent(self):
        return self.conditionalstyle.Parent

    @property
    def RightPadding(self):
        return self.conditionalstyle.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.conditionalstyle.RightPadding = value

    @property
    def Shading(self):
        return Shading(self.conditionalstyle.Shading)

    @property
    def TopPadding(self):
        return self.conditionalstyle.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.conditionalstyle.TopPadding = value

class Conflict:

    def __init__(self, conflict=None):
        self.conflict = conflict

    @property
    def Application(self):
        return self.conflict.Application

    @property
    def Creator(self):
        return self.conflict.Creator

    @property
    def Index(self):
        return self.conflict.Index

    @property
    def Parent(self):
        return self.conflict.Parent

    @property
    def Range(self):
        return self.conflict.Range

    @property
    def Type(self):
        return self.conflict.Type

    def Accept(self):
        return self.conflict.Accept()

    def Reject(self):
        return self.conflict.Reject()

class Conflicts:

    def __init__(self, conflicts=None):
        self.conflicts = conflicts

    def __call__(self, item):
        return Conflict(self.conflicts(item))

    @property
    def Application(self):
        return self.conflicts.Application

    @property
    def Count(self):
        return Conflicts(self.conflicts.Count)

    @property
    def Creator(self):
        return self.conflicts.Creator

    @property
    def Parent(self):
        return self.conflicts.Parent

    def AcceptAll(self):
        return self.conflicts.AcceptAll()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.conflicts.Item(*args, **arguments)

    def RejectAll(self):
        return self.conflicts.RejectAll()

class ContentControl:

    def __init__(self, contentcontrol=None):
        self.contentcontrol = contentcontrol

    @property
    def Application(self):
        return Application(self.contentcontrol.Application)

    @property
    def BuildingBlockCategory(self):
        return self.contentcontrol.BuildingBlockCategory

    @BuildingBlockCategory.setter
    def BuildingBlockCategory(self, value):
        self.contentcontrol.BuildingBlockCategory = value

    @property
    def BuildingBlockType(self):
        return WdBuildingBlockTypes(self.contentcontrol.BuildingBlockType)

    @BuildingBlockType.setter
    def BuildingBlockType(self, value):
        self.contentcontrol.BuildingBlockType = value

    @property
    def Checked(self):
        return self.contentcontrol.Checked

    @Checked.setter
    def Checked(self, value):
        self.contentcontrol.Checked = value

    @property
    def Creator(self):
        return self.contentcontrol.Creator

    @property
    def DateCalendarType(self):
        return WdCalendarType(self.contentcontrol.DateCalendarType)

    @DateCalendarType.setter
    def DateCalendarType(self, value):
        self.contentcontrol.DateCalendarType = value

    @property
    def DateDisplayFormat(self):
        return self.contentcontrol.DateDisplayFormat

    @DateDisplayFormat.setter
    def DateDisplayFormat(self, value):
        self.contentcontrol.DateDisplayFormat = value

    @property
    def DateDisplayLocale(self):
        return WdLanguageID(self.contentcontrol.DateDisplayLocale)

    @property
    def DateStorageFormat(self):
        return WdContentControlDateStorageFormat(self.contentcontrol.DateStorageFormat)

    @DateStorageFormat.setter
    def DateStorageFormat(self, value):
        self.contentcontrol.DateStorageFormat = value

    @property
    def DefaultTextStyle(self):
        return self.contentcontrol.DefaultTextStyle

    @DefaultTextStyle.setter
    def DefaultTextStyle(self, value):
        self.contentcontrol.DefaultTextStyle = value

    @property
    def DropdownListEntries(self):
        return ContentControlListEntries(self.contentcontrol.DropdownListEntries)

    @property
    def ID(self):
        return self.contentcontrol.ID

    @property
    def LockContentControl(self):
        return self.contentcontrol.LockContentControl

    @LockContentControl.setter
    def LockContentControl(self, value):
        self.contentcontrol.LockContentControl = value

    @property
    def LockContents(self):
        return self.contentcontrol.LockContents

    @LockContents.setter
    def LockContents(self, value):
        self.contentcontrol.LockContents = value

    @property
    def MultiLine(self):
        return self.contentcontrol.MultiLine

    @property
    def Parent(self):
        return self.contentcontrol.Parent

    @property
    def ParentContentControl(self):
        return ContentControl(self.contentcontrol.ParentContentControl)

    @property
    def PlaceholderText(self):
        return BuildingBlock(self.contentcontrol.PlaceholderText)

    @property
    def Range(self):
        return Range(self.contentcontrol.Range)

    @property
    def ShowingPlaceholderText(self):
        return self.contentcontrol.ShowingPlaceholderText

    @property
    def Tag(self):
        return self.contentcontrol.Tag

    @Tag.setter
    def Tag(self, value):
        self.contentcontrol.Tag = value

    @property
    def Temporary(self):
        return self.contentcontrol.Temporary

    @Temporary.setter
    def Temporary(self, value):
        self.contentcontrol.Temporary = value

    @property
    def Title(self):
        return self.contentcontrol.Title

    @Title.setter
    def Title(self, value):
        self.contentcontrol.Title = value

    @property
    def Type(self):
        return WdContentControlType(self.contentcontrol.Type)

    @Type.setter
    def Type(self, value):
        self.contentcontrol.Type = value

    @property
    def XMLMapping(self):
        return XMLMapping(self.contentcontrol.XMLMapping)

    def Copy(self):
        self.contentcontrol.Copy()

    def Cut(self):
        self.contentcontrol.Cut()

    def Delete(self, *args, DeleteContents=None):
        arguments = {"DeleteContents": DeleteContents}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contentcontrol.Delete(*args, **arguments)

    def SetCheckedSymbol(self, *args, CharacterNumber=None, Font=None):
        arguments = {"CharacterNumber": CharacterNumber, "Font": Font}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contentcontrol.SetCheckedSymbol(*args, **arguments)

    def SetPlaceholderText(self, *args, BuildingBlock=None, Range=None, Text=None):
        arguments = {"BuildingBlock": BuildingBlock, "Range": Range, "Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contentcontrol.SetPlaceholderText(*args, **arguments)

    def SetUncheckedSymbol(self, *args, CharacterNumber=None, Font=None):
        arguments = {"CharacterNumber": CharacterNumber, "Font": Font}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.contentcontrol.SetUncheckedSymbol(*args, **arguments)

    def Ungroup(self):
        self.contentcontrol.Ungroup()

class ContentControlListEntries:

    def __init__(self, contentcontrollistentries=None):
        self.contentcontrollistentries = contentcontrollistentries

    @property
    def Application(self):
        return Application(self.contentcontrollistentries.Application)

    @property
    def Count(self):
        return ContentControlListEntries(self.contentcontrollistentries.Count)

    @property
    def Creator(self):
        return self.contentcontrollistentries.Creator

    @property
    def Parent(self):
        return self.contentcontrollistentries.Parent

    def Add(self, *args, Text=None, Value=None, Index=None):
        arguments = {"Text": Text, "Value": Value, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.contentcontrollistentries.Add(*args, **arguments)

    def Clear(self):
        self.contentcontrollistentries.Clear()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.contentcontrollistentries.Item(*args, **arguments)

class ContentControlListEntry:

    def __init__(self, contentcontrollistentry=None):
        self.contentcontrollistentry = contentcontrollistentry

    @property
    def Application(self):
        return Application(self.contentcontrollistentry.Application)

    @property
    def Creator(self):
        return self.contentcontrollistentry.Creator

    @property
    def Index(self):
        return self.contentcontrollistentry.Index

    @Index.setter
    def Index(self, value):
        self.contentcontrollistentry.Index = value

    @property
    def Parent(self):
        return self.contentcontrollistentry.Parent

    @property
    def Text(self):
        return self.contentcontrollistentry.Text

    @Text.setter
    def Text(self, value):
        self.contentcontrollistentry.Text = value

    @property
    def Value(self):
        return self.contentcontrollistentry.Value

    @Value.setter
    def Value(self, value):
        self.contentcontrollistentry.Value = value

    def Delete(self):
        self.contentcontrollistentry.Delete()

    def MoveDown(self):
        self.contentcontrollistentry.MoveDown()

    def MoveUp(self):
        self.contentcontrollistentry.MoveUp()

    def Select(self):
        self.contentcontrollistentry.Select()

class ContentControls:

    def __init__(self, contentcontrols=None):
        self.contentcontrols = contentcontrols

    def __call__(self, item):
        return ContentControl(self.contentcontrols(item))

    @property
    def Application(self):
        return Application(self.contentcontrols.Application)

    @property
    def Count(self):
        return ContentControls(self.contentcontrols.Count)

    @property
    def Creator(self):
        return self.contentcontrols.Creator

    @property
    def Parent(self):
        return self.contentcontrols.Parent

    def Add(self, *args, Type=None, Range=None):
        arguments = {"Type": Type, "Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ContentControl(self.contentcontrols.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.contentcontrols.Item(*args, **arguments)

class CustomLabel:

    def __init__(self, customlabel=None):
        self.customlabel = customlabel

    @property
    def Application(self):
        return Application(self.customlabel.Application)

    @property
    def Creator(self):
        return self.customlabel.Creator

    @property
    def DotMatrix(self):
        return self.customlabel.DotMatrix

    @property
    def Height(self):
        return self.customlabel.Height

    @Height.setter
    def Height(self, value):
        self.customlabel.Height = value

    @property
    def HorizontalPitch(self):
        return self.customlabel.HorizontalPitch

    @HorizontalPitch.setter
    def HorizontalPitch(self, value):
        self.customlabel.HorizontalPitch = value

    @property
    def Index(self):
        return self.customlabel.Index

    @property
    def Name(self):
        return CustomLabel(self.customlabel.Name)

    @Name.setter
    def Name(self, value):
        self.customlabel.Name = value

    @property
    def NumberAcross(self):
        return self.customlabel.NumberAcross

    @NumberAcross.setter
    def NumberAcross(self, value):
        self.customlabel.NumberAcross = value

    @property
    def NumberDown(self):
        return self.customlabel.NumberDown

    @NumberDown.setter
    def NumberDown(self, value):
        self.customlabel.NumberDown = value

    @property
    def PageSize(self):
        return WdCustomLabelPageSize(self.customlabel.PageSize)

    @PageSize.setter
    def PageSize(self, value):
        self.customlabel.PageSize = value

    @property
    def Parent(self):
        return self.customlabel.Parent

    @property
    def SideMargin(self):
        return self.customlabel.SideMargin

    @SideMargin.setter
    def SideMargin(self, value):
        self.customlabel.SideMargin = value

    @property
    def TopMargin(self):
        return self.customlabel.TopMargin

    @TopMargin.setter
    def TopMargin(self, value):
        self.customlabel.TopMargin = value

    @property
    def Valid(self):
        return self.customlabel.Valid

    @property
    def VerticalPitch(self):
        return self.customlabel.VerticalPitch

    @VerticalPitch.setter
    def VerticalPitch(self, value):
        self.customlabel.VerticalPitch = value

    @property
    def Width(self):
        return self.customlabel.Width

    @Width.setter
    def Width(self, value):
        self.customlabel.Width = value

    def Delete(self):
        self.customlabel.Delete()

class CustomProperties:

    def __init__(self, customproperties=None):
        self.customproperties = customproperties

    def __call__(self, item):
        return CustomPropertie(self.customproperties(item))

    @property
    def Application(self):
        return Application(self.customproperties.Application)

    @property
    def Count(self):
        return self.customproperties.Count

    @property
    def Creator(self):
        return self.customproperties.Creator

    @property
    def Parent(self):
        return self.customproperties.Parent

    def Add(self, *args, Name=None, Value=None):
        arguments = {"Name": Name, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return CustomPropertie(self.customproperties.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.customproperties.Item(*args, **arguments)

class CustomProperty:

    def __init__(self, customproperty=None):
        self.customproperty = customproperty

    @property
    def Application(self):
        return Application(self.customproperty.Application)

    @property
    def Creator(self):
        return self.customproperty.Creator

    @property
    def Name(self):
        return self.customproperty.Name

    @property
    def Parent(self):
        return self.customproperty.Parent

    @property
    def Value(self):
        return self.customproperty.Value

    @Value.setter
    def Value(self, value):
        self.customproperty.Value = value

    def Delete(self):
        self.customproperty.Delete()

class DataLabel:

    def __init__(self, datalabel=None):
        self.datalabel = datalabel

    @property
    def Application(self):
        return self.datalabel.Application

    @property
    def AutoText(self):
        return self.datalabel.AutoText

    @property
    def Caption(self):
        return self.datalabel.Caption

    @Caption.setter
    def Caption(self, value):
        self.datalabel.Caption = value

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ChartCharacters(self.datalabel.Characters(*args, **arguments))

    @property
    def Creator(self):
        return self.datalabel.Creator

    @property
    def Format(self):
        return ChartFormat(self.datalabel.Format)

    @property
    def Formula(self):
        return self.datalabel.Formula

    @property
    def FormulaLocal(self):
        return self.datalabel.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.datalabel.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.datalabel.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.datalabel.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.datalabel.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.datalabel.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.datalabel.Height

    @property
    def HorizontalAlignment(self):
        return self.datalabel.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.datalabel.HorizontalAlignment = value

    @property
    def Left(self):
        return self.datalabel.Left

    @Left.setter
    def Left(self, value):
        self.datalabel.Left = value

    @property
    def Name(self):
        return self.datalabel.Name

    @property
    def NumberFormat(self):
        return self.datalabel.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.datalabel.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.datalabel.NumberFormatLinked

    @property
    def NumberFormatLocal(self):
        return self.datalabel.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.datalabel.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.datalabel.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.datalabel.Orientation = value

    @property
    def Parent(self):
        return self.datalabel.Parent

    @property
    def Position(self):
        return self.datalabel.Position

    @Position.setter
    def Position(self, value):
        self.datalabel.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.datalabel.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.datalabel.ReadingOrder = value

    @property
    def Separator(self):
        return self.datalabel.Separator

    @Separator.setter
    def Separator(self, value):
        self.datalabel.Separator = value

    @property
    def Shadow(self):
        return self.datalabel.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.datalabel.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.datalabel.ShowBubbleSize

    @property
    def ShowCategoryName(self):
        return self.datalabel.ShowCategoryName

    @property
    def ShowLegendKey(self):
        return self.datalabel.ShowLegendKey

    @property
    def ShowPercentage(self):
        return self.datalabel.ShowPercentage

    @property
    def ShowSeriesName(self):
        return self.datalabel.ShowSeriesName

    @property
    def ShowValue(self):
        return self.datalabel.ShowValue

    @property
    def Text(self):
        return self.datalabel.Text

    @Text.setter
    def Text(self, value):
        self.datalabel.Text = value

    @property
    def Top(self):
        return self.datalabel.Top

    @Top.setter
    def Top(self, value):
        self.datalabel.Top = value

    @property
    def VerticalAlignment(self):
        return self.datalabel.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.datalabel.VerticalAlignment = value

    @property
    def Width(self):
        return self.datalabel.Width

    def Delete(self):
        self.datalabel.Delete()

    def Select(self):
        self.datalabel.Select()

class DataLabels:

    def __init__(self, datalabels=None):
        self.datalabels = datalabels

    def __call__(self, item):
        return DataLabel(self.datalabels(item))

    @property
    def Application(self):
        return self.datalabels.Application

    @property
    def AutoText(self):
        return self.datalabels.AutoText

    @property
    def Count(self):
        return self.datalabels.Count

    @property
    def Creator(self):
        return self.datalabels.Creator

    @property
    def Format(self):
        return ChartFormat(self.datalabels.Format)

    @property
    def HorizontalAlignment(self):
        return self.datalabels.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.datalabels.HorizontalAlignment = value

    @property
    def Name(self):
        return self.datalabels.Name

    @property
    def NumberFormat(self):
        return self.datalabels.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.datalabels.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.datalabels.NumberFormatLinked

    @property
    def NumberFormatLocal(self):
        return self.datalabels.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.datalabels.NumberFormatLocal = value

    @property
    def Orientation(self):
        return self.datalabels.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.datalabels.Orientation = value

    @property
    def Parent(self):
        return self.datalabels.Parent

    @property
    def Position(self):
        return self.datalabels.Position

    @Position.setter
    def Position(self, value):
        self.datalabels.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.datalabels.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.datalabels.ReadingOrder = value

    @property
    def Separator(self):
        return self.datalabels.Separator

    @property
    def Shadow(self):
        return self.datalabels.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.datalabels.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.datalabels.ShowBubbleSize

    @property
    def ShowCategoryName(self):
        return self.datalabels.ShowCategoryName

    @property
    def ShowLegendKey(self):
        return self.datalabels.ShowLegendKey

    @property
    def ShowPercentage(self):
        return self.datalabels.ShowPercentage

    @property
    def ShowSeriesName(self):
        return self.datalabels.ShowSeriesName

    @property
    def ShowValue(self):
        return self.datalabels.ShowValue

    @property
    def VerticalAlignment(self):
        return self.datalabels.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.datalabels.VerticalAlignment = value

    def Delete(self):
        self.datalabels.Delete()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return DataLabel(self.datalabels.Item(*args, **arguments))

    def Select(self):
        self.datalabels.Select()

class DataTable:

    def __init__(self, datatable=None):
        self.datatable = datatable

    @property
    def Application(self):
        return self.datatable.Application

    @property
    def Border(self):
        return ChartBorder(self.datatable.Border)

    @property
    def Creator(self):
        return self.datatable.Creator

    @property
    def Font(self):
        return ChartFont(self.datatable.Font)

    @property
    def Format(self):
        return ChartFormat(self.datatable.Format)

    @property
    def HasBorderHorizontal(self):
        return self.datatable.HasBorderHorizontal

    @property
    def HasBorderOutline(self):
        return self.datatable.HasBorderOutline

    @property
    def HasBorderVertical(self):
        return self.datatable.HasBorderVertical

    @property
    def Parent(self):
        return self.datatable.Parent

    @property
    def ShowLegendKey(self):
        return self.datatable.ShowLegendKey

    def Delete(self):
        self.datatable.Delete()

    def Select(self):
        self.datatable.Select()

class DefaultWebOptions:

    def __init__(self, defaultweboptions=None):
        self.defaultweboptions = defaultweboptions

    @property
    def AllowPNG(self):
        return self.defaultweboptions.AllowPNG

    @property
    def AlwaysSaveInDefaultEncoding(self):
        return self.defaultweboptions.AlwaysSaveInDefaultEncoding

    @property
    def Application(self):
        return Application(self.defaultweboptions.Application)

    @property
    def BrowserLevel(self):
        return WdBrowserLevel(self.defaultweboptions.BrowserLevel)

    @BrowserLevel.setter
    def BrowserLevel(self, value):
        self.defaultweboptions.BrowserLevel = value

    @property
    def CheckIfOfficeIsHTMLEditor(self):
        return self.defaultweboptions.CheckIfOfficeIsHTMLEditor

    @property
    def CheckIfWordIsDefaultHTMLEditor(self):
        return self.defaultweboptions.CheckIfWordIsDefaultHTMLEditor

    @property
    def Creator(self):
        return self.defaultweboptions.Creator

    @property
    def Encoding(self):
        return self.defaultweboptions.Encoding

    @Encoding.setter
    def Encoding(self, value):
        self.defaultweboptions.Encoding = value

    @property
    def FolderSuffix(self):
        return self.defaultweboptions.FolderSuffix

    @property
    def Fonts(self):
        return self.defaultweboptions.Fonts

    @property
    def OptimizeForBrowser(self):
        return self.defaultweboptions.OptimizeForBrowser

    @property
    def OrganizeInFolder(self):
        return self.defaultweboptions.OrganizeInFolder

    @property
    def Parent(self):
        return self.defaultweboptions.Parent

    @property
    def PixelsPerInch(self):
        return self.defaultweboptions.PixelsPerInch

    @PixelsPerInch.setter
    def PixelsPerInch(self, value):
        self.defaultweboptions.PixelsPerInch = value

    @property
    def RelyOnCSS(self):
        return self.defaultweboptions.RelyOnCSS

    @property
    def RelyOnVML(self):
        return self.defaultweboptions.RelyOnVML

    @property
    def SaveNewWebPagesAsWebArchives(self):
        return self.defaultweboptions.SaveNewWebPagesAsWebArchives

    @property
    def ScreenSize(self):
        return self.defaultweboptions.ScreenSize

    @ScreenSize.setter
    def ScreenSize(self, value):
        self.defaultweboptions.ScreenSize = value

    @property
    def TargetBrowser(self):
        return self.defaultweboptions.TargetBrowser

    @property
    def UpdateLinksOnSave(self):
        return self.defaultweboptions.UpdateLinksOnSave

    @property
    def UseLongFileNames(self):
        return self.defaultweboptions.UseLongFileNames

class Dialog:

    def __init__(self, dialog=None):
        self.dialog = dialog

    @property
    def Application(self):
        return Application(self.dialog.Application)

    @property
    def CommandBarId(self):
        return self.dialog.CommandBarId

    @property
    def CommandName(self):
        return self.dialog.CommandName

    @property
    def Creator(self):
        return self.dialog.Creator

    @property
    def DefaultTab(self):
        return WdWordDialogTab(self.dialog.DefaultTab)

    @DefaultTab.setter
    def DefaultTab(self, value):
        self.dialog.DefaultTab = value

    @property
    def Parent(self):
        return self.dialog.Parent

    @property
    def Type(self):
        return WdWordDialog(self.dialog.Type)

    def Display(self, *args, TimeOut=None):
        arguments = {"TimeOut": TimeOut}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.dialog.Display(*args, **arguments)

    def Execute(self):
        self.dialog.Execute()

    def Show(self, *args, TimeOut=None):
        arguments = {"TimeOut": TimeOut}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.dialog.Show(*args, **arguments)

    def Update(self):
        self.dialog.Update()

class Dictionary:

    def __init__(self, dictionary=None):
        self.dictionary = dictionary

    @property
    def Application(self):
        return Application(self.dictionary.Application)

    @property
    def Creator(self):
        return self.dictionary.Creator

    @property
    def LanguageID(self):
        return WdLanguageID(self.dictionary.LanguageID)

    @LanguageID.setter
    def LanguageID(self, value):
        self.dictionary.LanguageID = value

    @property
    def LanguageSpecific(self):
        return self.dictionary.LanguageSpecific

    @property
    def Name(self):
        return self.dictionary.Name

    @property
    def Parent(self):
        return self.dictionary.Parent

    @property
    def Path(self):
        return self.dictionary.Path

    @property
    def ReadOnly(self):
        return self.dictionary.ReadOnly

    @property
    def Type(self):
        return WdDictionaryType(self.dictionary.Type)

    def Delete(self):
        self.dictionary.Delete()

class DisplayUnitLabel:

    def __init__(self, displayunitlabel=None):
        self.displayunitlabel = displayunitlabel

    @property
    def Application(self):
        return self.displayunitlabel.Application

    @property
    def Caption(self):
        return self.displayunitlabel.Caption

    @Caption.setter
    def Caption(self, value):
        self.displayunitlabel.Caption = value

    def Characters(self, *args, Start=None, Length=None):
        arguments = {"Start": Start, "Length": Length}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ChartCharacters(self.displayunitlabel.Characters(*args, **arguments))

    @property
    def Creator(self):
        return self.displayunitlabel.Creator

    @property
    def Format(self):
        return ChartFormat(self.displayunitlabel.Format)

    @property
    def Formula(self):
        return self.displayunitlabel.Formula

    @property
    def FormulaLocal(self):
        return self.displayunitlabel.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.displayunitlabel.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.displayunitlabel.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.displayunitlabel.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.displayunitlabel.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.displayunitlabel.FormulaR1C1Local = value

    @property
    def Height(self):
        return self.displayunitlabel.Height

    @property
    def HorizontalAlignment(self):
        return self.displayunitlabel.HorizontalAlignment

    @HorizontalAlignment.setter
    def HorizontalAlignment(self, value):
        self.displayunitlabel.HorizontalAlignment = value

    @property
    def IncludeInLayout(self):
        return self.displayunitlabel.IncludeInLayout

    @property
    def Left(self):
        return self.displayunitlabel.Left

    @Left.setter
    def Left(self, value):
        self.displayunitlabel.Left = value

    @property
    def Name(self):
        return self.displayunitlabel.Name

    @property
    def Orientation(self):
        return self.displayunitlabel.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.displayunitlabel.Orientation = value

    @property
    def Parent(self):
        return self.displayunitlabel.Parent

    @property
    def Position(self):
        return self.displayunitlabel.Position

    @Position.setter
    def Position(self, value):
        self.displayunitlabel.Position = value

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.displayunitlabel.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.displayunitlabel.ReadingOrder = value

    @property
    def Shadow(self):
        return self.displayunitlabel.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.displayunitlabel.Shadow = value

    @property
    def Text(self):
        return self.displayunitlabel.Text

    @Text.setter
    def Text(self, value):
        self.displayunitlabel.Text = value

    @property
    def Top(self):
        return self.displayunitlabel.Top

    @Top.setter
    def Top(self, value):
        self.displayunitlabel.Top = value

    @property
    def VerticalAlignment(self):
        return self.displayunitlabel.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.displayunitlabel.VerticalAlignment = value

    @property
    def Width(self):
        return self.displayunitlabel.Width

    def Delete(self):
        self.displayunitlabel.Delete()

    def Select(self):
        self.displayunitlabel.Select()

class Document:

    def __init__(self, document=None):
        self.document = document

    @property
    def ActiveTheme(self):
        return self.document.ActiveTheme

    @property
    def ActiveThemeDisplayName(self):
        return self.document.ActiveThemeDisplayName

    @property
    def ActiveWindow(self):
        return Window(self.document.ActiveWindow)

    @property
    def ActiveWritingStyle(self):
        return self.document.ActiveWritingStyle

    @ActiveWritingStyle.setter
    def ActiveWritingStyle(self, value):
        self.document.ActiveWritingStyle = value

    @property
    def Application(self):
        return Application(self.document.Application)

    @property
    def AttachedTemplate(self):
        return Template(self.document.AttachedTemplate)

    @property
    def AutoFormatOverride(self):
        return self.document.AutoFormatOverride

    @AutoFormatOverride.setter
    def AutoFormatOverride(self, value):
        self.document.AutoFormatOverride = value

    @property
    def AutoHyphenation(self):
        return self.document.AutoHyphenation

    @property
    def AutoSaveOn(self):
        return self.document.AutoSaveOn

    @property
    def Background(self):
        return Shape(self.document.Background)

    @property
    def Bibliography(self):
        return Bibliography(self.document.Bibliography)

    @property
    def Bookmarks(self):
        return self.document.Bookmarks

    @property
    def BuiltInDocumentProperties(self):
        return self.document.BuiltInDocumentProperties

    @property
    def Characters(self):
        return self.document.Characters

    @property
    def ClickAndTypeParagraphStyle(self):
        return self.document.ClickAndTypeParagraphStyle

    @ClickAndTypeParagraphStyle.setter
    def ClickAndTypeParagraphStyle(self, value):
        self.document.ClickAndTypeParagraphStyle = value

    @property
    def CoAuthoring(self):
        return self.document.CoAuthoring

    @property
    def CodeName(self):
        return self.document.CodeName

    @property
    def CommandBars(self):
        return self.document.CommandBars

    @property
    def Comments(self):
        return self.document.Comments

    def Compatibility(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.Compatibility(*args, **arguments)

    @property
    def CompatibilityMode(self):
        return self.document.CompatibilityMode

    @property
    def ConsecutiveHyphensLimit(self):
        return self.document.ConsecutiveHyphensLimit

    @ConsecutiveHyphensLimit.setter
    def ConsecutiveHyphensLimit(self, value):
        self.document.ConsecutiveHyphensLimit = value

    @property
    def Container(self):
        return self.document.Container

    @property
    def Content(self):
        return Range(self.document.Content)

    @property
    def ContentControls(self):
        return ContentControls(self.document.ContentControls)

    @property
    def ContentTypeProperties(self):
        return self.document.ContentTypeProperties

    @property
    def Creator(self):
        return self.document.Creator

    @property
    def CurrentRsid(self):
        return self.document.CurrentRsid

    @property
    def CustomDocumentProperties(self):
        return self.document.CustomDocumentProperties

    @property
    def CustomXMLParts(self):
        return self.document.CustomXMLParts

    @property
    def DefaultTableStyle(self):
        return self.document.DefaultTableStyle

    @property
    def DefaultTabStop(self):
        return self.document.DefaultTabStop

    @DefaultTabStop.setter
    def DefaultTabStop(self, value):
        self.document.DefaultTabStop = value

    @property
    def DefaultTargetFrame(self):
        return self.document.DefaultTargetFrame

    @DefaultTargetFrame.setter
    def DefaultTargetFrame(self, value):
        self.document.DefaultTargetFrame = value

    @property
    def DisableFeatures(self):
        return self.document.DisableFeatures

    @property
    def DisableFeaturesIntroducedAfter(self):
        return self.document.DisableFeaturesIntroducedAfter

    @property
    def DocumentInspectors(self):
        return self.document.DocumentInspectors

    @property
    def DocumentLibraryVersions(self):
        return self.document.DocumentLibraryVersions

    @property
    def DocumentTheme(self):
        return self.document.DocumentTheme

    @property
    def DoNotEmbedSystemFonts(self):
        return self.document.DoNotEmbedSystemFonts

    @property
    def Email(self):
        return Email(self.document.Email)

    @property
    def EmbedLinguisticData(self):
        return self.document.EmbedLinguisticData

    @property
    def EmbedTrueTypeFonts(self):
        return self.document.EmbedTrueTypeFonts

    @property
    def EncryptionProvider(self):
        return self.document.EncryptionProvider

    @property
    def Endnotes(self):
        return self.document.Endnotes

    @property
    def EnforceStyle(self):
        return self.document.EnforceStyle

    @EnforceStyle.setter
    def EnforceStyle(self, value):
        self.document.EnforceStyle = value

    @property
    def Envelope(self):
        return Envelope(self.document.Envelope)

    @property
    def FarEastLineBreakLanguage(self):
        return WdFarEastLineBreakLanguageID(self.document.FarEastLineBreakLanguage)

    @FarEastLineBreakLanguage.setter
    def FarEastLineBreakLanguage(self, value):
        self.document.FarEastLineBreakLanguage = value

    @property
    def FarEastLineBreakLevel(self):
        return WdFarEastLineBreakLevel(self.document.FarEastLineBreakLevel)

    @FarEastLineBreakLevel.setter
    def FarEastLineBreakLevel(self, value):
        self.document.FarEastLineBreakLevel = value

    @property
    def Fields(self):
        return self.document.Fields

    @property
    def Final(self):
        return self.document.Final

    @Final.setter
    def Final(self, value):
        self.document.Final = value

    @property
    def Footnotes(self):
        return self.document.Footnotes

    @property
    def FormattingShowClear(self):
        return self.document.FormattingShowClear

    @property
    def FormattingShowFilter(self):
        return self.document.FormattingShowFilter

    @property
    def FormattingShowFont(self):
        return self.document.FormattingShowFont

    @property
    def FormattingShowNextLevel(self):
        return self.document.FormattingShowNextLevel

    @FormattingShowNextLevel.setter
    def FormattingShowNextLevel(self, value):
        self.document.FormattingShowNextLevel = value

    @property
    def FormattingShowNumbering(self):
        return self.document.FormattingShowNumbering

    @property
    def FormattingShowParagraph(self):
        return self.document.FormattingShowParagraph

    @property
    def FormattingShowUserStyleName(self):
        return self.document.FormattingShowUserStyleName

    @FormattingShowUserStyleName.setter
    def FormattingShowUserStyleName(self, value):
        self.document.FormattingShowUserStyleName = value

    @property
    def FormFields(self):
        return self.document.FormFields

    @property
    def FormsDesign(self):
        return self.document.FormsDesign

    @property
    def Frames(self):
        return Frames(self.document.Frames)

    @property
    def Frameset(self):
        return Frameset(self.document.Frameset)

    @property
    def FullName(self):
        return self.document.FullName

    @property
    def GrammarChecked(self):
        return self.document.GrammarChecked

    @property
    def GrammaticalErrors(self):
        return self.document.GrammaticalErrors

    @property
    def GridDistanceHorizontal(self):
        return self.document.GridDistanceHorizontal

    @GridDistanceHorizontal.setter
    def GridDistanceHorizontal(self, value):
        self.document.GridDistanceHorizontal = value

    @property
    def GridDistanceVertical(self):
        return self.document.GridDistanceVertical

    @GridDistanceVertical.setter
    def GridDistanceVertical(self, value):
        self.document.GridDistanceVertical = value

    @property
    def GridOriginFromMargin(self):
        return self.document.GridOriginFromMargin

    @property
    def GridOriginHorizontal(self):
        return self.document.GridOriginHorizontal

    @GridOriginHorizontal.setter
    def GridOriginHorizontal(self, value):
        self.document.GridOriginHorizontal = value

    @property
    def GridOriginVertical(self):
        return self.document.GridOriginVertical

    @GridOriginVertical.setter
    def GridOriginVertical(self, value):
        self.document.GridOriginVertical = value

    @property
    def GridSpaceBetweenHorizontalLines(self):
        return self.document.GridSpaceBetweenHorizontalLines

    @GridSpaceBetweenHorizontalLines.setter
    def GridSpaceBetweenHorizontalLines(self, value):
        self.document.GridSpaceBetweenHorizontalLines = value

    @property
    def GridSpaceBetweenVerticalLines(self):
        return self.document.GridSpaceBetweenVerticalLines

    @GridSpaceBetweenVerticalLines.setter
    def GridSpaceBetweenVerticalLines(self, value):
        self.document.GridSpaceBetweenVerticalLines = value

    @property
    def HasPassword(self):
        return self.document.HasPassword

    @property
    def HasVBProject(self):
        return self.document.HasVBProject

    @property
    def HTMLDivisions(self):
        return HTMLDivisions(self.document.HTMLDivisions)

    @property
    def Hyperlinks(self):
        return self.document.Hyperlinks

    @property
    def HyphenateCaps(self):
        return self.document.HyphenateCaps

    @property
    def HyphenationZone(self):
        return self.document.HyphenationZone

    @HyphenationZone.setter
    def HyphenationZone(self, value):
        self.document.HyphenationZone = value

    @property
    def Indexes(self):
        return self.document.Indexes

    @property
    def InlineShapes(self):
        return self.document.InlineShapes

    @property
    def IsMasterDocument(self):
        return self.document.IsMasterDocument

    @property
    def IsSubdocument(self):
        return self.document.IsSubdocument

    @property
    def JustificationMode(self):
        return WdJustificationMode(self.document.JustificationMode)

    @JustificationMode.setter
    def JustificationMode(self, value):
        self.document.JustificationMode = value

    @property
    def KerningByAlgorithm(self):
        return self.document.KerningByAlgorithm

    @property
    def Kind(self):
        return WdDocumentKind(self.document.Kind)

    @Kind.setter
    def Kind(self, value):
        self.document.Kind = value

    @property
    def LanguageDetected(self):
        return self.document.LanguageDetected

    @LanguageDetected.setter
    def LanguageDetected(self, value):
        self.document.LanguageDetected = value

    @property
    def ListParagraphs(self):
        return self.document.ListParagraphs

    @property
    def Lists(self):
        return self.document.Lists

    @property
    def ListTemplates(self):
        return self.document.ListTemplates

    @property
    def LockQuickStyleSet(self):
        return self.document.LockQuickStyleSet

    @LockQuickStyleSet.setter
    def LockQuickStyleSet(self, value):
        self.document.LockQuickStyleSet = value

    @property
    def LockTheme(self):
        return self.document.LockTheme

    @LockTheme.setter
    def LockTheme(self, value):
        self.document.LockTheme = value

    @property
    def MailEnvelope(self):
        return self.document.MailEnvelope

    @property
    def MailMerge(self):
        return MailMerge(self.document.MailMerge)

    @property
    def Name(self):
        return self.document.Name

    @property
    def NoLineBreakAfter(self):
        return self.document.NoLineBreakAfter

    @NoLineBreakAfter.setter
    def NoLineBreakAfter(self, value):
        self.document.NoLineBreakAfter = value

    @property
    def NoLineBreakBefore(self):
        return self.document.NoLineBreakBefore

    @NoLineBreakBefore.setter
    def NoLineBreakBefore(self, value):
        self.document.NoLineBreakBefore = value

    @property
    def OMathBreakBin(self):
        return WdOMathBreakBin(self.document.OMathBreakBin)

    @OMathBreakBin.setter
    def OMathBreakBin(self, value):
        self.document.OMathBreakBin = value

    @property
    def OMathBreakSub(self):
        return WdOMathBreakSub(self.document.OMathBreakSub)

    @OMathBreakSub.setter
    def OMathBreakSub(self, value):
        self.document.OMathBreakSub = value

    @property
    def OMathFontName(self):
        return self.document.OMathFontName

    @OMathFontName.setter
    def OMathFontName(self, value):
        self.document.OMathFontName = value

    @property
    def OMathIntSubSupLim(self):
        return self.document.OMathIntSubSupLim

    @OMathIntSubSupLim.setter
    def OMathIntSubSupLim(self, value):
        self.document.OMathIntSubSupLim = value

    @property
    def OMathJc(self):
        return WdOMathJc(self.document.OMathJc)

    @OMathJc.setter
    def OMathJc(self, value):
        self.document.OMathJc = value

    @property
    def OMathLeftMargin(self):
        return self.document.OMathLeftMargin

    @OMathLeftMargin.setter
    def OMathLeftMargin(self, value):
        self.document.OMathLeftMargin = value

    @property
    def OMathNarySupSubLim(self):
        return self.document.OMathNarySupSubLim

    @OMathNarySupSubLim.setter
    def OMathNarySupSubLim(self, value):
        self.document.OMathNarySupSubLim = value

    @property
    def OMathRightMargin(self):
        return self.document.OMathRightMargin

    @OMathRightMargin.setter
    def OMathRightMargin(self, value):
        self.document.OMathRightMargin = value

    @property
    def OMaths(self):
        return OMaths(self.document.OMaths)

    @property
    def OMathSmallFrac(self):
        return self.document.OMathSmallFrac

    @OMathSmallFrac.setter
    def OMathSmallFrac(self, value):
        self.document.OMathSmallFrac = value

    @property
    def OMathWrap(self):
        return self.document.OMathWrap

    @OMathWrap.setter
    def OMathWrap(self, value):
        self.document.OMathWrap = value

    @property
    def OpenEncoding(self):
        return self.document.OpenEncoding

    @property
    def OptimizeForWord97(self):
        return self.document.OptimizeForWord97

    @property
    def OriginalDocumentTitle(self):
        return self.document.OriginalDocumentTitle

    @property
    def PageSetup(self):
        return PageSetup(self.document.PageSetup)

    @property
    def Paragraphs(self):
        return self.document.Paragraphs

    @property
    def Parent(self):
        return self.document.Parent

    @property
    def Password(self):
        return self.document.Password

    @property
    def PasswordEncryptionAlgorithm(self):
        return self.document.PasswordEncryptionAlgorithm

    @property
    def PasswordEncryptionFileProperties(self):
        return self.document.PasswordEncryptionFileProperties

    @property
    def PasswordEncryptionKeyLength(self):
        return self.document.PasswordEncryptionKeyLength

    @property
    def PasswordEncryptionProvider(self):
        return self.document.PasswordEncryptionProvider

    @property
    def Path(self):
        return self.document.Path

    @property
    def Permission(self):
        return self.document.Permission

    @property
    def PrintFormsData(self):
        return self.document.PrintFormsData

    @property
    def PrintPostScriptOverText(self):
        return self.document.PrintPostScriptOverText

    @property
    def PrintRevisions(self):
        return self.document.PrintRevisions

    @property
    def ProtectionType(self):
        return WdProtectionType(self.document.ProtectionType)

    @property
    def ReadabilityStatistics(self):
        return self.document.ReadabilityStatistics

    @property
    def ReadingLayoutSizeX(self):
        return self.document.ReadingLayoutSizeX

    @property
    def ReadingLayoutSizeY(self):
        return self.document.ReadingLayoutSizeY

    @property
    def ReadingModeLayoutFrozen(self):
        return self.document.ReadingModeLayoutFrozen

    @property
    def ReadOnly(self):
        return self.document.ReadOnly

    @property
    def ReadOnlyRecommended(self):
        return self.document.ReadOnlyRecommended

    @property
    def RemoveDateAndTime(self):
        return self.document.RemoveDateAndTime

    @property
    def RemovePersonalInformation(self):
        return self.document.RemovePersonalInformation

    @property
    def Research(self):
        return Research(self.document.Research)

    @property
    def RevisedDocumentTitle(self):
        return self.document.RevisedDocumentTitle

    @property
    def Revisions(self):
        return self.document.Revisions

    @property
    def Saved(self):
        return self.document.Saved

    @property
    def SaveEncoding(self):
        return self.document.SaveEncoding

    @SaveEncoding.setter
    def SaveEncoding(self, value):
        self.document.SaveEncoding = value

    @property
    def SaveFormat(self):
        return self.document.SaveFormat

    @property
    def SaveFormsData(self):
        return self.document.SaveFormsData

    @property
    def SaveSubsetFonts(self):
        return self.document.SaveSubsetFonts

    @property
    def Scripts(self):
        return self.document.Scripts

    @property
    def Sections(self):
        return Section(self.document.Sections)

    @property
    def SensitivityLabel(self):
        return self.document.SensitivityLabel

    @property
    def Sentences(self):
        return self.document.Sentences

    @property
    def ServerPolicy(self):
        return self.document.ServerPolicy

    @property
    def Shapes(self):
        return self.document.Shapes

    @property
    def ShowGrammaticalErrors(self):
        return self.document.ShowGrammaticalErrors

    @property
    def ShowSpellingErrors(self):
        return self.document.ShowSpellingErrors

    @property
    def Signatures(self):
        return self.document.Signatures

    @property
    def SmartDocument(self):
        return self.document.SmartDocument

    @property
    def SnapToGrid(self):
        return self.document.SnapToGrid

    @property
    def SnapToShapes(self):
        return self.document.SnapToShapes

    @property
    def SpellingChecked(self):
        return self.document.SpellingChecked

    @property
    def SpellingErrors(self):
        return self.document.SpellingErrors

    @property
    def StoryRanges(self):
        return self.document.StoryRanges

    @property
    def Styles(self):
        return self.document.Styles

    @property
    def StyleSheets(self):
        return StyleSheets(self.document.StyleSheets)

    @property
    def StyleSortMethod(self):
        return WdStyleSort(self.document.StyleSortMethod)

    @StyleSortMethod.setter
    def StyleSortMethod(self, value):
        self.document.StyleSortMethod = value

    @property
    def Subdocuments(self):
        return self.document.Subdocuments

    @property
    def Sync(self):
        return self.document.Sync

    @property
    def Tables(self):
        return Table(self.document.Tables)

    @property
    def TablesOfAuthorities(self):
        return TableOfAuthorities(self.document.TablesOfAuthorities)

    @property
    def TablesOfAuthoritiesCategories(self):
        return self.document.TablesOfAuthoritiesCategories

    @property
    def TablesOfContents(self):
        return self.document.TablesOfContents

    @property
    def TablesOfFigures(self):
        return self.document.TablesOfFigures

    @property
    def TextEncoding(self):
        return self.document.TextEncoding

    @TextEncoding.setter
    def TextEncoding(self, value):
        self.document.TextEncoding = value

    @property
    def TextLineEnding(self):
        return WdLineEndingType(self.document.TextLineEnding)

    @TextLineEnding.setter
    def TextLineEnding(self, value):
        self.document.TextLineEnding = value

    @property
    def TrackFormatting(self):
        return self.document.TrackFormatting

    @TrackFormatting.setter
    def TrackFormatting(self, value):
        self.document.TrackFormatting = value

    @property
    def TrackMoves(self):
        return self.document.TrackMoves

    @TrackMoves.setter
    def TrackMoves(self, value):
        self.document.TrackMoves = value

    @property
    def TrackRevisions(self):
        return self.document.TrackRevisions

    @property
    def Type(self):
        return WdDocumentType(self.document.Type)

    @property
    def UpdateStylesOnOpen(self):
        return self.document.UpdateStylesOnOpen

    @property
    def UseMathDefaults(self):
        return self.document.UseMathDefaults

    @UseMathDefaults.setter
    def UseMathDefaults(self, value):
        self.document.UseMathDefaults = value

    @property
    def UserControl(self):
        return self.document.UserControl

    @property
    def Variables(self):
        return self.document.Variables

    @property
    def VBASigned(self):
        return self.document.VBASigned

    @property
    def VBProject(self):
        return self.document.VBProject

    @property
    def WebOptions(self):
        return WebOptions(self.document.WebOptions)

    @property
    def Windows(self):
        return self.document.Windows

    @property
    def WordOpenXML(self):
        return self.document.WordOpenXML

    @property
    def Words(self):
        return self.document.Words

    @property
    def WritePassword(self):
        return self.document.WritePassword

    @property
    def WriteReserved(self):
        return self.document.WriteReserved

    @property
    def XMLSaveThroughXSLT(self):
        return self.document.XMLSaveThroughXSLT

    @property
    def XMLSchemaReferences(self):
        return self.document.XMLSchemaReferences

    @property
    def XMLShowAdvancedErrors(self):
        return self.document.XMLShowAdvancedErrors

    @XMLShowAdvancedErrors.setter
    def XMLShowAdvancedErrors(self, value):
        self.document.XMLShowAdvancedErrors = value

    @property
    def XMLUseXSLTWhenSaving(self):
        return self.document.XMLUseXSLTWhenSaving

    def AcceptAllRevisions(self):
        self.document.AcceptAllRevisions()

    def AcceptAllRevisionsShown(self):
        self.document.AcceptAllRevisionsShown()

    def Activate(self):
        self.document.Activate()

    def AddToFavorites(self):
        self.document.AddToFavorites()

    def ApplyQuickStyleSet2(self, *args, Style=None):
        arguments = {"Style": Style}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ApplyQuickStyleSet2(*args, **arguments)

    def ApplyTheme(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ApplyTheme(*args, **arguments)

    def AutoFormat(self):
        self.document.AutoFormat()

    def CanCheckin(self):
        return self.document.CanCheckin()

    def CheckConsistency(self):
        self.document.CheckConsistency()

    def CheckGrammar(self):
        self.document.CheckGrammar()

    def CheckIn(self, *args, SaveChanges=None, Comments=None, MakePublic=None):
        arguments = {"SaveChanges": SaveChanges, "Comments": Comments, "MakePublic": MakePublic}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.CheckIn(*args, **arguments)

    def CheckInWithVersion(self, *args, SaveChanges=None, Comments=None, MakePublic=None, VersionType=None):
        arguments = {"SaveChanges": SaveChanges, "Comments": Comments, "MakePublic": MakePublic, "VersionType": VersionType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.CheckInWithVersion(*args, **arguments)

    def CheckSpelling(self, *args, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = {"CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "AlwaysSuggest": AlwaysSuggest, "CustomDictionary2": CustomDictionary2, "CustomDictionary3": CustomDictionary3, "CustomDictionary4": CustomDictionary4, "CustomDictionary5": CustomDictionary5, "CustomDictionary6": CustomDictionary6, "CustomDictionary7": CustomDictionary7, "CustomDictionary8": CustomDictionary8, "CustomDictionary9": CustomDictionary9, "CustomDictionary10": CustomDictionary10}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.CheckSpelling(*args, **arguments)

    def Close(self, *args, SaveChanges=None, OriginalFormat=None, RouteDocument=None):
        arguments = {"SaveChanges": SaveChanges, "OriginalFormat": OriginalFormat, "RouteDocument": RouteDocument}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.Close(*args, **arguments)

    def ClosePrintPreview(self):
        self.document.ClosePrintPreview()

    def Compare(self, *args, Name=None, AuthorName=None, CompareTarget=None, DetectFormatChanges=None, IgnoreAllComparisonWarnings=None, AddToRecentFiles=None, RemovePersonalInformation=None, RemoveDateAndTime=None):
        arguments = {"Name": Name, "AuthorName": AuthorName, "CompareTarget": CompareTarget, "DetectFormatChanges": DetectFormatChanges, "IgnoreAllComparisonWarnings": IgnoreAllComparisonWarnings, "AddToRecentFiles": AddToRecentFiles, "RemovePersonalInformation": RemovePersonalInformation, "RemoveDateAndTime": RemoveDateAndTime}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.Compare(*args, **arguments)

    def ComputeStatistics(self, *args, Statistic=None, IncludeFootnotesAndEndnotes=None):
        arguments = {"Statistic": Statistic, "IncludeFootnotesAndEndnotes": IncludeFootnotesAndEndnotes}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ComputeStatistics(*args, **arguments)

    def Convert(self):
        self.document.Convert()

    def ConvertAutoHyphens(self):
        self.document.ConvertAutoHyphens()

    def ConvertNumbersToText(self):
        self.document.ConvertNumbersToText()

    def ConvertVietDoc(self, *args, CodePageOrigin=None):
        arguments = {"CodePageOrigin": CodePageOrigin}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ConvertVietDoc(*args, **arguments)

    def CopyStylesFromTemplate(self, *args, Template=None):
        arguments = {"Template": Template}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.CopyStylesFromTemplate(*args, **arguments)

    def CountNumberedItems(self, *args, NumberType=None, Level=None):
        arguments = {"NumberType": NumberType, "Level": Level}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.CountNumberedItems(*args, **arguments)

    def CreateLetterContent(self, *args, DateFormat=None, IncludeHeaderFooter=None, PageDesign=None, LetterStyle=None, Letterhead=None, LetterheadLocation=None, LetterheadSize=None, RecipientName=None, RecipientAddress=None, Salutation=None, SalutationType=None, RecipientReference=None, MailingInstructions=None, AttentionLine=None, Subject=None, CCList=None, ReturnAddress=None, SenderName=None, Closing=None, SenderCompany=None, SenderJobTitle=None, SenderInitials=None, EnclosureNumber=None, InfoBlock=None, RecipientCode=None, RecipientGender=None, ReturnAddressShortForm=None, SenderCity=None, SenderCode=None, SenderGender=None, SenderReference=None):
        arguments = {"DateFormat": DateFormat, "IncludeHeaderFooter": IncludeHeaderFooter, "PageDesign": PageDesign, "LetterStyle": LetterStyle, "Letterhead": Letterhead, "LetterheadLocation": LetterheadLocation, "LetterheadSize": LetterheadSize, "RecipientName": RecipientName, "RecipientAddress": RecipientAddress, "Salutation": Salutation, "SalutationType": SalutationType, "RecipientReference": RecipientReference, "MailingInstructions": MailingInstructions, "AttentionLine": AttentionLine, "Subject": Subject, "CCList": CCList, "ReturnAddress": ReturnAddress, "SenderName": SenderName, "Closing": Closing, "SenderCompany": SenderCompany, "SenderJobTitle": SenderJobTitle, "SenderInitials": SenderInitials, "EnclosureNumber": EnclosureNumber, "InfoBlock": InfoBlock, "RecipientCode": RecipientCode, "RecipientGender": RecipientGender, "ReturnAddressShortForm": ReturnAddressShortForm, "SenderCity": SenderCity, "SenderCode": SenderCode, "SenderGender": SenderGender, "SenderReference": SenderReference}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.CreateLetterContent(*args, **arguments)

    def DataForm(self):
        self.document.DataForm()

    def DeleteAllComments(self):
        self.document.DeleteAllComments()

    def DeleteAllCommentsShown(self):
        self.document.DeleteAllCommentsShown()

    def DeleteAllEditableRanges(self, *args, EditorID=None):
        arguments = {"EditorID": EditorID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.DeleteAllEditableRanges(*args, **arguments)

    def DeleteAllInkAnnotations(self):
        self.document.DeleteAllInkAnnotations()

    def DetectLanguage(self):
        self.document.DetectLanguage()

    def DowngradeDocument(self):
        self.document.DowngradeDocument()

    def EndReview(self):
        self.document.EndReview()

    def ExportAsFixedFormat(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, Range=None, From=None, To=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "Range": Range, "From": From, "To": To, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ExportAsFixedFormat(*args, **arguments)

    def ExportAsFixedFormat2(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, Range=None, From=None, To=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, OptimizeForImageQuality=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "Range": Range, "From": From, "To": To, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "OptimizeForImageQuality": OptimizeForImageQuality, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ExportAsFixedFormat2(*args, **arguments)

    def ExportAsFixedFormat3(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, Range=None, From=None, To=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, OptimizeForImageQuality=None, ImproveExportTagging=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "Range": Range, "From": From, "To": To, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "OptimizeForImageQuality": OptimizeForImageQuality, "ImproveExportTagging": ImproveExportTagging, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ExportAsFixedFormat3(*args, **arguments)

    def FitToPages(self):
        self.document.FitToPages()

    def FollowHyperlink(self, *args, Address=None, SubAddress=None, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = {"Address": Address, "SubAddress": SubAddress, "NewWindow": NewWindow, "AddHistory": AddHistory, "ExtraInfo": ExtraInfo, "Method": Method, "HeaderInfo": HeaderInfo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.FollowHyperlink(*args, **arguments)

    def FreezeLayout(self):
        self.document.FreezeLayout()

    def GetCrossReferenceItems(self, *args, ReferenceType=None):
        arguments = {"ReferenceType": ReferenceType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.GetCrossReferenceItems(*args, **arguments)

    def GetLetterContent(self):
        return self.document.GetLetterContent()

    def GetWorkflowTasks(self):
        return self.document.GetWorkflowTasks()

    def GetWorkflowTemplates(self):
        return self.document.GetWorkflowTemplates()

    def GoTo(self, *args, What=None, Which=None, Count=None, Name=None):
        arguments = {"What": What, "Which": Which, "Count": Count, "Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.GoTo(*args, **arguments)

    def LockServerFile(self):
        self.document.LockServerFile()

    def MakeCompatibilityDefault(self):
        self.document.MakeCompatibilityDefault()

    def ManualHyphenation(self):
        self.document.ManualHyphenation()

    def Merge(self, *args, Name=None, MergeTarget=None, DetectFormatChanges=None, UseFormattingFrom=None, AddToRecentFiles=None):
        arguments = {"Name": Name, "MergeTarget": MergeTarget, "DetectFormatChanges": DetectFormatChanges, "UseFormattingFrom": UseFormattingFrom, "AddToRecentFiles": AddToRecentFiles}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.Merge(*args, **arguments)

    def Post(self):
        self.document.Post()

    def PresentIt(self):
        self.document.PresentIt()

    def PrintOut(self, *args, Background=None, Append=None, Range=None, OutputFileName=None, From=None, To=None, Item=None, Copies=None, Pages=None, PageType=None, PrintToFile=None, Collate=None, FileName=None, ActivePrinterMacGX=None, ManualDuplexPrint=None, PrintZoomColumn=None, PrintZoomRow=None, PrintZoomPaperWidth=None, PrintZoomPaperHeight=None):
        arguments = {"Background": Background, "Append": Append, "Range": Range, "OutputFileName": OutputFileName, "From": From, "To": To, "Item": Item, "Copies": Copies, "Pages": Pages, "PageType": PageType, "PrintToFile": PrintToFile, "Collate": Collate, "FileName": FileName, "ActivePrinterMacGX": ActivePrinterMacGX, "ManualDuplexPrint": ManualDuplexPrint, "PrintZoomColumn": PrintZoomColumn, "PrintZoomRow": PrintZoomRow, "PrintZoomPaperWidth": PrintZoomPaperWidth, "PrintZoomPaperHeight": PrintZoomPaperHeight}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.PrintOut(*args, **arguments)

    def PrintPreview(self):
        self.document.PrintPreview()

    def Range(self, *args, Start=None, End=None):
        arguments = {"Start": Start, "End": End}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.Range(*args, **arguments)

    def Redo(self, *args, Times=None):
        arguments = {"Times": Times}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.Redo(*args, **arguments)

    def RejectAllRevisions(self):
        self.document.RejectAllRevisions()

    def RejectAllRevisionsShown(self):
        self.document.RejectAllRevisionsShown()

    def Reload(self):
        self.document.Reload()

    def ReloadAs(self, *args, Encoding=None):
        arguments = {"Encoding": Encoding}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ReloadAs(*args, **arguments)

    def RemoveDocumentInformation(self, *args, RemoveDocInfoType=None):
        arguments = {"RemoveDocInfoType": RemoveDocInfoType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.RemoveDocumentInformation(*args, **arguments)

    def RemoveLockedStyles(self):
        self.document.RemoveLockedStyles()

    def RemoveNumbers(self, *args, NumberType=None):
        arguments = {"NumberType": NumberType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.RemoveNumbers(*args, **arguments)

    def RemoveTheme(self):
        self.document.RemoveTheme()

    def Repaginate(self):
        self.document.Repaginate()

    def ReplyWithChanges(self, *args, ShowMessage=None):
        arguments = {"ShowMessage": ShowMessage}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.ReplyWithChanges(*args, **arguments)

    def ResetFormFields(self):
        self.document.ResetFormFields()

    def RunAutoMacro(self, *args, Which=None):
        arguments = {"Which": Which}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.RunAutoMacro(*args, **arguments)

    def RunLetterWizard(self, *args, LetterContent=None, WizardMode=None):
        arguments = {"LetterContent": LetterContent, "WizardMode": WizardMode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.RunLetterWizard(*args, **arguments)

    def Save(self):
        self.document.Save()

    def SaveAsQuickStyleSet(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.SaveAsQuickStyleSet(*args, **arguments)

    def Select(self):
        self.document.Select()

    def SelectAllEditableRanges(self, *args, EditorID=None):
        arguments = {"EditorID": EditorID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.SelectAllEditableRanges(*args, **arguments)

    def SelectContentControlsByTag(self, *args, Tag=None):
        arguments = {"Tag": Tag}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.SelectContentControlsByTag(*args, **arguments)

    def SelectContentControlsByTitle(self, *args, Title=None):
        arguments = {"Title": Title}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.SelectContentControlsByTitle(*args, **arguments)

    def SelectLinkedControls(self, *args, Node=None):
        arguments = {"Node": Node}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.SelectLinkedControls(*args, **arguments)

    def SelectNodes(self, *args, XPath=None, PrefixMapping=None, FastSearchSkippingTextNodes=None):
        arguments = {"XPath": XPath, "PrefixMapping": PrefixMapping, "FastSearchSkippingTextNodes": FastSearchSkippingTextNodes}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.SelectNodes(*args, **arguments)

    def SelectSingleNode(self, *args, XPath=None, PrefixMapping=None, FastSearchSkippingTextNodes=None):
        arguments = {"XPath": XPath, "PrefixMapping": PrefixMapping, "FastSearchSkippingTextNodes": FastSearchSkippingTextNodes}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.SelectSingleNode(*args, **arguments)

    def SelectUnlinkedControls(self, *args, Stream=None):
        arguments = {"Stream": Stream}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.SelectUnlinkedControls(*args, **arguments)

    def SendFax(self, *args, Address=None, Subject=None):
        arguments = {"Address": Address, "Subject": Subject}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.SendFax(*args, **arguments)

    def SendFaxOverInternet(self, *args, Recipients=None, Subject=None, ShowMessage=None):
        arguments = {"Recipients": Recipients, "Subject": Subject, "ShowMessage": ShowMessage}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.SendFaxOverInternet(*args, **arguments)

    def SendForReview(self, *args, Recipients=None, Subject=None, ShowMessage=None, IncludeAttachment=None):
        arguments = {"Recipients": Recipients, "Subject": Subject, "ShowMessage": ShowMessage, "IncludeAttachment": IncludeAttachment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.SendForReview(*args, **arguments)

    def SendMail(self):
        self.document.SendMail()

    def SetDefaultTableStyle(self, *args, Style=None, SetInTemplate=None):
        arguments = {"Style": Style, "SetInTemplate": SetInTemplate}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.SetDefaultTableStyle(*args, **arguments)

    def SetLetterContent(self, *args, LetterContent=None):
        arguments = {"LetterContent": LetterContent}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.SetLetterContent(*args, **arguments)

    def SetPasswordEncryptionOptions(self, *args, PasswordEncryptionProvider=None, PasswordEncryptionAlgorithm=None, PasswordEncryptionKeyLength=None, PasswordEncryptionFileProperties=None):
        arguments = {"PasswordEncryptionProvider": PasswordEncryptionProvider, "PasswordEncryptionAlgorithm": PasswordEncryptionAlgorithm, "PasswordEncryptionKeyLength": PasswordEncryptionKeyLength, "PasswordEncryptionFileProperties": PasswordEncryptionFileProperties}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.SetPasswordEncryptionOptions(*args, **arguments)

    def ToggleFormsDesign(self):
        self.document.ToggleFormsDesign()

    def TransformDocument(self, *args, Path=None, DataOnly=None):
        arguments = {"Path": Path, "DataOnly": DataOnly}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.TransformDocument(*args, **arguments)

    def Undo(self, *args, Times=None):
        arguments = {"Times": Times}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.document.Undo(*args, **arguments)

    def UndoClear(self):
        self.document.UndoClear()

    def Unprotect(self, *args, Password=None):
        arguments = {"Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.document.Unprotect(*args, **arguments)

    def UpdateStyles(self):
        self.document.UpdateStyles()

    def ViewCode(self):
        self.document.ViewCode()

    def ViewPropertyBrowser(self):
        self.document.ViewPropertyBrowser()

    def WebPagePreview(self):
        self.document.WebPagePreview()

class DownBars:

    def __init__(self, downbars=None):
        self.downbars = downbars

    @property
    def Application(self):
        return self.downbars.Application

    @property
    def Creator(self):
        return self.downbars.Creator

    @property
    def Format(self):
        return ChartFormat(self.downbars.Format)

    @property
    def Name(self):
        return self.downbars.Name

    @property
    def Parent(self):
        return self.downbars.Parent

    def Delete(self):
        self.downbars.Delete()

    def Select(self):
        self.downbars.Select()

class DropCap:

    def __init__(self, dropcap=None):
        self.dropcap = dropcap

    @property
    def Application(self):
        return Application(self.dropcap.Application)

    @property
    def Creator(self):
        return self.dropcap.Creator

    @property
    def DistanceFromText(self):
        return self.dropcap.DistanceFromText

    @DistanceFromText.setter
    def DistanceFromText(self, value):
        self.dropcap.DistanceFromText = value

    @property
    def FontName(self):
        return self.dropcap.FontName

    @FontName.setter
    def FontName(self, value):
        self.dropcap.FontName = value

    @property
    def LinesToDrop(self):
        return self.dropcap.LinesToDrop

    @LinesToDrop.setter
    def LinesToDrop(self, value):
        self.dropcap.LinesToDrop = value

    @property
    def Parent(self):
        return self.dropcap.Parent

    @property
    def Position(self):
        return WdDropPosition(self.dropcap.Position)

    @Position.setter
    def Position(self, value):
        self.dropcap.Position = value

    def Clear(self):
        self.dropcap.Clear()

    def Enable(self):
        self.dropcap.Enable()

class DropDown:

    def __init__(self, dropdown=None):
        self.dropdown = dropdown

    @property
    def Application(self):
        return Application(self.dropdown.Application)

    @property
    def Creator(self):
        return self.dropdown.Creator

    @property
    def Default(self):
        return self.dropdown.Default

    @Default.setter
    def Default(self, value):
        self.dropdown.Default = value

    @property
    def ListEntries(self):
        return self.dropdown.ListEntries

    @property
    def Parent(self):
        return self.dropdown.Parent

    @property
    def Valid(self):
        return self.dropdown.Valid

    @property
    def Value(self):
        return self.dropdown.Value

    @Value.setter
    def Value(self, value):
        self.dropdown.Value = value

class DropLines:

    def __init__(self, droplines=None):
        self.droplines = droplines

    @property
    def Application(self):
        return self.droplines.Application

    @property
    def Border(self):
        return ChartBorder(self.droplines.Border)

    @property
    def Creator(self):
        return self.droplines.Creator

    @property
    def Format(self):
        return ChartFormat(self.droplines.Format)

    @property
    def Name(self):
        return self.droplines.Name

    @property
    def Parent(self):
        return self.droplines.Parent

    def Delete(self):
        self.droplines.Delete()

    def Select(self):
        self.droplines.Select()

class Editor:

    def __init__(self, editor=None):
        self.editor = editor

    @property
    def Application(self):
        return Application(self.editor.Application)

    @property
    def Creator(self):
        return self.editor.Creator

    @property
    def ID(self):
        return self.editor.ID

    @ID.setter
    def ID(self, value):
        self.editor.ID = value

    @property
    def Name(self):
        return self.editor.Name

    @Name.setter
    def Name(self, value):
        self.editor.Name = value

    @property
    def NextRange(self):
        return Range(self.editor.NextRange)

    @property
    def Parent(self):
        return self.editor.Parent

    @property
    def Range(self):
        return Range(self.editor.Range)

    def Delete(self):
        self.editor.Delete()

    def DeleteAll(self):
        self.editor.DeleteAll()

    def SelectAll(self):
        self.editor.SelectAll()

class Editors:

    def __init__(self, editors=None):
        self.editors = editors

    def __call__(self, item):
        return Editor(self.editors(item))

    @property
    def Application(self):
        return Application(self.editors.Application)

    @property
    def Count(self):
        return self.editors.Count

    @property
    def Creator(self):
        return self.editors.Creator

    @property
    def Parent(self):
        return self.editors.Parent

    def Add(self, *args, EditorID=None):
        arguments = {"EditorID": EditorID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.editors.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.editors.Item(*args, **arguments)

class Email:

    def __init__(self, email=None):
        self.email = email

    @property
    def Application(self):
        return Application(self.email.Application)

    @property
    def Creator(self):
        return self.email.Creator

    @property
    def CurrentEmailAuthor(self):
        return EmailAuthor(self.email.CurrentEmailAuthor)

    @property
    def Parent(self):
        return self.email.Parent

class EmailAuthor:

    def __init__(self, emailauthor=None):
        self.emailauthor = emailauthor

    @property
    def Application(self):
        return Application(self.emailauthor.Application)

    @property
    def Creator(self):
        return self.emailauthor.Creator

    @property
    def Parent(self):
        return self.emailauthor.Parent

    @property
    def Style(self):
        return Style(self.emailauthor.Style)

class EmailOptions:

    def __init__(self, emailoptions=None):
        self.emailoptions = emailoptions

    @property
    def Application(self):
        return Application(self.emailoptions.Application)

    @property
    def AutoFormatAsYouTypeApplyBorders(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyBorders

    @property
    def AutoFormatAsYouTypeApplyBulletedLists(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyBulletedLists

    @property
    def AutoFormatAsYouTypeApplyClosings(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyClosings

    @property
    def AutoFormatAsYouTypeApplyDates(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyDates

    @property
    def AutoFormatAsYouTypeApplyFirstIndents(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyFirstIndents

    @property
    def AutoFormatAsYouTypeApplyHeadings(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyHeadings

    @property
    def AutoFormatAsYouTypeApplyNumberedLists(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyNumberedLists

    @property
    def AutoFormatAsYouTypeApplyTables(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyTables

    @property
    def AutoFormatAsYouTypeAutoLetterWizard(self):
        return self.emailoptions.AutoFormatAsYouTypeAutoLetterWizard

    @property
    def AutoFormatAsYouTypeDefineStyles(self):
        return self.emailoptions.AutoFormatAsYouTypeDefineStyles

    @property
    def AutoFormatAsYouTypeDeleteAutoSpaces(self):
        return self.emailoptions.AutoFormatAsYouTypeDeleteAutoSpaces

    @property
    def AutoFormatAsYouTypeFormatListItemBeginning(self):
        return self.emailoptions.AutoFormatAsYouTypeFormatListItemBeginning

    @property
    def AutoFormatAsYouTypeInsertClosings(self):
        return self.emailoptions.AutoFormatAsYouTypeInsertClosings

    @property
    def AutoFormatAsYouTypeInsertOvers(self):
        return self.emailoptions.AutoFormatAsYouTypeInsertOvers

    @property
    def AutoFormatAsYouTypeMatchParentheses(self):
        return self.emailoptions.AutoFormatAsYouTypeMatchParentheses

    @property
    def AutoFormatAsYouTypeReplaceFarEastDashes(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceFarEastDashes

    @property
    def AutoFormatAsYouTypeReplaceFractions(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceFractions

    @property
    def AutoFormatAsYouTypeReplaceHyperlinks(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceHyperlinks

    @property
    def AutoFormatAsYouTypeReplaceOrdinals(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceOrdinals

    @property
    def AutoFormatAsYouTypeReplacePlainTextEmphasis(self):
        return self.emailoptions.AutoFormatAsYouTypeReplacePlainTextEmphasis

    @property
    def AutoFormatAsYouTypeReplaceQuotes(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceQuotes

    @property
    def AutoFormatAsYouTypeReplaceSymbols(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceSymbols

    @property
    def ComposeStyle(self):
        return Style(self.emailoptions.ComposeStyle)

    @property
    def Creator(self):
        return self.emailoptions.Creator

    @property
    def EmailSignature(self):
        return EmailSignature(self.emailoptions.EmailSignature)

    @property
    def HTMLFidelity(self):
        return self.emailoptions.HTMLFidelity

    @property
    def MarkComments(self):
        return self.emailoptions.MarkComments

    @property
    def MarkCommentsWith(self):
        return self.emailoptions.MarkCommentsWith

    @MarkCommentsWith.setter
    def MarkCommentsWith(self, value):
        self.emailoptions.MarkCommentsWith = value

    @property
    def NewColorOnReply(self):
        return self.emailoptions.NewColorOnReply

    @property
    def Parent(self):
        return self.emailoptions.Parent

    @property
    def PlainTextStyle(self):
        return Style(self.emailoptions.PlainTextStyle)

    @property
    def RelyOnCSS(self):
        return self.emailoptions.RelyOnCSS

    @property
    def ReplyStyle(self):
        return Style(self.emailoptions.ReplyStyle)

    @property
    def TabIndentKey(self):
        return self.emailoptions.TabIndentKey

    @property
    def ThemeName(self):
        return self.emailoptions.ThemeName

    @ThemeName.setter
    def ThemeName(self, value):
        self.emailoptions.ThemeName = value

    @property
    def UseThemeStyle(self):
        return self.emailoptions.UseThemeStyle

    @property
    def UseThemeStyleOnReply(self):
        return self.emailoptions.UseThemeStyleOnReply

class EmailSignature:

    def __init__(self, emailsignature=None):
        self.emailsignature = emailsignature

    @property
    def Application(self):
        return Application(self.emailsignature.Application)

    @property
    def Creator(self):
        return self.emailsignature.Creator

    @property
    def EmailSignatureEntries(self):
        return EmailSignatureEntries(self.emailsignature.EmailSignatureEntries)

    @property
    def NewMessageSignature(self):
        return self.emailsignature.NewMessageSignature

    @NewMessageSignature.setter
    def NewMessageSignature(self, value):
        self.emailsignature.NewMessageSignature = value

    @property
    def Parent(self):
        return self.emailsignature.Parent

    @property
    def ReplyMessageSignature(self):
        return self.emailsignature.ReplyMessageSignature

    @ReplyMessageSignature.setter
    def ReplyMessageSignature(self, value):
        self.emailsignature.ReplyMessageSignature = value

class EmailSignatureEntries:

    def __init__(self, emailsignatureentries=None):
        self.emailsignatureentries = emailsignatureentries

    def __call__(self, item):
        return EmailSignatureEntrie(self.emailsignatureentries(item))

    @property
    def Application(self):
        return Application(self.emailsignatureentries.Application)

    @property
    def Count(self):
        return self.emailsignatureentries.Count

    @property
    def Creator(self):
        return self.emailsignatureentries.Creator

    @property
    def Parent(self):
        return self.emailsignatureentries.Parent

    def Add(self, *args, Name=None, Range=None):
        arguments = {"Name": Name, "Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return EmailSignatureEntrie(self.emailsignatureentries.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.emailsignatureentries.Item(*args, **arguments)

class EmailSignatureEntry:

    def __init__(self, emailsignatureentry=None):
        self.emailsignatureentry = emailsignatureentry

    @property
    def Application(self):
        return Application(self.emailsignatureentry.Application)

    @property
    def Creator(self):
        return self.emailsignatureentry.Creator

    @property
    def Index(self):
        return self.emailsignatureentry.Index

    @property
    def Name(self):
        return self.emailsignatureentry.Name

    @Name.setter
    def Name(self, value):
        self.emailsignatureentry.Name = value

    @property
    def Parent(self):
        return self.emailsignatureentry.Parent

    def Delete(self):
        self.emailsignatureentry.Delete()

class Endnote:

    def __init__(self, endnote=None):
        self.endnote = endnote

    @property
    def Application(self):
        return Application(self.endnote.Application)

    @property
    def Creator(self):
        return self.endnote.Creator

    @property
    def Index(self):
        return self.endnote.Index

    @property
    def Parent(self):
        return self.endnote.Parent

    @property
    def Range(self):
        return Range(self.endnote.Range)

    @property
    def Reference(self):
        return Range(self.endnote.Reference)

    def Delete(self):
        self.endnote.Delete()

class EndnoteOptions:

    def __init__(self, endnoteoptions=None):
        self.endnoteoptions = endnoteoptions

    @property
    def Application(self):
        return Application(self.endnoteoptions.Application)

    @property
    def Creator(self):
        return self.endnoteoptions.Creator

    @property
    def Location(self):
        return WdEndnoteLocation(self.endnoteoptions.Location)

    @Location.setter
    def Location(self, value):
        self.endnoteoptions.Location = value

    @property
    def NumberingRule(self):
        return WdNumberingRule(self.endnoteoptions.NumberingRule)

    @NumberingRule.setter
    def NumberingRule(self, value):
        self.endnoteoptions.NumberingRule = value

    @property
    def NumberStyle(self):
        return WdNoteNumberStyle(self.endnoteoptions.NumberStyle)

    @NumberStyle.setter
    def NumberStyle(self, value):
        self.endnoteoptions.NumberStyle = value

    @property
    def Parent(self):
        return self.endnoteoptions.Parent

    @property
    def StartingNumber(self):
        return self.endnoteoptions.StartingNumber

    @StartingNumber.setter
    def StartingNumber(self, value):
        self.endnoteoptions.StartingNumber = value

class Envelope:

    def __init__(self, envelope=None):
        self.envelope = envelope

    @property
    def Address(self):
        return Range(self.envelope.Address)

    @property
    def AddressFromLeft(self):
        return self.envelope.AddressFromLeft

    @AddressFromLeft.setter
    def AddressFromLeft(self, value):
        self.envelope.AddressFromLeft = value

    @property
    def AddressFromTop(self):
        return self.envelope.AddressFromTop

    @AddressFromTop.setter
    def AddressFromTop(self, value):
        self.envelope.AddressFromTop = value

    @property
    def AddressStyle(self):
        return Style(self.envelope.AddressStyle)

    @property
    def Application(self):
        return Application(self.envelope.Application)

    @property
    def Creator(self):
        return self.envelope.Creator

    @property
    def DefaultFaceUp(self):
        return self.envelope.DefaultFaceUp

    @property
    def DefaultHeight(self):
        return self.envelope.DefaultHeight

    @DefaultHeight.setter
    def DefaultHeight(self, value):
        self.envelope.DefaultHeight = value

    @property
    def DefaultOmitReturnAddress(self):
        return self.envelope.DefaultOmitReturnAddress

    @property
    def DefaultOrientation(self):
        return WdEnvelopeOrientation(self.envelope.DefaultOrientation)

    @DefaultOrientation.setter
    def DefaultOrientation(self, value):
        self.envelope.DefaultOrientation = value

    @property
    def DefaultPrintFIMA(self):
        return self.envelope.DefaultPrintFIMA

    @property
    def DefaultSize(self):
        return self.envelope.DefaultSize

    @DefaultSize.setter
    def DefaultSize(self, value):
        self.envelope.DefaultSize = value

    @property
    def DefaultWidth(self):
        return self.envelope.DefaultWidth

    @DefaultWidth.setter
    def DefaultWidth(self, value):
        self.envelope.DefaultWidth = value

    @property
    def FeedSource(self):
        return WdPaperTray(self.envelope.FeedSource)

    @FeedSource.setter
    def FeedSource(self, value):
        self.envelope.FeedSource = value

    @property
    def Parent(self):
        return self.envelope.Parent

    @property
    def RecipientNamefromLeft(self):
        return self.envelope.RecipientNamefromLeft

    @RecipientNamefromLeft.setter
    def RecipientNamefromLeft(self, value):
        self.envelope.RecipientNamefromLeft = value

    @property
    def RecipientNamefromTop(self):
        return self.envelope.RecipientNamefromTop

    @RecipientNamefromTop.setter
    def RecipientNamefromTop(self, value):
        self.envelope.RecipientNamefromTop = value

    @property
    def RecipientPostalfromLeft(self):
        return self.envelope.RecipientPostalfromLeft

    @RecipientPostalfromLeft.setter
    def RecipientPostalfromLeft(self, value):
        self.envelope.RecipientPostalfromLeft = value

    @property
    def RecipientPostalfromTop(self):
        return self.envelope.RecipientPostalfromTop

    @RecipientPostalfromTop.setter
    def RecipientPostalfromTop(self, value):
        self.envelope.RecipientPostalfromTop = value

    @property
    def ReturnAddress(self):
        return Range(self.envelope.ReturnAddress)

    @property
    def ReturnAddressFromLeft(self):
        return self.envelope.ReturnAddressFromLeft

    @ReturnAddressFromLeft.setter
    def ReturnAddressFromLeft(self, value):
        self.envelope.ReturnAddressFromLeft = value

    @property
    def ReturnAddressFromTop(self):
        return self.envelope.ReturnAddressFromTop

    @ReturnAddressFromTop.setter
    def ReturnAddressFromTop(self, value):
        self.envelope.ReturnAddressFromTop = value

    @property
    def ReturnAddressStyle(self):
        return Style(self.envelope.ReturnAddressStyle)

    @property
    def SenderNamefromLeft(self):
        return self.envelope.SenderNamefromLeft

    @SenderNamefromLeft.setter
    def SenderNamefromLeft(self, value):
        self.envelope.SenderNamefromLeft = value

    @property
    def SenderNamefromTop(self):
        return self.envelope.SenderNamefromTop

    @SenderNamefromTop.setter
    def SenderNamefromTop(self, value):
        self.envelope.SenderNamefromTop = value

    @property
    def SenderPostalfromLeft(self):
        return self.envelope.SenderPostalfromLeft

    @SenderPostalfromLeft.setter
    def SenderPostalfromLeft(self, value):
        self.envelope.SenderPostalfromLeft = value

    @property
    def SenderPostalfromTop(self):
        return self.envelope.SenderPostalfromTop

    @SenderPostalfromTop.setter
    def SenderPostalfromTop(self, value):
        self.envelope.SenderPostalfromTop = value

    @property
    def Vertical(self):
        return self.envelope.Vertical

    def Insert(self, *args, ExtractAddress=None, Address=None, AutoText=None, OmitReturnAddress=None, ReturnAddress=None, ReturnAutoText=None, PrintBarCode=None, PrintFIMA=None, Size=None, Height=None, Width=None, FeedSource=None, AddressFromLeft=None, AddressFromTop=None, ReturnAddressFromLeft=None, ReturnAddressFromTop=None, DefaultFaceUp=None, DefaultOrientation=None, PrintEPostage=None, Vertical=None, RecipientNamefromLeft=None, RecipientNamefromTop=None, RecipientPostalfromLeft=None, RecipientPostalfromTop=None, SenderNamefromLeft=None, SenderNamefromTop=None, SenderPostalfromLeft=None, SenderPostalfromTop=None):
        arguments = {"ExtractAddress": ExtractAddress, "Address": Address, "AutoText": AutoText, "OmitReturnAddress": OmitReturnAddress, "ReturnAddress": ReturnAddress, "ReturnAutoText": ReturnAutoText, "PrintBarCode": PrintBarCode, "PrintFIMA": PrintFIMA, "Size": Size, "Height": Height, "Width": Width, "FeedSource": FeedSource, "AddressFromLeft": AddressFromLeft, "AddressFromTop": AddressFromTop, "ReturnAddressFromLeft": ReturnAddressFromLeft, "ReturnAddressFromTop": ReturnAddressFromTop, "DefaultFaceUp": DefaultFaceUp, "DefaultOrientation": DefaultOrientation, "PrintEPostage": PrintEPostage, "Vertical": Vertical, "RecipientNamefromLeft": RecipientNamefromLeft, "RecipientNamefromTop": RecipientNamefromTop, "RecipientPostalfromLeft": RecipientPostalfromLeft, "RecipientPostalfromTop": RecipientPostalfromTop, "SenderNamefromLeft": SenderNamefromLeft, "SenderNamefromTop": SenderNamefromTop, "SenderPostalfromLeft": SenderPostalfromLeft, "SenderPostalfromTop": SenderPostalfromTop}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.envelope.Insert(*args, **arguments)

    def Options(self):
        self.envelope.Options()

    def PrintOut(self, *args, ExtractAddress=None, Address=None, AutoText=None, OmitReturnAddress=None, ReturnAddress=None, ReturnAutoText=None, PrintBarCode=None, PrintFIMA=None, Size=None, Height=None, Width=None, FeedSource=None, AddressFromLeft=None, AddressFromTop=None, ReturnAddressFromLeft=None, ReturnAddressFromTop=None, DefaultFaceUp=None, DefaultOrientation=None, PrintEPostage=None, Vertical=None, RecipientNamefromLeft=None, RecipientNamefromTop=None, RecipientPostalfromLeft=None, RecipientPostalfromTop=None, SenderNamefromLeft=None, SenderNamefromTop=None, SenderPostalfromLeft=None, SenderPostalfromTop=None):
        arguments = {"ExtractAddress": ExtractAddress, "Address": Address, "AutoText": AutoText, "OmitReturnAddress": OmitReturnAddress, "ReturnAddress": ReturnAddress, "ReturnAutoText": ReturnAutoText, "PrintBarCode": PrintBarCode, "PrintFIMA": PrintFIMA, "Size": Size, "Height": Height, "Width": Width, "FeedSource": FeedSource, "AddressFromLeft": AddressFromLeft, "AddressFromTop": AddressFromTop, "ReturnAddressFromLeft": ReturnAddressFromLeft, "ReturnAddressFromTop": ReturnAddressFromTop, "DefaultFaceUp": DefaultFaceUp, "DefaultOrientation": DefaultOrientation, "PrintEPostage": PrintEPostage, "Vertical": Vertical, "RecipientNamefromLeft": RecipientNamefromLeft, "RecipientNamefromTop": RecipientNamefromTop, "RecipientPostalfromLeft": RecipientPostalfromLeft, "RecipientPostalfromTop": RecipientPostalfromTop, "SenderNamefromLeft": SenderNamefromLeft, "SenderNamefromTop": SenderNamefromTop, "SenderPostalfromLeft": SenderPostalfromLeft, "SenderPostalfromTop": SenderPostalfromTop}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.envelope.PrintOut(*args, **arguments)

    def UpdateDocument(self):
        self.envelope.UpdateDocument()

class ErrorBars:

    def __init__(self, errorbars=None):
        self.errorbars = errorbars

    @property
    def Application(self):
        return self.errorbars.Application

    @property
    def Border(self):
        return ChartBorder(self.errorbars.Border)

    @property
    def Creator(self):
        return self.errorbars.Creator

    @property
    def EndStyle(self):
        return self.errorbars.EndStyle

    @EndStyle.setter
    def EndStyle(self, value):
        self.errorbars.EndStyle = value

    @property
    def Format(self):
        return ChartFormat(self.errorbars.Format)

    @property
    def Name(self):
        return self.errorbars.Name

    @property
    def Parent(self):
        return self.errorbars.Parent

    def ClearFormats(self):
        self.errorbars.ClearFormats()

    def Delete(self):
        self.errorbars.Delete()

    def Select(self):
        self.errorbars.Select()

class Field:

    def __init__(self, field=None):
        self.field = field

    @property
    def Application(self):
        return Application(self.field.Application)

    @property
    def Code(self):
        return Range(self.field.Code)

    @property
    def Creator(self):
        return self.field.Creator

    @property
    def Data(self):
        return self.field.Data

    @Data.setter
    def Data(self, value):
        self.field.Data = value

    @property
    def Index(self):
        return self.field.Index

    @property
    def InlineShape(self):
        return InlineShape(self.field.InlineShape)

    @property
    def Kind(self):
        return WdFieldKind(self.field.Kind)

    @property
    def LinkFormat(self):
        return LinkFormat(self.field.LinkFormat)

    @property
    def Locked(self):
        return self.field.Locked

    @property
    def Next(self):
        return self.field.Next

    @property
    def OLEFormat(self):
        return OLEFormat(self.field.OLEFormat)

    @property
    def Parent(self):
        return self.field.Parent

    @property
    def Previous(self):
        return self.field.Previous

    @property
    def Result(self):
        return Range(self.field.Result)

    @property
    def ShowCodes(self):
        return self.field.ShowCodes

    @property
    def Type(self):
        return WdFieldType(self.field.Type)

    def Copy(self):
        self.field.Copy()

    def Cut(self):
        self.field.Cut()

    def Delete(self):
        self.field.Delete()

    def DoClick(self):
        self.field.DoClick()

    def Select(self):
        self.field.Select()

    def Unlink(self):
        self.field.Unlink()

    def Update(self):
        return self.field.Update()

    def UpdateSource(self):
        self.field.UpdateSource()

class FileConverter:

    def __init__(self, fileconverter=None):
        self.fileconverter = fileconverter

    @property
    def Application(self):
        return Application(self.fileconverter.Application)

    @property
    def CanOpen(self):
        return self.fileconverter.CanOpen

    @property
    def CanSave(self):
        return self.fileconverter.CanSave

    @property
    def ClassName(self):
        return self.fileconverter.ClassName

    @property
    def Creator(self):
        return self.fileconverter.Creator

    @property
    def Extensions(self):
        return FileConverter(self.fileconverter.Extensions)

    @property
    def FormatName(self):
        return self.fileconverter.FormatName

    @property
    def Name(self):
        return self.fileconverter.Name

    @property
    def OpenFormat(self):
        return self.fileconverter.OpenFormat

    @property
    def Parent(self):
        return self.fileconverter.Parent

    @property
    def Path(self):
        return self.fileconverter.Path

    @property
    def SaveFormat(self):
        return self.fileconverter.SaveFormat

class FillFormat:

    def __init__(self, fillformat=None):
        self.fillformat = fillformat

    @property
    def Application(self):
        return Application(self.fillformat.Application)

    @property
    def BackColor(self):
        return ColorFormat(self.fillformat.BackColor)

    @BackColor.setter
    def BackColor(self, value):
        self.fillformat.BackColor = value

    @property
    def Creator(self):
        return self.fillformat.Creator

    @property
    def ForeColor(self):
        return ColorFormat(self.fillformat.ForeColor)

    @ForeColor.setter
    def ForeColor(self, value):
        self.fillformat.ForeColor = value

    @property
    def GradientAngle(self):
        return self.fillformat.GradientAngle

    @GradientAngle.setter
    def GradientAngle(self, value):
        self.fillformat.GradientAngle = value

    @property
    def GradientColorType(self):
        return self.fillformat.GradientColorType

    @property
    def GradientDegree(self):
        return self.fillformat.GradientDegree

    @property
    def GradientStops(self):
        return self.fillformat.GradientStops

    @property
    def GradientStyle(self):
        return self.fillformat.GradientStyle

    @property
    def GradientVariant(self):
        return self.fillformat.GradientVariant

    @property
    def Parent(self):
        return self.fillformat.Parent

    @property
    def Pattern(self):
        return self.fillformat.Pattern

    @Pattern.setter
    def Pattern(self, value):
        self.fillformat.Pattern = value

    @property
    def PictureEffects(self):
        return self.fillformat.PictureEffects

    @property
    def PresetGradientType(self):
        return self.fillformat.PresetGradientType

    @property
    def PresetTexture(self):
        return self.fillformat.PresetTexture

    @property
    def RotateWithObject(self):
        return self.fillformat.RotateWithObject

    @RotateWithObject.setter
    def RotateWithObject(self, value):
        self.fillformat.RotateWithObject = value

    @property
    def TextureAlignment(self):
        return self.fillformat.TextureAlignment

    @TextureAlignment.setter
    def TextureAlignment(self, value):
        self.fillformat.TextureAlignment = value

    @property
    def TextureHorizontalScale(self):
        return self.fillformat.TextureHorizontalScale

    @TextureHorizontalScale.setter
    def TextureHorizontalScale(self, value):
        self.fillformat.TextureHorizontalScale = value

    @property
    def TextureName(self):
        return self.fillformat.TextureName

    @property
    def TextureOffsetX(self):
        return self.fillformat.TextureOffsetX

    @TextureOffsetX.setter
    def TextureOffsetX(self, value):
        self.fillformat.TextureOffsetX = value

    @property
    def TextureOffsetY(self):
        return self.fillformat.TextureOffsetY

    @TextureOffsetY.setter
    def TextureOffsetY(self, value):
        self.fillformat.TextureOffsetY = value

    @property
    def TextureTile(self):
        return self.fillformat.TextureTile

    @TextureTile.setter
    def TextureTile(self, value):
        self.fillformat.TextureTile = value

    @property
    def TextureType(self):
        return self.fillformat.TextureType

    @property
    def TextureVerticalScale(self):
        return self.fillformat.TextureVerticalScale

    @TextureVerticalScale.setter
    def TextureVerticalScale(self, value):
        self.fillformat.TextureVerticalScale = value

    @property
    def Transparency(self):
        return self.fillformat.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.fillformat.Transparency = value

    @property
    def Type(self):
        return self.fillformat.Type

    @property
    def Visible(self):
        return self.fillformat.Visible

    def OneColorGradient(self, *args, Style=None, Variant=None, Degree=None):
        arguments = {"Style": Style, "Variant": Variant, "Degree": Degree}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.fillformat.OneColorGradient(*args, **arguments)

    def Patterned(self, *args, Pattern=None):
        arguments = {"Pattern": Pattern}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.fillformat.Patterned(*args, **arguments)

    def PresetGradient(self, *args, Style=None, Variant=None, PresetGradientType=None):
        arguments = {"Style": Style, "Variant": Variant, "PresetGradientType": PresetGradientType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.fillformat.PresetGradient(*args, **arguments)

    def PresetTextured(self, *args, PresetTexture=None):
        arguments = {"PresetTexture": PresetTexture}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.fillformat.PresetTextured(*args, **arguments)

    def Solid(self):
        self.fillformat.Solid()

    def TwoColorGradient(self, *args, Style=None, Variant=None):
        arguments = {"Style": Style, "Variant": Variant}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.fillformat.TwoColorGradient(*args, **arguments)

    def UserPicture(self, *args, PictureFile=None):
        arguments = {"PictureFile": PictureFile}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.fillformat.UserPicture(*args, **arguments)

    def UserTextured(self, *args, TextureFile=None):
        arguments = {"TextureFile": TextureFile}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.fillformat.UserTextured(*args, **arguments)

class Find:

    def __init__(self, find=None):
        self.find = find

    @property
    def Application(self):
        return Application(self.find.Application)

    @property
    def CorrectHangulEndings(self):
        return self.find.CorrectHangulEndings

    @property
    def Creator(self):
        return self.find.Creator

    @property
    def Font(self):
        return Font(self.find.Font)

    @Font.setter
    def Font(self, value):
        self.find.Font = value

    @property
    def Format(self):
        return self.find.Format

    @property
    def Forward(self):
        return self.find.Forward

    @property
    def Found(self):
        return self.find.Found

    @property
    def Frame(self):
        return Frame(self.find.Frame)

    @property
    def HanjaPhoneticHangul(self):
        return self.find.HanjaPhoneticHangul

    @HanjaPhoneticHangul.setter
    def HanjaPhoneticHangul(self, value):
        self.find.HanjaPhoneticHangul = value

    @property
    def Highlight(self):
        return self.find.Highlight

    @property
    def IgnorePunct(self):
        return self.find.IgnorePunct

    @IgnorePunct.setter
    def IgnorePunct(self, value):
        self.find.IgnorePunct = value

    @property
    def IgnoreSpace(self):
        return self.find.IgnoreSpace

    @IgnoreSpace.setter
    def IgnoreSpace(self, value):
        self.find.IgnoreSpace = value

    @property
    def LanguageID(self):
        return WdLanguageID(self.find.LanguageID)

    @LanguageID.setter
    def LanguageID(self, value):
        self.find.LanguageID = value

    @property
    def LanguageIDFarEast(self):
        return WdLanguageID(self.find.LanguageIDFarEast)

    @LanguageIDFarEast.setter
    def LanguageIDFarEast(self, value):
        self.find.LanguageIDFarEast = value

    @property
    def LanguageIDOther(self):
        return WdLanguageID(self.find.LanguageIDOther)

    @LanguageIDOther.setter
    def LanguageIDOther(self, value):
        self.find.LanguageIDOther = value

    @property
    def MatchAlefHamza(self):
        return self.find.MatchAlefHamza

    @property
    def MatchAllWordForms(self):
        return self.find.MatchAllWordForms

    @property
    def MatchByte(self):
        return self.find.MatchByte

    @property
    def MatchCase(self):
        return self.find.MatchCase

    @property
    def MatchControl(self):
        return self.find.MatchControl

    @property
    def MatchDiacritics(self):
        return self.find.MatchDiacritics

    @property
    def MatchFuzzy(self):
        return self.find.MatchFuzzy

    @property
    def MatchKashida(self):
        return self.find.MatchKashida

    @property
    def MatchPhrase(self):
        return self.find.MatchPhrase

    @property
    def MatchPrefix(self):
        return self.find.MatchPrefix

    @property
    def MatchSoundsLike(self):
        return self.find.MatchSoundsLike

    @property
    def MatchSuffix(self):
        return self.find.MatchSuffix

    @property
    def MatchWholeWord(self):
        return self.find.MatchWholeWord

    @property
    def MatchWildcards(self):
        return self.find.MatchWildcards

    @property
    def NoProofing(self):
        return self.find.NoProofing

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.find.ParagraphFormat)

    @ParagraphFormat.setter
    def ParagraphFormat(self, value):
        self.find.ParagraphFormat = value

    @property
    def Parent(self):
        return self.find.Parent

    @property
    def Replacement(self):
        return Replacement(self.find.Replacement)

    @property
    def Style(self):
        return self.find.Style

    @Style.setter
    def Style(self, value):
        self.find.Style = value

    @property
    def Text(self):
        return self.find.Text

    @Text.setter
    def Text(self, value):
        self.find.Text = value

    @property
    def Wrap(self):
        return self.find.Wrap

    @Wrap.setter
    def Wrap(self, value):
        self.find.Wrap = value

    def ClearAllFuzzyOptions(self):
        self.find.ClearAllFuzzyOptions()

    def ClearFormatting(self):
        self.find.ClearFormatting()

    def ClearHitHighlight(self):
        return self.find.ClearHitHighlight()

    def Execute(self, *args, FindText=None, MatchCase=None, MatchWholeWord=None, MatchWildcards=None, MatchSoundsLike=None, MatchAllWordForms=None, Forward=None, Wrap=None, Format=None, ReplaceWith=None, Replace=None, MatchKashida=None, MatchDiacritics=None, MatchAlefHamza=None, MatchControl=None):
        arguments = {"FindText": FindText, "MatchCase": MatchCase, "MatchWholeWord": MatchWholeWord, "MatchWildcards": MatchWildcards, "MatchSoundsLike": MatchSoundsLike, "MatchAllWordForms": MatchAllWordForms, "Forward": Forward, "Wrap": Wrap, "Format": Format, "ReplaceWith": ReplaceWith, "Replace": Replace, "MatchKashida": MatchKashida, "MatchDiacritics": MatchDiacritics, "MatchAlefHamza": MatchAlefHamza, "MatchControl": MatchControl}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.find.Execute(*args, **arguments)

    def Execute2007(self, *args, FindText=None, MatchCase=None, MatchWholeWord=None, MatchWildcards=None, MatchSoundsLike=None, MatchAllWordForms=None, Forward=None, Wrap=None, Format=None, ReplaceWith=None, Replace=None, MatchKashida=None, MatchDiacritics=None, MatchAlefHamza=None, MatchControl=None, MatchPrefix=None, MatchSuffix=None, MatchPhrase=None, IgnoreSpace=None, IgnorePunct=None):
        arguments = {"FindText": FindText, "MatchCase": MatchCase, "MatchWholeWord": MatchWholeWord, "MatchWildcards": MatchWildcards, "MatchSoundsLike": MatchSoundsLike, "MatchAllWordForms": MatchAllWordForms, "Forward": Forward, "Wrap": Wrap, "Format": Format, "ReplaceWith": ReplaceWith, "Replace": Replace, "MatchKashida": MatchKashida, "MatchDiacritics": MatchDiacritics, "MatchAlefHamza": MatchAlefHamza, "MatchControl": MatchControl, "MatchPrefix": MatchPrefix, "MatchSuffix": MatchSuffix, "MatchPhrase": MatchPhrase, "IgnoreSpace": IgnoreSpace, "IgnorePunct": IgnorePunct}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.find.Execute2007(*args, **arguments)

    def HitHighlight(self, *args, FindText=None, HighlightColor=None, TextColor=None, MatchCase=None, MatchWholeWord=None, MatchPrefix=None, MatchSuffix=None, MatchPhrase=None, MatchWildcards=None, MatchSoundsLike=None, MatchAllWordForms=None, MatchByte=None, MatchFuzzy=None, MatchKashida=None, MatchDiacritics=None, MatchAlefHamza=None, MatchControl=None, IgnoreSpace=None, IgnorePunct=None, HanjaPhoneticHangul=None):
        arguments = {"FindText": FindText, "HighlightColor": HighlightColor, "TextColor": TextColor, "MatchCase": MatchCase, "MatchWholeWord": MatchWholeWord, "MatchPrefix": MatchPrefix, "MatchSuffix": MatchSuffix, "MatchPhrase": MatchPhrase, "MatchWildcards": MatchWildcards, "MatchSoundsLike": MatchSoundsLike, "MatchAllWordForms": MatchAllWordForms, "MatchByte": MatchByte, "MatchFuzzy": MatchFuzzy, "MatchKashida": MatchKashida, "MatchDiacritics": MatchDiacritics, "MatchAlefHamza": MatchAlefHamza, "MatchControl": MatchControl, "IgnoreSpace": IgnoreSpace, "IgnorePunct": IgnorePunct, "HanjaPhoneticHangul": HanjaPhoneticHangul}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.find.HitHighlight(*args, **arguments)

    def SetAllFuzzyOptions(self):
        self.find.SetAllFuzzyOptions()

class FirstLetterException:

    def __init__(self, firstletterexception=None):
        self.firstletterexception = firstletterexception

    @property
    def Application(self):
        return Application(self.firstletterexception.Application)

    @property
    def Creator(self):
        return self.firstletterexception.Creator

    @property
    def Index(self):
        return self.firstletterexception.Index

    @property
    def Name(self):
        return self.firstletterexception.Name

    @property
    def Parent(self):
        return self.firstletterexception.Parent

    def Delete(self):
        self.firstletterexception.Delete()

class Floor:

    def __init__(self, floor=None):
        self.floor = floor

    @property
    def Application(self):
        return self.floor.Application

    @property
    def Creator(self):
        return self.floor.Creator

    @property
    def Format(self):
        return ChartFormat(self.floor.Format)

    @property
    def Name(self):
        return self.floor.Name

    @property
    def Parent(self):
        return self.floor.Parent

    @property
    def PictureType(self):
        return self.floor.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.floor.PictureType = value

    @property
    def Thickness(self):
        return self.floor.Thickness

    @Thickness.setter
    def Thickness(self, value):
        self.floor.Thickness = value

    def ClearFormats(self):
        self.floor.ClearFormats()

    def Paste(self):
        self.floor.Paste()

    def Select(self):
        self.floor.Select()

class Font:

    def __init__(self, font=None):
        self.font = font

    @property
    def AllCaps(self):
        return self.font.AllCaps

    @property
    def Application(self):
        return Application(self.font.Application)

    @property
    def Bold(self):
        return self.font.Bold

    @property
    def BoldBi(self):
        return self.font.BoldBi

    @property
    def Borders(self):
        return self.font.Borders

    @property
    def ColorIndex(self):
        return WdColorIndex(self.font.ColorIndex)

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.font.ColorIndex = value

    @property
    def ColorIndexBi(self):
        return Font(self.font.ColorIndexBi)

    @ColorIndexBi.setter
    def ColorIndexBi(self, value):
        self.font.ColorIndexBi = value

    @property
    def ContextualAlternates(self):
        return self.font.ContextualAlternates

    @property
    def Creator(self):
        return self.font.Creator

    @property
    def DiacriticColor(self):
        return Font(self.font.DiacriticColor)

    @DiacriticColor.setter
    def DiacriticColor(self, value):
        self.font.DiacriticColor = value

    @property
    def DisableCharacterSpaceGrid(self):
        return self.font.DisableCharacterSpaceGrid

    @property
    def DoubleStrikeThrough(self):
        return self.font.DoubleStrikeThrough

    @property
    def Duplicate(self):
        return Font(self.font.Duplicate)

    @property
    def Emboss(self):
        return self.font.Emboss

    @property
    def EmphasisMark(self):
        return WdEmphasisMark(self.font.EmphasisMark)

    @EmphasisMark.setter
    def EmphasisMark(self, value):
        self.font.EmphasisMark = value

    @property
    def Engrave(self):
        return self.font.Engrave

    @property
    def Fill(self):
        return self.font.Fill

    @property
    def Glow(self):
        return GlowFormat(self.font.Glow)

    @property
    def Hidden(self):
        return self.font.Hidden

    @property
    def Italic(self):
        return self.font.Italic

    @property
    def ItalicBi(self):
        return self.font.ItalicBi

    @property
    def Kerning(self):
        return self.font.Kerning

    @Kerning.setter
    def Kerning(self, value):
        self.font.Kerning = value

    @property
    def Ligatures(self):
        return Font(self.font.Ligatures)

    @Ligatures.setter
    def Ligatures(self, value):
        self.font.Ligatures = value

    @property
    def Line(self):
        return self.font.Line

    @property
    def Name(self):
        return self.font.Name

    @Name.setter
    def Name(self, value):
        self.font.Name = value

    @property
    def NameAscii(self):
        return self.font.NameAscii

    @NameAscii.setter
    def NameAscii(self, value):
        self.font.NameAscii = value

    @property
    def NameBi(self):
        return self.font.NameBi

    @NameBi.setter
    def NameBi(self, value):
        self.font.NameBi = value

    @property
    def NameFarEast(self):
        return self.font.NameFarEast

    @NameFarEast.setter
    def NameFarEast(self, value):
        self.font.NameFarEast = value

    @property
    def NameOther(self):
        return self.font.NameOther

    @NameOther.setter
    def NameOther(self, value):
        self.font.NameOther = value

    @property
    def NumberForm(self):
        return self.font.NumberForm

    @NumberForm.setter
    def NumberForm(self, value):
        self.font.NumberForm = value

    @property
    def NumberSpacing(self):
        return self.font.NumberSpacing

    @NumberSpacing.setter
    def NumberSpacing(self, value):
        self.font.NumberSpacing = value

    @property
    def Outline(self):
        return self.font.Outline

    @property
    def Parent(self):
        return self.font.Parent

    @property
    def Position(self):
        return self.font.Position

    @Position.setter
    def Position(self, value):
        self.font.Position = value

    @property
    def Reflection(self):
        return self.font.Reflection

    @property
    def Scaling(self):
        return self.font.Scaling

    @Scaling.setter
    def Scaling(self, value):
        self.font.Scaling = value

    @property
    def Shading(self):
        return Shading(self.font.Shading)

    @property
    def Shadow(self):
        return self.font.Shadow

    @property
    def Size(self):
        return self.font.Size

    @Size.setter
    def Size(self, value):
        self.font.Size = value

    @property
    def SizeBi(self):
        return self.font.SizeBi

    @SizeBi.setter
    def SizeBi(self, value):
        self.font.SizeBi = value

    @property
    def SmallCaps(self):
        return self.font.SmallCaps

    @property
    def Spacing(self):
        return self.font.Spacing

    @Spacing.setter
    def Spacing(self, value):
        self.font.Spacing = value

    @property
    def StrikeThrough(self):
        return self.font.StrikeThrough

    @property
    def StylisticSet(self):
        return self.font.StylisticSet

    @property
    def Subscript(self):
        return self.font.Subscript

    @property
    def Superscript(self):
        return self.font.Superscript

    @property
    def TextColor(self):
        return self.font.TextColor

    @property
    def TextShadow(self):
        return self.font.TextShadow

    @property
    def ThreeD(self):
        return self.font.ThreeD

    @property
    def Underline(self):
        return WdUnderline(self.font.Underline)

    @Underline.setter
    def Underline(self, value):
        self.font.Underline = value

    @property
    def UnderlineColor(self):
        return Font(self.font.UnderlineColor)

    @UnderlineColor.setter
    def UnderlineColor(self, value):
        self.font.UnderlineColor = value

    def Grow(self):
        self.font.Grow()

    def Reset(self):
        self.font.Reset()

    def SetAsTemplateDefault(self):
        self.font.SetAsTemplateDefault()

    def Shrink(self):
        self.font.Shrink()

class FontNames:

    def __init__(self, fontnames=None):
        self.fontnames = fontnames

    @property
    def Application(self):
        return Application(self.fontnames.Application)

    @property
    def Count(self):
        return self.fontnames.Count

    @property
    def Creator(self):
        return self.fontnames.Creator

    @property
    def Parent(self):
        return self.fontnames.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.fontnames.Item(*args, **arguments)

class Footnote:

    def __init__(self, footnote=None):
        self.footnote = footnote

    @property
    def Application(self):
        return Application(self.footnote.Application)

    @property
    def Creator(self):
        return self.footnote.Creator

    @property
    def Index(self):
        return self.footnote.Index

    @property
    def Parent(self):
        return self.footnote.Parent

    @property
    def Range(self):
        return Range(self.footnote.Range)

    @property
    def Reference(self):
        return Range(self.footnote.Reference)

    def Delete(self):
        self.footnote.Delete()

class FootnoteOptions:

    def __init__(self, footnoteoptions=None):
        self.footnoteoptions = footnoteoptions

    @property
    def Application(self):
        return Application(self.footnoteoptions.Application)

    @property
    def Creator(self):
        return self.footnoteoptions.Creator

    @property
    def Location(self):
        return WdFootnoteLocation(self.footnoteoptions.Location)

    @Location.setter
    def Location(self, value):
        self.footnoteoptions.Location = value

    @property
    def NumberingRule(self):
        return WdNumberingRule(self.footnoteoptions.NumberingRule)

    @NumberingRule.setter
    def NumberingRule(self, value):
        self.footnoteoptions.NumberingRule = value

    @property
    def NumberStyle(self):
        return WdNoteNumberStyle(self.footnoteoptions.NumberStyle)

    @NumberStyle.setter
    def NumberStyle(self, value):
        self.footnoteoptions.NumberStyle = value

    @property
    def Parent(self):
        return self.footnoteoptions.Parent

    @property
    def StartingNumber(self):
        return self.footnoteoptions.StartingNumber

    @StartingNumber.setter
    def StartingNumber(self, value):
        self.footnoteoptions.StartingNumber = value

class FormField:

    def __init__(self, formfield=None):
        self.formfield = formfield

    @property
    def Application(self):
        return Application(self.formfield.Application)

    @property
    def CalculateOnExit(self):
        return self.formfield.CalculateOnExit

    @property
    def CheckBox(self):
        return CheckBox(self.formfield.CheckBox)

    @property
    def Creator(self):
        return self.formfield.Creator

    @property
    def DropDown(self):
        return DropDown(self.formfield.DropDown)

    @property
    def Enabled(self):
        return self.formfield.Enabled

    @property
    def EntryMacro(self):
        return self.formfield.EntryMacro

    @EntryMacro.setter
    def EntryMacro(self, value):
        self.formfield.EntryMacro = value

    @property
    def ExitMacro(self):
        return self.formfield.ExitMacro

    @ExitMacro.setter
    def ExitMacro(self, value):
        self.formfield.ExitMacro = value

    @property
    def HelpText(self):
        return self.formfield.HelpText

    @HelpText.setter
    def HelpText(self, value):
        self.formfield.HelpText = value

    @property
    def Name(self):
        return self.formfield.Name

    @Name.setter
    def Name(self, value):
        self.formfield.Name = value

    @property
    def Next(self):
        return self.formfield.Next

    @property
    def OwnHelp(self):
        return self.formfield.OwnHelp

    @property
    def OwnStatus(self):
        return self.formfield.OwnStatus

    @property
    def Parent(self):
        return self.formfield.Parent

    @property
    def Previous(self):
        return self.formfield.Previous

    @property
    def Range(self):
        return Range(self.formfield.Range)

    @property
    def Result(self):
        return self.formfield.Result

    @property
    def StatusText(self):
        return self.formfield.StatusText

    @StatusText.setter
    def StatusText(self, value):
        self.formfield.StatusText = value

    @property
    def TextInput(self):
        return TextInput(self.formfield.TextInput)

    @property
    def Type(self):
        return WdFieldType(self.formfield.Type)

    def Copy(self):
        self.formfield.Copy()

    def Cut(self):
        self.formfield.Cut()

    def Delete(self):
        self.formfield.Delete()

    def Select(self):
        self.formfield.Select()

class Frame:

    def __init__(self, frame=None):
        self.frame = frame

    @property
    def Application(self):
        return Application(self.frame.Application)

    @property
    def Borders(self):
        return self.frame.Borders

    @property
    def Creator(self):
        return self.frame.Creator

    @property
    def Height(self):
        return self.frame.Height

    @Height.setter
    def Height(self, value):
        self.frame.Height = value

    @property
    def HeightRule(self):
        return WdFrameSizeRule(self.frame.HeightRule)

    @HeightRule.setter
    def HeightRule(self, value):
        self.frame.HeightRule = value

    @property
    def HorizontalDistanceFromText(self):
        return self.frame.HorizontalDistanceFromText

    @HorizontalDistanceFromText.setter
    def HorizontalDistanceFromText(self, value):
        self.frame.HorizontalDistanceFromText = value

    @property
    def HorizontalPosition(self):
        return self.frame.HorizontalPosition

    @HorizontalPosition.setter
    def HorizontalPosition(self, value):
        self.frame.HorizontalPosition = value

    @property
    def LockAnchor(self):
        return self.frame.LockAnchor

    @property
    def Parent(self):
        return self.frame.Parent

    @property
    def Range(self):
        return Range(self.frame.Range)

    @property
    def RelativeHorizontalPosition(self):
        return self.frame.RelativeHorizontalPosition

    @property
    def RelativeVerticalPosition(self):
        return self.frame.RelativeVerticalPosition

    @property
    def Shading(self):
        return Shading(self.frame.Shading)

    @property
    def TextWrap(self):
        return self.frame.TextWrap

    @property
    def VerticalDistanceFromText(self):
        return self.frame.VerticalDistanceFromText

    @VerticalDistanceFromText.setter
    def VerticalDistanceFromText(self, value):
        self.frame.VerticalDistanceFromText = value

    @property
    def VerticalPosition(self):
        return self.frame.VerticalPosition

    @VerticalPosition.setter
    def VerticalPosition(self, value):
        self.frame.VerticalPosition = value

    @property
    def Width(self):
        return self.frame.Width

    @Width.setter
    def Width(self, value):
        self.frame.Width = value

    @property
    def WidthRule(self):
        return WdFrameSizeRule(self.frame.WidthRule)

    @WidthRule.setter
    def WidthRule(self, value):
        self.frame.WidthRule = value

    def Copy(self):
        self.frame.Copy()

    def Cut(self):
        self.frame.Cut()

    def Delete(self):
        self.frame.Delete()

    def Select(self):
        self.frame.Select()

class Frameset:

    def __init__(self, frameset=None):
        self.frameset = frameset

    @property
    def Application(self):
        return Application(self.frameset.Application)

    @property
    def ChildFramesetCount(self):
        return Frameset(self.frameset.ChildFramesetCount)

    def ChildFramesetItem(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Frameset(self.frameset.ChildFramesetItem(*args, **arguments))

    @property
    def Creator(self):
        return self.frameset.Creator

    @property
    def FrameDefaultURL(self):
        return self.frameset.FrameDefaultURL

    @FrameDefaultURL.setter
    def FrameDefaultURL(self, value):
        self.frameset.FrameDefaultURL = value

    @property
    def FrameDisplayBorders(self):
        return self.frameset.FrameDisplayBorders

    @property
    def FrameLinkToFile(self):
        return self.frameset.FrameLinkToFile

    @property
    def FrameName(self):
        return self.frameset.FrameName

    @FrameName.setter
    def FrameName(self, value):
        self.frameset.FrameName = value

    @property
    def FrameResizable(self):
        return self.frameset.FrameResizable

    @property
    def FrameScrollbarType(self):
        return WdScrollbarType(self.frameset.FrameScrollbarType)

    @FrameScrollbarType.setter
    def FrameScrollbarType(self, value):
        self.frameset.FrameScrollbarType = value

    @property
    def FramesetBorderColor(self):
        return self.frameset.FramesetBorderColor

    @FramesetBorderColor.setter
    def FramesetBorderColor(self, value):
        self.frameset.FramesetBorderColor = value

    @property
    def FramesetBorderWidth(self):
        return self.frameset.FramesetBorderWidth

    @FramesetBorderWidth.setter
    def FramesetBorderWidth(self, value):
        self.frameset.FramesetBorderWidth = value

    @property
    def Height(self):
        return self.frameset.Height

    @Height.setter
    def Height(self, value):
        self.frameset.Height = value

    @property
    def HeightType(self):
        return WdFramesetSizeType(self.frameset.HeightType)

    @HeightType.setter
    def HeightType(self, value):
        self.frameset.HeightType = value

    @property
    def Parent(self):
        return self.frameset.Parent

    @property
    def ParentFrameset(self):
        return Frameset(self.frameset.ParentFrameset)

    @property
    def Type(self):
        return WdFramesetType(self.frameset.Type)

    @property
    def Width(self):
        return Frameset(self.frameset.Width)

    @Width.setter
    def Width(self, value):
        self.frameset.Width = value

    @property
    def WidthType(self):
        return Frameset(self.frameset.WidthType)

    @WidthType.setter
    def WidthType(self, value):
        self.frameset.WidthType = value

    def AddNewFrame(self, *args, Where=None):
        arguments = {"Where": Where}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.frameset.AddNewFrame(*args, **arguments)

    def Delete(self):
        self.frameset.Delete()

class FreeformBuilder:

    def __init__(self, freeformbuilder=None):
        self.freeformbuilder = freeformbuilder

    @property
    def Application(self):
        return Application(self.freeformbuilder.Application)

    @property
    def Creator(self):
        return self.freeformbuilder.Creator

    @property
    def Parent(self):
        return self.freeformbuilder.Parent

    def AddNodes(self, *args, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = {"SegmentType": SegmentType, "EditingType": EditingType, "X1": X1, "Y1": Y1, "X2": X2, "Y2": Y2, "X3": X3, "Y3": Y3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.freeformbuilder.AddNodes(*args, **arguments)

    def ConvertToShape(self, *args, Anchor=None):
        arguments = {"Anchor": Anchor}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.freeformbuilder.ConvertToShape(*args, **arguments)

class Global:

    def __init__(self, global=None):
        self.global = global

    @property
    def ActiveDocument(self):
        return Document(self.global.ActiveDocument)

    @property
    def ActivePrinter(self):
        return self.global.ActivePrinter

    @ActivePrinter.setter
    def ActivePrinter(self, value):
        self.global.ActivePrinter = value

    @property
    def ActiveProtectedViewWindow(self):
        return self.global.ActiveProtectedViewWindow

    @property
    def ActiveWindow(self):
        return Window(self.global.ActiveWindow)

    @property
    def AddIns(self):
        return self.global.AddIns

    @property
    def Application(self):
        return Application(self.global.Application)

    @property
    def AutoCaptions(self):
        return self.global.AutoCaptions

    @property
    def AutoCorrect(self):
        return AutoCorrect(self.global.AutoCorrect)

    @property
    def AutoCorrectEmail(self):
        return AutoCorrect(self.global.AutoCorrectEmail)

    @property
    def CaptionLabels(self):
        return self.global.CaptionLabels

    @property
    def CommandBars(self):
        return self.global.CommandBars

    @property
    def Creator(self):
        return self.global.Creator

    @property
    def CustomDictionaries(self):
        return self.global.CustomDictionaries

    @property
    def CustomizationContext(self):
        return Document(self.global.CustomizationContext)

    @CustomizationContext.setter
    def CustomizationContext(self, value):
        self.global.CustomizationContext = value

    @property
    def Dialogs(self):
        return self.global.Dialogs

    @property
    def Documents(self):
        return self.global.Documents

    @property
    def FileConverters(self):
        return self.global.FileConverters

    def FindKey(self, *args, KeyCode=None, KeyCode2=None):
        arguments = {"KeyCode": KeyCode, "KeyCode2": KeyCode2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return KeyBinding(self.global.FindKey(*args, **arguments))

    @property
    def FontNames(self):
        return FontNames(self.global.FontNames)

    @property
    def HangulHanjaDictionaries(self):
        return self.global.HangulHanjaDictionaries

    def IsObjectValid(self, *args, Object=None):
        arguments = {"Object": Object}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.IsObjectValid(*args, **arguments)

    @property
    def IsSandboxed(self):
        return self.global.IsSandboxed

    @property
    def KeyBindings(self):
        return self.global.KeyBindings

    def KeysBoundTo(self, *args, KeyCategory=None, Command=None, CommandParameter=None):
        arguments = {"KeyCategory": KeyCategory, "Command": Command, "CommandParameter": CommandParameter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.KeysBoundTo(*args, **arguments)

    @property
    def LandscapeFontNames(self):
        return FontNames(self.global.LandscapeFontNames)

    @property
    def Languages(self):
        return self.global.Languages

    @property
    def LanguageSettings(self):
        return self.global.LanguageSettings

    @property
    def ListGalleries(self):
        return self.global.ListGalleries

    @property
    def MacroContainer(self):
        return Template(self.global.MacroContainer)

    @property
    def Name(self):
        return self.global.Name

    @property
    def NormalTemplate(self):
        return Template(self.global.NormalTemplate)

    @property
    def Options(self):
        return Options(self.global.Options)

    @property
    def Parent(self):
        return self.global.Parent

    @property
    def PortraitFontNames(self):
        return FontNames(self.global.PortraitFontNames)

    @property
    def PrintPreview(self):
        return self.global.PrintPreview

    @property
    def ProtectedViewWindows(self):
        return self.global.ProtectedViewWindows

    @property
    def RecentFiles(self):
        return self.global.RecentFiles

    @property
    def Selection(self):
        return Selection(self.global.Selection)

    @property
    def ShowVisualBasicEditor(self):
        return self.global.ShowVisualBasicEditor

    @property
    def StatusBar(self):
        return self.global.StatusBar

    def SynonymInfo(self, *args, Word=None, LanguageID=None):
        arguments = {"Word": Word, "LanguageID": LanguageID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return SynonymInfo(self.global.SynonymInfo(*args, **arguments))

    @property
    def System(self):
        return System(self.global.System)

    @property
    def Tasks(self):
        return self.global.Tasks

    @property
    def Templates(self):
        return self.global.Templates

    @property
    def VBE(self):
        return self.global.VBE

    @property
    def Windows(self):
        return self.global.Windows

    @property
    def WordBasic(self):
        return self.global.WordBasic

    def BuildKeyCode(self, *args, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = {"Arg1": Arg1, "Arg2": Arg2, "Arg3": Arg3, "Arg4": Arg4}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.BuildKeyCode(*args, **arguments)

    def CentimetersToPoints(self, *args, Centimeters=None):
        arguments = {"Centimeters": Centimeters}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.CentimetersToPoints(*args, **arguments)

    def ChangeFileOpenDirectory(self, *args, Path=None):
        arguments = {"Path": Path}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.ChangeFileOpenDirectory(*args, **arguments)

    def CheckSpelling(self, *args, Word=None, CustomDictionary=None, IgnoreUppercase=None, MainDictionary=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = {"Word": Word, "CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "MainDictionary": MainDictionary, "CustomDictionary2": CustomDictionary2, "CustomDictionary3": CustomDictionary3, "CustomDictionary4": CustomDictionary4, "CustomDictionary5": CustomDictionary5, "CustomDictionary6": CustomDictionary6, "CustomDictionary7": CustomDictionary7, "CustomDictionary8": CustomDictionary8, "CustomDictionary9": CustomDictionary9, "CustomDictionary10": CustomDictionary10}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.CheckSpelling(*args, **arguments)

    def CleanString(self, *args, String=None):
        arguments = {"String": String}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.CleanString(*args, **arguments)

    def DDEExecute(self, *args, Channel=None, Command=None):
        arguments = {"Channel": Channel, "Command": Command}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.DDEExecute(*args, **arguments)

    def DDEInitiate(self, *args, App=None, Topic=None):
        arguments = {"App": App, "Topic": Topic}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.DDEInitiate(*args, **arguments)

    def DDEPoke(self, *args, Channel=None, Item=None, Data=None):
        arguments = {"Channel": Channel, "Item": Item, "Data": Data}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.DDEPoke(*args, **arguments)

    def DDERequest(self, *args, Channel=None, Item=None):
        arguments = {"Channel": Channel, "Item": Item}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.DDERequest(*args, **arguments)

    def DDETerminate(self, *args, Channel=None):
        arguments = {"Channel": Channel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.DDETerminate(*args, **arguments)

    def DDETerminateAll(self):
        self.global.DDETerminateAll()

    def GetSpellingSuggestions(self, *args, Word=None, CustomDictionary=None, IgnoreUppercase=None, MainDictionary=None, SuggestionMode=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = {"Word": Word, "CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "MainDictionary": MainDictionary, "SuggestionMode": SuggestionMode, "CustomDictionary2": CustomDictionary2, "CustomDictionary3": CustomDictionary3, "CustomDictionary4": CustomDictionary4, "CustomDictionary5": CustomDictionary5, "CustomDictionary6": CustomDictionary6, "CustomDictionary7": CustomDictionary7, "CustomDictionary8": CustomDictionary8, "CustomDictionary9": CustomDictionary9, "CustomDictionary10": CustomDictionary10}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.GetSpellingSuggestions(*args, **arguments)

    def Help(self, *args, HelpType=None):
        arguments = {"HelpType": HelpType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.global.Help(*args, **arguments)

    def InchesToPoints(self, *args, Inches=None):
        arguments = {"Inches": Inches}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.InchesToPoints(*args, **arguments)

    def KeyString(self, *args, KeyCode=None, KeyCode2=None):
        arguments = {"KeyCode": KeyCode, "KeyCode2": KeyCode2}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.KeyString(*args, **arguments)

    def LinesToPoints(self, *args, Lines=None):
        arguments = {"Lines": Lines}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.LinesToPoints(*args, **arguments)

    def MillimetersToPoints(self, *args, Millimeters=None):
        arguments = {"Millimeters": Millimeters}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.MillimetersToPoints(*args, **arguments)

    def NewWindow(self):
        return self.global.NewWindow()

    def PicasToPoints(self, *args, Picas=None):
        arguments = {"Picas": Picas}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.PicasToPoints(*args, **arguments)

    def PixelsToPoints(self, *args, Pixels=None, fVertical=None):
        arguments = {"Pixels": Pixels, "fVertical": fVertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.PixelsToPoints(*args, **arguments)

    def PointsToCentimeters(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.PointsToCentimeters(*args, **arguments)

    def PointsToInches(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.PointsToInches(*args, **arguments)

    def PointsToLines(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.PointsToLines(*args, **arguments)

    def PointsToMillimeters(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.PointsToMillimeters(*args, **arguments)

    def PointsToPicas(self, *args, Points=None):
        arguments = {"Points": Points}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.PointsToPicas(*args, **arguments)

    def PointsToPixels(self, *args, Points=None, fVertical=None):
        arguments = {"Points": Points, "fVertical": fVertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.PointsToPixels(*args, **arguments)

    def Repeat(self, *args, Times=None):
        arguments = {"Times": Times}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.global.Repeat(*args, **arguments)

class GlowFormat:

    def __init__(self, glowformat=None):
        self.glowformat = glowformat

    @property
    def Application(self):
        return self.glowformat.Application

    @property
    def Color(self):
        return ColorFormat(self.glowformat.Color)

    @property
    def Creator(self):
        return self.glowformat.Creator

    @property
    def Parent(self):
        return self.glowformat.Parent

    @property
    def Radius(self):
        return self.glowformat.Radius

    @Radius.setter
    def Radius(self, value):
        self.glowformat.Radius = value

class GridLines:

    def __init__(self, gridlines=None):
        self.gridlines = gridlines

    @property
    def Application(self):
        return self.gridlines.Application

    @property
    def Border(self):
        return ChartBorder(self.gridlines.Border)

    @property
    def Creator(self):
        return self.gridlines.Creator

    @property
    def Format(self):
        return ChartFormat(self.gridlines.Format)

    @property
    def Name(self):
        return self.gridlines.Name

    @property
    def Parent(self):
        return self.gridlines.Parent

    def Delete(self):
        self.gridlines.Delete()

    def Select(self):
        self.gridlines.Select()

class HangulAndAlphabetException:

    def __init__(self, hangulandalphabetexception=None):
        self.hangulandalphabetexception = hangulandalphabetexception

    @property
    def Application(self):
        return Application(self.hangulandalphabetexception.Application)

    @property
    def Creator(self):
        return self.hangulandalphabetexception.Creator

    @property
    def Index(self):
        return self.hangulandalphabetexception.Index

    @property
    def Name(self):
        return self.hangulandalphabetexception.Name

    @property
    def Parent(self):
        return self.hangulandalphabetexception.Parent

    def Delete(self):
        self.hangulandalphabetexception.Delete()

class HeaderFooter:

    def __init__(self, headerfooter=None):
        self.headerfooter = headerfooter

    @property
    def Application(self):
        return Application(self.headerfooter.Application)

    @property
    def Creator(self):
        return self.headerfooter.Creator

    @property
    def Exists(self):
        return self.headerfooter.Exists

    @property
    def Index(self):
        return WdHeaderFooterIndex(self.headerfooter.Index)

    @property
    def IsHeader(self):
        return self.headerfooter.IsHeader

    @property
    def LinkToPrevious(self):
        return self.headerfooter.LinkToPrevious

    @property
    def PageNumbers(self):
        return self.headerfooter.PageNumbers

    @property
    def Parent(self):
        return self.headerfooter.Parent

    @property
    def Range(self):
        return Range(self.headerfooter.Range)

    @property
    def Shapes(self):
        return self.headerfooter.Shapes

class HeadingStyle:

    def __init__(self, headingstyle=None):
        self.headingstyle = headingstyle

    @property
    def Application(self):
        return Application(self.headingstyle.Application)

    @property
    def Creator(self):
        return self.headingstyle.Creator

    @property
    def Level(self):
        return self.headingstyle.Level

    @Level.setter
    def Level(self, value):
        self.headingstyle.Level = value

    @property
    def Parent(self):
        return self.headingstyle.Parent

    @property
    def Style(self):
        return self.headingstyle.Style

    @Style.setter
    def Style(self, value):
        self.headingstyle.Style = value

    def Delete(self):
        self.headingstyle.Delete()

class HiLoLines:

    def __init__(self, hilolines=None):
        self.hilolines = hilolines

    @property
    def Application(self):
        return self.hilolines.Application

    @property
    def Border(self):
        return ChartBorder(self.hilolines.Border)

    @property
    def Creator(self):
        return self.hilolines.Creator

    @property
    def Format(self):
        return ChartFormat(self.hilolines.Format)

    @property
    def Name(self):
        return self.hilolines.Name

    @property
    def Parent(self):
        return self.hilolines.Parent

    def Delete(self):
        self.hilolines.Delete()

    def Select(self):
        self.hilolines.Select()

class HorizontalLineFormat:

    def __init__(self, horizontallineformat=None):
        self.horizontallineformat = horizontallineformat

    @property
    def Alignment(self):
        return WdHorizontalLineAlignment(self.horizontallineformat.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.horizontallineformat.Alignment = value

    @property
    def Application(self):
        return Application(self.horizontallineformat.Application)

    @property
    def Creator(self):
        return self.horizontallineformat.Creator

    @property
    def NoShade(self):
        return self.horizontallineformat.NoShade

    @property
    def Parent(self):
        return self.horizontallineformat.Parent

    @property
    def PercentWidth(self):
        return self.horizontallineformat.PercentWidth

    @PercentWidth.setter
    def PercentWidth(self, value):
        self.horizontallineformat.PercentWidth = value

    @property
    def WidthType(self):
        return HorizontalLineFormat(self.horizontallineformat.WidthType)

    @WidthType.setter
    def WidthType(self, value):
        self.horizontallineformat.WidthType = value

class HTMLDivision:

    def __init__(self, htmldivision=None):
        self.htmldivision = htmldivision

    @property
    def Application(self):
        return Application(self.htmldivision.Application)

    @property
    def Borders(self):
        return self.htmldivision.Borders

    @property
    def Creator(self):
        return self.htmldivision.Creator

    @property
    def HTMLDivisions(self):
        return HTMLDivisions(self.htmldivision.HTMLDivisions)

    @property
    def LeftIndent(self):
        return self.htmldivision.LeftIndent

    @LeftIndent.setter
    def LeftIndent(self, value):
        self.htmldivision.LeftIndent = value

    @property
    def Parent(self):
        return self.htmldivision.Parent

    @property
    def Range(self):
        return Range(self.htmldivision.Range)

    @property
    def RightIndent(self):
        return self.htmldivision.RightIndent

    @RightIndent.setter
    def RightIndent(self, value):
        self.htmldivision.RightIndent = value

    @property
    def SpaceAfter(self):
        return self.htmldivision.SpaceAfter

    @SpaceAfter.setter
    def SpaceAfter(self, value):
        self.htmldivision.SpaceAfter = value

    @property
    def SpaceBefore(self):
        return self.htmldivision.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.htmldivision.SpaceBefore = value

    def Delete(self):
        self.htmldivision.Delete()

    def HTMLDivisionParent(self, *args, LevelsUp=None):
        arguments = {"LevelsUp": LevelsUp}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.htmldivision.HTMLDivisionParent(*args, **arguments)

class HTMLDivisions:

    def __init__(self, htmldivisions=None):
        self.htmldivisions = htmldivisions

    def __call__(self, item):
        return HTMLDivision(self.htmldivisions(item))

    @property
    def Application(self):
        return Application(self.htmldivisions.Application)

    @property
    def Count(self):
        return self.htmldivisions.Count

    @property
    def Creator(self):
        return self.htmldivisions.Creator

    @property
    def NestingLevel(self):
        return self.htmldivisions.NestingLevel

    @property
    def Parent(self):
        return self.htmldivisions.Parent

    def Add(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return HTMLDivision(self.htmldivisions.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.htmldivisions.Item(*args, **arguments)

class Hyperlink:

    def __init__(self, hyperlink=None):
        self.hyperlink = hyperlink

    @property
    def Address(self):
        return self.hyperlink.Address

    @Address.setter
    def Address(self, value):
        self.hyperlink.Address = value

    @property
    def Application(self):
        return Application(self.hyperlink.Application)

    @property
    def Creator(self):
        return self.hyperlink.Creator

    @property
    def EmailSubject(self):
        return self.hyperlink.EmailSubject

    @EmailSubject.setter
    def EmailSubject(self, value):
        self.hyperlink.EmailSubject = value

    @property
    def ExtraInfoRequired(self):
        return self.hyperlink.ExtraInfoRequired

    @property
    def Name(self):
        return self.hyperlink.Name

    @property
    def Parent(self):
        return self.hyperlink.Parent

    @property
    def Range(self):
        return Range(self.hyperlink.Range)

    @property
    def ScreenTip(self):
        return self.hyperlink.ScreenTip

    @ScreenTip.setter
    def ScreenTip(self, value):
        self.hyperlink.ScreenTip = value

    @property
    def Shape(self):
        return Shape(self.hyperlink.Shape)

    @property
    def SubAddress(self):
        return self.hyperlink.SubAddress

    @SubAddress.setter
    def SubAddress(self, value):
        self.hyperlink.SubAddress = value

    @property
    def Target(self):
        return self.hyperlink.Target

    @Target.setter
    def Target(self, value):
        self.hyperlink.Target = value

    @property
    def TextToDisplay(self):
        return self.hyperlink.TextToDisplay

    @TextToDisplay.setter
    def TextToDisplay(self, value):
        self.hyperlink.TextToDisplay = value

    @property
    def Type(self):
        return self.hyperlink.Type

    def AddToFavorites(self):
        self.hyperlink.AddToFavorites()

    def CreateNewDocument(self, *args, FileName=None, EditNow=None, Overwrite=None):
        arguments = {"FileName": FileName, "EditNow": EditNow, "Overwrite": Overwrite}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.hyperlink.CreateNewDocument(*args, **arguments)

    def Delete(self):
        self.hyperlink.Delete()

    def Follow(self, *args, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = {"NewWindow": NewWindow, "AddHistory": AddHistory, "ExtraInfo": ExtraInfo, "Method": Method, "HeaderInfo": HeaderInfo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.hyperlink.Follow(*args, **arguments)

class Index:

    def __init__(self, index=None):
        self.index = index

    @property
    def AccentedLetters(self):
        return self.index.AccentedLetters

    @property
    def Application(self):
        return Application(self.index.Application)

    @property
    def Creator(self):
        return self.index.Creator

    @property
    def Filter(self):
        return self.index.Filter

    @Filter.setter
    def Filter(self, value):
        self.index.Filter = value

    @property
    def HeadingSeparator(self):
        return WdHeadingSeparator(self.index.HeadingSeparator)

    @HeadingSeparator.setter
    def HeadingSeparator(self, value):
        self.index.HeadingSeparator = value

    @property
    def IndexLanguage(self):
        return WdLanguageID(self.index.IndexLanguage)

    @IndexLanguage.setter
    def IndexLanguage(self, value):
        self.index.IndexLanguage = value

    @property
    def NumberOfColumns(self):
        return self.index.NumberOfColumns

    @property
    def Parent(self):
        return self.index.Parent

    @property
    def Range(self):
        return Range(self.index.Range)

    @property
    def RightAlignPageNumbers(self):
        return self.index.RightAlignPageNumbers

    @property
    def SortBy(self):
        return WdIndexSortBy(self.index.SortBy)

    @SortBy.setter
    def SortBy(self, value):
        self.index.SortBy = value

    @property
    def TabLeader(self):
        return WdTabLeader(self.index.TabLeader)

    @TabLeader.setter
    def TabLeader(self, value):
        self.index.TabLeader = value

    @property
    def Type(self):
        return WdIndexType(self.index.Type)

    @Type.setter
    def Type(self, value):
        self.index.Type = value

    def Delete(self):
        self.index.Delete()

    def Update(self):
        self.index.Update()

class InlineShape:

    def __init__(self, inlineshape=None):
        self.inlineshape = inlineshape

    @property
    def AlternativeText(self):
        return self.inlineshape.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.inlineshape.AlternativeText = value

    @property
    def Application(self):
        return Application(self.inlineshape.Application)

    @property
    def Borders(self):
        return self.inlineshape.Borders

    @property
    def Chart(self):
        return Chart(self.inlineshape.Chart)

    @property
    def Creator(self):
        return self.inlineshape.Creator

    @property
    def Field(self):
        return Field(self.inlineshape.Field)

    @property
    def Fill(self):
        return FillFormat(self.inlineshape.Fill)

    @property
    def Glow(self):
        return GlowFormat(self.inlineshape.Glow)

    @property
    def GroupItems(self):
        return self.inlineshape.GroupItems

    @property
    def HasChart(self):
        return self.inlineshape.HasChart

    @property
    def HasSmartArt(self):
        return self.inlineshape.HasSmartArt

    @property
    def Height(self):
        return self.inlineshape.Height

    @Height.setter
    def Height(self, value):
        self.inlineshape.Height = value

    @property
    def HorizontalLineFormat(self):
        return HorizontalLineFormat(self.inlineshape.HorizontalLineFormat)

    @property
    def Hyperlink(self):
        return Hyperlink(self.inlineshape.Hyperlink)

    @property
    def IsPictureBullet(self):
        return self.inlineshape.IsPictureBullet

    @property
    def Line(self):
        return LineFormat(self.inlineshape.Line)

    @property
    def LinkFormat(self):
        return LinkFormat(self.inlineshape.LinkFormat)

    @property
    def LockAspectRatio(self):
        return self.inlineshape.LockAspectRatio

    @property
    def OLEFormat(self):
        return OLEFormat(self.inlineshape.OLEFormat)

    @property
    def Parent(self):
        return self.inlineshape.Parent

    @property
    def PictureFormat(self):
        return PictureFormat(self.inlineshape.PictureFormat)

    @property
    def Range(self):
        return Range(self.inlineshape.Range)

    @property
    def Reflection(self):
        return ReflectionFormat(self.inlineshape.Reflection)

    @property
    def ScaleHeight(self):
        return self.inlineshape.ScaleHeight

    @property
    def ScaleWidth(self):
        return self.inlineshape.ScaleWidth

    @property
    def Script(self):
        return self.inlineshape.Script

    @property
    def Shadow(self):
        return ShadowFormat(self.inlineshape.Shadow)

    @property
    def SmartArt(self):
        return self.inlineshape.SmartArt

    @property
    def SoftEdge(self):
        return SoftEdgeFormat(self.inlineshape.SoftEdge)

    @property
    def TextEffect(self):
        return TextEffectFormat(self.inlineshape.TextEffect)

    @property
    def Title(self):
        return self.inlineshape.Title

    @Title.setter
    def Title(self, value):
        self.inlineshape.Title = value

    @property
    def Type(self):
        return WdInlineShapeType(self.inlineshape.Type)

    @property
    def Width(self):
        return self.inlineshape.Width

    @Width.setter
    def Width(self, value):
        self.inlineshape.Width = value

    def ConvertToShape(self):
        self.inlineshape.ConvertToShape()

    def Delete(self):
        self.inlineshape.Delete()

    def Reset(self):
        self.inlineshape.Reset()

    def Select(self):
        self.inlineshape.Select()

class Interior:

    def __init__(self, interior=None):
        self.interior = interior

    @property
    def Application(self):
        return self.interior.Application

    @property
    def Color(self):
        return self.interior.Color

    @Color.setter
    def Color(self, value):
        self.interior.Color = value

    @property
    def ColorIndex(self):
        return self.interior.ColorIndex

    @ColorIndex.setter
    def ColorIndex(self, value):
        self.interior.ColorIndex = value

    @property
    def Creator(self):
        return self.interior.Creator

    @property
    def InvertIfNegative(self):
        return self.interior.InvertIfNegative

    @property
    def Parent(self):
        return self.interior.Parent

    @property
    def Pattern(self):
        return self.interior.Pattern

    @Pattern.setter
    def Pattern(self, value):
        self.interior.Pattern = value

    @property
    def PatternColor(self):
        return self.interior.PatternColor

    @PatternColor.setter
    def PatternColor(self, value):
        self.interior.PatternColor = value

    @property
    def PatternColorIndex(self):
        return self.interior.PatternColorIndex

    @PatternColorIndex.setter
    def PatternColorIndex(self, value):
        self.interior.PatternColorIndex = value

class KeyBinding:

    def __init__(self, keybinding=None):
        self.keybinding = keybinding

    @property
    def Application(self):
        return Application(self.keybinding.Application)

    @property
    def Command(self):
        return self.keybinding.Command

    @property
    def CommandParameter(self):
        return self.keybinding.CommandParameter

    @property
    def Context(self):
        return self.keybinding.Context

    @property
    def Creator(self):
        return self.keybinding.Creator

    @property
    def KeyCategory(self):
        return WdKeyCategory(self.keybinding.KeyCategory)

    @property
    def KeyCode(self):
        return self.keybinding.KeyCode

    @property
    def KeyCode2(self):
        return self.keybinding.KeyCode2

    @property
    def KeyString(self):
        return self.keybinding.KeyString

    @property
    def Parent(self):
        return self.keybinding.Parent

    @property
    def Protected(self):
        return self.keybinding.Protected

    def Clear(self):
        self.keybinding.Clear()

    def Disable(self):
        self.keybinding.Disable()

    def Execute(self):
        self.keybinding.Execute()

    def Rebind(self, *args, KeyCategory=None, Command=None, CommandParameter=None):
        arguments = {"KeyCategory": KeyCategory, "Command": Command, "CommandParameter": CommandParameter}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.keybinding.Rebind(*args, **arguments)

class Language:

    def __init__(self, language=None):
        self.language = language

    @property
    def ActiveGrammarDictionary(self):
        return Dictionary(self.language.ActiveGrammarDictionary)

    @property
    def ActiveHyphenationDictionary(self):
        return Dictionary(self.language.ActiveHyphenationDictionary)

    @property
    def ActiveSpellingDictionary(self):
        return Dictionary(self.language.ActiveSpellingDictionary)

    @property
    def ActiveThesaurusDictionary(self):
        return Dictionary(self.language.ActiveThesaurusDictionary)

    @property
    def Application(self):
        return Application(self.language.Application)

    @property
    def Creator(self):
        return self.language.Creator

    @property
    def DefaultWritingStyle(self):
        return self.language.DefaultWritingStyle

    @DefaultWritingStyle.setter
    def DefaultWritingStyle(self, value):
        self.language.DefaultWritingStyle = value

    @property
    def ID(self):
        return WdLanguageID(self.language.ID)

    @property
    def Name(self):
        return self.language.Name

    @property
    def NameLocal(self):
        return self.language.NameLocal

    @property
    def Parent(self):
        return self.language.Parent

    @property
    def SpellingDictionaryType(self):
        return WdDictionaryType(self.language.SpellingDictionaryType)

    @SpellingDictionaryType.setter
    def SpellingDictionaryType(self, value):
        self.language.SpellingDictionaryType = value

    @property
    def WritingStyleList(self):
        return self.language.WritingStyleList

class LeaderLines:

    def __init__(self, leaderlines=None):
        self.leaderlines = leaderlines

    @property
    def Application(self):
        return self.leaderlines.Application

    @property
    def Border(self):
        return ChartBorder(self.leaderlines.Border)

    @property
    def Creator(self):
        return self.leaderlines.Creator

    @property
    def Format(self):
        return ChartFormat(self.leaderlines.Format)

    @property
    def Parent(self):
        return self.leaderlines.Parent

    def Delete(self):
        self.leaderlines.Delete()

    def Select(self):
        self.leaderlines.Select()

class Legend:

    def __init__(self, legend=None):
        self.legend = legend

    @property
    def Application(self):
        return self.legend.Application

    @property
    def Creator(self):
        return self.legend.Creator

    @property
    def Format(self):
        return ChartFormat(self.legend.Format)

    @property
    def Height(self):
        return self.legend.Height

    @Height.setter
    def Height(self, value):
        self.legend.Height = value

    @property
    def IncludeInLayout(self):
        return self.legend.IncludeInLayout

    @property
    def Left(self):
        return self.legend.Left

    @property
    def Name(self):
        return self.legend.Name

    @property
    def Parent(self):
        return self.legend.Parent

    @property
    def Position(self):
        return self.legend.Position

    @Position.setter
    def Position(self, value):
        self.legend.Position = value

    @property
    def Shadow(self):
        return self.legend.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.legend.Shadow = value

    @property
    def Top(self):
        return self.legend.Top

    @Top.setter
    def Top(self, value):
        self.legend.Top = value

    @property
    def Width(self):
        return self.legend.Width

    @Width.setter
    def Width(self, value):
        self.legend.Width = value

    def Clear(self):
        self.legend.Clear()

    def Delete(self):
        self.legend.Delete()

    def LegendEntries(self):
        return LegendEntries(self.legend.LegendEntries())

    def Select(self):
        self.legend.Select()

class LegendEntries:

    def __init__(self, legendentries=None):
        self.legendentries = legendentries

    def __call__(self, item):
        return LegendEntrie(self.legendentries(item))

    @property
    def Application(self):
        return self.legendentries.Application

    @property
    def Count(self):
        return self.legendentries.Count

    @property
    def Creator(self):
        return self.legendentries.Creator

    @property
    def Parent(self):
        return self.legendentries.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return LegendEntry(self.legendentries.Item(*args, **arguments))

class LegendEntry:

    def __init__(self, legendentry=None):
        self.legendentry = legendentry

    @property
    def Application(self):
        return self.legendentry.Application

    @property
    def Creator(self):
        return self.legendentry.Creator

    @property
    def Font(self):
        return ChartFont(self.legendentry.Font)

    @property
    def Format(self):
        return ChartFormat(self.legendentry.Format)

    @property
    def Height(self):
        return self.legendentry.Height

    @property
    def Index(self):
        return self.legendentry.Index

    @property
    def Left(self):
        return self.legendentry.Left

    @property
    def LegendKey(self):
        return LegendKey(self.legendentry.LegendKey)

    @property
    def Parent(self):
        return self.legendentry.Parent

    @property
    def Top(self):
        return self.legendentry.Top

    @property
    def Width(self):
        return self.legendentry.Width

    def Delete(self):
        self.legendentry.Delete()

    def Select(self):
        self.legendentry.Select()

class LegendKey:

    def __init__(self, legendkey=None):
        self.legendkey = legendkey

    @property
    def Application(self):
        return self.legendkey.Application

    @property
    def Creator(self):
        return self.legendkey.Creator

    @property
    def Format(self):
        return ChartFormat(self.legendkey.Format)

    @property
    def Height(self):
        return self.legendkey.Height

    @property
    def InvertIfNegative(self):
        return self.legendkey.InvertIfNegative

    @property
    def Left(self):
        return self.legendkey.Left

    @property
    def MarkerBackgroundColor(self):
        return self.legendkey.MarkerBackgroundColor

    @property
    def MarkerBackgroundColorIndex(self):
        return self.legendkey.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.legendkey.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.legendkey.MarkerForegroundColor

    @property
    def MarkerForegroundColorIndex(self):
        return self.legendkey.MarkerForegroundColorIndex

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.legendkey.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.legendkey.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.legendkey.MarkerSize = value

    @property
    def MarkerStyle(self):
        return self.legendkey.MarkerStyle

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.legendkey.MarkerStyle = value

    @property
    def Parent(self):
        return self.legendkey.Parent

    @property
    def PictureType(self):
        return self.legendkey.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.legendkey.PictureType = value

    @property
    def PictureUnit2(self):
        return self.legendkey.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.legendkey.PictureUnit2 = value

    @property
    def Shadow(self):
        return self.legendkey.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.legendkey.Shadow = value

    @property
    def Smooth(self):
        return self.legendkey.Smooth

    @property
    def Top(self):
        return self.legendkey.Top

    @property
    def Width(self):
        return self.legendkey.Width

    def ClearFormats(self):
        self.legendkey.ClearFormats()

    def Delete(self):
        self.legendkey.Delete()

class LetterContent:

    def __init__(self, lettercontent=None):
        self.lettercontent = lettercontent

    @property
    def Application(self):
        return Application(self.lettercontent.Application)

    @property
    def AttentionLine(self):
        return self.lettercontent.AttentionLine

    @AttentionLine.setter
    def AttentionLine(self, value):
        self.lettercontent.AttentionLine = value

    @property
    def CCList(self):
        return self.lettercontent.CCList

    @CCList.setter
    def CCList(self, value):
        self.lettercontent.CCList = value

    @property
    def Closing(self):
        return self.lettercontent.Closing

    @Closing.setter
    def Closing(self, value):
        self.lettercontent.Closing = value

    @property
    def Creator(self):
        return self.lettercontent.Creator

    @property
    def DateFormat(self):
        return self.lettercontent.DateFormat

    @DateFormat.setter
    def DateFormat(self, value):
        self.lettercontent.DateFormat = value

    @property
    def Duplicate(self):
        return LetterContent(self.lettercontent.Duplicate)

    @property
    def EnclosureNumber(self):
        return self.lettercontent.EnclosureNumber

    @EnclosureNumber.setter
    def EnclosureNumber(self, value):
        self.lettercontent.EnclosureNumber = value

    @property
    def IncludeHeaderFooter(self):
        return self.lettercontent.IncludeHeaderFooter

    @property
    def InfoBlock(self):
        return self.lettercontent.InfoBlock

    @property
    def Letterhead(self):
        return self.lettercontent.Letterhead

    @property
    def LetterheadLocation(self):
        return WdLetterheadLocation(self.lettercontent.LetterheadLocation)

    @LetterheadLocation.setter
    def LetterheadLocation(self, value):
        self.lettercontent.LetterheadLocation = value

    @property
    def LetterheadSize(self):
        return self.lettercontent.LetterheadSize

    @LetterheadSize.setter
    def LetterheadSize(self, value):
        self.lettercontent.LetterheadSize = value

    @property
    def LetterStyle(self):
        return WdLetterStyle(self.lettercontent.LetterStyle)

    @LetterStyle.setter
    def LetterStyle(self, value):
        self.lettercontent.LetterStyle = value

    @property
    def MailingInstructions(self):
        return self.lettercontent.MailingInstructions

    @MailingInstructions.setter
    def MailingInstructions(self, value):
        self.lettercontent.MailingInstructions = value

    @property
    def PageDesign(self):
        return self.lettercontent.PageDesign

    @PageDesign.setter
    def PageDesign(self, value):
        self.lettercontent.PageDesign = value

    @property
    def Parent(self):
        return self.lettercontent.Parent

    @property
    def RecipientAddress(self):
        return self.lettercontent.RecipientAddress

    @RecipientAddress.setter
    def RecipientAddress(self, value):
        self.lettercontent.RecipientAddress = value

    @property
    def RecipientCode(self):
        return self.lettercontent.RecipientCode

    @RecipientCode.setter
    def RecipientCode(self, value):
        self.lettercontent.RecipientCode = value

    @property
    def RecipientGender(self):
        return WdSalutationGender(self.lettercontent.RecipientGender)

    @RecipientGender.setter
    def RecipientGender(self, value):
        self.lettercontent.RecipientGender = value

    @property
    def RecipientName(self):
        return self.lettercontent.RecipientName

    @RecipientName.setter
    def RecipientName(self, value):
        self.lettercontent.RecipientName = value

    @property
    def RecipientReference(self):
        return self.lettercontent.RecipientReference

    @RecipientReference.setter
    def RecipientReference(self, value):
        self.lettercontent.RecipientReference = value

    @property
    def ReturnAddress(self):
        return self.lettercontent.ReturnAddress

    @ReturnAddress.setter
    def ReturnAddress(self, value):
        self.lettercontent.ReturnAddress = value

    @property
    def ReturnAddressShortForm(self):
        return self.lettercontent.ReturnAddressShortForm

    @ReturnAddressShortForm.setter
    def ReturnAddressShortForm(self, value):
        self.lettercontent.ReturnAddressShortForm = value

    @property
    def Salutation(self):
        return self.lettercontent.Salutation

    @Salutation.setter
    def Salutation(self, value):
        self.lettercontent.Salutation = value

    @property
    def SalutationType(self):
        return WdSalutationType(self.lettercontent.SalutationType)

    @SalutationType.setter
    def SalutationType(self, value):
        self.lettercontent.SalutationType = value

    @property
    def SenderCity(self):
        return self.lettercontent.SenderCity

    @SenderCity.setter
    def SenderCity(self, value):
        self.lettercontent.SenderCity = value

    @property
    def SenderCode(self):
        return self.lettercontent.SenderCode

    @SenderCode.setter
    def SenderCode(self, value):
        self.lettercontent.SenderCode = value

    @property
    def SenderCompany(self):
        return self.lettercontent.SenderCompany

    @SenderCompany.setter
    def SenderCompany(self, value):
        self.lettercontent.SenderCompany = value

    @property
    def SenderGender(self):
        return WdSalutationGender(self.lettercontent.SenderGender)

    @SenderGender.setter
    def SenderGender(self, value):
        self.lettercontent.SenderGender = value

    @property
    def SenderInitials(self):
        return self.lettercontent.SenderInitials

    @SenderInitials.setter
    def SenderInitials(self, value):
        self.lettercontent.SenderInitials = value

    @property
    def SenderJobTitle(self):
        return self.lettercontent.SenderJobTitle

    @SenderJobTitle.setter
    def SenderJobTitle(self, value):
        self.lettercontent.SenderJobTitle = value

    @property
    def SenderName(self):
        return self.lettercontent.SenderName

    @SenderName.setter
    def SenderName(self, value):
        self.lettercontent.SenderName = value

    @property
    def SenderReference(self):
        return self.lettercontent.SenderReference

    @property
    def Subject(self):
        return self.lettercontent.Subject

    @Subject.setter
    def Subject(self, value):
        self.lettercontent.Subject = value

class Line:

    def __init__(self, line=None):
        self.line = line

    @property
    def Application(self):
        return Application(self.line.Application)

    @property
    def Creator(self):
        return self.line.Creator

    @property
    def Height(self):
        return self.line.Height

    @Height.setter
    def Height(self, value):
        self.line.Height = value

    @property
    def Left(self):
        return self.line.Left

    @property
    def LineType(self):
        return wdLineType(self.line.LineType)

    @property
    def Parent(self):
        return self.line.Parent

    @property
    def Range(self):
        return Range(self.line.Range)

    @property
    def Rectangles(self):
        return Rectangles(self.line.Rectangles)

    @property
    def Top(self):
        return self.line.Top

    @property
    def Width(self):
        return self.line.Width

class LineFormat:

    def __init__(self, lineformat=None):
        self.lineformat = lineformat

    @property
    def Application(self):
        return Application(self.lineformat.Application)

    @property
    def BackColor(self):
        return ColorFormat(self.lineformat.BackColor)

    @BackColor.setter
    def BackColor(self, value):
        self.lineformat.BackColor = value

    @property
    def BeginArrowheadLength(self):
        return self.lineformat.BeginArrowheadLength

    @BeginArrowheadLength.setter
    def BeginArrowheadLength(self, value):
        self.lineformat.BeginArrowheadLength = value

    @property
    def BeginArrowheadStyle(self):
        return self.lineformat.BeginArrowheadStyle

    @BeginArrowheadStyle.setter
    def BeginArrowheadStyle(self, value):
        self.lineformat.BeginArrowheadStyle = value

    @property
    def BeginArrowheadWidth(self):
        return self.lineformat.BeginArrowheadWidth

    @BeginArrowheadWidth.setter
    def BeginArrowheadWidth(self, value):
        self.lineformat.BeginArrowheadWidth = value

    @property
    def Creator(self):
        return self.lineformat.Creator

    @property
    def DashStyle(self):
        return self.lineformat.DashStyle

    @DashStyle.setter
    def DashStyle(self, value):
        self.lineformat.DashStyle = value

    @property
    def EndArrowheadLength(self):
        return self.lineformat.EndArrowheadLength

    @EndArrowheadLength.setter
    def EndArrowheadLength(self, value):
        self.lineformat.EndArrowheadLength = value

    @property
    def EndArrowheadStyle(self):
        return self.lineformat.EndArrowheadStyle

    @EndArrowheadStyle.setter
    def EndArrowheadStyle(self, value):
        self.lineformat.EndArrowheadStyle = value

    @property
    def EndArrowheadWidth(self):
        return self.lineformat.EndArrowheadWidth

    @EndArrowheadWidth.setter
    def EndArrowheadWidth(self, value):
        self.lineformat.EndArrowheadWidth = value

    @property
    def ForeColor(self):
        return ColorFormat(self.lineformat.ForeColor)

    @ForeColor.setter
    def ForeColor(self, value):
        self.lineformat.ForeColor = value

    @property
    def InsetPen(self):
        return self.lineformat.InsetPen

    @property
    def Parent(self):
        return self.lineformat.Parent

    @property
    def Pattern(self):
        return self.lineformat.Pattern

    @Pattern.setter
    def Pattern(self, value):
        self.lineformat.Pattern = value

    @property
    def Style(self):
        return self.lineformat.Style

    @Style.setter
    def Style(self, value):
        self.lineformat.Style = value

    @property
    def Transparency(self):
        return self.lineformat.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.lineformat.Transparency = value

    @property
    def Visible(self):
        return self.lineformat.Visible

    @property
    def Weight(self):
        return self.lineformat.Weight

    @Weight.setter
    def Weight(self, value):
        self.lineformat.Weight = value

class LineNumbering:

    def __init__(self, linenumbering=None):
        self.linenumbering = linenumbering

    @property
    def Active(self):
        return self.linenumbering.Active

    @property
    def Application(self):
        return Application(self.linenumbering.Application)

    @property
    def CountBy(self):
        return self.linenumbering.CountBy

    @CountBy.setter
    def CountBy(self, value):
        self.linenumbering.CountBy = value

    @property
    def Creator(self):
        return self.linenumbering.Creator

    @property
    def DistanceFromText(self):
        return self.linenumbering.DistanceFromText

    @DistanceFromText.setter
    def DistanceFromText(self, value):
        self.linenumbering.DistanceFromText = value

    @property
    def Parent(self):
        return self.linenumbering.Parent

    @property
    def RestartMode(self):
        return WdNumberingRule(self.linenumbering.RestartMode)

    @RestartMode.setter
    def RestartMode(self, value):
        self.linenumbering.RestartMode = value

    @property
    def StartingNumber(self):
        return self.linenumbering.StartingNumber

    @StartingNumber.setter
    def StartingNumber(self, value):
        self.linenumbering.StartingNumber = value

class Lines:

    def __init__(self, lines=None):
        self.lines = lines

    def __call__(self, item):
        return Line(self.lines(item))

    @property
    def Application(self):
        return Application(self.lines.Application)

    @property
    def Count(self):
        return self.lines.Count

    @property
    def Creator(self):
        return self.lines.Creator

    @property
    def Parent(self):
        return self.lines.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.lines.Item(*args, **arguments)

class LinkFormat:

    def __init__(self, linkformat=None):
        self.linkformat = linkformat

    @property
    def Application(self):
        return Application(self.linkformat.Application)

    @property
    def AutoUpdate(self):
        return self.linkformat.AutoUpdate

    @property
    def Creator(self):
        return self.linkformat.Creator

    @property
    def Locked(self):
        return self.linkformat.Locked

    @property
    def Parent(self):
        return self.linkformat.Parent

    @property
    def SavePictureWithDocument(self):
        return self.linkformat.SavePictureWithDocument

    @property
    def SourceFullName(self):
        return self.linkformat.SourceFullName

    @SourceFullName.setter
    def SourceFullName(self, value):
        self.linkformat.SourceFullName = value

    @property
    def SourceName(self):
        return self.linkformat.SourceName

    @property
    def SourcePath(self):
        return self.linkformat.SourcePath

    @property
    def Type(self):
        return WdLinkType(self.linkformat.Type)

    def BreakLink(self):
        self.linkformat.BreakLink()

    def Update(self):
        self.linkformat.Update()

class List:

    def __init__(self, list=None):
        self.list = list

    @property
    def Application(self):
        return Application(self.list.Application)

    @property
    def Creator(self):
        return self.list.Creator

    @property
    def ListParagraphs(self):
        return self.list.ListParagraphs

    @property
    def Parent(self):
        return self.list.Parent

    @property
    def Range(self):
        return Range(self.list.Range)

    @property
    def SingleListTemplate(self):
        return self.list.SingleListTemplate

    @property
    def StyleName(self):
        return self.list.StyleName

    def ApplyListTemplate(self, *args, ListTemplate=None, ContinuePreviousList=None, ApplyTo=None, DefaultListBehavior=None):
        arguments = {"ListTemplate": ListTemplate, "ContinuePreviousList": ContinuePreviousList, "ApplyTo": ApplyTo, "DefaultListBehavior": DefaultListBehavior}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.list.ApplyListTemplate(*args, **arguments)

    def ApplyListTemplateWithLevel(self, *args, ListTemplate=None, ContinuePreviousList=None, DefaultListBehavior=None, ApplyLevel=None):
        arguments = {"ListTemplate": ListTemplate, "ContinuePreviousList": ContinuePreviousList, "DefaultListBehavior": DefaultListBehavior, "ApplyLevel": ApplyLevel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.list.ApplyListTemplateWithLevel(*args, **arguments)

    def CanContinuePreviousList(self, *args, ListTemplate=None):
        arguments = {"ListTemplate": ListTemplate}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.list.CanContinuePreviousList(*args, **arguments)

    def ConvertNumbersToText(self):
        self.list.ConvertNumbersToText()

    def CountNumberedItems(self):
        self.list.CountNumberedItems()

    def RemoveNumbers(self, *args, NumberType=None):
        arguments = {"NumberType": NumberType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.list.RemoveNumbers(*args, **arguments)

class ListEntry:

    def __init__(self, listentry=None):
        self.listentry = listentry

    @property
    def Application(self):
        return Application(self.listentry.Application)

    @property
    def Creator(self):
        return self.listentry.Creator

    @property
    def Index(self):
        return self.listentry.Index

    @property
    def Name(self):
        return self.listentry.Name

    @Name.setter
    def Name(self, value):
        self.listentry.Name = value

    @property
    def Parent(self):
        return self.listentry.Parent

    def Delete(self):
        self.listentry.Delete()

class ListFormat:

    def __init__(self, listformat=None):
        self.listformat = listformat

    @property
    def Application(self):
        return Application(self.listformat.Application)

    @property
    def Creator(self):
        return self.listformat.Creator

    @property
    def List(self):
        return List(self.listformat.List)

    @property
    def ListLevelNumber(self):
        return ListFormat(self.listformat.ListLevelNumber)

    @ListLevelNumber.setter
    def ListLevelNumber(self, value):
        self.listformat.ListLevelNumber = value

    @property
    def ListPictureBullet(self):
        return InlineShape(self.listformat.ListPictureBullet)

    @property
    def ListString(self):
        return self.listformat.ListString

    @property
    def ListTemplate(self):
        return ListTemplate(self.listformat.ListTemplate)

    @property
    def ListType(self):
        return ListFormat(self.listformat.ListType)

    @property
    def ListValue(self):
        return ListFormat(self.listformat.ListValue)

    @property
    def Parent(self):
        return self.listformat.Parent

    @property
    def SingleList(self):
        return self.listformat.SingleList

    @property
    def SingleListTemplate(self):
        return self.listformat.SingleListTemplate

    def ApplyBulletDefault(self, *args, DefaultListBehavior=None):
        arguments = {"DefaultListBehavior": DefaultListBehavior}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listformat.ApplyBulletDefault(*args, **arguments)

    def ApplyListTemplate(self, *args, ListTemplate=None, ContinuePreviousList=None, ApplyTo=None, DefaultListBehavior=None):
        arguments = {"ListTemplate": ListTemplate, "ContinuePreviousList": ContinuePreviousList, "ApplyTo": ApplyTo, "DefaultListBehavior": DefaultListBehavior}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listformat.ApplyListTemplate(*args, **arguments)

    def ApplyListTemplateWithLevel(self, *args, ListTemplate=None, ContinuePreviousList=None, ApplyTo=None, DefaultListBehavior=None, ApplyLevel=None):
        arguments = {"ListTemplate": ListTemplate, "ContinuePreviousList": ContinuePreviousList, "ApplyTo": ApplyTo, "DefaultListBehavior": DefaultListBehavior, "ApplyLevel": ApplyLevel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listformat.ApplyListTemplateWithLevel(*args, **arguments)

    def ApplyNumberDefault(self, *args, DefaultListBehavior=None):
        arguments = {"DefaultListBehavior": DefaultListBehavior}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listformat.ApplyNumberDefault(*args, **arguments)

    def ApplyOutlineNumberDefault(self, *args, DefaultListBehavior=None):
        arguments = {"DefaultListBehavior": DefaultListBehavior}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listformat.ApplyOutlineNumberDefault(*args, **arguments)

    def CanContinuePreviousList(self, *args, ListTemplate=None):
        arguments = {"ListTemplate": ListTemplate}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listformat.CanContinuePreviousList(*args, **arguments)

    def ConvertNumbersToText(self):
        self.listformat.ConvertNumbersToText()

    def CountNumberedItems(self):
        self.listformat.CountNumberedItems()

    def ListIndent(self):
        self.listformat.ListIndent()

    def ListOutdent(self):
        self.listformat.ListOutdent()

    def RemoveNumbers(self, *args, NumberType=None):
        arguments = {"NumberType": NumberType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listformat.RemoveNumbers(*args, **arguments)

class ListGallery:

    def __init__(self, listgallery=None):
        self.listgallery = listgallery

    @property
    def Application(self):
        return Application(self.listgallery.Application)

    @property
    def Creator(self):
        return self.listgallery.Creator

    @property
    def ListTemplates(self):
        return self.listgallery.ListTemplates

    def Modified(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.listgallery.Modified(*args, **arguments)

    @property
    def Parent(self):
        return self.listgallery.Parent

    def Reset(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listgallery.Reset(*args, **arguments)

class ListLevel:

    def __init__(self, listlevel=None):
        self.listlevel = listlevel

    @property
    def Alignment(self):
        return WdListLevelAlignment(self.listlevel.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.listlevel.Alignment = value

    @property
    def Application(self):
        return Application(self.listlevel.Application)

    @property
    def Creator(self):
        return self.listlevel.Creator

    @property
    def Font(self):
        return Font(self.listlevel.Font)

    @Font.setter
    def Font(self, value):
        self.listlevel.Font = value

    @property
    def Index(self):
        return self.listlevel.Index

    @property
    def LinkedStyle(self):
        return ListLevel(self.listlevel.LinkedStyle)

    @LinkedStyle.setter
    def LinkedStyle(self, value):
        self.listlevel.LinkedStyle = value

    @property
    def NumberFormat(self):
        return self.listlevel.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.listlevel.NumberFormat = value

    @property
    def NumberPosition(self):
        return ListLevel(self.listlevel.NumberPosition)

    @NumberPosition.setter
    def NumberPosition(self, value):
        self.listlevel.NumberPosition = value

    @property
    def NumberStyle(self):
        return ListLevel(self.listlevel.NumberStyle)

    @NumberStyle.setter
    def NumberStyle(self, value):
        self.listlevel.NumberStyle = value

    @property
    def Parent(self):
        return self.listlevel.Parent

    @property
    def PictureBullet(self):
        return InlineShape(self.listlevel.PictureBullet)

    @property
    def ResetOnHigher(self):
        return self.listlevel.ResetOnHigher

    @property
    def StartAt(self):
        return ListLevel(self.listlevel.StartAt)

    @StartAt.setter
    def StartAt(self, value):
        self.listlevel.StartAt = value

    @property
    def TabPosition(self):
        return ListLevel(self.listlevel.TabPosition)

    @TabPosition.setter
    def TabPosition(self, value):
        self.listlevel.TabPosition = value

    @property
    def TextPosition(self):
        return ListLevel(self.listlevel.TextPosition)

    @TextPosition.setter
    def TextPosition(self, value):
        self.listlevel.TextPosition = value

    @property
    def TrailingCharacter(self):
        return WdTrailingCharacter(self.listlevel.TrailingCharacter)

    @TrailingCharacter.setter
    def TrailingCharacter(self, value):
        self.listlevel.TrailingCharacter = value

    def ApplyPictureBullet(self, *args, FileName=None):
        arguments = {"FileName": FileName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listlevel.ApplyPictureBullet(*args, **arguments)

class ListTemplate:

    def __init__(self, listtemplate=None):
        self.listtemplate = listtemplate

    @property
    def Application(self):
        return Application(self.listtemplate.Application)

    @property
    def Creator(self):
        return self.listtemplate.Creator

    @property
    def ListLevels(self):
        return self.listtemplate.ListLevels

    @property
    def Name(self):
        return self.listtemplate.Name

    @Name.setter
    def Name(self, value):
        self.listtemplate.Name = value

    @property
    def OutlineNumbered(self):
        return self.listtemplate.OutlineNumbered

    @property
    def Parent(self):
        return self.listtemplate.Parent

    def Convert(self, *args, Level=None):
        arguments = {"Level": Level}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.listtemplate.Convert(*args, **arguments)

class MailingLabel:

    def __init__(self, mailinglabel=None):
        self.mailinglabel = mailinglabel

    @property
    def Application(self):
        return Application(self.mailinglabel.Application)

    @property
    def Creator(self):
        return self.mailinglabel.Creator

    @property
    def CustomLabels(self):
        return self.mailinglabel.CustomLabels

    @property
    def DefaultLabelName(self):
        return self.mailinglabel.DefaultLabelName

    @DefaultLabelName.setter
    def DefaultLabelName(self, value):
        self.mailinglabel.DefaultLabelName = value

    @property
    def DefaultLaserTray(self):
        return WdPaperTray(self.mailinglabel.DefaultLaserTray)

    @DefaultLaserTray.setter
    def DefaultLaserTray(self, value):
        self.mailinglabel.DefaultLaserTray = value

    @property
    def Parent(self):
        return self.mailinglabel.Parent

    @property
    def Vertical(self):
        return self.mailinglabel.Vertical

    def CreateNewDocument(self, *args, Name=None, Address=None, AutoText=None, ExtractAddress=None, LaserTray=None, PrintEPostageLabel=None, Vertical=None):
        arguments = {"Name": Name, "Address": Address, "AutoText": AutoText, "ExtractAddress": ExtractAddress, "LaserTray": LaserTray, "PrintEPostageLabel": PrintEPostageLabel, "Vertical": Vertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mailinglabel.CreateNewDocument(*args, **arguments)

    def CreateNewDocumentByID(self, *args, LabelID=None, Address=None, AutoText=None, ExtractAddress=None, LaserTray=None, PrintEPostageLabel=None, Vertical=None):
        arguments = {"LabelID": LabelID, "Address": Address, "AutoText": AutoText, "ExtractAddress": ExtractAddress, "LaserTray": LaserTray, "PrintEPostageLabel": PrintEPostageLabel, "Vertical": Vertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mailinglabel.CreateNewDocumentByID(*args, **arguments)

    def LabelOptions(self):
        self.mailinglabel.LabelOptions()

    def PrintOut(self, *args, Name=None, Address=None, ExtractAddress=None, LaserTray=None, SingleLabel=None, Row=None, Column=None, PrintEPostageLabel=None, Vertical=None):
        arguments = {"Name": Name, "Address": Address, "ExtractAddress": ExtractAddress, "LaserTray": LaserTray, "SingleLabel": SingleLabel, "Row": Row, "Column": Column, "PrintEPostageLabel": PrintEPostageLabel, "Vertical": Vertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailinglabel.PrintOut(*args, **arguments)

    def PrintOutByID(self, *args, LabelID=None, Address=None, ExtractAddress=None, LaserTray=None, SingleLabel=None, Row=None, Column=None, PrintEPostageLabel=None, Vertical=None):
        arguments = {"LabelID": LabelID, "Address": Address, "ExtractAddress": ExtractAddress, "LaserTray": LaserTray, "SingleLabel": SingleLabel, "Row": Row, "Column": Column, "PrintEPostageLabel": PrintEPostageLabel, "Vertical": Vertical}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailinglabel.PrintOutByID(*args, **arguments)

class MailMerge:

    def __init__(self, mailmerge=None):
        self.mailmerge = mailmerge

    @property
    def Application(self):
        return Application(self.mailmerge.Application)

    @property
    def Creator(self):
        return self.mailmerge.Creator

    @property
    def DataSource(self):
        return MailMergeDataSource(self.mailmerge.DataSource)

    @property
    def Destination(self):
        return WdMailMergeDestination(self.mailmerge.Destination)

    @Destination.setter
    def Destination(self, value):
        self.mailmerge.Destination = value

    @property
    def Fields(self):
        return self.mailmerge.Fields

    @property
    def HighlightMergeFields(self):
        return self.mailmerge.HighlightMergeFields

    @property
    def MailAddressFieldName(self):
        return self.mailmerge.MailAddressFieldName

    @MailAddressFieldName.setter
    def MailAddressFieldName(self, value):
        self.mailmerge.MailAddressFieldName = value

    @property
    def MailAsAttachment(self):
        return self.mailmerge.MailAsAttachment

    @property
    def MailFormat(self):
        return WdMailMergeMailFormat(self.mailmerge.MailFormat)

    @property
    def MailSubject(self):
        return self.mailmerge.MailSubject

    @MailSubject.setter
    def MailSubject(self, value):
        self.mailmerge.MailSubject = value

    @property
    def MainDocumentType(self):
        return WdMailMergeMainDocType(self.mailmerge.MainDocumentType)

    @MainDocumentType.setter
    def MainDocumentType(self, value):
        self.mailmerge.MainDocumentType = value

    @property
    def Parent(self):
        return self.mailmerge.Parent

    @property
    def ShowSendToCustom(self):
        return self.mailmerge.ShowSendToCustom

    @ShowSendToCustom.setter
    def ShowSendToCustom(self, value):
        self.mailmerge.ShowSendToCustom = value

    @property
    def State(self):
        return WdMailMergeState(self.mailmerge.State)

    @property
    def SuppressBlankLines(self):
        return self.mailmerge.SuppressBlankLines

    @property
    def ViewMailMergeFieldCodes(self):
        return self.mailmerge.ViewMailMergeFieldCodes

    @property
    def WizardState(self):
        return self.mailmerge.WizardState

    @WizardState.setter
    def WizardState(self, value):
        self.mailmerge.WizardState = value

    def Check(self):
        self.mailmerge.Check()

    def CreateDataSource(self, *args, Name=None, PasswordDocument=None, WritePasswordDocument=None, HeaderRecord=None, MSQuery=None, SQLStatement=None, SQLStatement1=None, Connection=None, LinkToSource=None):
        arguments = {"Name": Name, "PasswordDocument": PasswordDocument, "WritePasswordDocument": WritePasswordDocument, "HeaderRecord": HeaderRecord, "MSQuery": MSQuery, "SQLStatement": SQLStatement, "SQLStatement1": SQLStatement1, "Connection": Connection, "LinkToSource": LinkToSource}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailmerge.CreateDataSource(*args, **arguments)

    def CreateHeaderSource(self, *args, Name=None, PasswordDocument=None, WritePasswordDocument=None, HeaderRecord=None):
        arguments = {"Name": Name, "PasswordDocument": PasswordDocument, "WritePasswordDocument": WritePasswordDocument, "HeaderRecord": HeaderRecord}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailmerge.CreateHeaderSource(*args, **arguments)

    def EditDataSource(self):
        self.mailmerge.EditDataSource()

    def EditHeaderSource(self):
        self.mailmerge.EditHeaderSource()

    def EditMainDocument(self):
        self.mailmerge.EditMainDocument()

    def Execute(self, *args, Pause=None):
        arguments = {"Pause": Pause}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailmerge.Execute(*args, **arguments)

    def OpenDataSource(self, *args, Name=None, Format=None, ConfirmConversions=None, ReadOnly=None, LinkToSource=None, AddToRecentFiles=None, PasswordDocument=None, PasswordTemplate=None, Revert=None, WritePasswordDocument=None, WritePasswordTemplate=None, Connection=None, SQLStatement=None, SQLStatement1=None, OpenExclusive=None, SubType=None):
        arguments = {"Name": Name, "Format": Format, "ConfirmConversions": ConfirmConversions, "ReadOnly": ReadOnly, "LinkToSource": LinkToSource, "AddToRecentFiles": AddToRecentFiles, "PasswordDocument": PasswordDocument, "PasswordTemplate": PasswordTemplate, "Revert": Revert, "WritePasswordDocument": WritePasswordDocument, "WritePasswordTemplate": WritePasswordTemplate, "Connection": Connection, "SQLStatement": SQLStatement, "SQLStatement1": SQLStatement1, "OpenExclusive": OpenExclusive, "SubType": SubType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailmerge.OpenDataSource(*args, **arguments)

    def OpenHeaderSource(self, *args, Name=None, Format=None, ConfirmConversions=None, ReadOnly=None, AddToRecentFiles=None, PasswordDocument=None, PasswordTemplate=None, Revert=None, WritePasswordDocument=None, WritePasswordTemplate=None, OpenExclusive=None):
        arguments = {"Name": Name, "Format": Format, "ConfirmConversions": ConfirmConversions, "ReadOnly": ReadOnly, "AddToRecentFiles": AddToRecentFiles, "PasswordDocument": PasswordDocument, "PasswordTemplate": PasswordTemplate, "Revert": Revert, "WritePasswordDocument": WritePasswordDocument, "WritePasswordTemplate": WritePasswordTemplate, "OpenExclusive": OpenExclusive}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailmerge.OpenHeaderSource(*args, **arguments)

    def ShowWizard(self, *args, InitialState=None, ShowDocumentStep=None, ShowTemplateStep=None, ShowDataStep=None, ShowWriteStep=None, ShowPreviewStep=None, ShowMergeStep=None):
        arguments = {"InitialState": InitialState, "ShowDocumentStep": ShowDocumentStep, "ShowTemplateStep": ShowTemplateStep, "ShowDataStep": ShowDataStep, "ShowWriteStep": ShowWriteStep, "ShowPreviewStep": ShowPreviewStep, "ShowMergeStep": ShowMergeStep}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailmerge.ShowWizard(*args, **arguments)

class MailMergeDataField:

    def __init__(self, mailmergedatafield=None):
        self.mailmergedatafield = mailmergedatafield

    @property
    def Application(self):
        return Application(self.mailmergedatafield.Application)

    @property
    def Creator(self):
        return self.mailmergedatafield.Creator

    @property
    def Index(self):
        return self.mailmergedatafield.Index

    @property
    def Name(self):
        return self.mailmergedatafield.Name

    @property
    def Parent(self):
        return self.mailmergedatafield.Parent

    @property
    def Value(self):
        return self.mailmergedatafield.Value

class MailMergeDataSource:

    def __init__(self, mailmergedatasource=None):
        self.mailmergedatasource = mailmergedatasource

    @property
    def ActiveRecord(self):
        return WdMailMergeActiveRecord(self.mailmergedatasource.ActiveRecord)

    @ActiveRecord.setter
    def ActiveRecord(self, value):
        self.mailmergedatasource.ActiveRecord = value

    @property
    def Application(self):
        return Application(self.mailmergedatasource.Application)

    @property
    def ConnectString(self):
        return self.mailmergedatasource.ConnectString

    @property
    def Creator(self):
        return self.mailmergedatasource.Creator

    @property
    def DataFields(self):
        return self.mailmergedatasource.DataFields

    @property
    def FieldNames(self):
        return MailMergeFieldNames(self.mailmergedatasource.FieldNames)

    @property
    def FirstRecord(self):
        return self.mailmergedatasource.FirstRecord

    @FirstRecord.setter
    def FirstRecord(self, value):
        self.mailmergedatasource.FirstRecord = value

    @property
    def HeaderSourceName(self):
        return self.mailmergedatasource.HeaderSourceName

    @property
    def HeaderSourceType(self):
        return WdMailMergeDataSource(self.mailmergedatasource.HeaderSourceType)

    @property
    def Included(self):
        return self.mailmergedatasource.Included

    @property
    def InvalidAddress(self):
        return self.mailmergedatasource.InvalidAddress

    @property
    def InvalidComments(self):
        return self.mailmergedatasource.InvalidComments

    @property
    def LastRecord(self):
        return self.mailmergedatasource.LastRecord

    @LastRecord.setter
    def LastRecord(self, value):
        self.mailmergedatasource.LastRecord = value

    @property
    def MappedDataFields(self):
        return MappedDataFields(self.mailmergedatasource.MappedDataFields)

    @property
    def Name(self):
        return self.mailmergedatasource.Name

    @property
    def Parent(self):
        return self.mailmergedatasource.Parent

    @property
    def QueryString(self):
        return self.mailmergedatasource.QueryString

    @QueryString.setter
    def QueryString(self, value):
        self.mailmergedatasource.QueryString = value

    @property
    def RecordCount(self):
        return self.mailmergedatasource.RecordCount

    @property
    def TableName(self):
        return self.mailmergedatasource.TableName

    @property
    def Type(self):
        return WdMailMergeDataSource(self.mailmergedatasource.Type)

    def Close(self):
        self.mailmergedatasource.Close()

    def FindRecord(self, *args, FindText=None, Field=None):
        arguments = {"FindText": FindText, "Field": Field}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mailmergedatasource.FindRecord(*args, **arguments)

    def SetAllErrorFlags(self, *args, Invalid=None, InvalidComment=None):
        arguments = {"Invalid": Invalid, "InvalidComment": InvalidComment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailmergedatasource.SetAllErrorFlags(*args, **arguments)

    def SetAllIncludedFlags(self, *args, Included=None):
        arguments = {"Included": Included}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.mailmergedatasource.SetAllIncludedFlags(*args, **arguments)

class MailMergeField:

    def __init__(self, mailmergefield=None):
        self.mailmergefield = mailmergefield

    @property
    def Application(self):
        return Application(self.mailmergefield.Application)

    @property
    def Code(self):
        return Range(self.mailmergefield.Code)

    @property
    def Creator(self):
        return self.mailmergefield.Creator

    @property
    def Locked(self):
        return self.mailmergefield.Locked

    @property
    def Next(self):
        return self.mailmergefield.Next

    @property
    def Parent(self):
        return self.mailmergefield.Parent

    @property
    def Previous(self):
        return self.mailmergefield.Previous

    @property
    def Type(self):
        return WdFieldType(self.mailmergefield.Type)

    def Copy(self):
        self.mailmergefield.Copy()

    def Cut(self):
        self.mailmergefield.Cut()

    def Delete(self):
        self.mailmergefield.Delete()

    def Select(self):
        self.mailmergefield.Select()

class MailMergeFieldName:

    def __init__(self, mailmergefieldname=None):
        self.mailmergefieldname = mailmergefieldname

    @property
    def Application(self):
        return Application(self.mailmergefieldname.Application)

    @property
    def Creator(self):
        return self.mailmergefieldname.Creator

    @property
    def Index(self):
        return self.mailmergefieldname.Index

    @property
    def Name(self):
        return self.mailmergefieldname.Name

    @property
    def Parent(self):
        return self.mailmergefieldname.Parent

class MailMessage:

    def __init__(self, mailmessage=None):
        self.mailmessage = mailmessage

    @property
    def Application(self):
        return Application(self.mailmessage.Application)

    @property
    def Creator(self):
        return self.mailmessage.Creator

    @property
    def Parent(self):
        return self.mailmessage.Parent

    def CheckName(self):
        self.mailmessage.CheckName()

    def Delete(self):
        self.mailmessage.Delete()

    def DisplayMoveDialog(self):
        self.mailmessage.DisplayMoveDialog()

    def DisplayProperties(self):
        self.mailmessage.DisplayProperties()

    def DisplaySelectNamesDialog(self):
        self.mailmessage.DisplaySelectNamesDialog()

    def Forward(self):
        self.mailmessage.Forward()

    def GoToNext(self):
        self.mailmessage.GoToNext()

    def GoToPrevious(self):
        self.mailmessage.GoToPrevious()

    def Reply(self):
        self.mailmessage.Reply()

    def ReplyAll(self):
        self.mailmessage.ReplyAll()

    def ToggleHeader(self):
        self.mailmessage.ToggleHeader()

class MappedDataField:

    def __init__(self, mappeddatafield=None):
        self.mappeddatafield = mappeddatafield

    @property
    def Application(self):
        return Application(self.mappeddatafield.Application)

    @property
    def Creator(self):
        return self.mappeddatafield.Creator

    @property
    def DataFieldIndex(self):
        return self.mappeddatafield.DataFieldIndex

    @DataFieldIndex.setter
    def DataFieldIndex(self, value):
        self.mappeddatafield.DataFieldIndex = value

    @property
    def DataFieldName(self):
        return self.mappeddatafield.DataFieldName

    @property
    def Index(self):
        return self.mappeddatafield.Index

    @property
    def Name(self):
        return self.mappeddatafield.Name

    @property
    def Parent(self):
        return self.mappeddatafield.Parent

    @property
    def Value(self):
        return self.mappeddatafield.Value

class MappedDataFields:

    def __init__(self, mappeddatafields=None):
        self.mappeddatafields = mappeddatafields

    def __call__(self, item):
        return MappedDataField(self.mappeddatafields(item))

    @property
    def Application(self):
        return Application(self.mappeddatafields.Application)

    @property
    def Count(self):
        return self.mappeddatafields.Count

    @property
    def Creator(self):
        return self.mappeddatafields.Creator

    @property
    def Parent(self):
        return self.mappeddatafields.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.mappeddatafields.Item(*args, **arguments)

class Model3DFormat:

    def __init__(self, model3dformat=None):
        self.model3dformat = model3dformat

    @property
    def Application(self):
        return Application(self.model3dformat.Application)

    @property
    def AutoFit(self):
        return self.model3dformat.AutoFit

    @property
    def CameraPositionX(self):
        return self.model3dformat.CameraPositionX

    @property
    def CameraPositionY(self):
        return self.model3dformat.CameraPositionY

    @property
    def CameraPositionZ(self):
        return self.model3dformat.CameraPositionZ

    @property
    def Creator(self):
        return self.model3dformat.Creator

    @property
    def FieldOfView(self):
        return self.model3dformat.FieldOfView

    @property
    def LookAtPointX(self):
        return self.model3dformat.LookAtPointX

    @property
    def LookAtPointY(self):
        return self.model3dformat.LookAtPointY

    @property
    def LookAtPointZ(self):
        return self.model3dformat.LookAtPointZ

    @property
    def Parent(self):
        return self.model3dformat.Parent

    @property
    def RotationX(self):
        return self.model3dformat.RotationX

    @property
    def RotationY(self):
        return self.model3dformat.RotationY

    @property
    def RotationZ(self):
        return self.model3dformat.RotationZ

    def IncrementRotationX(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.model3dformat.IncrementRotationX(*args, **arguments)

    def IncrementRotationY(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.model3dformat.IncrementRotationY(*args, **arguments)

    def IncrementRotationZ(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.model3dformat.IncrementRotationZ(*args, **arguments)

    def ResetModel(self, *args, ResetSize=None):
        arguments = {"ResetSize": ResetSize}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.model3dformat.ResetModel(*args, **arguments)

class OLEFormat:

    def __init__(self, oleformat=None):
        self.oleformat = oleformat

    @property
    def Application(self):
        return Application(self.oleformat.Application)

    @property
    def ClassType(self):
        return self.oleformat.ClassType

    @ClassType.setter
    def ClassType(self, value):
        self.oleformat.ClassType = value

    @property
    def Creator(self):
        return self.oleformat.Creator

    @property
    def DisplayAsIcon(self):
        return self.oleformat.DisplayAsIcon

    @property
    def IconIndex(self):
        return self.oleformat.IconIndex

    @IconIndex.setter
    def IconIndex(self, value):
        self.oleformat.IconIndex = value

    @property
    def IconLabel(self):
        return self.oleformat.IconLabel

    @IconLabel.setter
    def IconLabel(self, value):
        self.oleformat.IconLabel = value

    @property
    def IconName(self):
        return self.oleformat.IconName

    @IconName.setter
    def IconName(self, value):
        self.oleformat.IconName = value

    @property
    def IconPath(self):
        return self.oleformat.IconPath

    @property
    def Label(self):
        return self.oleformat.Label

    @property
    def Object(self):
        return self.oleformat.Object

    @property
    def Parent(self):
        return self.oleformat.Parent

    @property
    def PreserveFormattingOnUpdate(self):
        return self.oleformat.PreserveFormattingOnUpdate

    @property
    def ProgID(self):
        return self.oleformat.ProgID

    def Activate(self):
        self.oleformat.Activate()

    def ActivateAs(self, *args, ClassType=None):
        arguments = {"ClassType": ClassType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.oleformat.ActivateAs(*args, **arguments)

    def ConvertTo(self, *args, ClassType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None):
        arguments = {"ClassType": ClassType, "DisplayAsIcon": DisplayAsIcon, "IconFileName": IconFileName, "IconIndex": IconIndex, "IconLabel": IconLabel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.oleformat.ConvertTo(*args, **arguments)

    def DoVerb(self, *args, VerbIndex=None):
        arguments = {"VerbIndex": VerbIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.oleformat.DoVerb(*args, **arguments)

    def Edit(self):
        self.oleformat.Edit()

    def Open(self):
        self.oleformat.Open()

class OMath:

    def __init__(self, omath=None):
        self.omath = omath

    @property
    def AlignPoint(self):
        return self.omath.AlignPoint

    @AlignPoint.setter
    def AlignPoint(self, value):
        self.omath.AlignPoint = value

    @property
    def Application(self):
        return Application(self.omath.Application)

    @property
    def ArgIndex(self):
        return self.omath.ArgIndex

    @property
    def ArgSize(self):
        return self.omath.ArgSize

    @ArgSize.setter
    def ArgSize(self, value):
        self.omath.ArgSize = value

    @property
    def Breaks(self):
        return OMathBreaks(self.omath.Breaks)

    @property
    def Creator(self):
        return self.omath.Creator

    @property
    def Functions(self):
        return OMathFunctions(self.omath.Functions)

    @property
    def Justification(self):
        return WdOMathJc(self.omath.Justification)

    @Justification.setter
    def Justification(self, value):
        self.omath.Justification = value

    @property
    def NestingLevel(self):
        return self.omath.NestingLevel

    @property
    def Parent(self):
        return self.omath.Parent

    @property
    def ParentArg(self):
        return OMath(self.omath.ParentArg)

    @property
    def ParentCol(self):
        return OMathMatCol(self.omath.ParentCol)

    @property
    def ParentFunction(self):
        return OMathFunction(self.omath.ParentFunction)

    @property
    def ParentOMath(self):
        return OMath(self.omath.ParentOMath)

    @property
    def ParentRow(self):
        return OMathMatRow(self.omath.ParentRow)

    @property
    def Range(self):
        return Range(self.omath.Range)

    @property
    def Type(self):
        return WdOMathType(self.omath.Type)

    @Type.setter
    def Type(self, value):
        self.omath.Type = value

    def BuildUp(self):
        return self.omath.BuildUp()

    def ConvertToLiteralText(self):
        self.omath.ConvertToLiteralText()

    def ConvertToMathText(self):
        self.omath.ConvertToMathText()

    def ConvertToNormalText(self):
        self.omath.ConvertToNormalText()

    def Linearize(self):
        return self.omath.Linearize()

    def Remove(self):
        self.omath.Remove()

class OMathAcc:

    def __init__(self, omathacc=None):
        self.omathacc = omathacc

    @property
    def Application(self):
        return Application(self.omathacc.Application)

    @property
    def Char(self):
        return self.omathacc.Char

    @Char.setter
    def Char(self, value):
        self.omathacc.Char = value

    @property
    def Creator(self):
        return self.omathacc.Creator

    @property
    def E(self):
        return OMath(self.omathacc.E)

    @property
    def Parent(self):
        return self.omathacc.Parent

class OMathArgs:

    def __init__(self, omathargs=None):
        self.omathargs = omathargs

    @property
    def Application(self):
        return Application(self.omathargs.Application)

    @property
    def Count(self):
        return OMathArgs(self.omathargs.Count)

    @property
    def Creator(self):
        return self.omathargs.Creator

    @property
    def Parent(self):
        return self.omathargs.Parent

    def Add(self, *args, BeforeArg=None):
        arguments = {"BeforeArg": BeforeArg}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathargs.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathargs.Item(*args, **arguments)

class OMathAutoCorrect:

    def __init__(self, omathautocorrect=None):
        self.omathautocorrect = omathautocorrect

    @property
    def Application(self):
        return Application(self.omathautocorrect.Application)

    @property
    def Creator(self):
        return self.omathautocorrect.Creator

    @property
    def Entries(self):
        return OMathAutoCorrectEntries(self.omathautocorrect.Entries)

    @property
    def Functions(self):
        return OMathRecognizedFunctions(self.omathautocorrect.Functions)

    @property
    def Parent(self):
        return self.omathautocorrect.Parent

    @property
    def ReplaceText(self):
        return self.omathautocorrect.ReplaceText

    @ReplaceText.setter
    def ReplaceText(self, value):
        self.omathautocorrect.ReplaceText = value

    @property
    def UseOutsideOMath(self):
        return self.omathautocorrect.UseOutsideOMath

    @UseOutsideOMath.setter
    def UseOutsideOMath(self, value):
        self.omathautocorrect.UseOutsideOMath = value

class OMathAutoCorrectEntries:

    def __init__(self, omathautocorrectentries=None):
        self.omathautocorrectentries = omathautocorrectentries

    @property
    def Application(self):
        return Application(self.omathautocorrectentries.Application)

    @property
    def Count(self):
        return OMathAutoCorrectEntries(self.omathautocorrectentries.Count)

    @property
    def Creator(self):
        return self.omathautocorrectentries.Creator

    @property
    def Parent(self):
        return self.omathautocorrectentries.Parent

    def Add(self, *args, Name=None, Value=None):
        arguments = {"Name": Name, "Value": Value}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathautocorrectentries.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathautocorrectentries.Item(*args, **arguments)

class OMathAutoCorrectEntry:

    def __init__(self, omathautocorrectentry=None):
        self.omathautocorrectentry = omathautocorrectentry

    @property
    def Application(self):
        return Application(self.omathautocorrectentry.Application)

    @property
    def Creator(self):
        return self.omathautocorrectentry.Creator

    @property
    def Index(self):
        return self.omathautocorrectentry.Index

    @property
    def Name(self):
        return self.omathautocorrectentry.Name

    @Name.setter
    def Name(self, value):
        self.omathautocorrectentry.Name = value

    @property
    def Parent(self):
        return self.omathautocorrectentry.Parent

    @property
    def Value(self):
        return self.omathautocorrectentry.Value

    @Value.setter
    def Value(self, value):
        self.omathautocorrectentry.Value = value

    def Delete(self):
        self.omathautocorrectentry.Delete()

class OMathBar:

    def __init__(self, omathbar=None):
        self.omathbar = omathbar

    @property
    def Application(self):
        return Application(self.omathbar.Application)

    @property
    def BarTop(self):
        return self.omathbar.BarTop

    @BarTop.setter
    def BarTop(self, value):
        self.omathbar.BarTop = value

    @property
    def Creator(self):
        return self.omathbar.Creator

    @property
    def E(self):
        return OMath(self.omathbar.E)

    @property
    def Parent(self):
        return self.omathbar.Parent

class OMathBorderBox:

    def __init__(self, omathborderbox=None):
        self.omathborderbox = omathborderbox

    @property
    def Application(self):
        return Application(self.omathborderbox.Application)

    @property
    def Creator(self):
        return self.omathborderbox.Creator

    @property
    def E(self):
        return OMath(self.omathborderbox.E)

    @property
    def HideBot(self):
        return self.omathborderbox.HideBot

    @HideBot.setter
    def HideBot(self, value):
        self.omathborderbox.HideBot = value

    @property
    def HideLeft(self):
        return self.omathborderbox.HideLeft

    @HideLeft.setter
    def HideLeft(self, value):
        self.omathborderbox.HideLeft = value

    @property
    def HideRight(self):
        return self.omathborderbox.HideRight

    @HideRight.setter
    def HideRight(self, value):
        self.omathborderbox.HideRight = value

    @property
    def HideTop(self):
        return self.omathborderbox.HideTop

    @HideTop.setter
    def HideTop(self, value):
        self.omathborderbox.HideTop = value

    @property
    def Parent(self):
        return self.omathborderbox.Parent

    @property
    def StrikeBLTR(self):
        return self.omathborderbox.StrikeBLTR

    @StrikeBLTR.setter
    def StrikeBLTR(self, value):
        self.omathborderbox.StrikeBLTR = value

    @property
    def StrikeH(self):
        return self.omathborderbox.StrikeH

    @StrikeH.setter
    def StrikeH(self, value):
        self.omathborderbox.StrikeH = value

    @property
    def StrikeTLBR(self):
        return self.omathborderbox.StrikeTLBR

    @StrikeTLBR.setter
    def StrikeTLBR(self, value):
        self.omathborderbox.StrikeTLBR = value

    @property
    def StrikeV(self):
        return self.omathborderbox.StrikeV

    @StrikeV.setter
    def StrikeV(self, value):
        self.omathborderbox.StrikeV = value

class OMathBox:

    def __init__(self, omathbox=None):
        self.omathbox = omathbox

    @property
    def Application(self):
        return Application(self.omathbox.Application)

    @property
    def Creator(self):
        return self.omathbox.Creator

    @property
    def Diff(self):
        return self.omathbox.Diff

    @Diff.setter
    def Diff(self, value):
        self.omathbox.Diff = value

    @property
    def E(self):
        return OMath(self.omathbox.E)

    @property
    def NoBreak(self):
        return self.omathbox.NoBreak

    @NoBreak.setter
    def NoBreak(self, value):
        self.omathbox.NoBreak = value

    @property
    def OpEmu(self):
        return self.omathbox.OpEmu

    @OpEmu.setter
    def OpEmu(self, value):
        self.omathbox.OpEmu = value

    @property
    def Parent(self):
        return self.omathbox.Parent

class OMathBreak:

    def __init__(self, omathbreak=None):
        self.omathbreak = omathbreak

    @property
    def AlignAt(self):
        return self.omathbreak.AlignAt

    @AlignAt.setter
    def AlignAt(self, value):
        self.omathbreak.AlignAt = value

    @property
    def Application(self):
        return Application(self.omathbreak.Application)

    @property
    def Creator(self):
        return self.omathbreak.Creator

    @property
    def Parent(self):
        return self.omathbreak.Parent

    @property
    def Range(self):
        return Range(self.omathbreak.Range)

    def Delete(self):
        self.omathbreak.Delete()

class OMathBreaks:

    def __init__(self, omathbreaks=None):
        self.omathbreaks = omathbreaks

    @property
    def Application(self):
        return Application(self.omathbreaks.Application)

    @property
    def Count(self):
        return OMathBreaks(self.omathbreaks.Count)

    @property
    def Creator(self):
        return self.omathbreaks.Creator

    @property
    def Parent(self):
        return self.omathbreaks.Parent

    def Add(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathbreaks.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathbreaks.Item(*args, **arguments)

class OMathDelim:

    def __init__(self, omathdelim=None):
        self.omathdelim = omathdelim

    @property
    def Application(self):
        return Application(self.omathdelim.Application)

    @property
    def BegChar(self):
        return self.omathdelim.BegChar

    @BegChar.setter
    def BegChar(self, value):
        self.omathdelim.BegChar = value

    @property
    def Creator(self):
        return self.omathdelim.Creator

    @property
    def E(self):
        return OMathArgs(self.omathdelim.E)

    @property
    def EndChar(self):
        return self.omathdelim.EndChar

    @EndChar.setter
    def EndChar(self, value):
        self.omathdelim.EndChar = value

    @property
    def Grow(self):
        return self.omathdelim.Grow

    @Grow.setter
    def Grow(self, value):
        self.omathdelim.Grow = value

    @property
    def NoLeftChar(self):
        return self.omathdelim.NoLeftChar

    @NoLeftChar.setter
    def NoLeftChar(self, value):
        self.omathdelim.NoLeftChar = value

    @property
    def NoRightChar(self):
        return self.omathdelim.NoRightChar

    @NoRightChar.setter
    def NoRightChar(self, value):
        self.omathdelim.NoRightChar = value

    @property
    def Parent(self):
        return self.omathdelim.Parent

    @property
    def SepChar(self):
        return self.omathdelim.SepChar

    @SepChar.setter
    def SepChar(self, value):
        self.omathdelim.SepChar = value

    @property
    def Shape(self):
        return WdOMathShapeType(self.omathdelim.Shape)

    @Shape.setter
    def Shape(self, value):
        self.omathdelim.Shape = value

class OMathEqArray:

    def __init__(self, omatheqarray=None):
        self.omatheqarray = omatheqarray

    @property
    def Align(self):
        return WdOMathVertAlignType(self.omatheqarray.Align)

    @Align.setter
    def Align(self, value):
        self.omatheqarray.Align = value

    @property
    def Application(self):
        return Application(self.omatheqarray.Application)

    @property
    def Creator(self):
        return self.omatheqarray.Creator

    @property
    def E(self):
        return OMathArgs(self.omatheqarray.E)

    @property
    def MaxDist(self):
        return self.omatheqarray.MaxDist

    @MaxDist.setter
    def MaxDist(self, value):
        self.omatheqarray.MaxDist = value

    @property
    def ObjDist(self):
        return self.omatheqarray.ObjDist

    @ObjDist.setter
    def ObjDist(self, value):
        self.omatheqarray.ObjDist = value

    @property
    def Parent(self):
        return self.omatheqarray.Parent

    @property
    def RowSpacing(self):
        return self.omatheqarray.RowSpacing

    @RowSpacing.setter
    def RowSpacing(self, value):
        self.omatheqarray.RowSpacing = value

    @property
    def RowSpacingRule(self):
        return WdOMathSpacingRule(self.omatheqarray.RowSpacingRule)

    @RowSpacingRule.setter
    def RowSpacingRule(self, value):
        self.omatheqarray.RowSpacingRule = value

class OMathFrac:

    def __init__(self, omathfrac=None):
        self.omathfrac = omathfrac

    @property
    def Application(self):
        return Application(self.omathfrac.Application)

    @property
    def Creator(self):
        return self.omathfrac.Creator

    @property
    def Den(self):
        return OMath(self.omathfrac.Den)

    @property
    def Num(self):
        return OMath(self.omathfrac.Num)

    @property
    def Parent(self):
        return self.omathfrac.Parent

    @property
    def Type(self):
        return WdOMathFracType(self.omathfrac.Type)

    @Type.setter
    def Type(self, value):
        self.omathfrac.Type = value

class OMathFunc:

    def __init__(self, omathfunc=None):
        self.omathfunc = omathfunc

    @property
    def Application(self):
        return Application(self.omathfunc.Application)

    @property
    def Creator(self):
        return self.omathfunc.Creator

    @property
    def E(self):
        return OMath(self.omathfunc.E)

    @property
    def FName(self):
        return OMath(self.omathfunc.FName)

    @property
    def Parent(self):
        return self.omathfunc.Parent

class OMathFunction:

    def __init__(self, omathfunction=None):
        self.omathfunction = omathfunction

    @property
    def Acc(self):
        return OMathAcc(self.omathfunction.Acc)

    @property
    def Application(self):
        return Application(self.omathfunction.Application)

    @property
    def Args(self):
        return OMathArgs(self.omathfunction.Args)

    @property
    def Bar(self):
        return OMathBar(self.omathfunction.Bar)

    @property
    def BorderBox(self):
        return OMathBorderBox(self.omathfunction.BorderBox)

    @property
    def Box(self):
        return OMathBox(self.omathfunction.Box)

    @property
    def Creator(self):
        return self.omathfunction.Creator

    @property
    def Delim(self):
        return OMathDelim(self.omathfunction.Delim)

    @property
    def EqArray(self):
        return OMathEqArray(self.omathfunction.EqArray)

    @property
    def Frac(self):
        return OMathFrac(self.omathfunction.Frac)

    @property
    def Func(self):
        return OMathFunc(self.omathfunction.Func)

    @property
    def GroupChar(self):
        return OMathGroupChar(self.omathfunction.GroupChar)

    @property
    def LimLow(self):
        return OMathLimLow(self.omathfunction.LimLow)

    @property
    def LimUpp(self):
        return OMathLimUpp(self.omathfunction.LimUpp)

    @property
    def Mat(self):
        return OMathMat(self.omathfunction.Mat)

    @property
    def Nary(self):
        return OMathNary(self.omathfunction.Nary)

    @property
    def OMath(self):
        return OMath(self.omathfunction.OMath)

    @property
    def Parent(self):
        return self.omathfunction.Parent

    @property
    def Phantom(self):
        return OMathPhantom(self.omathfunction.Phantom)

    @property
    def Rad(self):
        return OMathRad(self.omathfunction.Rad)

    @property
    def Range(self):
        return Range(self.omathfunction.Range)

    @property
    def ScrPre(self):
        return OMathScrPre(self.omathfunction.ScrPre)

    @property
    def ScrSub(self):
        return self.omathfunction.ScrSub

    @property
    def ScrSubSup(self):
        return OMathScrSubSup(self.omathfunction.ScrSubSup)

    @property
    def ScrSup(self):
        return OMathScrSup(self.omathfunction.ScrSup)

    @property
    def Type(self):
        return WdOMathFunctionType(self.omathfunction.Type)

    def Remove(self):
        return self.omathfunction.Remove()

class OMathFunctions:

    def __init__(self, omathfunctions=None):
        self.omathfunctions = omathfunctions

    @property
    def Application(self):
        return Application(self.omathfunctions.Application)

    @property
    def Count(self):
        return OMathFunctions(self.omathfunctions.Count)

    @property
    def Creator(self):
        return self.omathfunctions.Creator

    @property
    def Parent(self):
        return self.omathfunctions.Parent

    def Add(self, *args, Range=None, Type=None, NumArgs=None, NumCols=None):
        arguments = {"Range": Range, "Type": Type, "NumArgs": NumArgs, "NumCols": NumCols}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathfunctions.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathfunctions.Item(*args, **arguments)

class OMathGroupChar:

    def __init__(self, omathgroupchar=None):
        self.omathgroupchar = omathgroupchar

    @property
    def AlignTop(self):
        return self.omathgroupchar.AlignTop

    @AlignTop.setter
    def AlignTop(self, value):
        self.omathgroupchar.AlignTop = value

    @property
    def Application(self):
        return Application(self.omathgroupchar.Application)

    @property
    def Char(self):
        return self.omathgroupchar.Char

    @Char.setter
    def Char(self, value):
        self.omathgroupchar.Char = value

    @property
    def CharTop(self):
        return self.omathgroupchar.CharTop

    @CharTop.setter
    def CharTop(self, value):
        self.omathgroupchar.CharTop = value

    @property
    def Creator(self):
        return self.omathgroupchar.Creator

    @property
    def E(self):
        return OMath(self.omathgroupchar.E)

    @property
    def Parent(self):
        return self.omathgroupchar.Parent

class OMathLimLow:

    def __init__(self, omathlimlow=None):
        self.omathlimlow = omathlimlow

    @property
    def Application(self):
        return Application(self.omathlimlow.Application)

    @property
    def Creator(self):
        return self.omathlimlow.Creator

    @property
    def E(self):
        return OMath(self.omathlimlow.E)

    @property
    def Lim(self):
        return OMath(self.omathlimlow.Lim)

    @property
    def Parent(self):
        return self.omathlimlow.Parent

    def ToLimUpp(self):
        return self.omathlimlow.ToLimUpp()

class OMathLimUpp:

    def __init__(self, omathlimupp=None):
        self.omathlimupp = omathlimupp

    @property
    def Application(self):
        return Application(self.omathlimupp.Application)

    @property
    def Creator(self):
        return self.omathlimupp.Creator

    @property
    def E(self):
        return OMath(self.omathlimupp.E)

    @property
    def Lim(self):
        return OMath(self.omathlimupp.Lim)

    @property
    def Parent(self):
        return self.omathlimupp.Parent

    def ToLimLow(self):
        return self.omathlimupp.ToLimLow()

class OMathMat:

    def __init__(self, omathmat=None):
        self.omathmat = omathmat

    @property
    def Align(self):
        return WdOMathVertAlignType(self.omathmat.Align)

    @Align.setter
    def Align(self, value):
        self.omathmat.Align = value

    @property
    def Application(self):
        return Application(self.omathmat.Application)

    def Cell(self, *args, Row=None, Col=None):
        arguments = {"Row": Row, "Col": Col}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return OMath(self.omathmat.Cell(*args, **arguments))

    @property
    def ColGap(self):
        return self.omathmat.ColGap

    @ColGap.setter
    def ColGap(self, value):
        self.omathmat.ColGap = value

    @property
    def ColGapRule(self):
        return WdOMathSpacingRule(self.omathmat.ColGapRule)

    @ColGapRule.setter
    def ColGapRule(self, value):
        self.omathmat.ColGapRule = value

    @property
    def Cols(self):
        return OMathMatCols(self.omathmat.Cols)

    @property
    def ColSpacing(self):
        return self.omathmat.ColSpacing

    @ColSpacing.setter
    def ColSpacing(self, value):
        self.omathmat.ColSpacing = value

    @property
    def Creator(self):
        return self.omathmat.Creator

    @property
    def Parent(self):
        return self.omathmat.Parent

    @property
    def PlcHoldHidden(self):
        return self.omathmat.PlcHoldHidden

    @PlcHoldHidden.setter
    def PlcHoldHidden(self, value):
        self.omathmat.PlcHoldHidden = value

    @property
    def Rows(self):
        return OMathMatRows(self.omathmat.Rows)

    @property
    def RowSpacing(self):
        return self.omathmat.RowSpacing

    @RowSpacing.setter
    def RowSpacing(self, value):
        self.omathmat.RowSpacing = value

    @property
    def RowSpacingRule(self):
        return WdOMathSpacingRule(self.omathmat.RowSpacingRule)

    @RowSpacingRule.setter
    def RowSpacingRule(self, value):
        self.omathmat.RowSpacingRule = value

class OMathMatCol:

    def __init__(self, omathmatcol=None):
        self.omathmatcol = omathmatcol

    @property
    def Align(self):
        return WdOMathHorizAlignType(self.omathmatcol.Align)

    @Align.setter
    def Align(self, value):
        self.omathmatcol.Align = value

    @property
    def Application(self):
        return Application(self.omathmatcol.Application)

    @property
    def Args(self):
        return OMathArgs(self.omathmatcol.Args)

    @property
    def ColIndex(self):
        return self.omathmatcol.ColIndex

    @property
    def Creator(self):
        return self.omathmatcol.Creator

    @property
    def Parent(self):
        return self.omathmatcol.Parent

    def Delete(self):
        self.omathmatcol.Delete()

class OMathMatCols:

    def __init__(self, omathmatcols=None):
        self.omathmatcols = omathmatcols

    @property
    def Application(self):
        return Application(self.omathmatcols.Application)

    @property
    def Count(self):
        return OMathMatCols(self.omathmatcols.Count)

    @property
    def Creator(self):
        return self.omathmatcols.Creator

    @property
    def Parent(self):
        return self.omathmatcols.Parent

    def Add(self, *args, BeforeCol=None):
        arguments = {"BeforeCol": BeforeCol}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathmatcols.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathmatcols.Item(*args, **arguments)

class OMathMatRow:

    def __init__(self, omathmatrow=None):
        self.omathmatrow = omathmatrow

    @property
    def Application(self):
        return Application(self.omathmatrow.Application)

    @property
    def Args(self):
        return OMathArgs(self.omathmatrow.Args)

    @property
    def Creator(self):
        return self.omathmatrow.Creator

    @property
    def Parent(self):
        return self.omathmatrow.Parent

    @property
    def RowIndex(self):
        return self.omathmatrow.RowIndex

    @RowIndex.setter
    def RowIndex(self, value):
        self.omathmatrow.RowIndex = value

    def Delete(self):
        self.omathmatrow.Delete()

class OMathMatRows:

    def __init__(self, omathmatrows=None):
        self.omathmatrows = omathmatrows

    @property
    def Application(self):
        return Application(self.omathmatrows.Application)

    @property
    def Count(self):
        return OMathMatRows(self.omathmatrows.Count)

    @property
    def Creator(self):
        return self.omathmatrows.Creator

    @property
    def Parent(self):
        return self.omathmatrows.Parent

    def Add(self, *args, BeforeRow=None):
        arguments = {"BeforeRow": BeforeRow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathmatrows.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathmatrows.Item(*args, **arguments)

class OMathNary:

    def __init__(self, omathnary=None):
        self.omathnary = omathnary

    @property
    def Application(self):
        return Application(self.omathnary.Application)

    @property
    def Char(self):
        return self.omathnary.Char

    @Char.setter
    def Char(self, value):
        self.omathnary.Char = value

    @property
    def Creator(self):
        return self.omathnary.Creator

    @property
    def E(self):
        return OMath(self.omathnary.E)

    @property
    def Grow(self):
        return self.omathnary.Grow

    @Grow.setter
    def Grow(self, value):
        self.omathnary.Grow = value

    @property
    def HideSub(self):
        return self.omathnary.HideSub

    @HideSub.setter
    def HideSub(self, value):
        self.omathnary.HideSub = value

    @property
    def HideSup(self):
        return self.omathnary.HideSup

    @HideSup.setter
    def HideSup(self, value):
        self.omathnary.HideSup = value

    @property
    def Parent(self):
        return self.omathnary.Parent

    @property
    def Sub(self):
        return OMath(self.omathnary.Sub)

    @property
    def SubSupLim(self):
        return self.omathnary.SubSupLim

    @SubSupLim.setter
    def SubSupLim(self, value):
        self.omathnary.SubSupLim = value

    @property
    def Sup(self):
        return OMath(self.omathnary.Sup)

class OMathPhantom:

    def __init__(self, omathphantom=None):
        self.omathphantom = omathphantom

    @property
    def Application(self):
        return Application(self.omathphantom.Application)

    @property
    def Creator(self):
        return self.omathphantom.Creator

    @property
    def E(self):
        return OMath(self.omathphantom.E)

    @property
    def Parent(self):
        return self.omathphantom.Parent

    @property
    def Show(self):
        return self.omathphantom.Show

    @Show.setter
    def Show(self, value):
        self.omathphantom.Show = value

    @property
    def Smash(self):
        return self.omathphantom.Smash

    @Smash.setter
    def Smash(self, value):
        self.omathphantom.Smash = value

    @property
    def Transp(self):
        return self.omathphantom.Transp

    @Transp.setter
    def Transp(self, value):
        self.omathphantom.Transp = value

    @property
    def ZeroAsc(self):
        return self.omathphantom.ZeroAsc

    @ZeroAsc.setter
    def ZeroAsc(self, value):
        self.omathphantom.ZeroAsc = value

    @property
    def ZeroDesc(self):
        return self.omathphantom.ZeroDesc

    @ZeroDesc.setter
    def ZeroDesc(self, value):
        self.omathphantom.ZeroDesc = value

    @property
    def ZeroWid(self):
        return self.omathphantom.ZeroWid

    @ZeroWid.setter
    def ZeroWid(self, value):
        self.omathphantom.ZeroWid = value

class OMathRad:

    def __init__(self, omathrad=None):
        self.omathrad = omathrad

    @property
    def Application(self):
        return Application(self.omathrad.Application)

    @property
    def Creator(self):
        return self.omathrad.Creator

    @property
    def Deg(self):
        return OMath(self.omathrad.Deg)

    @property
    def E(self):
        return OMath(self.omathrad.E)

    @property
    def HideDeg(self):
        return self.omathrad.HideDeg

    @HideDeg.setter
    def HideDeg(self, value):
        self.omathrad.HideDeg = value

    @property
    def Parent(self):
        return self.omathrad.Parent

class OMathRecognizedFunction:

    def __init__(self, omathrecognizedfunction=None):
        self.omathrecognizedfunction = omathrecognizedfunction

    @property
    def Application(self):
        return Application(self.omathrecognizedfunction.Application)

    @property
    def Creator(self):
        return self.omathrecognizedfunction.Creator

    @property
    def Index(self):
        return self.omathrecognizedfunction.Index

    @property
    def Name(self):
        return self.omathrecognizedfunction.Name

    @property
    def Parent(self):
        return self.omathrecognizedfunction.Parent

    def Delete(self):
        self.omathrecognizedfunction.Delete()

class OMathRecognizedFunctions:

    def __init__(self, omathrecognizedfunctions=None):
        self.omathrecognizedfunctions = omathrecognizedfunctions

    @property
    def Application(self):
        return Application(self.omathrecognizedfunctions.Application)

    @property
    def Count(self):
        return OMathRecognizedFunctions(self.omathrecognizedfunctions.Count)

    @property
    def Creator(self):
        return self.omathrecognizedfunctions.Creator

    @property
    def Parent(self):
        return self.omathrecognizedfunctions.Parent

    def Add(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathrecognizedfunctions.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omathrecognizedfunctions.Item(*args, **arguments)

class OMaths:

    def __init__(self, omaths=None):
        self.omaths = omaths

    def __call__(self, item):
        return OMath(self.omaths(item))

    @property
    def Application(self):
        return Application(self.omaths.Application)

    @property
    def Count(self):
        return OMaths(self.omaths.Count)

    @property
    def Creator(self):
        return self.omaths.Creator

    @property
    def Parent(self):
        return self.omaths.Parent

    def Add(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return OMath(self.omaths.Add(*args, **arguments))

    def BuildUp(self):
        return self.omaths.BuildUp()

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.omaths.Item(*args, **arguments)

    def Linearize(self):
        return self.omaths.Linearize()

class OMathScrPre:

    def __init__(self, omathscrpre=None):
        self.omathscrpre = omathscrpre

    @property
    def Application(self):
        return Application(self.omathscrpre.Application)

    @property
    def Creator(self):
        return self.omathscrpre.Creator

    @property
    def E(self):
        return OMath(self.omathscrpre.E)

    @property
    def Parent(self):
        return self.omathscrpre.Parent

    @property
    def Sub(self):
        return OMath(self.omathscrpre.Sub)

    @property
    def Sup(self):
        return OMath(self.omathscrpre.Sup)

    def ToScrSubSup(self):
        return self.omathscrpre.ToScrSubSup()

class OMathScrSub:

    def __init__(self, omathscrsub=None):
        self.omathscrsub = omathscrsub

    @property
    def Application(self):
        return Application(self.omathscrsub.Application)

    @property
    def Creator(self):
        return self.omathscrsub.Creator

    @property
    def E(self):
        return OMath(self.omathscrsub.E)

    @property
    def Parent(self):
        return self.omathscrsub.Parent

    @property
    def Sub(self):
        return OMath(self.omathscrsub.Sub)

class OMathScrSubSup:

    def __init__(self, omathscrsubsup=None):
        self.omathscrsubsup = omathscrsubsup

    @property
    def AlignScripts(self):
        return self.omathscrsubsup.AlignScripts

    @AlignScripts.setter
    def AlignScripts(self, value):
        self.omathscrsubsup.AlignScripts = value

    @property
    def Application(self):
        return Application(self.omathscrsubsup.Application)

    @property
    def Creator(self):
        return self.omathscrsubsup.Creator

    @property
    def E(self):
        return OMath(self.omathscrsubsup.E)

    @property
    def Parent(self):
        return self.omathscrsubsup.Parent

    @property
    def Sub(self):
        return OMath(self.omathscrsubsup.Sub)

    @property
    def Sup(self):
        return OMath(self.omathscrsubsup.Sup)

    def RemoveSub(self):
        return self.omathscrsubsup.RemoveSub()

    def RemoveSup(self):
        return self.omathscrsubsup.RemoveSup()

    def ToScrPre(self):
        return self.omathscrsubsup.ToScrPre()

class OMathScrSup:

    def __init__(self, omathscrsup=None):
        self.omathscrsup = omathscrsup

    @property
    def Application(self):
        return Application(self.omathscrsup.Application)

    @property
    def Creator(self):
        return self.omathscrsup.Creator

    @property
    def E(self):
        return OMath(self.omathscrsup.E)

    @property
    def Parent(self):
        return self.omathscrsup.Parent

    @property
    def Sup(self):
        return OMath(self.omathscrsup.Sup)

class Options:

    def __init__(self, options=None):
        self.options = options

    @property
    def AddBiDirectionalMarksWhenSavingTextFile(self):
        return self.options.AddBiDirectionalMarksWhenSavingTextFile

    @property
    def AddControlCharacters(self):
        return self.options.AddControlCharacters

    @property
    def AddHebDoubleQuote(self):
        return self.options.AddHebDoubleQuote

    @property
    def AllowAccentedUppercase(self):
        return self.options.AllowAccentedUppercase

    @property
    def AllowClickAndTypeMouse(self):
        return self.options.AllowClickAndTypeMouse

    @property
    def AllowCombinedAuxiliaryForms(self):
        return self.options.AllowCombinedAuxiliaryForms

    @property
    def AllowCompoundNounProcessing(self):
        return self.options.AllowCompoundNounProcessing

    @property
    def AllowDragAndDrop(self):
        return self.options.AllowDragAndDrop

    @property
    def AllowOpenInDraftView(self):
        return self.options.AllowOpenInDraftView

    @AllowOpenInDraftView.setter
    def AllowOpenInDraftView(self, value):
        self.options.AllowOpenInDraftView = value

    @property
    def AllowPixelUnits(self):
        return self.options.AllowPixelUnits

    @property
    def AllowReadingMode(self):
        return self.options.AllowReadingMode

    @property
    def AnimateScreenMovements(self):
        return self.options.AnimateScreenMovements

    @property
    def Application(self):
        return Application(self.options.Application)

    @property
    def ApplyFarEastFontsToAscii(self):
        return self.options.ApplyFarEastFontsToAscii

    @property
    def ArabicMode(self):
        return WdAraSpeller(self.options.ArabicMode)

    @ArabicMode.setter
    def ArabicMode(self, value):
        self.options.ArabicMode = value

    @property
    def ArabicNumeral(self):
        return WdArabicNumeral(self.options.ArabicNumeral)

    @ArabicNumeral.setter
    def ArabicNumeral(self, value):
        self.options.ArabicNumeral = value

    @property
    def AutoCreateNewDrawings(self):
        return self.options.AutoCreateNewDrawings

    @property
    def AutoFormatApplyBulletedLists(self):
        return self.options.AutoFormatApplyBulletedLists

    @property
    def AutoFormatApplyFirstIndents(self):
        return self.options.AutoFormatApplyFirstIndents

    @property
    def AutoFormatApplyHeadings(self):
        return self.options.AutoFormatApplyHeadings

    @property
    def AutoFormatApplyLists(self):
        return self.options.AutoFormatApplyLists

    @property
    def AutoFormatApplyOtherParas(self):
        return self.options.AutoFormatApplyOtherParas

    @property
    def AutoFormatAsYouTypeApplyBorders(self):
        return self.options.AutoFormatAsYouTypeApplyBorders

    @property
    def AutoFormatAsYouTypeApplyBulletedLists(self):
        return self.options.AutoFormatAsYouTypeApplyBulletedLists

    @property
    def AutoFormatAsYouTypeApplyClosings(self):
        return self.options.AutoFormatAsYouTypeApplyClosings

    @property
    def AutoFormatAsYouTypeApplyDates(self):
        return self.options.AutoFormatAsYouTypeApplyDates

    @property
    def AutoFormatAsYouTypeApplyFirstIndents(self):
        return self.options.AutoFormatAsYouTypeApplyFirstIndents

    @property
    def AutoFormatAsYouTypeApplyHeadings(self):
        return self.options.AutoFormatAsYouTypeApplyHeadings

    @property
    def AutoFormatAsYouTypeApplyNumberedLists(self):
        return self.options.AutoFormatAsYouTypeApplyNumberedLists

    @property
    def AutoFormatAsYouTypeApplyTables(self):
        return self.options.AutoFormatAsYouTypeApplyTables

    @property
    def AutoFormatAsYouTypeAutoLetterWizard(self):
        return self.options.AutoFormatAsYouTypeAutoLetterWizard

    @property
    def AutoFormatAsYouTypeDefineStyles(self):
        return self.options.AutoFormatAsYouTypeDefineStyles

    @property
    def AutoFormatAsYouTypeDeleteAutoSpaces(self):
        return self.options.AutoFormatAsYouTypeDeleteAutoSpaces

    @property
    def AutoFormatAsYouTypeFormatListItemBeginning(self):
        return self.options.AutoFormatAsYouTypeFormatListItemBeginning

    @property
    def AutoFormatAsYouTypeInsertClosings(self):
        return self.options.AutoFormatAsYouTypeInsertClosings

    @property
    def AutoFormatAsYouTypeInsertOvers(self):
        return self.options.AutoFormatAsYouTypeInsertOvers

    @property
    def AutoFormatAsYouTypeMatchParentheses(self):
        return self.options.AutoFormatAsYouTypeMatchParentheses

    @property
    def AutoFormatAsYouTypeReplaceFarEastDashes(self):
        return self.options.AutoFormatAsYouTypeReplaceFarEastDashes

    @property
    def AutoFormatAsYouTypeReplaceFractions(self):
        return self.options.AutoFormatAsYouTypeReplaceFractions

    @property
    def AutoFormatAsYouTypeReplaceHyperlinks(self):
        return self.options.AutoFormatAsYouTypeReplaceHyperlinks

    @property
    def AutoFormatAsYouTypeReplaceOrdinals(self):
        return self.options.AutoFormatAsYouTypeReplaceOrdinals

    @property
    def AutoFormatAsYouTypeReplacePlainTextEmphasis(self):
        return self.options.AutoFormatAsYouTypeReplacePlainTextEmphasis

    @property
    def AutoFormatAsYouTypeReplaceQuotes(self):
        return self.options.AutoFormatAsYouTypeReplaceQuotes

    @property
    def AutoFormatAsYouTypeReplaceSymbols(self):
        return self.options.AutoFormatAsYouTypeReplaceSymbols

    @property
    def AutoFormatDeleteAutoSpaces(self):
        return self.options.AutoFormatDeleteAutoSpaces

    @property
    def AutoFormatMatchParentheses(self):
        return self.options.AutoFormatMatchParentheses

    @property
    def AutoFormatPlainTextWordMail(self):
        return self.options.AutoFormatPlainTextWordMail

    @property
    def AutoFormatPreserveStyles(self):
        return self.options.AutoFormatPreserveStyles

    @property
    def AutoFormatReplaceFarEastDashes(self):
        return self.options.AutoFormatReplaceFarEastDashes

    @property
    def AutoFormatReplaceFractions(self):
        return self.options.AutoFormatReplaceFractions

    @property
    def AutoFormatReplaceHyperlinks(self):
        return self.options.AutoFormatReplaceHyperlinks

    @property
    def AutoFormatReplaceOrdinals(self):
        return self.options.AutoFormatReplaceOrdinals

    @property
    def AutoFormatReplacePlainTextEmphasis(self):
        return self.options.AutoFormatReplacePlainTextEmphasis

    @property
    def AutoFormatReplaceQuotes(self):
        return self.options.AutoFormatReplaceQuotes

    @property
    def AutoFormatReplaceSymbols(self):
        return self.options.AutoFormatReplaceSymbols

    @property
    def AutoKeyboardSwitching(self):
        return self.options.AutoKeyboardSwitching

    @property
    def AutoWordSelection(self):
        return self.options.AutoWordSelection

    @property
    def BackgroundSave(self):
        return self.options.BackgroundSave

    @property
    def BibliographySort(self):
        return self.options.BibliographySort

    @BibliographySort.setter
    def BibliographySort(self, value):
        self.options.BibliographySort = value

    @property
    def BibliographyStyle(self):
        return self.options.BibliographyStyle

    @BibliographyStyle.setter
    def BibliographyStyle(self, value):
        self.options.BibliographyStyle = value

    @property
    def BrazilReform(self):
        return self.options.BrazilReform

    @BrazilReform.setter
    def BrazilReform(self, value):
        self.options.BrazilReform = value

    @property
    def ButtonFieldClicks(self):
        return self.options.ButtonFieldClicks

    @ButtonFieldClicks.setter
    def ButtonFieldClicks(self, value):
        self.options.ButtonFieldClicks = value

    @property
    def CheckGrammarAsYouType(self):
        return self.options.CheckGrammarAsYouType

    @property
    def CheckGrammarWithSpelling(self):
        return self.options.CheckGrammarWithSpelling

    @property
    def CheckHangulEndings(self):
        return self.options.CheckHangulEndings

    @property
    def CheckSpellingAsYouType(self):
        return self.options.CheckSpellingAsYouType

    @property
    def CommentsColor(self):
        return WdColorIndex(self.options.CommentsColor)

    @CommentsColor.setter
    def CommentsColor(self, value):
        self.options.CommentsColor = value

    @property
    def ConfirmConversions(self):
        return self.options.ConfirmConversions

    @property
    def ContextualSpeller(self):
        return self.options.ContextualSpeller

    @ContextualSpeller.setter
    def ContextualSpeller(self, value):
        self.options.ContextualSpeller = value

    @property
    def ConvertHighAnsiToFarEast(self):
        return self.options.ConvertHighAnsiToFarEast

    @property
    def CreateBackup(self):
        return self.options.CreateBackup

    @property
    def Creator(self):
        return self.options.Creator

    @property
    def CtrlClickHyperlinkToOpen(self):
        return self.options.CtrlClickHyperlinkToOpen

    @property
    def CursorMovement(self):
        return WdCursorMovement(self.options.CursorMovement)

    @CursorMovement.setter
    def CursorMovement(self, value):
        self.options.CursorMovement = value

    @property
    def DefaultBorderColor(self):
        return Border(self.options.DefaultBorderColor)

    @DefaultBorderColor.setter
    def DefaultBorderColor(self, value):
        self.options.DefaultBorderColor = value

    @property
    def DefaultBorderColorIndex(self):
        return WdColorIndex(self.options.DefaultBorderColorIndex)

    @DefaultBorderColorIndex.setter
    def DefaultBorderColorIndex(self, value):
        self.options.DefaultBorderColorIndex = value

    @property
    def DefaultBorderLineStyle(self):
        return WdLineStyle(self.options.DefaultBorderLineStyle)

    @DefaultBorderLineStyle.setter
    def DefaultBorderLineStyle(self, value):
        self.options.DefaultBorderLineStyle = value

    @property
    def DefaultBorderLineWidth(self):
        return WdLineWidth(self.options.DefaultBorderLineWidth)

    @DefaultBorderLineWidth.setter
    def DefaultBorderLineWidth(self, value):
        self.options.DefaultBorderLineWidth = value

    @property
    def DefaultEPostageApp(self):
        return self.options.DefaultEPostageApp

    @property
    def DefaultFilePath(self):
        return self.options.DefaultFilePath

    @DefaultFilePath.setter
    def DefaultFilePath(self, value):
        self.options.DefaultFilePath = value

    @property
    def DefaultHighlightColorIndex(self):
        return self.options.DefaultHighlightColorIndex

    @DefaultHighlightColorIndex.setter
    def DefaultHighlightColorIndex(self, value):
        self.options.DefaultHighlightColorIndex = value

    @property
    def DefaultOpenFormat(self):
        return self.options.DefaultOpenFormat

    @DefaultOpenFormat.setter
    def DefaultOpenFormat(self, value):
        self.options.DefaultOpenFormat = value

    @property
    def DefaultTextEncoding(self):
        return self.options.DefaultTextEncoding

    @DefaultTextEncoding.setter
    def DefaultTextEncoding(self, value):
        self.options.DefaultTextEncoding = value

    @property
    def DefaultTray(self):
        return self.options.DefaultTray

    @DefaultTray.setter
    def DefaultTray(self, value):
        self.options.DefaultTray = value

    @property
    def DefaultTrayID(self):
        return WdPaperTray(self.options.DefaultTrayID)

    @DefaultTrayID.setter
    def DefaultTrayID(self, value):
        self.options.DefaultTrayID = value

    @property
    def DeletedCellColor(self):
        return WdCellColor(self.options.DeletedCellColor)

    @DeletedCellColor.setter
    def DeletedCellColor(self, value):
        self.options.DeletedCellColor = value

    @property
    def DeletedTextColor(self):
        return WdColorIndex(self.options.DeletedTextColor)

    @DeletedTextColor.setter
    def DeletedTextColor(self, value):
        self.options.DeletedTextColor = value

    @property
    def DeletedTextMark(self):
        return WdDeletedTextMark(self.options.DeletedTextMark)

    @DeletedTextMark.setter
    def DeletedTextMark(self, value):
        self.options.DeletedTextMark = value

    @property
    def DiacriticColorVal(self):
        return self.options.DiacriticColorVal

    @DiacriticColorVal.setter
    def DiacriticColorVal(self, value):
        self.options.DiacriticColorVal = value

    @property
    def DisableFeaturesbyDefault(self):
        return self.options.DisableFeaturesbyDefault

    @property
    def DisableFeaturesIntroducedAfterbyDefault(self):
        return self.options.DisableFeaturesIntroducedAfterbyDefault

    @property
    def DisplayGridLines(self):
        return self.options.DisplayGridLines

    @property
    def DisplayPasteOptions(self):
        return self.options.DisplayPasteOptions

    @property
    def DocumentViewDirection(self):
        return WdDocumentViewDirection(self.options.DocumentViewDirection)

    @DocumentViewDirection.setter
    def DocumentViewDirection(self, value):
        self.options.DocumentViewDirection = value

    @property
    def DoNotPromptForConvert(self):
        return self.options.DoNotPromptForConvert

    @property
    def EnableHangulHanjaRecentOrdering(self):
        return self.options.EnableHangulHanjaRecentOrdering

    @property
    def EnableLegacyIMEMode(self):
        return self.options.EnableLegacyIMEMode

    @EnableLegacyIMEMode.setter
    def EnableLegacyIMEMode(self, value):
        self.options.EnableLegacyIMEMode = value

    @property
    def EnableLivePreview(self):
        return self.options.EnableLivePreview

    @property
    def EnableMisusedWordsDictionary(self):
        return self.options.EnableMisusedWordsDictionary

    @property
    def EnableSound(self):
        return self.options.EnableSound

    @property
    def EnvelopeFeederInstalled(self):
        return self.options.EnvelopeFeederInstalled

    @property
    def FormatScanning(self):
        return self.options.FormatScanning

    @property
    def FrenchReform(self):
        return WdFrenchSpeller(self.options.FrenchReform)

    @FrenchReform.setter
    def FrenchReform(self, value):
        self.options.FrenchReform = value

    @property
    def GridDistanceHorizontal(self):
        return self.options.GridDistanceHorizontal

    @GridDistanceHorizontal.setter
    def GridDistanceHorizontal(self, value):
        self.options.GridDistanceHorizontal = value

    @property
    def GridDistanceVertical(self):
        return self.options.GridDistanceVertical

    @GridDistanceVertical.setter
    def GridDistanceVertical(self, value):
        self.options.GridDistanceVertical = value

    @property
    def GridOriginHorizontal(self):
        return self.options.GridOriginHorizontal

    @GridOriginHorizontal.setter
    def GridOriginHorizontal(self, value):
        self.options.GridOriginHorizontal = value

    @property
    def GridOriginVertical(self):
        return self.options.GridOriginVertical

    @GridOriginVertical.setter
    def GridOriginVertical(self, value):
        self.options.GridOriginVertical = value

    @property
    def HangulHanjaFastConversion(self):
        return self.options.HangulHanjaFastConversion

    @property
    def HebrewMode(self):
        return WdHebSpellStart(self.options.HebrewMode)

    @HebrewMode.setter
    def HebrewMode(self, value):
        self.options.HebrewMode = value

    @property
    def IgnoreInternetAndFileAddresses(self):
        return self.options.IgnoreInternetAndFileAddresses

    @property
    def IgnoreMixedDigits(self):
        return self.options.IgnoreMixedDigits

    @property
    def IgnoreUppercase(self):
        return self.options.IgnoreUppercase

    @property
    def IMEAutomaticControl(self):
        return self.options.IMEAutomaticControl

    @property
    def InlineConversion(self):
        return self.options.InlineConversion

    @property
    def InsertedCellColor(self):
        return WdCellColor(self.options.InsertedCellColor)

    @InsertedCellColor.setter
    def InsertedCellColor(self, value):
        self.options.InsertedCellColor = value

    @property
    def InsertedTextColor(self):
        return WdColorIndex(self.options.InsertedTextColor)

    @InsertedTextColor.setter
    def InsertedTextColor(self, value):
        self.options.InsertedTextColor = value

    @property
    def InsertedTextMark(self):
        return self.options.InsertedTextMark

    @InsertedTextMark.setter
    def InsertedTextMark(self, value):
        self.options.InsertedTextMark = value

    @property
    def INSKeyForOvertype(self):
        return self.options.INSKeyForOvertype

    @property
    def INSKeyForPaste(self):
        return self.options.INSKeyForPaste

    @property
    def InterpretHighAnsi(self):
        return WdHighAnsiText(self.options.InterpretHighAnsi)

    @InterpretHighAnsi.setter
    def InterpretHighAnsi(self, value):
        self.options.InterpretHighAnsi = value

    @property
    def LocalNetworkFile(self):
        return self.options.LocalNetworkFile

    @property
    def MapPaperSize(self):
        return self.options.MapPaperSize

    @property
    def MatchFuzzyAY(self):
        return self.options.MatchFuzzyAY

    @property
    def MatchFuzzyBV(self):
        return self.options.MatchFuzzyBV

    @property
    def MatchFuzzyByte(self):
        return self.options.MatchFuzzyByte

    @property
    def MatchFuzzyCase(self):
        return self.options.MatchFuzzyCase

    @property
    def MatchFuzzyDash(self):
        return self.options.MatchFuzzyDash

    @property
    def MatchFuzzyDZ(self):
        return self.options.MatchFuzzyDZ

    @property
    def MatchFuzzyHF(self):
        return self.options.MatchFuzzyHF

    @property
    def MatchFuzzyHiragana(self):
        return self.options.MatchFuzzyHiragana

    @property
    def MatchFuzzyIterationMark(self):
        return self.options.MatchFuzzyIterationMark

    @property
    def MatchFuzzyKanji(self):
        return self.options.MatchFuzzyKanji

    @property
    def MatchFuzzyKiKu(self):
        return self.options.MatchFuzzyKiKu

    @property
    def MatchFuzzyOldKana(self):
        return self.options.MatchFuzzyOldKana

    @property
    def MatchFuzzyProlongedSoundMark(self):
        return self.options.MatchFuzzyProlongedSoundMark

    @property
    def MatchFuzzyPunctuation(self):
        return self.options.MatchFuzzyPunctuation

    @property
    def MatchFuzzySmallKana(self):
        return self.options.MatchFuzzySmallKana

    @property
    def MatchFuzzySpace(self):
        return self.options.MatchFuzzySpace

    @property
    def MatchFuzzyTC(self):
        return self.options.MatchFuzzyTC

    @property
    def MatchFuzzyZJ(self):
        return self.options.MatchFuzzyZJ

    @property
    def MeasurementUnit(self):
        return WdMeasurementUnits(self.options.MeasurementUnit)

    @MeasurementUnit.setter
    def MeasurementUnit(self, value):
        self.options.MeasurementUnit = value

    @property
    def MergedCellColor(self):
        return WdCellColor(self.options.MergedCellColor)

    @MergedCellColor.setter
    def MergedCellColor(self, value):
        self.options.MergedCellColor = value

    @property
    def MonthNames(self):
        return WdMonthNames(self.options.MonthNames)

    @MonthNames.setter
    def MonthNames(self, value):
        self.options.MonthNames = value

    @property
    def MoveFromTextColor(self):
        return WdColorIndex(self.options.MoveFromTextColor)

    @MoveFromTextColor.setter
    def MoveFromTextColor(self, value):
        self.options.MoveFromTextColor = value

    @property
    def MoveFromTextMark(self):
        return WdMoveFromTextMark(self.options.MoveFromTextMark)

    @MoveFromTextMark.setter
    def MoveFromTextMark(self, value):
        self.options.MoveFromTextMark = value

    @property
    def MoveToTextColor(self):
        return WdColorIndex(self.options.MoveToTextColor)

    @MoveToTextColor.setter
    def MoveToTextColor(self, value):
        self.options.MoveToTextColor = value

    @property
    def MoveToTextMark(self):
        return WdMoveToTextMark(self.options.MoveToTextMark)

    @MoveToTextMark.setter
    def MoveToTextMark(self, value):
        self.options.MoveToTextMark = value

    @property
    def MultipleWordConversionsMode(self):
        return WdMultipleWordConversionsMode(self.options.MultipleWordConversionsMode)

    @MultipleWordConversionsMode.setter
    def MultipleWordConversionsMode(self, value):
        self.options.MultipleWordConversionsMode = value

    @property
    def OMathAutoBuildUp(self):
        return self.options.OMathAutoBuildUp

    @OMathAutoBuildUp.setter
    def OMathAutoBuildUp(self, value):
        self.options.OMathAutoBuildUp = value

    @property
    def OMathCopyLF(self):
        return self.options.OMathCopyLF

    @OMathCopyLF.setter
    def OMathCopyLF(self, value):
        self.options.OMathCopyLF = value

    @property
    def OptimizeForWord97byDefault(self):
        return self.options.OptimizeForWord97byDefault

    @property
    def Options(self):
        return self.options.Options

    @property
    def Overtype(self):
        return self.options.Overtype

    @property
    def Pagination(self):
        return self.options.Pagination

    @property
    def Parent(self):
        return self.options.Parent

    @property
    def PasteAdjustParagraphSpacing(self):
        return self.options.PasteAdjustParagraphSpacing

    @property
    def PasteAdjustTableFormatting(self):
        return self.options.PasteAdjustTableFormatting

    @property
    def PasteAdjustWordSpacing(self):
        return self.options.PasteAdjustWordSpacing

    @property
    def PasteFormatBetweenDocuments(self):
        return WdPasteOptions(self.options.PasteFormatBetweenDocuments)

    @PasteFormatBetweenDocuments.setter
    def PasteFormatBetweenDocuments(self, value):
        self.options.PasteFormatBetweenDocuments = value

    @property
    def PasteFormatBetweenStyledDocuments(self):
        return WdPasteOptions(self.options.PasteFormatBetweenStyledDocuments)

    @PasteFormatBetweenStyledDocuments.setter
    def PasteFormatBetweenStyledDocuments(self, value):
        self.options.PasteFormatBetweenStyledDocuments = value

    @property
    def PasteFormatFromExternalSource(self):
        return WdPasteOptions(self.options.PasteFormatFromExternalSource)

    @PasteFormatFromExternalSource.setter
    def PasteFormatFromExternalSource(self, value):
        self.options.PasteFormatFromExternalSource = value

    @property
    def PasteFormatWithinDocument(self):
        return WdPasteOptions(self.options.PasteFormatWithinDocument)

    @PasteFormatWithinDocument.setter
    def PasteFormatWithinDocument(self, value):
        self.options.PasteFormatWithinDocument = value

    @property
    def PasteMergeFromPPT(self):
        return self.options.PasteMergeFromPPT

    @property
    def PasteMergeFromXL(self):
        return self.options.PasteMergeFromXL

    @property
    def PasteMergeLists(self):
        return self.options.PasteMergeLists

    @property
    def PasteOptionKeepBulletsAndNumbers(self):
        return self.options.PasteOptionKeepBulletsAndNumbers

    @PasteOptionKeepBulletsAndNumbers.setter
    def PasteOptionKeepBulletsAndNumbers(self, value):
        self.options.PasteOptionKeepBulletsAndNumbers = value

    @property
    def PasteSmartCutPaste(self):
        return self.options.PasteSmartCutPaste

    @property
    def PasteSmartStyleBehavior(self):
        return self.options.PasteSmartStyleBehavior

    @property
    def PictureEditor(self):
        return self.options.PictureEditor

    @PictureEditor.setter
    def PictureEditor(self, value):
        self.options.PictureEditor = value

    @property
    def PictureWrapType(self):
        return self.options.PictureWrapType

    @property
    def PortugalReform(self):
        return self.options.PortugalReform

    @PortugalReform.setter
    def PortugalReform(self, value):
        self.options.PortugalReform = value

    @property
    def PrecisePositioning(self):
        return self.options.PrecisePositioning

    @PrecisePositioning.setter
    def PrecisePositioning(self, value):
        self.options.PrecisePositioning = value

    @property
    def PrintBackground(self):
        return self.options.PrintBackground

    @property
    def PrintBackgrounds(self):
        return self.options.PrintBackgrounds

    @property
    def PrintComments(self):
        return self.options.PrintComments

    @property
    def PrintDraft(self):
        return self.options.PrintDraft

    @property
    def PrintDrawingObjects(self):
        return self.options.PrintDrawingObjects

    @property
    def PrintEvenPagesInAscendingOrder(self):
        return self.options.PrintEvenPagesInAscendingOrder

    @property
    def PrintFieldCodes(self):
        return self.options.PrintFieldCodes

    @property
    def PrintHiddenText(self):
        return self.options.PrintHiddenText

    @property
    def PrintOddPagesInAscendingOrder(self):
        return self.options.PrintOddPagesInAscendingOrder

    @property
    def PrintProperties(self):
        return self.options.PrintProperties

    @property
    def PrintReverse(self):
        return self.options.PrintReverse

    @property
    def PrintXMLTag(self):
        return self.options.PrintXMLTag

    @property
    def PromptUpdateStyle(self):
        return self.options.PromptUpdateStyle

    @property
    def RepeatWord(self):
        return self.options.RepeatWord

    @RepeatWord.setter
    def RepeatWord(self, value):
        self.options.RepeatWord = value

    @property
    def ReplaceSelection(self):
        return self.options.ReplaceSelection

    @property
    def RevisedLinesColor(self):
        return WdColorIndex(self.options.RevisedLinesColor)

    @RevisedLinesColor.setter
    def RevisedLinesColor(self, value):
        self.options.RevisedLinesColor = value

    @property
    def RevisedLinesMark(self):
        return WdRevisedLinesMark(self.options.RevisedLinesMark)

    @RevisedLinesMark.setter
    def RevisedLinesMark(self, value):
        self.options.RevisedLinesMark = value

    @property
    def RevisedPropertiesColor(self):
        return WdColorIndex(self.options.RevisedPropertiesColor)

    @RevisedPropertiesColor.setter
    def RevisedPropertiesColor(self, value):
        self.options.RevisedPropertiesColor = value

    @property
    def RevisedPropertiesMark(self):
        return WdRevisedPropertiesMark(self.options.RevisedPropertiesMark)

    @RevisedPropertiesMark.setter
    def RevisedPropertiesMark(self, value):
        self.options.RevisedPropertiesMark = value

    @property
    def RevisionsBalloonPrintOrientation(self):
        return WdRevisionsBalloonPrintOrientation(self.options.RevisionsBalloonPrintOrientation)

    @RevisionsBalloonPrintOrientation.setter
    def RevisionsBalloonPrintOrientation(self, value):
        self.options.RevisionsBalloonPrintOrientation = value

    @property
    def SaveInterval(self):
        return self.options.SaveInterval

    @SaveInterval.setter
    def SaveInterval(self, value):
        self.options.SaveInterval = value

    @property
    def SaveNormalPrompt(self):
        return self.options.SaveNormalPrompt

    @property
    def SavePropertiesPrompt(self):
        return self.options.SavePropertiesPrompt

    @property
    def SendMailAttach(self):
        return self.options.SendMailAttach

    @property
    def SequenceCheck(self):
        return self.options.SequenceCheck

    @property
    def ShowControlCharacters(self):
        return self.options.ShowControlCharacters

    @property
    def ShowDevTools(self):
        return self.options.ShowDevTools

    @ShowDevTools.setter
    def ShowDevTools(self, value):
        self.options.ShowDevTools = value

    @property
    def ShowDiacritics(self):
        return self.options.ShowDiacritics

    @property
    def ShowFormatError(self):
        return self.options.ShowFormatError

    @property
    def ShowMarkupOpenSave(self):
        return self.options.ShowMarkupOpenSave

    @ShowMarkupOpenSave.setter
    def ShowMarkupOpenSave(self, value):
        self.options.ShowMarkupOpenSave = value

    @property
    def ShowMenuFloaties(self):
        return self.options.ShowMenuFloaties

    @ShowMenuFloaties.setter
    def ShowMenuFloaties(self, value):
        self.options.ShowMenuFloaties = value

    @property
    def ShowReadabilityStatistics(self):
        return self.options.ShowReadabilityStatistics

    @property
    def ShowSelectionFloaties(self):
        return self.options.ShowSelectionFloaties

    @ShowSelectionFloaties.setter
    def ShowSelectionFloaties(self, value):
        self.options.ShowSelectionFloaties = value

    @property
    def SmartCursoring(self):
        return self.options.SmartCursoring

    @SmartCursoring.setter
    def SmartCursoring(self, value):
        self.options.SmartCursoring = value

    @property
    def SmartCutPaste(self):
        return self.options.SmartCutPaste

    @property
    def SmartParaSelection(self):
        return self.options.SmartParaSelection

    @property
    def SnapToGrid(self):
        return self.options.SnapToGrid

    @property
    def SnapToShapes(self):
        return self.options.SnapToShapes

    @property
    def SpanishMode(self):
        return self.options.SpanishMode

    @SpanishMode.setter
    def SpanishMode(self, value):
        self.options.SpanishMode = value

    @property
    def SplitCellColor(self):
        return WdCellColor(self.options.SplitCellColor)

    @SplitCellColor.setter
    def SplitCellColor(self, value):
        self.options.SplitCellColor = value

    @property
    def StoreRSIDOnSave(self):
        return self.options.StoreRSIDOnSave

    @property
    def StrictFinalYaa(self):
        return self.options.StrictFinalYaa

    @property
    def StrictInitialAlefHamza(self):
        return self.options.StrictInitialAlefHamza

    @property
    def StrictRussianE(self):
        return self.options.StrictRussianE

    @property
    def StrictTaaMarboota(self):
        return self.options.StrictTaaMarboota

    @property
    def SuggestFromMainDictionaryOnly(self):
        return self.options.SuggestFromMainDictionaryOnly

    @property
    def SuggestSpellingCorrections(self):
        return self.options.SuggestSpellingCorrections

    @property
    def TabIndentKey(self):
        return self.options.TabIndentKey

    @property
    def TypeNReplace(self):
        return self.options.TypeNReplace

    @property
    def UpdateFieldsAtPrint(self):
        return self.options.UpdateFieldsAtPrint

    @property
    def UpdateFieldsWithTrackedChangesAtPrint(self):
        return self.options.UpdateFieldsWithTrackedChangesAtPrint

    @property
    def UpdateLinksAtOpen(self):
        return self.options.UpdateLinksAtOpen

    @property
    def UpdateLinksAtPrint(self):
        return self.options.UpdateLinksAtPrint

    @property
    def UpdateStyleListBehavior(self):
        return self.options.UpdateStyleListBehavior

    @UpdateStyleListBehavior.setter
    def UpdateStyleListBehavior(self, value):
        self.options.UpdateStyleListBehavior = value

    @property
    def UseCharacterUnit(self):
        return self.options.UseCharacterUnit

    @property
    def UseDiffDiacColor(self):
        return self.options.UseDiffDiacColor

    @property
    def UseGermanSpellingReform(self):
        return self.options.UseGermanSpellingReform

    @property
    def UseNormalStyleForList(self):
        return self.options.UseNormalStyleForList

    @UseNormalStyleForList.setter
    def UseNormalStyleForList(self, value):
        self.options.UseNormalStyleForList = value

    @property
    def VisualSelection(self):
        return WdVisualSelection(self.options.VisualSelection)

    @VisualSelection.setter
    def VisualSelection(self, value):
        self.options.VisualSelection = value

    @property
    def WarnBeforeSavingPrintingSendingMarkup(self):
        return self.options.WarnBeforeSavingPrintingSendingMarkup

class OtherCorrectionsException:

    def __init__(self, othercorrectionsexception=None):
        self.othercorrectionsexception = othercorrectionsexception

    @property
    def Application(self):
        return Application(self.othercorrectionsexception.Application)

    @property
    def Creator(self):
        return self.othercorrectionsexception.Creator

    @property
    def Index(self):
        return self.othercorrectionsexception.Index

    @property
    def Name(self):
        return self.othercorrectionsexception.Name

    @property
    def Parent(self):
        return self.othercorrectionsexception.Parent

    def Delete(self):
        self.othercorrectionsexception.Delete()

class Page:

    def __init__(self, page=None):
        self.page = page

    @property
    def Application(self):
        return Application(self.page.Application)

    @property
    def Breaks(self):
        return Breaks(self.page.Breaks)

    @property
    def Creator(self):
        return self.page.Creator

    @property
    def EnhMetaFileBits(self):
        return self.page.EnhMetaFileBits

    @property
    def Height(self):
        return self.page.Height

    @property
    def Left(self):
        return self.page.Left

    @property
    def Parent(self):
        return self.page.Parent

    @property
    def Rectangles(self):
        return Rectangles(self.page.Rectangles)

    @property
    def Top(self):
        return self.page.Top

    @property
    def Width(self):
        return self.page.Width

class PageNumber:

    def __init__(self, pagenumber=None):
        self.pagenumber = pagenumber

    @property
    def Alignment(self):
        return WdPageNumberAlignment(self.pagenumber.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.pagenumber.Alignment = value

    @property
    def Application(self):
        return Application(self.pagenumber.Application)

    @property
    def Creator(self):
        return self.pagenumber.Creator

    @property
    def Index(self):
        return self.pagenumber.Index

    @property
    def Parent(self):
        return self.pagenumber.Parent

    def Copy(self):
        self.pagenumber.Copy()

    def Cut(self):
        self.pagenumber.Cut()

    def Delete(self):
        self.pagenumber.Delete()

    def Select(self):
        self.pagenumber.Select()

class PageSetup:

    def __init__(self, pagesetup=None):
        self.pagesetup = pagesetup

    @property
    def Application(self):
        return Application(self.pagesetup.Application)

    @property
    def BookFoldPrinting(self):
        return self.pagesetup.BookFoldPrinting

    @property
    def BookFoldPrintingSheets(self):
        return self.pagesetup.BookFoldPrintingSheets

    @BookFoldPrintingSheets.setter
    def BookFoldPrintingSheets(self, value):
        self.pagesetup.BookFoldPrintingSheets = value

    @property
    def BookFoldRevPrinting(self):
        return self.pagesetup.BookFoldRevPrinting

    @property
    def BottomMargin(self):
        return self.pagesetup.BottomMargin

    @BottomMargin.setter
    def BottomMargin(self, value):
        self.pagesetup.BottomMargin = value

    @property
    def CharsLine(self):
        return self.pagesetup.CharsLine

    @CharsLine.setter
    def CharsLine(self, value):
        self.pagesetup.CharsLine = value

    @property
    def Creator(self):
        return self.pagesetup.Creator

    @property
    def DifferentFirstPageHeaderFooter(self):
        return self.pagesetup.DifferentFirstPageHeaderFooter

    @property
    def FirstPageTray(self):
        return WdPaperTray(self.pagesetup.FirstPageTray)

    @FirstPageTray.setter
    def FirstPageTray(self, value):
        self.pagesetup.FirstPageTray = value

    @property
    def FooterDistance(self):
        return self.pagesetup.FooterDistance

    @FooterDistance.setter
    def FooterDistance(self, value):
        self.pagesetup.FooterDistance = value

    @property
    def Gutter(self):
        return self.pagesetup.Gutter

    @Gutter.setter
    def Gutter(self, value):
        self.pagesetup.Gutter = value

    @property
    def GutterPos(self):
        return WdGutterStyle(self.pagesetup.GutterPos)

    @GutterPos.setter
    def GutterPos(self, value):
        self.pagesetup.GutterPos = value

    @property
    def GutterStyle(self):
        return WdGutterStyleOld(self.pagesetup.GutterStyle)

    @GutterStyle.setter
    def GutterStyle(self, value):
        self.pagesetup.GutterStyle = value

    @property
    def HeaderDistance(self):
        return self.pagesetup.HeaderDistance

    @HeaderDistance.setter
    def HeaderDistance(self, value):
        self.pagesetup.HeaderDistance = value

    @property
    def LayoutMode(self):
        return WdLayoutMode(self.pagesetup.LayoutMode)

    @LayoutMode.setter
    def LayoutMode(self, value):
        self.pagesetup.LayoutMode = value

    @property
    def LeftMargin(self):
        return self.pagesetup.LeftMargin

    @LeftMargin.setter
    def LeftMargin(self, value):
        self.pagesetup.LeftMargin = value

    @property
    def LineNumbering(self):
        return LineNumbering(self.pagesetup.LineNumbering)

    @LineNumbering.setter
    def LineNumbering(self, value):
        self.pagesetup.LineNumbering = value

    @property
    def LinesPage(self):
        return self.pagesetup.LinesPage

    @LinesPage.setter
    def LinesPage(self, value):
        self.pagesetup.LinesPage = value

    @property
    def MirrorMargins(self):
        return self.pagesetup.MirrorMargins

    @property
    def OddAndEvenPagesHeaderFooter(self):
        return self.pagesetup.OddAndEvenPagesHeaderFooter

    @property
    def Orientation(self):
        return WdOrientation(self.pagesetup.Orientation)

    @Orientation.setter
    def Orientation(self, value):
        self.pagesetup.Orientation = value

    @property
    def OtherPagesTray(self):
        return WdPaperTray(self.pagesetup.OtherPagesTray)

    @OtherPagesTray.setter
    def OtherPagesTray(self, value):
        self.pagesetup.OtherPagesTray = value

    @property
    def PageHeight(self):
        return self.pagesetup.PageHeight

    @PageHeight.setter
    def PageHeight(self, value):
        self.pagesetup.PageHeight = value

    @property
    def PageWidth(self):
        return self.pagesetup.PageWidth

    @PageWidth.setter
    def PageWidth(self, value):
        self.pagesetup.PageWidth = value

    @property
    def PaperSize(self):
        return WdPaperSize(self.pagesetup.PaperSize)

    @PaperSize.setter
    def PaperSize(self, value):
        self.pagesetup.PaperSize = value

    @property
    def Parent(self):
        return self.pagesetup.Parent

    @property
    def RightMargin(self):
        return self.pagesetup.RightMargin

    @RightMargin.setter
    def RightMargin(self, value):
        self.pagesetup.RightMargin = value

    @property
    def SectionDirection(self):
        return WdSectionDirection(self.pagesetup.SectionDirection)

    @SectionDirection.setter
    def SectionDirection(self, value):
        self.pagesetup.SectionDirection = value

    @property
    def SectionStart(self):
        return WdSectionStart(self.pagesetup.SectionStart)

    @SectionStart.setter
    def SectionStart(self, value):
        self.pagesetup.SectionStart = value

    @property
    def SuppressEndnotes(self):
        return self.pagesetup.SuppressEndnotes

    @property
    def TextColumns(self):
        return self.pagesetup.TextColumns

    @property
    def TopMargin(self):
        return self.pagesetup.TopMargin

    @TopMargin.setter
    def TopMargin(self, value):
        self.pagesetup.TopMargin = value

    @property
    def TwoPagesOnOne(self):
        return self.pagesetup.TwoPagesOnOne

    @property
    def VerticalAlignment(self):
        return WdVerticalAlignment(self.pagesetup.VerticalAlignment)

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.pagesetup.VerticalAlignment = value

    def SetAsTemplateDefault(self):
        self.pagesetup.SetAsTemplateDefault()

    def TogglePortrait(self):
        self.pagesetup.TogglePortrait()

class Pane:

    def __init__(self, pane=None):
        self.pane = pane

    @property
    def Application(self):
        return Application(self.pane.Application)

    @property
    def BrowseWidth(self):
        return self.pane.BrowseWidth

    @property
    def Creator(self):
        return self.pane.Creator

    @property
    def DisplayRulers(self):
        return self.pane.DisplayRulers

    @property
    def DisplayVerticalRuler(self):
        return self.pane.DisplayVerticalRuler

    @property
    def Document(self):
        return Document(self.pane.Document)

    @property
    def Frameset(self):
        return Frameset(self.pane.Frameset)

    @property
    def HorizontalPercentScrolled(self):
        return self.pane.HorizontalPercentScrolled

    @HorizontalPercentScrolled.setter
    def HorizontalPercentScrolled(self, value):
        self.pane.HorizontalPercentScrolled = value

    @property
    def Index(self):
        return self.pane.Index

    @property
    def MinimumFontSize(self):
        return self.pane.MinimumFontSize

    @MinimumFontSize.setter
    def MinimumFontSize(self, value):
        self.pane.MinimumFontSize = value

    @property
    def Next(self):
        return Pane(self.pane.Next)

    @property
    def Pages(self):
        return self.pane.Pages

    @property
    def Parent(self):
        return self.pane.Parent

    @property
    def Previous(self):
        return Pane(self.pane.Previous)

    @property
    def Selection(self):
        return Selection(self.pane.Selection)

    @property
    def VerticalPercentScrolled(self):
        return self.pane.VerticalPercentScrolled

    @VerticalPercentScrolled.setter
    def VerticalPercentScrolled(self, value):
        self.pane.VerticalPercentScrolled = value

    @property
    def View(self):
        return View(self.pane.View)

    @property
    def Zooms(self):
        return self.pane.Zooms

    def Activate(self):
        self.pane.Activate()

    def AutoScroll(self, *args, Velocity=None):
        arguments = {"Velocity": Velocity}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pane.AutoScroll(*args, **arguments)

    def Close(self):
        self.pane.Close()

    def LargeScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pane.LargeScroll(*args, **arguments)

    def NewFrameset(self):
        self.pane.NewFrameset()

    def PageScroll(self, *args, Down=None, Up=None):
        arguments = {"Down": Down, "Up": Up}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pane.PageScroll(*args, **arguments)

    def SmallScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pane.SmallScroll(*args, **arguments)

    def TOCInFrameset(self):
        self.pane.TOCInFrameset()

class Paragraph:

    def __init__(self, paragraph=None):
        self.paragraph = paragraph

    @property
    def AddSpaceBetweenFarEastAndAlpha(self):
        return self.paragraph.AddSpaceBetweenFarEastAndAlpha

    @property
    def AddSpaceBetweenFarEastAndDigit(self):
        return self.paragraph.AddSpaceBetweenFarEastAndDigit

    @property
    def Alignment(self):
        return WdParagraphAlignment(self.paragraph.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.paragraph.Alignment = value

    @property
    def Application(self):
        return Application(self.paragraph.Application)

    @property
    def AutoAdjustRightIndent(self):
        return self.paragraph.AutoAdjustRightIndent

    @property
    def BaseLineAlignment(self):
        return WdBaselineAlignment(self.paragraph.BaseLineAlignment)

    @BaseLineAlignment.setter
    def BaseLineAlignment(self, value):
        self.paragraph.BaseLineAlignment = value

    @property
    def Borders(self):
        return self.paragraph.Borders

    @property
    def CharacterUnitFirstLineIndent(self):
        return self.paragraph.CharacterUnitFirstLineIndent

    @CharacterUnitFirstLineIndent.setter
    def CharacterUnitFirstLineIndent(self, value):
        self.paragraph.CharacterUnitFirstLineIndent = value

    @property
    def CharacterUnitLeftIndent(self):
        return self.paragraph.CharacterUnitLeftIndent

    @CharacterUnitLeftIndent.setter
    def CharacterUnitLeftIndent(self, value):
        self.paragraph.CharacterUnitLeftIndent = value

    @property
    def CharacterUnitRightIndent(self):
        return self.paragraph.CharacterUnitRightIndent

    @CharacterUnitRightIndent.setter
    def CharacterUnitRightIndent(self, value):
        self.paragraph.CharacterUnitRightIndent = value

    @property
    def Creator(self):
        return self.paragraph.Creator

    @property
    def DisableLineHeightGrid(self):
        return self.paragraph.DisableLineHeightGrid

    @property
    def DropCap(self):
        return DropCap(self.paragraph.DropCap)

    @property
    def FarEastLineBreakControl(self):
        return self.paragraph.FarEastLineBreakControl

    @property
    def FirstLineIndent(self):
        return self.paragraph.FirstLineIndent

    @FirstLineIndent.setter
    def FirstLineIndent(self, value):
        self.paragraph.FirstLineIndent = value

    @property
    def Format(self):
        return ParagraphFormat(self.paragraph.Format)

    @Format.setter
    def Format(self, value):
        self.paragraph.Format = value

    @property
    def HalfWidthPunctuationOnTopOfLine(self):
        return self.paragraph.HalfWidthPunctuationOnTopOfLine

    @property
    def HangingPunctuation(self):
        return self.paragraph.HangingPunctuation

    @property
    def Hyphenation(self):
        return self.paragraph.Hyphenation

    @property
    def ID(self):
        return self.paragraph.ID

    @ID.setter
    def ID(self, value):
        self.paragraph.ID = value

    @property
    def IsStyleSeparator(self):
        return self.paragraph.IsStyleSeparator

    @property
    def KeepTogether(self):
        return self.paragraph.KeepTogether

    @property
    def KeepWithNext(self):
        return self.paragraph.KeepWithNext

    @property
    def LeftIndent(self):
        return self.paragraph.LeftIndent

    @LeftIndent.setter
    def LeftIndent(self, value):
        self.paragraph.LeftIndent = value

    @property
    def LineSpacing(self):
        return self.paragraph.LineSpacing

    @LineSpacing.setter
    def LineSpacing(self, value):
        self.paragraph.LineSpacing = value

    @property
    def LineSpacingRule(self):
        return WdLineSpacing(self.paragraph.LineSpacingRule)

    @LineSpacingRule.setter
    def LineSpacingRule(self, value):
        self.paragraph.LineSpacingRule = value

    @property
    def LineUnitAfter(self):
        return self.paragraph.LineUnitAfter

    @LineUnitAfter.setter
    def LineUnitAfter(self, value):
        self.paragraph.LineUnitAfter = value

    @property
    def LineUnitBefore(self):
        return self.paragraph.LineUnitBefore

    @LineUnitBefore.setter
    def LineUnitBefore(self, value):
        self.paragraph.LineUnitBefore = value

    def ListNumberOriginal(self, *args, Level=None):
        arguments = {"Level": Level}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.paragraph.ListNumberOriginal(*args, **arguments)

    @property
    def MirrorIndents(self):
        return self.paragraph.MirrorIndents

    @MirrorIndents.setter
    def MirrorIndents(self, value):
        self.paragraph.MirrorIndents = value

    @property
    def NoLineNumber(self):
        return self.paragraph.NoLineNumber

    @property
    def OutlineLevel(self):
        return WdOutlineLevel(self.paragraph.OutlineLevel)

    @OutlineLevel.setter
    def OutlineLevel(self, value):
        self.paragraph.OutlineLevel = value

    @property
    def PageBreakBefore(self):
        return self.paragraph.PageBreakBefore

    @property
    def Parent(self):
        return self.paragraph.Parent

    @property
    def Range(self):
        return Range(self.paragraph.Range)

    @property
    def ReadingOrder(self):
        return WdReadingOrder(self.paragraph.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.paragraph.ReadingOrder = value

    @property
    def RightIndent(self):
        return self.paragraph.RightIndent

    @RightIndent.setter
    def RightIndent(self, value):
        self.paragraph.RightIndent = value

    @property
    def Shading(self):
        return Shading(self.paragraph.Shading)

    @property
    def SpaceAfter(self):
        return self.paragraph.SpaceAfter

    @SpaceAfter.setter
    def SpaceAfter(self, value):
        self.paragraph.SpaceAfter = value

    @property
    def SpaceAfterAuto(self):
        return self.paragraph.SpaceAfterAuto

    @property
    def SpaceBefore(self):
        return self.paragraph.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.paragraph.SpaceBefore = value

    @property
    def SpaceBeforeAuto(self):
        return self.paragraph.SpaceBeforeAuto

    @property
    def Style(self):
        return self.paragraph.Style

    @Style.setter
    def Style(self, value):
        self.paragraph.Style = value

    @property
    def TabStops(self):
        return self.paragraph.TabStops

    @TabStops.setter
    def TabStops(self, value):
        self.paragraph.TabStops = value

    @property
    def TextboxTightWrap(self):
        return WdTextboxTightWrap(self.paragraph.TextboxTightWrap)

    @TextboxTightWrap.setter
    def TextboxTightWrap(self, value):
        self.paragraph.TextboxTightWrap = value

    @property
    def WidowControl(self):
        return self.paragraph.WidowControl

    @property
    def WordWrap(self):
        return self.paragraph.WordWrap

    def CloseUp(self):
        self.paragraph.CloseUp()

    def Indent(self):
        self.paragraph.Indent()

    def IndentCharWidth(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraph.IndentCharWidth(*args, **arguments)

    def IndentFirstLineCharWidth(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraph.IndentFirstLineCharWidth(*args, **arguments)

    def JoinList(self):
        self.paragraph.JoinList()

    def ListAdvanceTo(self, *args, Level1=None, Level2=None, Level3=None, Level4=None, Level5=None, Level6=None, Level7=None, Level8=None, Level9=None):
        arguments = {"Level1": Level1, "Level2": Level2, "Level3": Level3, "Level4": Level4, "Level5": Level5, "Level6": Level6, "Level7": Level7, "Level8": Level8, "Level9": Level9}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraph.ListAdvanceTo(*args, **arguments)

    def Next(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.paragraph.Next(*args, **arguments)

    def OpenOrCloseUp(self):
        self.paragraph.OpenOrCloseUp()

    def OpenUp(self):
        self.paragraph.OpenUp()

    def Outdent(self):
        self.paragraph.Outdent()

    def OutlineDemote(self):
        self.paragraph.OutlineDemote()

    def OutlineDemoteToBody(self):
        self.paragraph.OutlineDemoteToBody()

    def OutlinePromote(self):
        self.paragraph.OutlinePromote()

    def Previous(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.paragraph.Previous(*args, **arguments)

    def Reset(self):
        self.paragraph.Reset()

    def ResetAdvanceTo(self):
        self.paragraph.ResetAdvanceTo()

    def SelectNumber(self):
        self.paragraph.SelectNumber()

    def SeparateList(self):
        self.paragraph.SeparateList()

    def Space1(self):
        self.paragraph.Space1()

    def Space15(self):
        self.paragraph.Space15()

    def Space2(self):
        self.paragraph.Space2()

    def TabHangingIndent(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraph.TabHangingIndent(*args, **arguments)

    def TabIndent(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraph.TabIndent(*args, **arguments)

class ParagraphFormat:

    def __init__(self, paragraphformat=None):
        self.paragraphformat = paragraphformat

    @property
    def AddSpaceBetweenFarEastAndAlpha(self):
        return self.paragraphformat.AddSpaceBetweenFarEastAndAlpha

    @property
    def AddSpaceBetweenFarEastAndDigit(self):
        return self.paragraphformat.AddSpaceBetweenFarEastAndDigit

    @property
    def Alignment(self):
        return WdParagraphAlignment(self.paragraphformat.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.paragraphformat.Alignment = value

    @property
    def Application(self):
        return Application(self.paragraphformat.Application)

    @property
    def AutoAdjustRightIndent(self):
        return self.paragraphformat.AutoAdjustRightIndent

    @property
    def BaseLineAlignment(self):
        return WdBaselineAlignment(self.paragraphformat.BaseLineAlignment)

    @BaseLineAlignment.setter
    def BaseLineAlignment(self, value):
        self.paragraphformat.BaseLineAlignment = value

    @property
    def Borders(self):
        return self.paragraphformat.Borders

    @property
    def CharacterUnitFirstLineIndent(self):
        return self.paragraphformat.CharacterUnitFirstLineIndent

    @CharacterUnitFirstLineIndent.setter
    def CharacterUnitFirstLineIndent(self, value):
        self.paragraphformat.CharacterUnitFirstLineIndent = value

    @property
    def CharacterUnitLeftIndent(self):
        return self.paragraphformat.CharacterUnitLeftIndent

    @CharacterUnitLeftIndent.setter
    def CharacterUnitLeftIndent(self, value):
        self.paragraphformat.CharacterUnitLeftIndent = value

    @property
    def CharacterUnitRightIndent(self):
        return self.paragraphformat.CharacterUnitRightIndent

    @CharacterUnitRightIndent.setter
    def CharacterUnitRightIndent(self, value):
        self.paragraphformat.CharacterUnitRightIndent = value

    @property
    def Creator(self):
        return self.paragraphformat.Creator

    @property
    def DisableLineHeightGrid(self):
        return self.paragraphformat.DisableLineHeightGrid

    @property
    def Duplicate(self):
        return ParagraphFormat(self.paragraphformat.Duplicate)

    @property
    def FarEastLineBreakControl(self):
        return self.paragraphformat.FarEastLineBreakControl

    @property
    def FirstLineIndent(self):
        return self.paragraphformat.FirstLineIndent

    @FirstLineIndent.setter
    def FirstLineIndent(self, value):
        self.paragraphformat.FirstLineIndent = value

    @property
    def HalfWidthPunctuationOnTopOfLine(self):
        return self.paragraphformat.HalfWidthPunctuationOnTopOfLine

    @property
    def HangingPunctuation(self):
        return self.paragraphformat.HangingPunctuation

    @property
    def Hyphenation(self):
        return self.paragraphformat.Hyphenation

    @property
    def KeepTogether(self):
        return self.paragraphformat.KeepTogether

    @property
    def KeepWithNext(self):
        return self.paragraphformat.KeepWithNext

    @property
    def LeftIndent(self):
        return self.paragraphformat.LeftIndent

    @LeftIndent.setter
    def LeftIndent(self, value):
        self.paragraphformat.LeftIndent = value

    @property
    def LineSpacing(self):
        return self.paragraphformat.LineSpacing

    @LineSpacing.setter
    def LineSpacing(self, value):
        self.paragraphformat.LineSpacing = value

    @property
    def LineSpacingRule(self):
        return WdLineSpacing(self.paragraphformat.LineSpacingRule)

    @LineSpacingRule.setter
    def LineSpacingRule(self, value):
        self.paragraphformat.LineSpacingRule = value

    @property
    def LineUnitAfter(self):
        return self.paragraphformat.LineUnitAfter

    @LineUnitAfter.setter
    def LineUnitAfter(self, value):
        self.paragraphformat.LineUnitAfter = value

    @property
    def LineUnitBefore(self):
        return self.paragraphformat.LineUnitBefore

    @LineUnitBefore.setter
    def LineUnitBefore(self, value):
        self.paragraphformat.LineUnitBefore = value

    @property
    def MirrorIndents(self):
        return self.paragraphformat.MirrorIndents

    @MirrorIndents.setter
    def MirrorIndents(self, value):
        self.paragraphformat.MirrorIndents = value

    @property
    def NoLineNumber(self):
        return self.paragraphformat.NoLineNumber

    @property
    def OutlineLevel(self):
        return WdOutlineLevel(self.paragraphformat.OutlineLevel)

    @OutlineLevel.setter
    def OutlineLevel(self, value):
        self.paragraphformat.OutlineLevel = value

    @property
    def PageBreakBefore(self):
        return self.paragraphformat.PageBreakBefore

    @property
    def Parent(self):
        return self.paragraphformat.Parent

    @property
    def ReadingOrder(self):
        return WdReadingOrder(self.paragraphformat.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.paragraphformat.ReadingOrder = value

    @property
    def RightIndent(self):
        return self.paragraphformat.RightIndent

    @RightIndent.setter
    def RightIndent(self, value):
        self.paragraphformat.RightIndent = value

    @property
    def Shading(self):
        return Shading(self.paragraphformat.Shading)

    @property
    def SpaceAfter(self):
        return self.paragraphformat.SpaceAfter

    @SpaceAfter.setter
    def SpaceAfter(self, value):
        self.paragraphformat.SpaceAfter = value

    @property
    def SpaceAfterAuto(self):
        return self.paragraphformat.SpaceAfterAuto

    @property
    def SpaceBefore(self):
        return self.paragraphformat.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.paragraphformat.SpaceBefore = value

    @property
    def SpaceBeforeAuto(self):
        return self.paragraphformat.SpaceBeforeAuto

    @property
    def Style(self):
        return self.paragraphformat.Style

    @Style.setter
    def Style(self, value):
        self.paragraphformat.Style = value

    @property
    def TabStops(self):
        return self.paragraphformat.TabStops

    @TabStops.setter
    def TabStops(self, value):
        self.paragraphformat.TabStops = value

    @property
    def TextboxTightWrap(self):
        return WdTextboxTightWrap(self.paragraphformat.TextboxTightWrap)

    @TextboxTightWrap.setter
    def TextboxTightWrap(self, value):
        self.paragraphformat.TextboxTightWrap = value

    @property
    def WidowControl(self):
        return self.paragraphformat.WidowControl

    @property
    def WordWrap(self):
        return self.paragraphformat.WordWrap

    def CloseUp(self):
        self.paragraphformat.CloseUp()

    def IndentCharWidth(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraphformat.IndentCharWidth(*args, **arguments)

    def IndentFirstLineCharWidth(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraphformat.IndentFirstLineCharWidth(*args, **arguments)

    def OpenOrCloseUp(self):
        self.paragraphformat.OpenOrCloseUp()

    def OpenUp(self):
        self.paragraphformat.OpenUp()

    def Reset(self):
        self.paragraphformat.Reset()

    def Space1(self):
        self.paragraphformat.Space1()

    def Space15(self):
        self.paragraphformat.Space15()

    def Space2(self):
        self.paragraphformat.Space2()

    def TabHangingIndent(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraphformat.TabHangingIndent(*args, **arguments)

    def TabIndent(self, *args, Count=None):
        arguments = {"Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.paragraphformat.TabIndent(*args, **arguments)

class PictureFormat:

    def __init__(self, pictureformat=None):
        self.pictureformat = pictureformat

    @property
    def Application(self):
        return Application(self.pictureformat.Application)

    @property
    def Brightness(self):
        return self.pictureformat.Brightness

    @Brightness.setter
    def Brightness(self, value):
        self.pictureformat.Brightness = value

    @property
    def ColorType(self):
        return self.pictureformat.ColorType

    @ColorType.setter
    def ColorType(self, value):
        self.pictureformat.ColorType = value

    @property
    def Contrast(self):
        return self.pictureformat.Contrast

    @Contrast.setter
    def Contrast(self, value):
        self.pictureformat.Contrast = value

    @property
    def Creator(self):
        return self.pictureformat.Creator

    @property
    def Crop(self):
        return self.pictureformat.Crop

    @Crop.setter
    def Crop(self, value):
        self.pictureformat.Crop = value

    @property
    def CropBottom(self):
        return self.pictureformat.CropBottom

    @CropBottom.setter
    def CropBottom(self, value):
        self.pictureformat.CropBottom = value

    @property
    def CropLeft(self):
        return self.pictureformat.CropLeft

    @CropLeft.setter
    def CropLeft(self, value):
        self.pictureformat.CropLeft = value

    @property
    def CropRight(self):
        return self.pictureformat.CropRight

    @CropRight.setter
    def CropRight(self, value):
        self.pictureformat.CropRight = value

    @property
    def CropTop(self):
        return self.pictureformat.CropTop

    @CropTop.setter
    def CropTop(self, value):
        self.pictureformat.CropTop = value

    @property
    def Parent(self):
        return self.pictureformat.Parent

    @property
    def TransparencyColor(self):
        return self.pictureformat.TransparencyColor

    @TransparencyColor.setter
    def TransparencyColor(self, value):
        self.pictureformat.TransparencyColor = value

    @property
    def TransparentBackground(self):
        return self.pictureformat.TransparentBackground

    def IncrementBrightness(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pictureformat.IncrementBrightness(*args, **arguments)

    def IncrementContrast(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.pictureformat.IncrementContrast(*args, **arguments)

class PlotArea:

    def __init__(self, plotarea=None):
        self.plotarea = plotarea

    @property
    def Application(self):
        return self.plotarea.Application

    @property
    def Creator(self):
        return self.plotarea.Creator

    @property
    def Format(self):
        return ChartFormat(self.plotarea.Format)

    @property
    def Height(self):
        return self.plotarea.Height

    @Height.setter
    def Height(self, value):
        self.plotarea.Height = value

    @property
    def InsideHeight(self):
        return self.plotarea.InsideHeight

    @InsideHeight.setter
    def InsideHeight(self, value):
        self.plotarea.InsideHeight = value

    @property
    def InsideLeft(self):
        return self.plotarea.InsideLeft

    @InsideLeft.setter
    def InsideLeft(self, value):
        self.plotarea.InsideLeft = value

    @property
    def InsideTop(self):
        return self.plotarea.InsideTop

    @InsideTop.setter
    def InsideTop(self, value):
        self.plotarea.InsideTop = value

    @property
    def InsideWidth(self):
        return self.plotarea.InsideWidth

    @InsideWidth.setter
    def InsideWidth(self, value):
        self.plotarea.InsideWidth = value

    @property
    def Left(self):
        return self.plotarea.Left

    @Left.setter
    def Left(self, value):
        self.plotarea.Left = value

    @property
    def Name(self):
        return self.plotarea.Name

    @property
    def Parent(self):
        return self.plotarea.Parent

    @property
    def Position(self):
        return self.plotarea.Position

    @Position.setter
    def Position(self, value):
        self.plotarea.Position = value

    @property
    def Top(self):
        return self.plotarea.Top

    @Top.setter
    def Top(self, value):
        self.plotarea.Top = value

    @property
    def Width(self):
        return self.plotarea.Width

    @Width.setter
    def Width(self, value):
        self.plotarea.Width = value

    def ClearFormats(self):
        self.plotarea.ClearFormats()

    def Select(self):
        self.plotarea.Select()

class Point:

    def __init__(self, point=None):
        self.point = point

    @property
    def Application(self):
        return self.point.Application

    @property
    def ApplyPictToEnd(self):
        return self.point.ApplyPictToEnd

    @property
    def ApplyPictToFront(self):
        return self.point.ApplyPictToFront

    @property
    def ApplyPictToSides(self):
        return self.point.ApplyPictToSides

    @property
    def Creator(self):
        return self.point.Creator

    @property
    def DataLabel(self):
        return DataLabel(self.point.DataLabel)

    @property
    def Explosion(self):
        return self.point.Explosion

    @Explosion.setter
    def Explosion(self, value):
        self.point.Explosion = value

    @property
    def Format(self):
        return ChartFormat(self.point.Format)

    @property
    def Has3DEffect(self):
        return self.point.Has3DEffect

    @property
    def HasDataLabel(self):
        return self.point.HasDataLabel

    @property
    def Height(self):
        return self.point.Height

    @property
    def InvertIfNegative(self):
        return self.point.InvertIfNegative

    @property
    def Left(self):
        return self.point.Left

    @property
    def MarkerBackgroundColor(self):
        return self.point.MarkerBackgroundColor

    @property
    def MarkerBackgroundColorIndex(self):
        return self.point.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.point.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.point.MarkerForegroundColor

    @property
    def MarkerForegroundColorIndex(self):
        return self.point.MarkerForegroundColorIndex

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.point.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.point.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.point.MarkerSize = value

    @property
    def MarkerStyle(self):
        return self.point.MarkerStyle

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.point.MarkerStyle = value

    @property
    def Name(self):
        return self.point.Name

    @property
    def Parent(self):
        return self.point.Parent

    @property
    def PictureType(self):
        return self.point.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.point.PictureType = value

    @property
    def PictureUnit2(self):
        return self.point.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.point.PictureUnit2 = value

    @property
    def SecondaryPlot(self):
        return self.point.SecondaryPlot

    @property
    def Shadow(self):
        return self.point.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.point.Shadow = value

    @property
    def Top(self):
        return self.point.Top

    @property
    def Width(self):
        return self.point.Width

    def ApplyDataLabels(self, *args, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = {"Type": Type, "LegendKey": LegendKey, "AutoText": AutoText, "HasLeaderLines": HasLeaderLines, "ShowSeriesName": ShowSeriesName, "ShowCategoryName": ShowCategoryName, "ShowValue": ShowValue, "ShowPercentage": ShowPercentage, "ShowBubbleSize": ShowBubbleSize, "Separator": Separator}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.point.ApplyDataLabels(*args, **arguments)

    def ClearFormats(self):
        self.point.ClearFormats()

    def Copy(self):
        self.point.Copy()

    def Delete(self):
        self.point.Delete()

    def Paste(self):
        self.point.Paste()

    def PieSliceLocation(self, *args, loc=None, Index=None):
        arguments = {"loc": loc, "Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.point.PieSliceLocation(*args, **arguments)

    def Select(self):
        self.point.Select()

class Points:

    def __init__(self, points=None):
        self.points = points

    def __call__(self, item):
        return Point(self.points(item))

    @property
    def Application(self):
        return self.points.Application

    @property
    def Count(self):
        return self.points.Count

    @property
    def Creator(self):
        return self.points.Creator

    @property
    def Parent(self):
        return self.points.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Point(self.points.Item(*args, **arguments))

class ProtectedViewWindow:

    def __init__(self, protectedviewwindow=None):
        self.protectedviewwindow = protectedviewwindow

    @property
    def Active(self):
        return self.protectedviewwindow.Active

    @property
    def Application(self):
        return self.protectedviewwindow.Application

    @property
    def Caption(self):
        return self.protectedviewwindow.Caption

    @Caption.setter
    def Caption(self, value):
        self.protectedviewwindow.Caption = value

    @property
    def Creator(self):
        return self.protectedviewwindow.Creator

    @property
    def Document(self):
        return self.protectedviewwindow.Document

    @property
    def Height(self):
        return self.protectedviewwindow.Height

    @Height.setter
    def Height(self, value):
        self.protectedviewwindow.Height = value

    @property
    def Index(self):
        return self.protectedviewwindow.Index

    @property
    def Left(self):
        return self.protectedviewwindow.Left

    @Left.setter
    def Left(self, value):
        self.protectedviewwindow.Left = value

    @property
    def Parent(self):
        return self.protectedviewwindow.Parent

    @property
    def SourceName(self):
        return self.protectedviewwindow.SourceName

    @property
    def SourcePath(self):
        return self.protectedviewwindow.SourcePath

    @property
    def Top(self):
        return self.protectedviewwindow.Top

    @Top.setter
    def Top(self, value):
        self.protectedviewwindow.Top = value

    @property
    def Visible(self):
        return self.protectedviewwindow.Visible

    @property
    def Width(self):
        return self.protectedviewwindow.Width

    @Width.setter
    def Width(self, value):
        self.protectedviewwindow.Width = value

    @property
    def WindowState(self):
        return self.protectedviewwindow.WindowState

    @WindowState.setter
    def WindowState(self, value):
        self.protectedviewwindow.WindowState = value

    def Activate(self):
        return self.protectedviewwindow.Activate()

    def Close(self):
        return self.protectedviewwindow.Close()

    def ToggleRibbon(self):
        self.protectedviewwindow.ToggleRibbon()

class ProtectedViewWindows:

    def __init__(self, protectedviewwindows=None):
        self.protectedviewwindows = protectedviewwindows

    def __call__(self, item):
        return ProtectedViewWindow(self.protectedviewwindows(item))

    @property
    def Application(self):
        return self.protectedviewwindows.Application

    @property
    def Count(self):
        return self.protectedviewwindows.Count

    @property
    def Creator(self):
        return self.protectedviewwindows.Creator

    @property
    def Parent(self):
        return self.protectedviewwindows.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.protectedviewwindows.Item(*args, **arguments)

    def Open(self, *args, FileName=None, AddToRecentFiles=None, PasswordDocument=None, Visible=None, OpenAndRepair=None):
        arguments = {"FileName": FileName, "AddToRecentFiles": AddToRecentFiles, "PasswordDocument": PasswordDocument, "Visible": Visible, "OpenAndRepair": OpenAndRepair}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return ProtectedViewWindow(self.protectedviewwindows.Open(*args, **arguments))

class Range:

    def __init__(self, range=None):
        self.range = range

    @property
    def Application(self):
        return Application(self.range.Application)

    @property
    def Bold(self):
        return self.range.Bold

    @property
    def BoldBi(self):
        return self.range.BoldBi

    @property
    def BookmarkID(self):
        return self.range.BookmarkID

    @property
    def Bookmarks(self):
        return self.range.Bookmarks

    @property
    def Borders(self):
        return self.range.Borders

    @property
    def Case(self):
        return WdCharacterCase(self.range.Case)

    @Case.setter
    def Case(self, value):
        self.range.Case = value

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Cells(*args, **arguments)

    @property
    def Characters(self):
        return self.range.Characters

    @property
    def CharacterStyle(self):
        return self.range.CharacterStyle

    @property
    def CharacterWidth(self):
        return WdCharacterWidth(self.range.CharacterWidth)

    @CharacterWidth.setter
    def CharacterWidth(self, value):
        self.range.CharacterWidth = value

    @property
    def Columns(self):
        return self.range.Columns

    @property
    def CombineCharacters(self):
        return self.range.CombineCharacters

    @property
    def Comments(self):
        return self.range.Comments

    @property
    def Conflicts(self):
        return self.range.Conflicts

    @property
    def ContentControls(self):
        return ContentControls(self.range.ContentControls)

    @property
    def Creator(self):
        return self.range.Creator

    @property
    def DisableCharacterSpaceGrid(self):
        return self.range.DisableCharacterSpaceGrid

    @property
    def Document(self):
        return Document(self.range.Document)

    @property
    def Duplicate(self):
        return Range(self.range.Duplicate)

    @property
    def Editors(self):
        return Editors(self.range.Editors)

    @property
    def EmphasisMark(self):
        return WdEmphasisMark(self.range.EmphasisMark)

    @EmphasisMark.setter
    def EmphasisMark(self, value):
        self.range.EmphasisMark = value

    @property
    def End(self):
        return self.range.End

    @End.setter
    def End(self, value):
        self.range.End = value

    @property
    def EndnoteOptions(self):
        return EndnoteOptions(self.range.EndnoteOptions)

    @property
    def Endnotes(self):
        return self.range.Endnotes

    @property
    def EnhMetaFileBits(self):
        return self.range.EnhMetaFileBits

    @property
    def Fields(self):
        return self.range.Fields

    @property
    def Find(self):
        return Find(self.range.Find)

    @property
    def FitTextWidth(self):
        return self.range.FitTextWidth

    @FitTextWidth.setter
    def FitTextWidth(self, value):
        self.range.FitTextWidth = value

    @property
    def Font(self):
        return Font(self.range.Font)

    @Font.setter
    def Font(self, value):
        self.range.Font = value

    @property
    def FootnoteOptions(self):
        return FootnoteOptions(self.range.FootnoteOptions)

    @property
    def Footnotes(self):
        return self.range.Footnotes

    @property
    def FormattedText(self):
        return Range(self.range.FormattedText)

    @FormattedText.setter
    def FormattedText(self, value):
        self.range.FormattedText = value

    @property
    def FormFields(self):
        return self.range.FormFields

    @property
    def Frames(self):
        return Frames(self.range.Frames)

    @property
    def GrammarChecked(self):
        return self.range.GrammarChecked

    @property
    def GrammaticalErrors(self):
        return self.range.GrammaticalErrors

    @property
    def HighlightColorIndex(self):
        return WdColorIndex(self.range.HighlightColorIndex)

    @HighlightColorIndex.setter
    def HighlightColorIndex(self, value):
        self.range.HighlightColorIndex = value

    @property
    def HorizontalInVertical(self):
        return WdHorizontalInVerticalType(self.range.HorizontalInVertical)

    @HorizontalInVertical.setter
    def HorizontalInVertical(self, value):
        self.range.HorizontalInVertical = value

    @property
    def HTMLDivisions(self):
        return HTMLDivisions(self.range.HTMLDivisions)

    @property
    def Hyperlinks(self):
        return self.range.Hyperlinks

    @property
    def ID(self):
        return self.range.ID

    @ID.setter
    def ID(self, value):
        self.range.ID = value

    def Information(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Information(*args, **arguments)

    @property
    def InlineShapes(self):
        return self.range.InlineShapes

    @property
    def IsEndOfRowMark(self):
        return self.range.IsEndOfRowMark

    @property
    def Italic(self):
        return self.range.Italic

    @property
    def ItalicBi(self):
        return self.range.ItalicBi

    @property
    def Kana(self):
        return WdKana(self.range.Kana)

    @Kana.setter
    def Kana(self, value):
        self.range.Kana = value

    @property
    def LanguageDetected(self):
        return self.range.LanguageDetected

    @LanguageDetected.setter
    def LanguageDetected(self, value):
        self.range.LanguageDetected = value

    @property
    def LanguageID(self):
        return WdLanguageID(self.range.LanguageID)

    @LanguageID.setter
    def LanguageID(self, value):
        self.range.LanguageID = value

    @property
    def LanguageIDFarEast(self):
        return WdLanguageID(self.range.LanguageIDFarEast)

    @LanguageIDFarEast.setter
    def LanguageIDFarEast(self, value):
        self.range.LanguageIDFarEast = value

    @property
    def LanguageIDOther(self):
        return WdLanguageID(self.range.LanguageIDOther)

    @LanguageIDOther.setter
    def LanguageIDOther(self, value):
        self.range.LanguageIDOther = value

    @property
    def ListFormat(self):
        return ListFormat(self.range.ListFormat)

    @property
    def ListParagraphs(self):
        return self.range.ListParagraphs

    @property
    def ListStyle(self):
        return self.range.ListStyle

    @property
    def Locks(self):
        return CoAuthLocks(self.range.Locks)

    @property
    def NextStoryRange(self):
        return Range(self.range.NextStoryRange)

    @property
    def NoProofing(self):
        return self.range.NoProofing

    @property
    def OMaths(self):
        return OMaths(self.range.OMaths)

    @property
    def Orientation(self):
        return WdTextOrientation(self.range.Orientation)

    @Orientation.setter
    def Orientation(self, value):
        self.range.Orientation = value

    @property
    def PageSetup(self):
        return PageSetup(self.range.PageSetup)

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.range.ParagraphFormat)

    @ParagraphFormat.setter
    def ParagraphFormat(self, value):
        self.range.ParagraphFormat = value

    @property
    def Paragraphs(self):
        return self.range.Paragraphs

    @property
    def ParagraphStyle(self):
        return self.range.ParagraphStyle

    @property
    def Parent(self):
        return self.range.Parent

    @property
    def ParentContentControl(self):
        return ContentControl(self.range.ParentContentControl)

    @property
    def PreviousBookmarkID(self):
        return self.range.PreviousBookmarkID

    @property
    def ReadabilityStatistics(self):
        return self.range.ReadabilityStatistics

    @property
    def Revisions(self):
        return self.range.Revisions

    @property
    def Rows(self):
        return self.range.Rows

    @property
    def Scripts(self):
        return self.range.Scripts

    @property
    def Sections(self):
        return self.range.Sections

    @property
    def Sentences(self):
        return self.range.Sentences

    @property
    def Shading(self):
        return Shading(self.range.Shading)

    @property
    def ShapeRange(self):
        return self.range.ShapeRange

    @property
    def ShowAll(self):
        return self.range.ShowAll

    @property
    def SpellingChecked(self):
        return self.range.SpellingChecked

    @property
    def SpellingErrors(self):
        return self.range.SpellingErrors

    @property
    def Start(self):
        return self.range.Start

    @Start.setter
    def Start(self, value):
        self.range.Start = value

    @property
    def StoryLength(self):
        return self.range.StoryLength

    @property
    def StoryType(self):
        return WdStoryType(self.range.StoryType)

    @property
    def Style(self):
        return self.range.Style

    @Style.setter
    def Style(self, value):
        self.range.Style = value

    @property
    def Subdocuments(self):
        return self.range.Subdocuments

    @property
    def SynonymInfo(self):
        return SynonymInfo(self.range.SynonymInfo)

    @property
    def Tables(self):
        return self.range.Tables

    @property
    def TableStyle(self):
        return self.range.TableStyle

    @property
    def Text(self):
        return self.range.Text

    @Text.setter
    def Text(self, value):
        self.range.Text = value

    @property
    def TextRetrievalMode(self):
        return TextRetrievalMode(self.range.TextRetrievalMode)

    @property
    def TopLevelTables(self):
        return self.range.TopLevelTables

    @property
    def TwoLinesInOne(self):
        return WdTwoLinesInOneType(self.range.TwoLinesInOne)

    @TwoLinesInOne.setter
    def TwoLinesInOne(self, value):
        self.range.TwoLinesInOne = value

    @property
    def Underline(self):
        return WdUnderline(self.range.Underline)

    @Underline.setter
    def Underline(self, value):
        self.range.Underline = value

    @property
    def Updates(self):
        return self.range.Updates

    @property
    def WordOpenXML(self):
        return self.range.WordOpenXML

    @property
    def Words(self):
        return self.range.Words

    def XML(self, *args, DataOnly=None):
        arguments = {"DataOnly": DataOnly}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.XML(*args, **arguments)

    def AutoFormat(self):
        self.range.AutoFormat()

    def Calculate(self):
        self.range.Calculate()

    def CheckGrammar(self):
        self.range.CheckGrammar()

    def CheckSpelling(self, *args, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = {"CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "AlwaysSuggest": AlwaysSuggest, "CustomDictionary2": CustomDictionary2, "CustomDictionary3": CustomDictionary3, "CustomDictionary4": CustomDictionary4, "CustomDictionary5": CustomDictionary5, "CustomDictionary6": CustomDictionary6, "CustomDictionary7": CustomDictionary7, "CustomDictionary8": CustomDictionary8, "CustomDictionary9": CustomDictionary9, "CustomDictionary10": CustomDictionary10}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.CheckSpelling(*args, **arguments)

    def CheckSynonyms(self):
        self.range.CheckSynonyms()

    def Collapse(self, *args, Direction=None):
        arguments = {"Direction": Direction}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.Collapse(*args, **arguments)

    def ComputeStatistics(self, *args, Statistic=None):
        arguments = {"Statistic": Statistic}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.ComputeStatistics(*args, **arguments)

    def ConvertHangulAndHanja(self, *args, ConversionsMode=None, FastConversion=None, CheckHangulEnding=None, EnableRecentOrdering=None, CustomDictionary=None):
        arguments = {"ConversionsMode": ConversionsMode, "FastConversion": FastConversion, "CheckHangulEnding": CheckHangulEnding, "EnableRecentOrdering": EnableRecentOrdering, "CustomDictionary": CustomDictionary}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.ConvertHangulAndHanja(*args, **arguments)

    def ConvertToTable(self, *args, Separator=None, NumRows=None, NumColumns=None, InitialColumnWidth=None, Format=None, ApplyBorders=None, ApplyShading=None, ApplyFont=None, ApplyColor=None, ApplyHeadingRows=None, ApplyLastRow=None, ApplyFirstColumn=None, ApplyLastColumn=None, AutoFit=None, AutoFitBehavior=None, DefaultTableBehavior=None):
        arguments = {"Separator": Separator, "NumRows": NumRows, "NumColumns": NumColumns, "InitialColumnWidth": InitialColumnWidth, "Format": Format, "ApplyBorders": ApplyBorders, "ApplyShading": ApplyShading, "ApplyFont": ApplyFont, "ApplyColor": ApplyColor, "ApplyHeadingRows": ApplyHeadingRows, "ApplyLastRow": ApplyLastRow, "ApplyFirstColumn": ApplyFirstColumn, "ApplyLastColumn": ApplyLastColumn, "AutoFit": AutoFit, "AutoFitBehavior": AutoFitBehavior, "DefaultTableBehavior": DefaultTableBehavior}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.ConvertToTable(*args, **arguments)

    def Copy(self):
        self.range.Copy()

    def CopyAsPicture(self):
        self.range.CopyAsPicture()

    def Cut(self):
        self.range.Cut()

    def Delete(self, *args, \[_Unit_\]=None, \[_Count_\]=None):
        arguments = {"\[_Unit_\]": \[_Unit_\], "\[_Count_\]": \[_Count_\]}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Delete(*args, **arguments)

    def DetectLanguage(self):
        self.range.DetectLanguage()

    def EndOf(self, *args, Unit=None, Extend=None):
        arguments = {"Unit": Unit, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.EndOf(*args, **arguments)

    def Expand(self, *args, Unit=None):
        arguments = {"Unit": Unit}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.Expand(*args, **arguments)

    def ExportAsFixedFormat(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "ExportCurrentPage": ExportCurrentPage, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.ExportAsFixedFormat(*args, **arguments)

    def ExportAsFixedFormat2(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, OptimizeForImageQuality=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "ExportCurrentPage": ExportCurrentPage, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "OptimizeForImageQuality": OptimizeForImageQuality, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.ExportAsFixedFormat2(*args, **arguments)

    def ExportAsFixedFormat3(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, OptimizeForImageQuality=None, ImproveExportTagging=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "ExportCurrentPage": ExportCurrentPage, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "OptimizeForImageQuality": OptimizeForImageQuality, "ImproveExportTagging": ImproveExportTagging, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.ExportAsFixedFormat3(*args, **arguments)

    def ExportFragment(self, *args, FileName=None, Format=None):
        arguments = {"FileName": FileName, "Format": Format}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.ExportFragment(*args, **arguments)

    def GetSpellingSuggestions(self, *args, CustomDictionary=None, IgnoreUppercase=None, MainDictionary=None, SuggestionMode=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = {"CustomDictionary": CustomDictionary, "IgnoreUppercase": IgnoreUppercase, "MainDictionary": MainDictionary, "SuggestionMode": SuggestionMode, "CustomDictionary2": CustomDictionary2, "CustomDictionary3": CustomDictionary3, "CustomDictionary4": CustomDictionary4, "CustomDictionary5": CustomDictionary5, "CustomDictionary6": CustomDictionary6, "CustomDictionary7": CustomDictionary7, "CustomDictionary8": CustomDictionary8, "CustomDictionary9": CustomDictionary9, "CustomDictionary10": CustomDictionary10}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.GetSpellingSuggestions(*args, **arguments)

    def GoTo(self, *args, What=None, Which=None, Count=None, Name=None):
        arguments = {"What": What, "Which": Which, "Count": Count, "Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.GoTo(*args, **arguments)

    def GoToEditableRange(self, *args, EditorID=None):
        arguments = {"EditorID": EditorID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.GoToEditableRange(*args, **arguments)

    def GoToNext(self, *args, What=None):
        arguments = {"What": What}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.GoToNext(*args, **arguments)

    def GoToPrevious(self, *args, What=None):
        arguments = {"What": What}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.GoToPrevious(*args, **arguments)

    def ImportFragment(self, *args, FileName=None, MatchDestination=None):
        arguments = {"FileName": FileName, "MatchDestination": MatchDestination}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.ImportFragment(*args, **arguments)

    def InRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.InRange(*args, **arguments)

    def InsertAfter(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertAfter(*args, **arguments)

    def InsertAlignmentTab(self, *args, Alignment=None, RelativeTo=None):
        arguments = {"Alignment": Alignment, "RelativeTo": RelativeTo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertAlignmentTab(*args, **arguments)

    def InsertAutoText(self):
        self.range.InsertAutoText()

    def InsertBefore(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertBefore(*args, **arguments)

    def InsertBreak(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertBreak(*args, **arguments)

    def InsertCaption(self, *args, Label=None, Title=None, TitleAutoText=None, Position=None, ExcludeLabel=None):
        arguments = {"Label": Label, "Title": Title, "TitleAutoText": TitleAutoText, "Position": Position, "ExcludeLabel": ExcludeLabel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertCaption(*args, **arguments)

    def InsertCrossReference(self, *args, ReferenceType=None, ReferenceKind=None, ReferenceItem=None, InsertAsHyperlink=None, IncludePosition=None, SeparateNumbers=None, SeparatorString=None):
        arguments = {"ReferenceType": ReferenceType, "ReferenceKind": ReferenceKind, "ReferenceItem": ReferenceItem, "InsertAsHyperlink": InsertAsHyperlink, "IncludePosition": IncludePosition, "SeparateNumbers": SeparateNumbers, "SeparatorString": SeparatorString}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertCrossReference(*args, **arguments)

    def InsertDatabase(self, *args, Format=None, Style=None, LinkToSource=None, Connection=None, SQLStatement=None, SQLStatement1=None, PasswordDocument=None, PasswordTemplate=None, WritePasswordDocument=None, WritePasswordTemplate=None, DataSource=None, From=None, To=None, IncludeFields=None):
        arguments = {"Format": Format, "Style": Style, "LinkToSource": LinkToSource, "Connection": Connection, "SQLStatement": SQLStatement, "SQLStatement1": SQLStatement1, "PasswordDocument": PasswordDocument, "PasswordTemplate": PasswordTemplate, "WritePasswordDocument": WritePasswordDocument, "WritePasswordTemplate": WritePasswordTemplate, "DataSource": DataSource, "From": From, "To": To, "IncludeFields": IncludeFields}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertDatabase(*args, **arguments)

    def InsertDateTime(self, *args, DateTimeFormat=None, InsertAsField=None, InsertAsFullWidth=None, DateLanguage=None, CalendarType=None):
        arguments = {"DateTimeFormat": DateTimeFormat, "InsertAsField": InsertAsField, "InsertAsFullWidth": InsertAsFullWidth, "DateLanguage": DateLanguage, "CalendarType": CalendarType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertDateTime(*args, **arguments)

    def InsertFile(self, *args, FileName=None, Range=None, ConfirmConversions=None, Link=None, Attachment=None):
        arguments = {"FileName": FileName, "Range": Range, "ConfirmConversions": ConfirmConversions, "Link": Link, "Attachment": Attachment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertFile(*args, **arguments)

    def InsertParagraph(self):
        self.range.InsertParagraph()

    def InsertParagraphAfter(self):
        self.range.InsertParagraphAfter()

    def InsertParagraphBefore(self):
        self.range.InsertParagraphBefore()

    def InsertSymbol(self, *args, CharacterNumber=None, Font=None, Unicode=None, Bias=None):
        arguments = {"CharacterNumber": CharacterNumber, "Font": Font, "Unicode": Unicode, "Bias": Bias}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.InsertSymbol(*args, **arguments)

    def InsertXML(self, *args, XML=None, Transform=None):
        arguments = {"XML": XML, "Transform": Transform}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.InsertXML(*args, **arguments)

    def InStory(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.InStory(*args, **arguments)

    def IsEqual(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.IsEqual(*args, **arguments)

    def LookupNameProperties(self):
        self.range.LookupNameProperties()

    def ModifyEnclosure(self, *args, Style=None, Symbol=None, EnclosedText=None):
        arguments = {"Style": Style, "Symbol": Symbol, "EnclosedText": EnclosedText}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.ModifyEnclosure(*args, **arguments)

    def Move(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Move(*args, **arguments)

    def MoveEnd(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.MoveEnd(*args, **arguments)

    def MoveEndUntil(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.MoveEndUntil(*args, **arguments)

    def MoveEndWhile(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.MoveEndWhile(*args, **arguments)

    def MoveStart(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.MoveStart(*args, **arguments)

    def MoveStartUntil(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.MoveStartUntil(*args, **arguments)

    def MoveStartWhile(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.MoveStartWhile(*args, **arguments)

    def MoveUntil(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.MoveUntil(*args, **arguments)

    def MoveWhile(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.MoveWhile(*args, **arguments)

    def Next(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Next(*args, **arguments)

    def NextSubdocument(self):
        self.range.NextSubdocument()

    def Paste(self):
        self.range.Paste()

    def PasteAndFormat(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.PasteAndFormat(*args, **arguments)

    def PasteAppendTable(self):
        self.range.PasteAppendTable()

    def PasteAsNestedTable(self):
        self.range.PasteAsNestedTable()

    def PasteExcelTable(self, *args, LinkedToExcel=None, WordFormatting=None, RTF=None):
        arguments = {"LinkedToExcel": LinkedToExcel, "WordFormatting": WordFormatting, "RTF": RTF}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.PasteExcelTable(*args, **arguments)

    def PasteSpecial(self, *args, IconIndex=None, Link=None, Placement=None, DisplayAsIcon=None, DataType=None, IconFileName=None, IconLabel=None):
        arguments = {"IconIndex": IconIndex, "Link": Link, "Placement": Placement, "DisplayAsIcon": DisplayAsIcon, "DataType": DataType, "IconFileName": IconFileName, "IconLabel": IconLabel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.PasteSpecial(*args, **arguments)

    def PhoneticGuide(self, *args, Text=None, Alignment=None, Raise=None, FontSize=None, FontName=None):
        arguments = {"Text": Text, "Alignment": Alignment, "Raise": Raise, "FontSize": FontSize, "FontName": FontName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.PhoneticGuide(*args, **arguments)

    def Previous(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.range.Previous(*args, **arguments)

    def PreviousSubdocument(self):
        self.range.PreviousSubdocument()

    def Relocate(self, *args, Direction=None):
        arguments = {"Direction": Direction}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.Relocate(*args, **arguments)

    def Select(self):
        self.range.Select()

    def SetListLevel(self, *args, Level=None):
        arguments = {"Level": Level}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.SetListLevel(*args, **arguments)

    def SetRange(self, *args, Start=None, End=None):
        arguments = {"Start": Start, "End": End}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.SetRange(*args, **arguments)

    def Sort(self, *args, ExcludeHeader=None, FieldNumber=None, SortFieldType=None, SortOrder=None, FieldNumber2=None, SortFieldType2=None, SortOrder2=None, FieldNumber3=None, SortFieldType3=None, SortOrder3=None, SortColumn=None, Separator=None, CaseSensitive=None, BidiSort=None, IgnoreThe=None, IgnoreKashida=None, IgnoreDiacritics=None, IgnoreHe=None, LanguageID=None):
        arguments = {"ExcludeHeader": ExcludeHeader, "FieldNumber": FieldNumber, "SortFieldType": SortFieldType, "SortOrder": SortOrder, "FieldNumber2": FieldNumber2, "SortFieldType2": SortFieldType2, "SortOrder2": SortOrder2, "FieldNumber3": FieldNumber3, "SortFieldType3": SortFieldType3, "SortOrder3": SortOrder3, "SortColumn": SortColumn, "Separator": Separator, "CaseSensitive": CaseSensitive, "BidiSort": BidiSort, "IgnoreThe": IgnoreThe, "IgnoreKashida": IgnoreKashida, "IgnoreDiacritics": IgnoreDiacritics, "IgnoreHe": IgnoreHe, "LanguageID": LanguageID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.Sort(*args, **arguments)

    def SortAscending(self):
        self.range.SortAscending()

    def SortDescending(self):
        self.range.SortDescending()

    def StartOf(self, *args, Unit=None, Extend=None):
        arguments = {"Unit": Unit, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.StartOf(*args, **arguments)

    def TCSCConverter(self, *args, WdTCSCConverterDirection=None, CommonTerms=None, UseVariants=None):
        arguments = {"WdTCSCConverterDirection": WdTCSCConverterDirection, "CommonTerms": CommonTerms, "UseVariants": UseVariants}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.range.TCSCConverter(*args, **arguments)

    def WholeStory(self):
        self.range.WholeStory()

class ReadabilityStatistic:

    def __init__(self, readabilitystatistic=None):
        self.readabilitystatistic = readabilitystatistic

    @property
    def Application(self):
        return Application(self.readabilitystatistic.Application)

    @property
    def Creator(self):
        return self.readabilitystatistic.Creator

    @property
    def Name(self):
        return self.readabilitystatistic.Name

    @property
    def Parent(self):
        return self.readabilitystatistic.Parent

    @property
    def Value(self):
        return self.readabilitystatistic.Value

class RecentFile:

    def __init__(self, recentfile=None):
        self.recentfile = recentfile

    @property
    def Application(self):
        return Application(self.recentfile.Application)

    @property
    def Creator(self):
        return self.recentfile.Creator

    @property
    def Index(self):
        return self.recentfile.Index

    @property
    def Name(self):
        return self.recentfile.Name

    @property
    def Parent(self):
        return self.recentfile.Parent

    @property
    def Path(self):
        return self.recentfile.Path

    @property
    def ReadOnly(self):
        return self.recentfile.ReadOnly

    def Delete(self):
        self.recentfile.Delete()

    def Open(self):
        return self.recentfile.Open()

class Rectangle:

    def __init__(self, rectangle=None):
        self.rectangle = rectangle

    @property
    def Application(self):
        return Application(self.rectangle.Application)

    @property
    def Creator(self):
        return self.rectangle.Creator

    @property
    def Height(self):
        return self.rectangle.Height

    @property
    def Left(self):
        return self.rectangle.Left

    @property
    def Lines(self):
        return Lines(self.rectangle.Lines)

    @property
    def Parent(self):
        return self.rectangle.Parent

    @property
    def Range(self):
        return Range(self.rectangle.Range)

    @property
    def RectangleType(self):
        return WdRectangleType(self.rectangle.RectangleType)

    @property
    def Top(self):
        return self.rectangle.Top

    @property
    def Width(self):
        return self.rectangle.Width

    @Width.setter
    def Width(self, value):
        self.rectangle.Width = value

class Rectangles:

    def __init__(self, rectangles=None):
        self.rectangles = rectangles

    def __call__(self, item):
        return Rectangle(self.rectangles(item))

    @property
    def Application(self):
        return Application(self.rectangles.Application)

    @property
    def Count(self):
        return self.rectangles.Count

    @property
    def Creator(self):
        return self.rectangles.Creator

    @property
    def Parent(self):
        return self.rectangles.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.rectangles.Item(*args, **arguments)

class ReflectionFormat:

    def __init__(self, reflectionformat=None):
        self.reflectionformat = reflectionformat

    @property
    def Application(self):
        return self.reflectionformat.Application

    @property
    def Blur(self):
        return self.reflectionformat.Blur

    @Blur.setter
    def Blur(self, value):
        self.reflectionformat.Blur = value

    @property
    def Creator(self):
        return self.reflectionformat.Creator

    @property
    def Offset(self):
        return self.reflectionformat.Offset

    @Offset.setter
    def Offset(self, value):
        self.reflectionformat.Offset = value

    @property
    def Parent(self):
        return self.reflectionformat.Parent

    @property
    def Size(self):
        return self.reflectionformat.Size

    @Size.setter
    def Size(self, value):
        self.reflectionformat.Size = value

    @property
    def Transparency(self):
        return self.reflectionformat.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.reflectionformat.Transparency = value

    @property
    def Type(self):
        return self.reflectionformat.Type

    @Type.setter
    def Type(self, value):
        self.reflectionformat.Type = value

class Replacement:

    def __init__(self, replacement=None):
        self.replacement = replacement

    @property
    def Application(self):
        return Application(self.replacement.Application)

    @property
    def Creator(self):
        return self.replacement.Creator

    @property
    def Font(self):
        return Font(self.replacement.Font)

    @Font.setter
    def Font(self, value):
        self.replacement.Font = value

    @property
    def Frame(self):
        return Frame(self.replacement.Frame)

    @property
    def Highlight(self):
        return self.replacement.Highlight

    @property
    def LanguageID(self):
        return WdLanguageID(self.replacement.LanguageID)

    @LanguageID.setter
    def LanguageID(self, value):
        self.replacement.LanguageID = value

    @property
    def LanguageIDFarEast(self):
        return WdLanguageID(self.replacement.LanguageIDFarEast)

    @LanguageIDFarEast.setter
    def LanguageIDFarEast(self, value):
        self.replacement.LanguageIDFarEast = value

    @property
    def NoProofing(self):
        return self.replacement.NoProofing

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.replacement.ParagraphFormat)

    @ParagraphFormat.setter
    def ParagraphFormat(self, value):
        self.replacement.ParagraphFormat = value

    @property
    def Parent(self):
        return self.replacement.Parent

    @property
    def Style(self):
        return WdBuiltinStyle(self.replacement.Style)

    @Style.setter
    def Style(self, value):
        self.replacement.Style = value

    @property
    def Text(self):
        return self.replacement.Text

    @Text.setter
    def Text(self, value):
        self.replacement.Text = value

    def ClearFormatting(self):
        self.replacement.ClearFormatting()

class Research:

    def __init__(self, research=None):
        self.research = research

    @property
    def Application(self):
        return Application(self.research.Application)

    @property
    def Creator(self):
        return self.research.Creator

    @property
    def FavoriteService(self):
        return self.research.FavoriteService

    @FavoriteService.setter
    def FavoriteService(self, value):
        self.research.FavoriteService = value

    @property
    def Parent(self):
        return self.research.Parent

    def IsResearchService(self, *args, ServiceID=None):
        arguments = {"ServiceID": ServiceID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.research.IsResearchService(*args, **arguments)

    def Query(self, *args, ServiceID=None, QueryString=None, QueryLanguage=None, UseSelection=None, RequeryContextXML=None, NewQueryContextXML=None, LaunchQuery=None):
        arguments = {"ServiceID": ServiceID, "QueryString": QueryString, "QueryLanguage": QueryLanguage, "UseSelection": UseSelection, "RequeryContextXML": RequeryContextXML, "NewQueryContextXML": NewQueryContextXML, "LaunchQuery": LaunchQuery}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.research.Query(*args, **arguments)

    def SetLanguagePair(self, *args, LanguageFrom=None, LanguageTo=None):
        arguments = {"LanguageFrom": LanguageFrom, "LanguageTo": LanguageTo}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.research.SetLanguagePair(*args, **arguments)

class Reviewer:

    def __init__(self, reviewer=None):
        self.reviewer = reviewer

    @property
    def Application(self):
        return self.reviewer.Application

    @property
    def Creator(self):
        return self.reviewer.Creator

    @property
    def Parent(self):
        return self.reviewer.Parent

    @property
    def Visible(self):
        return self.reviewer.Visible

class Reviewers:

    def __init__(self, reviewers=None):
        self.reviewers = reviewers

    def __call__(self, item):
        return Reviewer(self.reviewers(item))

    @property
    def Application(self):
        return Application(self.reviewers.Application)

    @property
    def Count(self):
        return self.reviewers.Count

    @property
    def Creator(self):
        return self.reviewers.Creator

    @property
    def Parent(self):
        return self.reviewers.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.reviewers.Item(*args, **arguments)

class Revision:

    def __init__(self, revision=None):
        self.revision = revision

    @property
    def Application(self):
        return Application(self.revision.Application)

    @property
    def Author(self):
        return self.revision.Author

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.revision.Cells(*args, **arguments)

    @property
    def Creator(self):
        return self.revision.Creator

    @property
    def Date(self):
        return self.revision.Date

    @property
    def FormatDescription(self):
        return self.revision.FormatDescription

    @property
    def Index(self):
        return self.revision.Index

    @property
    def MovedRange(self):
        return Range(self.revision.MovedRange)

    @property
    def Parent(self):
        return self.revision.Parent

    @property
    def Range(self):
        return Range(self.revision.Range)

    @property
    def Style(self):
        return Style(self.revision.Style)

    @property
    def Type(self):
        return WdRevisionType(self.revision.Type)

    def Accept(self):
        self.revision.Accept()

    def Reject(self):
        self.revision.Reject()

class Row:

    def __init__(self, row=None):
        self.row = row

    @property
    def Alignment(self):
        return WdRowAlignment(self.row.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.row.Alignment = value

    @property
    def AllowBreakAcrossPages(self):
        return self.row.AllowBreakAcrossPages

    @property
    def Application(self):
        return Application(self.row.Application)

    @property
    def Borders(self):
        return self.row.Borders

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.row.Cells(*args, **arguments)

    @property
    def Creator(self):
        return self.row.Creator

    @property
    def HeadingFormat(self):
        return self.row.HeadingFormat

    @property
    def Height(self):
        return self.row.Height

    @Height.setter
    def Height(self, value):
        self.row.Height = value

    @property
    def HeightRule(self):
        return WdRowHeightRule(self.row.HeightRule)

    @HeightRule.setter
    def HeightRule(self, value):
        self.row.HeightRule = value

    @property
    def ID(self):
        return self.row.ID

    @ID.setter
    def ID(self, value):
        self.row.ID = value

    @property
    def Index(self):
        return self.row.Index

    @property
    def IsFirst(self):
        return self.row.IsFirst

    @property
    def IsLast(self):
        return self.row.IsLast

    @property
    def LeftIndent(self):
        return self.row.LeftIndent

    @LeftIndent.setter
    def LeftIndent(self, value):
        self.row.LeftIndent = value

    @property
    def NestingLevel(self):
        return self.row.NestingLevel

    @property
    def Next(self):
        return Row(self.row.Next)

    @property
    def Parent(self):
        return self.row.Parent

    @property
    def Previous(self):
        return Row(self.row.Previous)

    @property
    def Range(self):
        return Range(self.row.Range)

    @property
    def Shading(self):
        return Shading(self.row.Shading)

    @property
    def SpaceBetweenColumns(self):
        return self.row.SpaceBetweenColumns

    @SpaceBetweenColumns.setter
    def SpaceBetweenColumns(self, value):
        self.row.SpaceBetweenColumns = value

    def ConvertToText(self, *args, Separator=None, NestedTables=None):
        arguments = {"Separator": Separator, "NestedTables": NestedTables}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.row.ConvertToText(*args, **arguments)

    def Delete(self):
        self.row.Delete()

    def Select(self):
        self.row.Select()

    def SetHeight(self, *args, RowHeight=None, HeightRule=None):
        arguments = {"RowHeight": RowHeight, "HeightRule": HeightRule}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.row.SetHeight(*args, **arguments)

    def SetLeftIndent(self, *args, LeftIndent=None, RulerStyle=None):
        arguments = {"LeftIndent": LeftIndent, "RulerStyle": RulerStyle}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.row.SetLeftIndent(*args, **arguments)

class Section:

    def __init__(self, section=None):
        self.section = section

    @property
    def Application(self):
        return Application(self.section.Application)

    @property
    def Borders(self):
        return self.section.Borders

    @property
    def Creator(self):
        return self.section.Creator

    @property
    def Footers(self):
        return self.section.Footers

    @property
    def Headers(self):
        return self.section.Headers

    @property
    def Index(self):
        return self.section.Index

    @property
    def PageSetup(self):
        return PageSetup(self.section.PageSetup)

    @property
    def Parent(self):
        return self.section.Parent

    @property
    def ProtectedForForms(self):
        return self.section.ProtectedForForms

    @property
    def Range(self):
        return Range(self.section.Range)

class Selection:

    def __init__(self, selection=None):
        self.selection = selection

    @property
    def Active(self):
        return self.selection.Active

    @property
    def Application(self):
        return Application(self.selection.Application)

    @property
    def BookmarkID(self):
        return self.selection.BookmarkID

    @property
    def Bookmarks(self):
        return self.selection.Bookmarks

    @property
    def Borders(self):
        return self.selection.Borders

    def Cells(self, *args, RowIndex=None, ColumnIndex=None):
        arguments = {"RowIndex": RowIndex, "ColumnIndex": ColumnIndex}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.Cells(*args, **arguments)

    @property
    def Characters(self):
        return self.selection.Characters

    @property
    def ChildShapeRange(self):
        return self.selection.ChildShapeRange

    @property
    def Columns(self):
        return self.selection.Columns

    @property
    def ColumnSelectMode(self):
        return self.selection.ColumnSelectMode

    @property
    def Comments(self):
        return self.selection.Comments

    @property
    def Creator(self):
        return self.selection.Creator

    @property
    def Document(self):
        return Document(self.selection.Document)

    @property
    def Editors(self):
        return Editors(self.selection.Editors)

    @property
    def End(self):
        return self.selection.End

    @End.setter
    def End(self, value):
        self.selection.End = value

    @property
    def EndnoteOptions(self):
        return EndnoteOptions(self.selection.EndnoteOptions)

    @property
    def Endnotes(self):
        return self.selection.Endnotes

    @property
    def EnhMetaFileBits(self):
        return self.selection.EnhMetaFileBits

    @property
    def ExtendMode(self):
        return self.selection.ExtendMode

    @property
    def Fields(self):
        return self.selection.Fields

    @property
    def Find(self):
        return Find(self.selection.Find)

    @property
    def FitTextWidth(self):
        return self.selection.FitTextWidth

    @FitTextWidth.setter
    def FitTextWidth(self, value):
        self.selection.FitTextWidth = value

    @property
    def Flags(self):
        return WdSelectionFlags(self.selection.Flags)

    @Flags.setter
    def Flags(self, value):
        self.selection.Flags = value

    @property
    def Font(self):
        return Font(self.selection.Font)

    @Font.setter
    def Font(self, value):
        self.selection.Font = value

    @property
    def FootnoteOptions(self):
        return FootnoteOptions(self.selection.FootnoteOptions)

    @property
    def Footnotes(self):
        return self.selection.Footnotes

    @property
    def FormattedText(self):
        return Range(self.selection.FormattedText)

    @FormattedText.setter
    def FormattedText(self, value):
        self.selection.FormattedText = value

    @property
    def FormFields(self):
        return self.selection.FormFields

    @property
    def Frames(self):
        return Frames(self.selection.Frames)

    @property
    def HasChildShapeRange(self):
        return self.selection.HasChildShapeRange

    @property
    def HeaderFooter(self):
        return HeaderFooter(self.selection.HeaderFooter)

    @property
    def HTMLDivisions(self):
        return HTMLDivisions(self.selection.HTMLDivisions)

    @property
    def Hyperlinks(self):
        return self.selection.Hyperlinks

    def Information(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.Information(*args, **arguments)

    @property
    def InlineShapes(self):
        return self.selection.InlineShapes

    @property
    def IPAtEndOfLine(self):
        return self.selection.IPAtEndOfLine

    @property
    def IsEndOfRowMark(self):
        return self.selection.IsEndOfRowMark

    @property
    def LanguageDetected(self):
        return self.selection.LanguageDetected

    @LanguageDetected.setter
    def LanguageDetected(self, value):
        self.selection.LanguageDetected = value

    @property
    def LanguageID(self):
        return self.selection.LanguageID

    @LanguageID.setter
    def LanguageID(self, value):
        self.selection.LanguageID = value

    @property
    def LanguageIDFarEast(self):
        return WdLanguageID(self.selection.LanguageIDFarEast)

    @LanguageIDFarEast.setter
    def LanguageIDFarEast(self, value):
        self.selection.LanguageIDFarEast = value

    @property
    def LanguageIDOther(self):
        return WdLanguageID(self.selection.LanguageIDOther)

    @LanguageIDOther.setter
    def LanguageIDOther(self, value):
        self.selection.LanguageIDOther = value

    @property
    def NoProofing(self):
        return self.selection.NoProofing

    @property
    def OMaths(self):
        return OMaths(self.selection.OMaths)

    @property
    def Orientation(self):
        return WdTextOrientation(self.selection.Orientation)

    @Orientation.setter
    def Orientation(self, value):
        self.selection.Orientation = value

    @property
    def PageSetup(self):
        return PageSetup(self.selection.PageSetup)

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.selection.ParagraphFormat)

    @ParagraphFormat.setter
    def ParagraphFormat(self, value):
        self.selection.ParagraphFormat = value

    @property
    def Paragraphs(self):
        return self.selection.Paragraphs

    @property
    def Parent(self):
        return Selection(self.selection.Parent)

    @property
    def PreviousBookmarkID(self):
        return self.selection.PreviousBookmarkID

    @property
    def Range(self):
        return Range(self.selection.Range)

    @property
    def Rows(self):
        return self.selection.Rows

    @property
    def Sections(self):
        return self.selection.Sections

    @property
    def Sentences(self):
        return self.selection.Sentences

    @property
    def Shading(self):
        return Shading(self.selection.Shading)

    @property
    def ShapeRange(self):
        return self.selection.ShapeRange

    @property
    def Start(self):
        return self.selection.Start

    @Start.setter
    def Start(self, value):
        self.selection.Start = value

    @property
    def StartIsActive(self):
        return self.selection.StartIsActive

    @property
    def StoryLength(self):
        return self.selection.StoryLength

    @property
    def StoryType(self):
        return WdStoryType(self.selection.StoryType)

    @property
    def Style(self):
        return WdBuiltinStyle(self.selection.Style)

    @Style.setter
    def Style(self, value):
        self.selection.Style = value

    @property
    def Tables(self):
        return self.selection.Tables

    @property
    def Text(self):
        return self.selection.Text

    @Text.setter
    def Text(self, value):
        self.selection.Text = value

    @property
    def TopLevelTables(self):
        return self.selection.TopLevelTables

    @property
    def Type(self):
        return WdSelectionType(self.selection.Type)

    @property
    def WordOpenXML(self):
        return self.selection.WordOpenXML

    @property
    def Words(self):
        return self.selection.Words

    def XML(self, *args, DataOnly=None):
        arguments = {"DataOnly": DataOnly}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.XML(*args, **arguments)

    def BoldRun(self):
        self.selection.BoldRun()

    def Calculate(self):
        self.selection.Calculate()

    def ClearCharacterAllFormatting(self):
        self.selection.ClearCharacterAllFormatting()

    def ClearCharacterDirectFormatting(self):
        self.selection.ClearCharacterDirectFormatting()

    def ClearCharacterStyle(self):
        self.selection.ClearCharacterStyle()

    def ClearFormatting(self):
        self.selection.ClearFormatting()

    def ClearParagraphAllFormatting(self):
        self.selection.ClearParagraphAllFormatting()

    def ClearParagraphDirectFormatting(self):
        self.selection.ClearParagraphDirectFormatting()

    def ClearParagraphStyle(self):
        self.selection.ClearParagraphStyle()

    def Collapse(self, *args, Direction=None):
        arguments = {"Direction": Direction}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.Collapse(*args, **arguments)

    def ConvertToTable(self, *args, Separator=None, NumRows=None, NumColumns=None, InitialColumnWidth=None, Format=None, ApplyBorders=None, ApplyShading=None, ApplyFont=None, ApplyColor=None, ApplyHeadingRows=None, ApplyLastRow=None, ApplyFirstColumn=None, ApplyLastColumn=None, AutoFit=None, AutoFitBehavior=None, DefaultTableBehavior=None):
        arguments = {"Separator": Separator, "NumRows": NumRows, "NumColumns": NumColumns, "InitialColumnWidth": InitialColumnWidth, "Format": Format, "ApplyBorders": ApplyBorders, "ApplyShading": ApplyShading, "ApplyFont": ApplyFont, "ApplyColor": ApplyColor, "ApplyHeadingRows": ApplyHeadingRows, "ApplyLastRow": ApplyLastRow, "ApplyFirstColumn": ApplyFirstColumn, "ApplyLastColumn": ApplyLastColumn, "AutoFit": AutoFit, "AutoFitBehavior": AutoFitBehavior, "DefaultTableBehavior": DefaultTableBehavior}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.ConvertToTable(*args, **arguments)

    def Copy(self):
        self.selection.Copy()

    def CopyAsPicture(self):
        self.selection.CopyAsPicture()

    def CopyFormat(self):
        self.selection.CopyFormat()

    def CreateAutoTextEntry(self, *args, Name=None, StyleName=None):
        arguments = {"Name": Name, "StyleName": StyleName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.CreateAutoTextEntry(*args, **arguments)

    def CreateTextbox(self):
        self.selection.CreateTextbox()

    def Cut(self):
        self.selection.Cut()

    def Delete(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.Delete(*args, **arguments)

    def DetectLanguage(self):
        self.selection.DetectLanguage()

    def EndKey(self, *args, Unit=None, Extend=None):
        arguments = {"Unit": Unit, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.EndKey(*args, **arguments)

    def EndOf(self, *args, Unit=None, Extend=None):
        arguments = {"Unit": Unit, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.EndOf(*args, **arguments)

    def EscapeKey(self):
        self.selection.EscapeKey()

    def Expand(self, *args, Unit=None):
        arguments = {"Unit": Unit}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.Expand(*args, **arguments)

    def ExportAsFixedFormat(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "ExportCurrentPage": ExportCurrentPage, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.ExportAsFixedFormat(*args, **arguments)

    def ExportAsFixedFormat2(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, OptimizeForImageQuality=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "ExportCurrentPage": ExportCurrentPage, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "OptimizeForImageQuality": OptimizeForImageQuality, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.ExportAsFixedFormat2(*args, **arguments)

    def ExportAsFixedFormat3(self, *args, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005\_1=None, OptimizeForImageQuality=None, ImproveExportTagging=None, FixedFormatExtClassPtr=None):
        arguments = {"OutputFileName": OutputFileName, "ExportFormat": ExportFormat, "OpenAfterExport": OpenAfterExport, "OptimizeFor": OptimizeFor, "ExportCurrentPage": ExportCurrentPage, "Item": Item, "IncludeDocProps": IncludeDocProps, "KeepIRM": KeepIRM, "CreateBookmarks": CreateBookmarks, "DocStructureTags": DocStructureTags, "BitmapMissingFonts": BitmapMissingFonts, "UseISO19005\_1": UseISO19005\_1, "OptimizeForImageQuality": OptimizeForImageQuality, "ImproveExportTagging": ImproveExportTagging, "FixedFormatExtClassPtr": FixedFormatExtClassPtr}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.ExportAsFixedFormat3(*args, **arguments)

    def Extend(self, *args, Character=None):
        arguments = {"Character": Character}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.Extend(*args, **arguments)

    def GoTo(self, *args, What=None, Which=None, Count=None, Name=None):
        arguments = {"What": What, "Which": Which, "Count": Count, "Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.GoTo(*args, **arguments)

    def GoToEditableRange(self, *args, EditorID=None):
        arguments = {"EditorID": EditorID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.GoToEditableRange(*args, **arguments)

    def GoToNext(self, *args, What=None):
        arguments = {"What": What}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.GoToNext(*args, **arguments)

    def GoToPrevious(self, *args, What=None):
        arguments = {"What": What}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.GoToPrevious(*args, **arguments)

    def HomeKey(self, *args, Unit=None, Extend=None):
        arguments = {"Unit": Unit, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.HomeKey(*args, **arguments)

    def InRange(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.InRange(*args, **arguments)

    def InsertAfter(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertAfter(*args, **arguments)

    def InsertBefore(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertBefore(*args, **arguments)

    def InsertBreak(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertBreak(*args, **arguments)

    def InsertCaption(self, *args, Label=None, Title=None, TitleAutoText=None, Position=None, ExcludeLabel=None):
        arguments = {"Label": Label, "Title": Title, "TitleAutoText": TitleAutoText, "Position": Position, "ExcludeLabel": ExcludeLabel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertCaption(*args, **arguments)

    def InsertCells(self, *args, ShiftCells=None):
        arguments = {"ShiftCells": ShiftCells}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertCells(*args, **arguments)

    def InsertColumns(self):
        self.selection.InsertColumns()

    def InsertColumnsRight(self):
        self.selection.InsertColumnsRight()

    def InsertCrossReference(self, *args, ReferenceType=None, ReferenceKind=None, ReferenceItem=None, InsertAsHyperlink=None, IncludePosition=None, SeparateNumbers=None, SeparatorString=None):
        arguments = {"ReferenceType": ReferenceType, "ReferenceKind": ReferenceKind, "ReferenceItem": ReferenceItem, "InsertAsHyperlink": InsertAsHyperlink, "IncludePosition": IncludePosition, "SeparateNumbers": SeparateNumbers, "SeparatorString": SeparatorString}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertCrossReference(*args, **arguments)

    def InsertDateTime(self, *args, DateTimeFormat=None, InsertAsField=None, InsertAsFullWidth=None, DateLanguage=None, CalendarType=None):
        arguments = {"DateTimeFormat": DateTimeFormat, "InsertAsField": InsertAsField, "InsertAsFullWidth": InsertAsFullWidth, "DateLanguage": DateLanguage, "CalendarType": CalendarType}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertDateTime(*args, **arguments)

    def InsertFile(self, *args, FileName=None, Range=None, ConfirmConversions=None, Link=None, Attachment=None):
        arguments = {"FileName": FileName, "Range": Range, "ConfirmConversions": ConfirmConversions, "Link": Link, "Attachment": Attachment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertFile(*args, **arguments)

    def InsertFormula(self, *args, Formula=None, NumberFormat=None):
        arguments = {"Formula": Formula, "NumberFormat": NumberFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertFormula(*args, **arguments)

    def InsertNewPage(self):
        self.selection.InsertNewPage()

    def InsertParagraph(self):
        self.selection.InsertParagraph()

    def InsertParagraphAfter(self):
        self.selection.InsertParagraphAfter()

    def InsertParagraphBefore(self):
        self.selection.InsertParagraphBefore()

    def InsertRows(self, *args, NumRows=None):
        arguments = {"NumRows": NumRows}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertRows(*args, **arguments)

    def InsertRowsAbove(self):
        self.selection.InsertRowsAbove()

    def InsertRowsBelow(self):
        self.selection.InsertRowsBelow()

    def InsertStyleSeparator(self):
        self.selection.InsertStyleSeparator()

    def InsertSymbol(self, *args, CharacterNumber=None, Font=None, Unicode=None, Bias=None):
        arguments = {"CharacterNumber": CharacterNumber, "Font": Font, "Unicode": Unicode, "Bias": Bias}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.InsertSymbol(*args, **arguments)

    def InsertXML(self, *args, XML=None, Transform=None):
        arguments = {"XML": XML, "Transform": Transform}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.InsertXML(*args, **arguments)

    def InStory(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.InStory(*args, **arguments)

    def IsEqual(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.IsEqual(*args, **arguments)

    def ItalicRun(self):
        self.selection.ItalicRun()

    def LtrPara(self):
        self.selection.LtrPara()

    def LtrRun(self):
        self.selection.LtrRun()

    def Move(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.Move(*args, **arguments)

    def MoveDown(self, *args, Unit=None, Count=None, Extend=None):
        arguments = {"Unit": Unit, "Count": Count, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.MoveDown(*args, **arguments)

    def MoveEnd(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.MoveEnd(*args, **arguments)

    def MoveEndUntil(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.MoveEndUntil(*args, **arguments)

    def MoveEndWhile(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.MoveEndWhile(*args, **arguments)

    def MoveLeft(self, *args, Unit=None, Count=None, Extend=None):
        arguments = {"Unit": Unit, "Count": Count, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.MoveLeft(*args, **arguments)

    def MoveRight(self, *args, Unit=None, Count=None, Extend=None):
        arguments = {"Unit": Unit, "Count": Count, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.MoveRight(*args, **arguments)

    def MoveStart(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.MoveStart(*args, **arguments)

    def MoveStartUntil(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.MoveStartUntil(*args, **arguments)

    def MoveStartWhile(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.MoveStartWhile(*args, **arguments)

    def MoveUntil(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.MoveUntil(*args, **arguments)

    def MoveUp(self, *args, Unit=None, Count=None, Extend=None):
        arguments = {"Unit": Unit, "Count": Count, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.MoveUp(*args, **arguments)

    def MoveWhile(self, *args, Cset=None, Count=None):
        arguments = {"Cset": Cset, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.MoveWhile(*args, **arguments)

    def Next(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.Next(*args, **arguments)

    def NextField(self):
        return self.selection.NextField()

    def NextRevision(self, *args, Wrap=None):
        arguments = {"Wrap": Wrap}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.NextRevision(*args, **arguments)

    def NextSubdocument(self):
        self.selection.NextSubdocument()

    def Paste(self):
        self.selection.Paste()

    def PasteAndFormat(self, *args, Type=None):
        arguments = {"Type": Type}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.PasteAndFormat(*args, **arguments)

    def PasteAppendTable(self):
        self.selection.PasteAppendTable()

    def PasteAsNestedTable(self):
        self.selection.PasteAsNestedTable()

    def PasteExcelTable(self, *args, LinkedToExcel=None, WordFormatting=None, RTF=None):
        arguments = {"LinkedToExcel": LinkedToExcel, "WordFormatting": WordFormatting, "RTF": RTF}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.PasteExcelTable(*args, **arguments)

    def PasteFormat(self):
        self.selection.PasteFormat()

    def PasteSpecial(self, *args, IconIndex=None, Link=None, Placement=None, DisplayAsIcon=None, DataType=None, IconFileName=None, IconLabel=None):
        arguments = {"IconIndex": IconIndex, "Link": Link, "Placement": Placement, "DisplayAsIcon": DisplayAsIcon, "DataType": DataType, "IconFileName": IconFileName, "IconLabel": IconLabel}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.PasteSpecial(*args, **arguments)

    def Previous(self, *args, Unit=None, Count=None):
        arguments = {"Unit": Unit, "Count": Count}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.Previous(*args, **arguments)

    def PreviousField(self):
        return self.selection.PreviousField()

    def PreviousRevision(self, *args, Wrap=None):
        arguments = {"Wrap": Wrap}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.selection.PreviousRevision(*args, **arguments)

    def PreviousSubdocument(self):
        self.selection.PreviousSubdocument()

    def ReadingModeGrowFont(self):
        return self.selection.ReadingModeGrowFont()

    def ReadingModeShrinkFont(self):
        return self.selection.ReadingModeShrinkFont()

    def RtlPara(self):
        self.selection.RtlPara()

    def RtlRun(self):
        self.selection.RtlRun()

    def Select(self):
        self.selection.Select()

    def SelectCell(self):
        self.selection.SelectCell()

    def SelectColumn(self):
        self.selection.SelectColumn()

    def SelectCurrentAlignment(self):
        self.selection.SelectCurrentAlignment()

    def SelectCurrentColor(self):
        self.selection.SelectCurrentColor()

    def SelectCurrentFont(self):
        self.selection.SelectCurrentFont()

    def SelectCurrentIndent(self):
        self.selection.SelectCurrentIndent()

    def SelectCurrentSpacing(self):
        self.selection.SelectCurrentSpacing()

    def SelectCurrentTabs(self):
        self.selection.SelectCurrentTabs()

    def SelectRow(self):
        self.selection.SelectRow()

    def SetRange(self, *args, Start=None, End=None):
        arguments = {"Start": Start, "End": End}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.SetRange(*args, **arguments)

    def Shrink(self):
        self.selection.Shrink()

    def ShrinkDiscontiguousSelection(self):
        self.selection.ShrinkDiscontiguousSelection()

    def Sort(self, *args, ExcludeHeader=None, FieldNumber=None, SortFieldType=None, SortOrder=None, FieldNumber2=None, SortFieldType2=None, SortOrder2=None, FieldNumber3=None, SortFieldType3=None, SortOrder3=None, SortColumn=None, Separator=None, CaseSensitive=None, BidiSort=None, IgnoreThe=None, IgnoreKashida=None, IgnoreDiacritics=None, IgnoreHe=None, LanguageID=None, SubFieldNumber=None, SubFieldNumber2=None, SubFieldNumber3=None):
        arguments = {"ExcludeHeader": ExcludeHeader, "FieldNumber": FieldNumber, "SortFieldType": SortFieldType, "SortOrder": SortOrder, "FieldNumber2": FieldNumber2, "SortFieldType2": SortFieldType2, "SortOrder2": SortOrder2, "FieldNumber3": FieldNumber3, "SortFieldType3": SortFieldType3, "SortOrder3": SortOrder3, "SortColumn": SortColumn, "Separator": Separator, "CaseSensitive": CaseSensitive, "BidiSort": BidiSort, "IgnoreThe": IgnoreThe, "IgnoreKashida": IgnoreKashida, "IgnoreDiacritics": IgnoreDiacritics, "IgnoreHe": IgnoreHe, "LanguageID": LanguageID, "SubFieldNumber": SubFieldNumber, "SubFieldNumber2": SubFieldNumber2, "SubFieldNumber3": SubFieldNumber3}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.Sort(*args, **arguments)

    def SortAscending(self):
        self.selection.SortAscending()

    def SortDescending(self):
        self.selection.SortDescending()

    def SplitTable(self):
        self.selection.SplitTable()

    def StartOf(self, *args, Unit=None, Extend=None):
        arguments = {"Unit": Unit, "Extend": Extend}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.StartOf(*args, **arguments)

    def ToggleCharacterCode(self):
        self.selection.ToggleCharacterCode()

    def TypeBackspace(self):
        self.selection.TypeBackspace()

    def TypeParagraph(self):
        self.selection.TypeParagraph()

    def TypeText(self, *args, Text=None):
        arguments = {"Text": Text}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.selection.TypeText(*args, **arguments)

    def WholeStory(self):
        self.selection.WholeStory()

class Series:

    def __init__(self, series=None):
        self.series = series

    @property
    def Application(self):
        return self.series.Application

    @property
    def ApplyPictToEnd(self):
        return self.series.ApplyPictToEnd

    @property
    def ApplyPictToFront(self):
        return self.series.ApplyPictToFront

    @property
    def ApplyPictToSides(self):
        return self.series.ApplyPictToSides

    @property
    def AxisGroup(self):
        return self.series.AxisGroup

    @property
    def BarShape(self):
        return self.series.BarShape

    @BarShape.setter
    def BarShape(self, value):
        self.series.BarShape = value

    @property
    def Border(self):
        return ChartBorder(self.series.Border)

    @property
    def BubbleSizes(self):
        return self.series.BubbleSizes

    @BubbleSizes.setter
    def BubbleSizes(self, value):
        self.series.BubbleSizes = value

    @property
    def ChartType(self):
        return self.series.ChartType

    @ChartType.setter
    def ChartType(self, value):
        self.series.ChartType = value

    @property
    def Creator(self):
        return self.series.Creator

    @property
    def ErrorBars(self):
        return ErrorBars(self.series.ErrorBars)

    @property
    def Explosion(self):
        return self.series.Explosion

    @Explosion.setter
    def Explosion(self, value):
        self.series.Explosion = value

    @property
    def Format(self):
        return ChartFormat(self.series.Format)

    @property
    def Formula(self):
        return self.series.Formula

    @Formula.setter
    def Formula(self, value):
        self.series.Formula = value

    @property
    def FormulaLocal(self):
        return self.series.FormulaLocal

    @FormulaLocal.setter
    def FormulaLocal(self, value):
        self.series.FormulaLocal = value

    @property
    def FormulaR1C1(self):
        return self.series.FormulaR1C1

    @FormulaR1C1.setter
    def FormulaR1C1(self, value):
        self.series.FormulaR1C1 = value

    @property
    def FormulaR1C1Local(self):
        return self.series.FormulaR1C1Local

    @FormulaR1C1Local.setter
    def FormulaR1C1Local(self, value):
        self.series.FormulaR1C1Local = value

    @property
    def Has3DEffect(self):
        return self.series.Has3DEffect

    @property
    def HasDataLabels(self):
        return self.series.HasDataLabels

    @property
    def HasErrorBars(self):
        return self.series.HasErrorBars

    @property
    def HasLeaderLines(self):
        return self.series.HasLeaderLines

    @property
    def InvertColor(self):
        return self.series.InvertColor

    @InvertColor.setter
    def InvertColor(self, value):
        self.series.InvertColor = value

    @property
    def InvertColorIndex(self):
        return self.series.InvertColorIndex

    @InvertColorIndex.setter
    def InvertColorIndex(self, value):
        self.series.InvertColorIndex = value

    @property
    def InvertIfNegative(self):
        return self.series.InvertIfNegative

    @property
    def LeaderLines(self):
        return LeaderLines(self.series.LeaderLines)

    @property
    def MarkerBackgroundColor(self):
        return self.series.MarkerBackgroundColor

    @property
    def MarkerBackgroundColorIndex(self):
        return self.series.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.series.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.series.MarkerForegroundColor

    @property
    def MarkerForegroundColorIndex(self):
        return self.series.MarkerForegroundColorIndex

    @MarkerForegroundColorIndex.setter
    def MarkerForegroundColorIndex(self, value):
        self.series.MarkerForegroundColorIndex = value

    @property
    def MarkerSize(self):
        return self.series.MarkerSize

    @MarkerSize.setter
    def MarkerSize(self, value):
        self.series.MarkerSize = value

    @property
    def MarkerStyle(self):
        return self.series.MarkerStyle

    @MarkerStyle.setter
    def MarkerStyle(self, value):
        self.series.MarkerStyle = value

    @property
    def Name(self):
        return self.series.Name

    @Name.setter
    def Name(self, value):
        self.series.Name = value

    @property
    def Parent(self):
        return self.series.Parent

    @property
    def PictureType(self):
        return self.series.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.series.PictureType = value

    @property
    def PictureUnit2(self):
        return self.series.PictureUnit2

    @PictureUnit2.setter
    def PictureUnit2(self, value):
        self.series.PictureUnit2 = value

    @property
    def PlotColorIndex(self):
        return self.series.PlotColorIndex

    @property
    def PlotOrder(self):
        return self.series.PlotOrder

    @PlotOrder.setter
    def PlotOrder(self, value):
        self.series.PlotOrder = value

    @property
    def Shadow(self):
        return self.series.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.series.Shadow = value

    @property
    def Smooth(self):
        return self.series.Smooth

    @property
    def Type(self):
        return self.series.Type

    @Type.setter
    def Type(self, value):
        self.series.Type = value

    @property
    def Values(self):
        return self.series.Values

    @Values.setter
    def Values(self, value):
        self.series.Values = value

    @property
    def XValues(self):
        return self.series.XValues

    @XValues.setter
    def XValues(self, value):
        self.series.XValues = value

    def ApplyDataLabels(self, *args, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = {"Type": Type, "LegendKey": LegendKey, "AutoText": AutoText, "HasLeaderLines": HasLeaderLines, "ShowSeriesName": ShowSeriesName, "ShowCategoryName": ShowCategoryName, "ShowValue": ShowValue, "ShowPercentage": ShowPercentage, "ShowBubbleSize": ShowBubbleSize, "Separator": Separator}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.series.ApplyDataLabels(*args, **arguments)

    def ClearFormats(self):
        self.series.ClearFormats()

    def Copy(self):
        self.series.Copy()

    def DataLabels(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.series.DataLabels(*args, **arguments)

    def Delete(self):
        self.series.Delete()

    def ErrorBar(self, *args, Direction=None, Include=None, Type=None, Amount=None, MinusValues=None):
        arguments = {"Direction": Direction, "Include": Include, "Type": Type, "Amount": Amount, "MinusValues": MinusValues}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.series.ErrorBar(*args, **arguments)

    def Paste(self):
        self.series.Paste()

    def Points(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Points(self.series.Points(*args, **arguments))

    def Select(self):
        self.series.Select()

    def Trendlines(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Trendlines(self.series.Trendlines(*args, **arguments))

class SeriesCollection:

    def __init__(self, seriescollection=None):
        self.seriescollection = seriescollection

    @property
    def Application(self):
        return self.seriescollection.Application

    @property
    def Count(self):
        return self.seriescollection.Count

    @property
    def Creator(self):
        return self.seriescollection.Creator

    @property
    def Parent(self):
        return self.seriescollection.Parent

    def Add(self, *args, Source=None, Rowcol=None, SeriesLabels=None, CategoryLabels=None, Replace=None):
        arguments = {"Source": Source, "Rowcol": Rowcol, "SeriesLabels": SeriesLabels, "CategoryLabels": CategoryLabels, "Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Series(self.seriescollection.Add(*args, **arguments))

    def Extend(self, *args, Source=None, Rowcol=None, CategoryLabels=None):
        arguments = {"Source": Source, "Rowcol": Rowcol, "CategoryLabels": CategoryLabels}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.seriescollection.Extend(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Series(self.seriescollection.Item(*args, **arguments))

    def NewSeries(self):
        return Series(self.seriescollection.NewSeries())

class SeriesLines:

    def __init__(self, serieslines=None):
        self.serieslines = serieslines

    @property
    def Application(self):
        return self.serieslines.Application

    @property
    def Border(self):
        return ChartBorder(self.serieslines.Border)

    @property
    def Creator(self):
        return self.serieslines.Creator

    @property
    def Format(self):
        return ChartFormat(self.serieslines.Format)

    @property
    def Name(self):
        return self.serieslines.Name

    @property
    def Parent(self):
        return self.serieslines.Parent

    def Delete(self):
        self.serieslines.Delete()

    def Select(self):
        self.serieslines.Select()

class Shading:

    def __init__(self, shading=None):
        self.shading = shading

    @property
    def Application(self):
        return Application(self.shading.Application)

    @property
    def BackgroundPatternColor(self):
        return Shading(self.shading.BackgroundPatternColor)

    @BackgroundPatternColor.setter
    def BackgroundPatternColor(self, value):
        self.shading.BackgroundPatternColor = value

    @property
    def BackgroundPatternColorIndex(self):
        return Shading(self.shading.BackgroundPatternColorIndex)

    @BackgroundPatternColorIndex.setter
    def BackgroundPatternColorIndex(self, value):
        self.shading.BackgroundPatternColorIndex = value

    @property
    def Creator(self):
        return self.shading.Creator

    @property
    def ForegroundPatternColor(self):
        return Shading(self.shading.ForegroundPatternColor)

    @ForegroundPatternColor.setter
    def ForegroundPatternColor(self, value):
        self.shading.ForegroundPatternColor = value

    @property
    def ForegroundPatternColorIndex(self):
        return Shading(self.shading.ForegroundPatternColorIndex)

    @ForegroundPatternColorIndex.setter
    def ForegroundPatternColorIndex(self, value):
        self.shading.ForegroundPatternColorIndex = value

    @property
    def Parent(self):
        return self.shading.Parent

    @property
    def Texture(self):
        return WdTextureIndex(self.shading.Texture)

    @Texture.setter
    def Texture(self, value):
        self.shading.Texture = value

class ShadowFormat:

    def __init__(self, shadowformat=None):
        self.shadowformat = shadowformat

    @property
    def Application(self):
        return Application(self.shadowformat.Application)

    @property
    def Blur(self):
        return self.shadowformat.Blur

    @Blur.setter
    def Blur(self, value):
        self.shadowformat.Blur = value

    @property
    def Creator(self):
        return self.shadowformat.Creator

    @property
    def ForeColor(self):
        return ColorFormat(self.shadowformat.ForeColor)

    @ForeColor.setter
    def ForeColor(self, value):
        self.shadowformat.ForeColor = value

    @property
    def Obscured(self):
        return self.shadowformat.Obscured

    @property
    def OffsetX(self):
        return self.shadowformat.OffsetX

    @OffsetX.setter
    def OffsetX(self, value):
        self.shadowformat.OffsetX = value

    @property
    def OffsetY(self):
        return self.shadowformat.OffsetY

    @OffsetY.setter
    def OffsetY(self, value):
        self.shadowformat.OffsetY = value

    @property
    def Parent(self):
        return self.shadowformat.Parent

    @property
    def RotateWithShape(self):
        return self.shadowformat.RotateWithShape

    @RotateWithShape.setter
    def RotateWithShape(self, value):
        self.shadowformat.RotateWithShape = value

    @property
    def Size(self):
        return self.shadowformat.Size

    @Size.setter
    def Size(self, value):
        self.shadowformat.Size = value

    @property
    def Style(self):
        return self.shadowformat.Style

    @Style.setter
    def Style(self, value):
        self.shadowformat.Style = value

    @property
    def Transparency(self):
        return self.shadowformat.Transparency

    @Transparency.setter
    def Transparency(self, value):
        self.shadowformat.Transparency = value

    @property
    def Type(self):
        return self.shadowformat.Type

    @Type.setter
    def Type(self, value):
        self.shadowformat.Type = value

    @property
    def Visible(self):
        return self.shadowformat.Visible

    def IncrementOffsetX(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shadowformat.IncrementOffsetX(*args, **arguments)

    def IncrementOffsetY(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shadowformat.IncrementOffsetY(*args, **arguments)

class Shape:

    def __init__(self, shape=None):
        self.shape = shape

    @property
    def Adjustments(self):
        return Adjustments(self.shape.Adjustments)

    @property
    def AlternativeText(self):
        return self.shape.AlternativeText

    @AlternativeText.setter
    def AlternativeText(self, value):
        self.shape.AlternativeText = value

    @property
    def Anchor(self):
        return Range(self.shape.Anchor)

    @property
    def Application(self):
        return Application(self.shape.Application)

    @property
    def AutoShapeType(self):
        return Shape(self.shape.AutoShapeType)

    @AutoShapeType.setter
    def AutoShapeType(self, value):
        self.shape.AutoShapeType = value

    @property
    def BackgroundStyle(self):
        return self.shape.BackgroundStyle

    @property
    def Callout(self):
        return CalloutFormat(self.shape.Callout)

    @property
    def CanvasItems(self):
        return CanvasShapes(self.shape.CanvasItems)

    @property
    def Chart(self):
        return Chart(self.shape.Chart)

    @property
    def Child(self):
        return self.shape.Child

    @property
    def Creator(self):
        return self.shape.Creator

    @property
    def Decorative(self):
        return self.shape.Decorative

    @property
    def Fill(self):
        return FillFormat(self.shape.Fill)

    @property
    def Glow(self):
        return GlowFormat(self.shape.Glow)

    @property
    def GraphicStyle(self):
        return self.shape.GraphicStyle

    @GraphicStyle.setter
    def GraphicStyle(self, value):
        self.shape.GraphicStyle = value

    @property
    def GroupItems(self):
        return self.shape.GroupItems

    @property
    def HasChart(self):
        return self.shape.HasChart

    @property
    def HasSmartArt(self):
        return self.shape.HasSmartArt

    @property
    def Height(self):
        return self.shape.Height

    @Height.setter
    def Height(self, value):
        self.shape.Height = value

    @property
    def HeightRelative(self):
        return self.shape.HeightRelative

    @HeightRelative.setter
    def HeightRelative(self, value):
        self.shape.HeightRelative = value

    @property
    def HorizontalFlip(self):
        return self.shape.HorizontalFlip

    @property
    def Hyperlink(self):
        return Hyperlink(self.shape.Hyperlink)

    @property
    def ID(self):
        return self.shape.ID

    @property
    def LayoutInCell(self):
        return self.shape.LayoutInCell

    @property
    def Left(self):
        return WdShapePosition(self.shape.Left)

    @Left.setter
    def Left(self, value):
        self.shape.Left = value

    @property
    def LeftRelative(self):
        return self.shape.LeftRelative

    @LeftRelative.setter
    def LeftRelative(self, value):
        self.shape.LeftRelative = value

    @property
    def Line(self):
        return LineFormat(self.shape.Line)

    @property
    def LinkFormat(self):
        return LinkFormat(self.shape.LinkFormat)

    @property
    def LockAnchor(self):
        return self.shape.LockAnchor

    @property
    def LockAspectRatio(self):
        return self.shape.LockAspectRatio

    @property
    def Model3D(self):
        return Model3DFormat(self.shape.Model3D)

    @property
    def Name(self):
        return self.shape.Name

    @Name.setter
    def Name(self, value):
        self.shape.Name = value

    @property
    def Nodes(self):
        return self.shape.Nodes

    @property
    def OLEFormat(self):
        return OLEFormat(self.shape.OLEFormat)

    @property
    def Parent(self):
        return self.shape.Parent

    @property
    def ParentGroup(self):
        return Shape(self.shape.ParentGroup)

    @property
    def PictureFormat(self):
        return PictureFormat(self.shape.PictureFormat)

    @property
    def Reflection(self):
        return ReflectionFormat(self.shape.Reflection)

    @property
    def RelativeHorizontalPosition(self):
        return self.shape.RelativeHorizontalPosition

    @property
    def RelativeHorizontalSize(self):
        return WdRelativeVerticalSize(self.shape.RelativeHorizontalSize)

    @RelativeHorizontalSize.setter
    def RelativeHorizontalSize(self, value):
        self.shape.RelativeHorizontalSize = value

    @property
    def RelativeVerticalPosition(self):
        return self.shape.RelativeVerticalPosition

    @property
    def RelativeVerticalSize(self):
        return WdRelativeVerticalSize(self.shape.RelativeVerticalSize)

    @RelativeVerticalSize.setter
    def RelativeVerticalSize(self, value):
        self.shape.RelativeVerticalSize = value

    @property
    def Rotation(self):
        return self.shape.Rotation

    @Rotation.setter
    def Rotation(self, value):
        self.shape.Rotation = value

    @property
    def Script(self):
        return self.shape.Script

    @property
    def Shadow(self):
        return ShadowFormat(self.shape.Shadow)

    @property
    def ShapeStyle(self):
        return self.shape.ShapeStyle

    @ShapeStyle.setter
    def ShapeStyle(self, value):
        self.shape.ShapeStyle = value

    @property
    def SmartArt(self):
        return self.shape.SmartArt

    @property
    def SoftEdge(self):
        return SoftEdgeFormat(self.shape.SoftEdge)

    @property
    def TextEffect(self):
        return TextEffectFormat(self.shape.TextEffect)

    @property
    def TextFrame(self):
        return TextFrame(self.shape.TextFrame)

    @property
    def TextFrame2(self):
        return self.shape.TextFrame2

    @property
    def ThreeD(self):
        return ThreeDFormat(self.shape.ThreeD)

    @property
    def Title(self):
        return self.shape.Title

    @Title.setter
    def Title(self, value):
        self.shape.Title = value

    @property
    def Top(self):
        return self.shape.Top

    @Top.setter
    def Top(self, value):
        self.shape.Top = value

    @property
    def TopRelative(self):
        return self.shape.TopRelative

    @TopRelative.setter
    def TopRelative(self, value):
        self.shape.TopRelative = value

    @property
    def Type(self):
        return self.shape.Type

    @property
    def VerticalFlip(self):
        return self.shape.VerticalFlip

    @property
    def Vertices(self):
        return self.shape.Vertices

    @property
    def Visible(self):
        return self.shape.Visible

    @property
    def Width(self):
        return self.shape.Width

    @Width.setter
    def Width(self, value):
        self.shape.Width = value

    @property
    def WidthRelative(self):
        return self.shape.WidthRelative

    @WidthRelative.setter
    def WidthRelative(self, value):
        self.shape.WidthRelative = value

    @property
    def WrapFormat(self):
        return WrapFormat(self.shape.WrapFormat)

    @property
    def ZOrderPosition(self):
        return self.shape.ZOrderPosition

    def Apply(self):
        self.shape.Apply()

    def CanvasCropBottom(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.CanvasCropBottom(*args, **arguments)

    def CanvasCropLeft(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.CanvasCropLeft(*args, **arguments)

    def CanvasCropRight(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.CanvasCropRight(*args, **arguments)

    def CanvasCropTop(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.CanvasCropTop(*args, **arguments)

    def ConvertToInlineShape(self):
        self.shape.ConvertToInlineShape()

    def Delete(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.Delete(*args, **arguments)

    def Duplicate(self):
        self.shape.Duplicate()

    def Flip(self, *args, FlipCmd=None):
        arguments = {"FlipCmd": FlipCmd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.Flip(*args, **arguments)

    def IncrementLeft(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.IncrementLeft(*args, **arguments)

    def IncrementRotation(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.IncrementRotation(*args, **arguments)

    def IncrementTop(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.IncrementTop(*args, **arguments)

    def PickUp(self):
        self.shape.PickUp()

    def ScaleHeight(self, *args, Factor=None, RelativeToOriginalSize=None, Scale=None):
        arguments = {"Factor": Factor, "RelativeToOriginalSize": RelativeToOriginalSize, "Scale": Scale}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.ScaleHeight(*args, **arguments)

    def ScaleWidth(self, *args, Factor=None, RelativeToOriginalSize=None, Scale=None):
        arguments = {"Factor": Factor, "RelativeToOriginalSize": RelativeToOriginalSize, "Scale": Scale}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.ScaleWidth(*args, **arguments)

    def Select(self, *args, Replace=None):
        arguments = {"Replace": Replace}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.shape.Select(*args, **arguments)

    def SetShapesDefaultProperties(self):
        self.shape.SetShapesDefaultProperties()

    def Ungroup(self):
        return self.shape.Ungroup()

    def ZOrder(self, *args, ZOrderCmd=None):
        arguments = {"ZOrderCmd": ZOrderCmd}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.shape.ZOrder(*args, **arguments)

class ShapeNode:

    def __init__(self, shapenode=None):
        self.shapenode = shapenode

    @property
    def Application(self):
        return Application(self.shapenode.Application)

    @property
    def Creator(self):
        return self.shapenode.Creator

    @property
    def EditingType(self):
        return self.shapenode.EditingType

    @property
    def Parent(self):
        return self.shapenode.Parent

    @property
    def Points(self):
        return self.shapenode.Points

    @property
    def SegmentType(self):
        return self.shapenode.SegmentType

class SoftEdgeFormat:

    def __init__(self, softedgeformat=None):
        self.softedgeformat = softedgeformat

    @property
    def Application(self):
        return self.softedgeformat.Application

    @property
    def Creator(self):
        return self.softedgeformat.Creator

    @property
    def Parent(self):
        return self.softedgeformat.Parent

    @property
    def Radius(self):
        return self.softedgeformat.Radius

    @Radius.setter
    def Radius(self, value):
        self.softedgeformat.Radius = value

    @property
    def Type(self):
        return self.softedgeformat.Type

    @Type.setter
    def Type(self, value):
        self.softedgeformat.Type = value

class Source:

    def __init__(self, source=None):
        self.source = source

    @property
    def Application(self):
        return Application(self.source.Application)

    @property
    def Cited(self):
        return self.source.Cited

    @property
    def Creator(self):
        return self.source.Creator

    def Field(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.source.Field(*args, **arguments)

    @property
    def Parent(self):
        return self.source.Parent

    @property
    def Tag(self):
        return self.source.Tag

    @property
    def XML(self):
        return self.source.XML

    def Delete(self):
        self.source.Delete()

class Sources:

    def __init__(self, sources=None):
        self.sources = sources

    @property
    def Application(self):
        return Application(self.sources.Application)

    @property
    def Count(self):
        return Sources(self.sources.Count)

    @property
    def Creator(self):
        return self.sources.Creator

    @property
    def Parent(self):
        return self.sources.Parent

    def Add(self, *args, Data=None):
        arguments = {"Data": Data}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.sources.Add(*args, **arguments)

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.sources.Item(*args, **arguments)

class SpellingSuggestion:

    def __init__(self, spellingsuggestion=None):
        self.spellingsuggestion = spellingsuggestion

    @property
    def Application(self):
        return Application(self.spellingsuggestion.Application)

    @property
    def Creator(self):
        return self.spellingsuggestion.Creator

    @property
    def Name(self):
        return self.spellingsuggestion.Name

    @property
    def Parent(self):
        return self.spellingsuggestion.Parent

class Style:

    def __init__(self, style=None):
        self.style = style

    @property
    def Application(self):
        return Application(self.style.Application)

    @property
    def AutomaticallyUpdate(self):
        return self.style.AutomaticallyUpdate

    @property
    def BaseStyle(self):
        return self.style.BaseStyle

    @BaseStyle.setter
    def BaseStyle(self, value):
        self.style.BaseStyle = value

    @property
    def Borders(self):
        return self.style.Borders

    @property
    def BuiltIn(self):
        return self.style.BuiltIn

    @property
    def Creator(self):
        return self.style.Creator

    @property
    def Description(self):
        return self.style.Description

    @property
    def Font(self):
        return Font(self.style.Font)

    @Font.setter
    def Font(self, value):
        self.style.Font = value

    @property
    def Frame(self):
        return Frame(self.style.Frame)

    @property
    def Hidden(self):
        return self.style.Hidden

    @property
    def InUse(self):
        return self.style.InUse

    @property
    def LanguageID(self):
        return WdLanguageID(self.style.LanguageID)

    @LanguageID.setter
    def LanguageID(self, value):
        self.style.LanguageID = value

    @property
    def LanguageIDFarEast(self):
        return WdLanguageID(self.style.LanguageIDFarEast)

    @LanguageIDFarEast.setter
    def LanguageIDFarEast(self, value):
        self.style.LanguageIDFarEast = value

    @property
    def Linked(self):
        return self.style.Linked

    @property
    def LinkStyle(self):
        return self.style.LinkStyle

    @property
    def ListLevelNumber(self):
        return self.style.ListLevelNumber

    @property
    def ListTemplate(self):
        return ListTemplate(self.style.ListTemplate)

    @property
    def Locked(self):
        return self.style.Locked

    @property
    def NameLocal(self):
        return self.style.NameLocal

    @property
    def NextParagraphStyle(self):
        return self.style.NextParagraphStyle

    @NextParagraphStyle.setter
    def NextParagraphStyle(self, value):
        self.style.NextParagraphStyle = value

    @property
    def NoProofing(self):
        return self.style.NoProofing

    @property
    def NoSpaceBetweenParagraphsOfSameStyle(self):
        return self.style.NoSpaceBetweenParagraphsOfSameStyle

    @property
    def ParagraphFormat(self):
        return ParagraphFormat(self.style.ParagraphFormat)

    @ParagraphFormat.setter
    def ParagraphFormat(self, value):
        self.style.ParagraphFormat = value

    @property
    def Parent(self):
        return self.style.Parent

    @property
    def Priority(self):
        return self.style.Priority

    @Priority.setter
    def Priority(self, value):
        self.style.Priority = value

    @property
    def QuickStyle(self):
        return self.style.QuickStyle

    @QuickStyle.setter
    def QuickStyle(self, value):
        self.style.QuickStyle = value

    @property
    def Shading(self):
        return Shading(self.style.Shading)

    @property
    def Table(self):
        return TableStyle(self.style.Table)

    @property
    def Type(self):
        return WdStyleType(self.style.Type)

    @property
    def UnhideWhenUsed(self):
        return self.style.UnhideWhenUsed

    @property
    def Visibility(self):
        return self.style.Visibility

    def Delete(self):
        self.style.Delete()

    def LinkToListTemplate(self, *args, ListTemplate=None, ListLevelNumber=None):
        arguments = {"ListTemplate": ListTemplate, "ListLevelNumber": ListLevelNumber}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.style.LinkToListTemplate(*args, **arguments)

class StyleSheet:

    def __init__(self, stylesheet=None):
        self.stylesheet = stylesheet

    @property
    def Application(self):
        return Application(self.stylesheet.Application)

    @property
    def Creator(self):
        return self.stylesheet.Creator

    @property
    def FullName(self):
        return self.stylesheet.FullName

    @property
    def Index(self):
        return self.stylesheet.Index

    @property
    def Name(self):
        return self.stylesheet.Name

    @property
    def Parent(self):
        return self.stylesheet.Parent

    @property
    def Path(self):
        return self.stylesheet.Path

    @property
    def Title(self):
        return self.stylesheet.Title

    @Title.setter
    def Title(self, value):
        self.stylesheet.Title = value

    @property
    def Type(self):
        return WdStyleSheetLinkType(self.stylesheet.Type)

    @Type.setter
    def Type(self, value):
        self.stylesheet.Type = value

    def Delete(self):
        self.stylesheet.Delete()

    def Move(self, *args, Precedence=None):
        arguments = {"Precedence": Precedence}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.stylesheet.Move(*args, **arguments)

class StyleSheets:

    def __init__(self, stylesheets=None):
        self.stylesheets = stylesheets

    def __call__(self, item):
        return StyleSheet(self.stylesheets(item))

    @property
    def Application(self):
        return Application(self.stylesheets.Application)

    @property
    def Count(self):
        return self.stylesheets.Count

    @property
    def Creator(self):
        return self.stylesheets.Creator

    @property
    def Parent(self):
        return self.stylesheets.Parent

    def Add(self, *args, FileName=None, LinkType=None, Title=None, Precedence=None):
        arguments = {"FileName": FileName, "LinkType": LinkType, "Title": Title, "Precedence": Precedence}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return StyleSheet(self.stylesheets.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.stylesheets.Item(*args, **arguments)

class Subdocument:

    def __init__(self, subdocument=None):
        self.subdocument = subdocument

    @property
    def Application(self):
        return Application(self.subdocument.Application)

    @property
    def Creator(self):
        return self.subdocument.Creator

    @property
    def HasFile(self):
        return self.subdocument.HasFile

    @property
    def Level(self):
        return self.subdocument.Level

    @property
    def Locked(self):
        return self.subdocument.Locked

    @property
    def Name(self):
        return self.subdocument.Name

    @property
    def Parent(self):
        return self.subdocument.Parent

    @property
    def Path(self):
        return self.subdocument.Path

    @property
    def Range(self):
        return Range(self.subdocument.Range)

    def Delete(self):
        self.subdocument.Delete()

    def Open(self):
        return self.subdocument.Open()

    def Split(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.subdocument.Split(*args, **arguments)

class SynonymInfo:

    def __init__(self, synonyminfo=None):
        self.synonyminfo = synonyminfo

    @property
    def AntonymList(self):
        return self.synonyminfo.AntonymList

    @property
    def Application(self):
        return Application(self.synonyminfo.Application)

    @property
    def Creator(self):
        return self.synonyminfo.Creator

    @property
    def Found(self):
        return self.synonyminfo.Found

    @property
    def MeaningCount(self):
        return self.synonyminfo.MeaningCount

    @property
    def MeaningList(self):
        return self.synonyminfo.MeaningList

    @property
    def Parent(self):
        return self.synonyminfo.Parent

    @property
    def PartOfSpeechList(self):
        return self.synonyminfo.PartOfSpeechList

    @property
    def RelatedExpressionList(self):
        return self.synonyminfo.RelatedExpressionList

    @property
    def RelatedWordList(self):
        return self.synonyminfo.RelatedWordList

    def SynonymList(self, *args, Meaning=None):
        arguments = {"Meaning": Meaning}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.synonyminfo.SynonymList(*args, **arguments)

    @property
    def Word(self):
        return self.synonyminfo.Word

class System:

    def __init__(self, system=None):
        self.system = system

    @property
    def Application(self):
        return Application(self.system.Application)

    @property
    def CountryRegion(self):
        return WdCountry(self.system.CountryRegion)

    @property
    def Creator(self):
        return self.system.Creator

    @property
    def Cursor(self):
        return WdCursorType(self.system.Cursor)

    @Cursor.setter
    def Cursor(self, value):
        self.system.Cursor = value

    @property
    def FreeDiskSpace(self):
        return self.system.FreeDiskSpace

    @property
    def HorizontalResolution(self):
        return self.system.HorizontalResolution

    @property
    def LanguageDesignation(self):
        return self.system.LanguageDesignation

    @property
    def MathCoprocessorInstalled(self):
        return self.system.MathCoprocessorInstalled

    @property
    def OperatingSystem(self):
        return self.system.OperatingSystem

    @property
    def Parent(self):
        return self.system.Parent

    @property
    def PrivateProfileString(self):
        return self.system.PrivateProfileString

    @PrivateProfileString.setter
    def PrivateProfileString(self, value):
        self.system.PrivateProfileString = value

    @property
    def ProfileString(self):
        return self.system.ProfileString

    @ProfileString.setter
    def ProfileString(self, value):
        self.system.ProfileString = value

    @property
    def System(self):
        return self.system.System

    @property
    def System(self):
        return self.system.System

    @property
    def System(self):
        return self.system.System

    @property
    def Version(self):
        return self.system.Version

    @property
    def VerticalResolution(self):
        return self.system.VerticalResolution

    def Connect(self, *args, Path=None, Drive=None, Password=None):
        arguments = {"Path": Path, "Drive": Drive, "Password": Password}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.system.Connect(*args, **arguments)

    def MSInfo(self):
        self.system.MSInfo()

class Table:

    def __init__(self, table=None):
        self.table = table

    @property
    def AllowAutoFit(self):
        return self.table.AllowAutoFit

    @property
    def Application(self):
        return Application(self.table.Application)

    @property
    def ApplyStyleColumnBands(self):
        return self.table.ApplyStyleColumnBands

    @ApplyStyleColumnBands.setter
    def ApplyStyleColumnBands(self, value):
        self.table.ApplyStyleColumnBands = value

    @property
    def ApplyStyleFirstColumn(self):
        return self.table.ApplyStyleFirstColumn

    @property
    def ApplyStyleHeadingRows(self):
        return self.table.ApplyStyleHeadingRows

    @property
    def ApplyStyleLastColumn(self):
        return self.table.ApplyStyleLastColumn

    @property
    def ApplyStyleLastRow(self):
        return self.table.ApplyStyleLastRow

    @property
    def ApplyStyleRowBands(self):
        return self.table.ApplyStyleRowBands

    @ApplyStyleRowBands.setter
    def ApplyStyleRowBands(self, value):
        self.table.ApplyStyleRowBands = value

    @property
    def AutoFormatType(self):
        return self.table.AutoFormatType

    @property
    def Borders(self):
        return self.table.Borders

    @property
    def BottomPadding(self):
        return self.table.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.table.BottomPadding = value

    @property
    def Columns(self):
        return self.table.Columns

    @property
    def Creator(self):
        return self.table.Creator

    @property
    def Descr(self):
        return self.table.Descr

    @Descr.setter
    def Descr(self, value):
        self.table.Descr = value

    @property
    def ID(self):
        return self.table.ID

    @ID.setter
    def ID(self, value):
        self.table.ID = value

    @property
    def LeftPadding(self):
        return self.table.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.table.LeftPadding = value

    @property
    def NestingLevel(self):
        return self.table.NestingLevel

    @property
    def Parent(self):
        return self.table.Parent

    @property
    def PreferredWidth(self):
        return self.table.PreferredWidth

    @PreferredWidth.setter
    def PreferredWidth(self, value):
        self.table.PreferredWidth = value

    @property
    def PreferredWidthType(self):
        return WdPreferredWidthType(self.table.PreferredWidthType)

    @PreferredWidthType.setter
    def PreferredWidthType(self, value):
        self.table.PreferredWidthType = value

    @property
    def Range(self):
        return Range(self.table.Range)

    @property
    def RightPadding(self):
        return self.table.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.table.RightPadding = value

    @property
    def Rows(self):
        return self.table.Rows

    @property
    def Shading(self):
        return Shading(self.table.Shading)

    @property
    def Spacing(self):
        return self.table.Spacing

    @Spacing.setter
    def Spacing(self, value):
        self.table.Spacing = value

    @property
    def Style(self):
        return self.table.Style

    @Style.setter
    def Style(self, value):
        self.table.Style = value

    @property
    def TableDirection(self):
        return WdTableDirection(self.table.TableDirection)

    @TableDirection.setter
    def TableDirection(self, value):
        self.table.TableDirection = value

    @property
    def Tables(self):
        return self.table.Tables

    @property
    def Title(self):
        return self.table.Title

    @Title.setter
    def Title(self, value):
        self.table.Title = value

    @property
    def TopPadding(self):
        return self.table.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.table.TopPadding = value

    @property
    def Uniform(self):
        return self.table.Uniform

    def ApplyStyleDirectFormatting(self, *args, StyleName=None):
        arguments = {"StyleName": StyleName}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.table.ApplyStyleDirectFormatting(*args, **arguments)

    def AutoFitBehavior(self, *args, Behavior=None):
        arguments = {"Behavior": Behavior}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.table.AutoFitBehavior(*args, **arguments)

    def AutoFormat(self, *args, Format=None, ApplyBorders=None, ApplyShading=None, ApplyFont=None, ApplyColor=None, ApplyHeadingRows=None, ApplyLastRow=None, ApplyFirstColumn=None, ApplyLastColumn=None, AutoFit=None):
        arguments = {"Format": Format, "ApplyBorders": ApplyBorders, "ApplyShading": ApplyShading, "ApplyFont": ApplyFont, "ApplyColor": ApplyColor, "ApplyHeadingRows": ApplyHeadingRows, "ApplyLastRow": ApplyLastRow, "ApplyFirstColumn": ApplyFirstColumn, "ApplyLastColumn": ApplyLastColumn, "AutoFit": AutoFit}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.table.AutoFormat(*args, **arguments)

    def Cell(self, *args, Row=None, Column=None):
        arguments = {"Row": Row, "Column": Column}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.table.Cell(*args, **arguments)

    def ConvertToText(self, *args, Separator=None, NestedTables=None):
        arguments = {"Separator": Separator, "NestedTables": NestedTables}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.table.ConvertToText(*args, **arguments)

    def Delete(self):
        self.table.Delete()

    def Select(self):
        self.table.Select()

    def Sort(self, *args, ExcludeHeader=None, FieldNumber=None, SortFieldType=None, SortOrder=None, FieldNumber2=None, SortFieldType2=None, SortOrder2=None, FieldNumber3=None, SortFieldType3=None, SortOrder3=None, CaseSensitive=None, BidiSort=None, IgnoreThe=None, IgnoreKashida=None, IgnoreDiacritics=None, IgnoreHe=None, LanguageID=None):
        arguments = {"ExcludeHeader": ExcludeHeader, "FieldNumber": FieldNumber, "SortFieldType": SortFieldType, "SortOrder": SortOrder, "FieldNumber2": FieldNumber2, "SortFieldType2": SortFieldType2, "SortOrder2": SortOrder2, "FieldNumber3": FieldNumber3, "SortFieldType3": SortFieldType3, "SortOrder3": SortOrder3, "CaseSensitive": CaseSensitive, "BidiSort": BidiSort, "IgnoreThe": IgnoreThe, "IgnoreKashida": IgnoreKashida, "IgnoreDiacritics": IgnoreDiacritics, "IgnoreHe": IgnoreHe, "LanguageID": LanguageID}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.table.Sort(*args, **arguments)

    def SortAscending(self):
        self.table.SortAscending()

    def SortDescending(self):
        self.table.SortDescending()

    def Split(self, *args, BeforeRow=None):
        arguments = {"BeforeRow": BeforeRow}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.table.Split(*args, **arguments)

    def UpdateAutoFormat(self):
        self.table.UpdateAutoFormat()

class TableOfAuthorities:

    def __init__(self, tableofauthorities=None):
        self.tableofauthorities = tableofauthorities

    @property
    def Application(self):
        return Application(self.tableofauthorities.Application)

    @property
    def Bookmark(self):
        return self.tableofauthorities.Bookmark

    @Bookmark.setter
    def Bookmark(self, value):
        self.tableofauthorities.Bookmark = value

    @property
    def Category(self):
        return self.tableofauthorities.Category

    @Category.setter
    def Category(self, value):
        self.tableofauthorities.Category = value

    @property
    def Creator(self):
        return self.tableofauthorities.Creator

    @property
    def EntrySeparator(self):
        return self.tableofauthorities.EntrySeparator

    @EntrySeparator.setter
    def EntrySeparator(self, value):
        self.tableofauthorities.EntrySeparator = value

    @property
    def IncludeCategoryHeader(self):
        return self.tableofauthorities.IncludeCategoryHeader

    @property
    def IncludeSequenceName(self):
        return self.tableofauthorities.IncludeSequenceName

    @IncludeSequenceName.setter
    def IncludeSequenceName(self, value):
        self.tableofauthorities.IncludeSequenceName = value

    @property
    def KeepEntryFormatting(self):
        return self.tableofauthorities.KeepEntryFormatting

    @property
    def PageNumberSeparator(self):
        return self.tableofauthorities.PageNumberSeparator

    @PageNumberSeparator.setter
    def PageNumberSeparator(self, value):
        self.tableofauthorities.PageNumberSeparator = value

    @property
    def PageRangeSeparator(self):
        return self.tableofauthorities.PageRangeSeparator

    @PageRangeSeparator.setter
    def PageRangeSeparator(self, value):
        self.tableofauthorities.PageRangeSeparator = value

    @property
    def Parent(self):
        return self.tableofauthorities.Parent

    @property
    def Passim(self):
        return self.tableofauthorities.Passim

    @property
    def Range(self):
        return Range(self.tableofauthorities.Range)

    @property
    def Separator(self):
        return self.tableofauthorities.Separator

    @Separator.setter
    def Separator(self, value):
        self.tableofauthorities.Separator = value

    @property
    def TabLeader(self):
        return WdTabLeader(self.tableofauthorities.TabLeader)

    @TabLeader.setter
    def TabLeader(self, value):
        self.tableofauthorities.TabLeader = value

    def Delete(self):
        self.tableofauthorities.Delete()

    def Update(self):
        self.tableofauthorities.Update()

class TableOfAuthoritiesCategory:

    def __init__(self, tableofauthoritiescategory=None):
        self.tableofauthoritiescategory = tableofauthoritiescategory

    @property
    def Application(self):
        return Application(self.tableofauthoritiescategory.Application)

    @property
    def Creator(self):
        return self.tableofauthoritiescategory.Creator

    @property
    def Index(self):
        return self.tableofauthoritiescategory.Index

    @property
    def Name(self):
        return self.tableofauthoritiescategory.Name

    @property
    def Parent(self):
        return self.tableofauthoritiescategory.Parent

class TableOfContents:

    def __init__(self, tableofcontents=None):
        self.tableofcontents = tableofcontents

    @property
    def Application(self):
        return Application(self.tableofcontents.Application)

    @property
    def Creator(self):
        return self.tableofcontents.Creator

    @property
    def HeadingStyles(self):
        return self.tableofcontents.HeadingStyles

    @property
    def HidePageNumbersInWeb(self):
        return self.tableofcontents.HidePageNumbersInWeb

    @HidePageNumbersInWeb.setter
    def HidePageNumbersInWeb(self, value):
        self.tableofcontents.HidePageNumbersInWeb = value

    @property
    def IncludePageNumbers(self):
        return self.tableofcontents.IncludePageNumbers

    @property
    def LowerHeadingLevel(self):
        return self.tableofcontents.LowerHeadingLevel

    @LowerHeadingLevel.setter
    def LowerHeadingLevel(self, value):
        self.tableofcontents.LowerHeadingLevel = value

    @property
    def Parent(self):
        return self.tableofcontents.Parent

    @property
    def Range(self):
        return Range(self.tableofcontents.Range)

    @property
    def RightAlignPageNumbers(self):
        return self.tableofcontents.RightAlignPageNumbers

    @property
    def TabLeader(self):
        return WdTabLeader(self.tableofcontents.TabLeader)

    @TabLeader.setter
    def TabLeader(self, value):
        self.tableofcontents.TabLeader = value

    @property
    def TableID(self):
        return self.tableofcontents.TableID

    @TableID.setter
    def TableID(self, value):
        self.tableofcontents.TableID = value

    @property
    def UpperHeadingLevel(self):
        return self.tableofcontents.UpperHeadingLevel

    @UpperHeadingLevel.setter
    def UpperHeadingLevel(self, value):
        self.tableofcontents.UpperHeadingLevel = value

    @property
    def UseFields(self):
        return self.tableofcontents.UseFields

    @property
    def UseHeadingStyles(self):
        return self.tableofcontents.UseHeadingStyles

    @property
    def UseHyperlinks(self):
        return self.tableofcontents.UseHyperlinks

    @UseHyperlinks.setter
    def UseHyperlinks(self, value):
        self.tableofcontents.UseHyperlinks = value

    def Delete(self):
        self.tableofcontents.Delete()

    def Update(self):
        self.tableofcontents.Update()

    def UpdatePageNumbers(self):
        self.tableofcontents.UpdatePageNumbers()

class TableOfFigures:

    def __init__(self, tableoffigures=None):
        self.tableoffigures = tableoffigures

    @property
    def Application(self):
        return Application(self.tableoffigures.Application)

    @property
    def Caption(self):
        return self.tableoffigures.Caption

    @Caption.setter
    def Caption(self, value):
        self.tableoffigures.Caption = value

    @property
    def Creator(self):
        return self.tableoffigures.Creator

    @property
    def HeadingStyles(self):
        return self.tableoffigures.HeadingStyles

    @property
    def HidePageNumbersInWeb(self):
        return self.tableoffigures.HidePageNumbersInWeb

    @HidePageNumbersInWeb.setter
    def HidePageNumbersInWeb(self, value):
        self.tableoffigures.HidePageNumbersInWeb = value

    @property
    def IncludeLabel(self):
        return self.tableoffigures.IncludeLabel

    @property
    def IncludePageNumbers(self):
        return self.tableoffigures.IncludePageNumbers

    @property
    def LowerHeadingLevel(self):
        return self.tableoffigures.LowerHeadingLevel

    @LowerHeadingLevel.setter
    def LowerHeadingLevel(self, value):
        self.tableoffigures.LowerHeadingLevel = value

    @property
    def Parent(self):
        return self.tableoffigures.Parent

    @property
    def Range(self):
        return Range(self.tableoffigures.Range)

    @property
    def RightAlignPageNumbers(self):
        return self.tableoffigures.RightAlignPageNumbers

    @property
    def TabLeader(self):
        return WdTabLeader(self.tableoffigures.TabLeader)

    @TabLeader.setter
    def TabLeader(self, value):
        self.tableoffigures.TabLeader = value

    @property
    def TableID(self):
        return self.tableoffigures.TableID

    @TableID.setter
    def TableID(self, value):
        self.tableoffigures.TableID = value

    @property
    def UpperHeadingLevel(self):
        return self.tableoffigures.UpperHeadingLevel

    @UpperHeadingLevel.setter
    def UpperHeadingLevel(self, value):
        self.tableoffigures.UpperHeadingLevel = value

    @property
    def UseFields(self):
        return self.tableoffigures.UseFields

    @property
    def UseHeadingStyles(self):
        return self.tableoffigures.UseHeadingStyles

    @property
    def UseHyperlinks(self):
        return self.tableoffigures.UseHyperlinks

    @UseHyperlinks.setter
    def UseHyperlinks(self, value):
        self.tableoffigures.UseHyperlinks = value

    def Delete(self):
        self.tableoffigures.Delete()

    def Update(self):
        self.tableoffigures.Update()

    def UpdatePageNumbers(self):
        self.tableoffigures.UpdatePageNumbers()

class TableStyle:

    def __init__(self, tablestyle=None):
        self.tablestyle = tablestyle

    @property
    def Alignment(self):
        return WdRowAlignment(self.tablestyle.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.tablestyle.Alignment = value

    @property
    def AllowBreakAcrossPage(self):
        return self.tablestyle.AllowBreakAcrossPage

    @property
    def AllowPageBreaks(self):
        return self.tablestyle.AllowPageBreaks

    @property
    def Application(self):
        return Application(self.tablestyle.Application)

    @property
    def Borders(self):
        return self.tablestyle.Borders

    @property
    def BottomPadding(self):
        return self.tablestyle.BottomPadding

    @BottomPadding.setter
    def BottomPadding(self, value):
        self.tablestyle.BottomPadding = value

    @property
    def ColumnStripe(self):
        return self.tablestyle.ColumnStripe

    @ColumnStripe.setter
    def ColumnStripe(self, value):
        self.tablestyle.ColumnStripe = value

    @property
    def Creator(self):
        return self.tablestyle.Creator

    @property
    def LeftIndent(self):
        return self.tablestyle.LeftIndent

    @LeftIndent.setter
    def LeftIndent(self, value):
        self.tablestyle.LeftIndent = value

    @property
    def LeftPadding(self):
        return self.tablestyle.LeftPadding

    @LeftPadding.setter
    def LeftPadding(self, value):
        self.tablestyle.LeftPadding = value

    @property
    def Parent(self):
        return self.tablestyle.Parent

    @property
    def RightPadding(self):
        return self.tablestyle.RightPadding

    @RightPadding.setter
    def RightPadding(self, value):
        self.tablestyle.RightPadding = value

    @property
    def RowStripe(self):
        return self.tablestyle.RowStripe

    @RowStripe.setter
    def RowStripe(self, value):
        self.tablestyle.RowStripe = value

    @property
    def Shading(self):
        return Shading(self.tablestyle.Shading)

    @property
    def Spacing(self):
        return self.tablestyle.Spacing

    @Spacing.setter
    def Spacing(self, value):
        self.tablestyle.Spacing = value

    @property
    def TableDirection(self):
        return WdTableDirection(self.tablestyle.TableDirection)

    @TableDirection.setter
    def TableDirection(self, value):
        self.tablestyle.TableDirection = value

    @property
    def TopPadding(self):
        return self.tablestyle.TopPadding

    @TopPadding.setter
    def TopPadding(self, value):
        self.tablestyle.TopPadding = value

    def Condition(self, *args, ConditionCode=None):
        arguments = {"ConditionCode": ConditionCode}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.tablestyle.Condition(*args, **arguments)

class TabStop:

    def __init__(self, tabstop=None):
        self.tabstop = tabstop

    @property
    def Alignment(self):
        return WdTabAlignment(self.tabstop.Alignment)

    @Alignment.setter
    def Alignment(self, value):
        self.tabstop.Alignment = value

    @property
    def Application(self):
        return Application(self.tabstop.Application)

    @property
    def Creator(self):
        return self.tabstop.Creator

    @property
    def CustomTab(self):
        return self.tabstop.CustomTab

    @property
    def Leader(self):
        return TabStop(self.tabstop.Leader)

    @Leader.setter
    def Leader(self, value):
        self.tabstop.Leader = value

    @property
    def Next(self):
        return self.tabstop.Next

    @property
    def Parent(self):
        return self.tabstop.Parent

    @property
    def Position(self):
        return self.tabstop.Position

    @Position.setter
    def Position(self, value):
        self.tabstop.Position = value

    @property
    def Previous(self):
        return self.tabstop.Previous

    def Clear(self):
        self.tabstop.Clear()

class Task:

    def __init__(self, task=None):
        self.task = task

    @property
    def Application(self):
        return Application(self.task.Application)

    @property
    def Creator(self):
        return self.task.Creator

    @property
    def Height(self):
        return self.task.Height

    @Height.setter
    def Height(self, value):
        self.task.Height = value

    @property
    def Left(self):
        return self.task.Left

    @Left.setter
    def Left(self, value):
        self.task.Left = value

    @property
    def Name(self):
        return self.task.Name

    @property
    def Parent(self):
        return self.task.Parent

    @property
    def Top(self):
        return self.task.Top

    @Top.setter
    def Top(self, value):
        self.task.Top = value

    @property
    def Visible(self):
        return self.task.Visible

    @property
    def Width(self):
        return Task(self.task.Width)

    @Width.setter
    def Width(self, value):
        self.task.Width = value

    @property
    def WindowState(self):
        return WdWindowState(self.task.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.task.WindowState = value

    def Activate(self, *args, Wait=None):
        arguments = {"Wait": Wait}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.task.Activate(*args, **arguments)

    def Close(self):
        self.task.Close()

    def Move(self, *args, Left=None, Top=None):
        arguments = {"Left": Left, "Top": Top}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.task.Move(*args, **arguments)

    def Resize(self, *args, Width=None, Height=None):
        arguments = {"Width": Width, "Height": Height}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.task.Resize(*args, **arguments)

    def SendWindowMessage(self, *args, Message=None, wParam=None, IParam=None):
        arguments = {"Message": Message, "wParam": wParam, "IParam": IParam}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.task.SendWindowMessage(*args, **arguments)

class TaskPane:

    def __init__(self, taskpane=None):
        self.taskpane = taskpane

    @property
    def Application(self):
        return Application(self.taskpane.Application)

    @property
    def Creator(self):
        return self.taskpane.Creator

    @property
    def Parent(self):
        return self.taskpane.Parent

    @property
    def Visible(self):
        return self.taskpane.Visible

class TaskPanes:

    def __init__(self, taskpanes=None):
        self.taskpanes = taskpanes

    def __call__(self, item):
        return TaskPane(self.taskpanes(item))

    @property
    def Application(self):
        return Application(self.taskpanes.Application)

    @property
    def Count(self):
        return self.taskpanes.Count

    @property
    def Creator(self):
        return self.taskpanes.Creator

    @property
    def Parent(self):
        return self.taskpanes.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.taskpanes.Item(*args, **arguments)

class Template:

    def __init__(self, template=None):
        self.template = template

    @property
    def Application(self):
        return Application(self.template.Application)

    @property
    def BuildingBlockEntries(self):
        return BuildingBlockEntries(self.template.BuildingBlockEntries)

    @property
    def BuildingBlockTypes(self):
        return BuildingBlockTypes(self.template.BuildingBlockTypes)

    @property
    def BuiltInDocumentProperties(self):
        return self.template.BuiltInDocumentProperties

    @property
    def Creator(self):
        return self.template.Creator

    @property
    def CustomDocumentProperties(self):
        return self.template.CustomDocumentProperties

    @property
    def FarEastLineBreakLanguage(self):
        return WdFarEastLineBreakLanguageID(self.template.FarEastLineBreakLanguage)

    @FarEastLineBreakLanguage.setter
    def FarEastLineBreakLanguage(self, value):
        self.template.FarEastLineBreakLanguage = value

    @property
    def FarEastLineBreakLevel(self):
        return WdFarEastLineBreakLevel(self.template.FarEastLineBreakLevel)

    @FarEastLineBreakLevel.setter
    def FarEastLineBreakLevel(self, value):
        self.template.FarEastLineBreakLevel = value

    @property
    def FullName(self):
        return self.template.FullName

    @property
    def JustificationMode(self):
        return WdJustificationMode(self.template.JustificationMode)

    @JustificationMode.setter
    def JustificationMode(self, value):
        self.template.JustificationMode = value

    @property
    def KerningByAlgorithm(self):
        return self.template.KerningByAlgorithm

    @property
    def LanguageID(self):
        return WdLanguageID(self.template.LanguageID)

    @LanguageID.setter
    def LanguageID(self, value):
        self.template.LanguageID = value

    @property
    def LanguageIDFarEast(self):
        return WdLanguageID(self.template.LanguageIDFarEast)

    @LanguageIDFarEast.setter
    def LanguageIDFarEast(self, value):
        self.template.LanguageIDFarEast = value

    @property
    def ListTemplates(self):
        return self.template.ListTemplates

    @property
    def Name(self):
        return self.template.Name

    @property
    def NoLineBreakAfter(self):
        return self.template.NoLineBreakAfter

    @NoLineBreakAfter.setter
    def NoLineBreakAfter(self, value):
        self.template.NoLineBreakAfter = value

    @property
    def NoLineBreakBefore(self):
        return self.template.NoLineBreakBefore

    @NoLineBreakBefore.setter
    def NoLineBreakBefore(self, value):
        self.template.NoLineBreakBefore = value

    @property
    def NoProofing(self):
        return self.template.NoProofing

    @property
    def Parent(self):
        return self.template.Parent

    @property
    def Path(self):
        return self.template.Path

    @property
    def Saved(self):
        return self.template.Saved

    @property
    def Type(self):
        return WdTemplateType(self.template.Type)

    @property
    def VBProject(self):
        return self.template.VBProject

    def OpenAsDocument(self):
        return self.template.OpenAsDocument()

    def Save(self):
        self.template.Save()

class TextColumn:

    def __init__(self, textcolumn=None):
        self.textcolumn = textcolumn

    @property
    def Application(self):
        return Application(self.textcolumn.Application)

    @property
    def Creator(self):
        return self.textcolumn.Creator

    @property
    def Parent(self):
        return self.textcolumn.Parent

    @property
    def SpaceAfter(self):
        return self.textcolumn.SpaceAfter

    @SpaceAfter.setter
    def SpaceAfter(self, value):
        self.textcolumn.SpaceAfter = value

    @property
    def Width(self):
        return self.textcolumn.Width

    @Width.setter
    def Width(self, value):
        self.textcolumn.Width = value

class TextEffectFormat:

    def __init__(self, texteffectformat=None):
        self.texteffectformat = texteffectformat

    @property
    def Alignment(self):
        return self.texteffectformat.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.texteffectformat.Alignment = value

    @property
    def Application(self):
        return Application(self.texteffectformat.Application)

    @property
    def Creator(self):
        return self.texteffectformat.Creator

    @property
    def FontBold(self):
        return self.texteffectformat.FontBold

    @property
    def FontItalic(self):
        return self.texteffectformat.FontItalic

    @property
    def FontName(self):
        return self.texteffectformat.FontName

    @FontName.setter
    def FontName(self, value):
        self.texteffectformat.FontName = value

    @property
    def FontSize(self):
        return self.texteffectformat.FontSize

    @FontSize.setter
    def FontSize(self, value):
        self.texteffectformat.FontSize = value

    @property
    def KernedPairs(self):
        return self.texteffectformat.KernedPairs

    @property
    def NormalizedHeight(self):
        return self.texteffectformat.NormalizedHeight

    @property
    def Parent(self):
        return self.texteffectformat.Parent

    @property
    def PresetShape(self):
        return self.texteffectformat.PresetShape

    @PresetShape.setter
    def PresetShape(self, value):
        self.texteffectformat.PresetShape = value

    @property
    def PresetTextEffect(self):
        return self.texteffectformat.PresetTextEffect

    @PresetTextEffect.setter
    def PresetTextEffect(self, value):
        self.texteffectformat.PresetTextEffect = value

    @property
    def RotatedChars(self):
        return self.texteffectformat.RotatedChars

    @property
    def Text(self):
        return self.texteffectformat.Text

    @Text.setter
    def Text(self, value):
        self.texteffectformat.Text = value

    @property
    def Tracking(self):
        return self.texteffectformat.Tracking

    @Tracking.setter
    def Tracking(self, value):
        self.texteffectformat.Tracking = value

    def ToggleVerticalText(self):
        self.texteffectformat.ToggleVerticalText()

class TextFrame:

    def __init__(self, textframe=None):
        self.textframe = textframe

    @property
    def Application(self):
        return Application(self.textframe.Application)

    @property
    def AutoSize(self):
        return self.textframe.AutoSize

    @AutoSize.setter
    def AutoSize(self, value):
        self.textframe.AutoSize = value

    @property
    def Column(self):
        return self.textframe.Column

    @property
    def ContainingRange(self):
        return Range(self.textframe.ContainingRange)

    @property
    def Creator(self):
        return self.textframe.Creator

    @property
    def HasText(self):
        return self.textframe.HasText

    @property
    def HorizontalAnchor(self):
        return self.textframe.HorizontalAnchor

    @HorizontalAnchor.setter
    def HorizontalAnchor(self, value):
        self.textframe.HorizontalAnchor = value

    @property
    def MarginBottom(self):
        return self.textframe.MarginBottom

    @MarginBottom.setter
    def MarginBottom(self, value):
        self.textframe.MarginBottom = value

    @property
    def MarginLeft(self):
        return self.textframe.MarginLeft

    @MarginLeft.setter
    def MarginLeft(self, value):
        self.textframe.MarginLeft = value

    @property
    def MarginRight(self):
        return self.textframe.MarginRight

    @MarginRight.setter
    def MarginRight(self, value):
        self.textframe.MarginRight = value

    @property
    def MarginTop(self):
        return self.textframe.MarginTop

    @MarginTop.setter
    def MarginTop(self, value):
        self.textframe.MarginTop = value

    @property
    def Next(self):
        return TextFrame(self.textframe.Next)

    @property
    def NoTextRotation(self):
        return self.textframe.NoTextRotation

    @property
    def Orientation(self):
        return self.textframe.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.textframe.Orientation = value

    @property
    def Overflowing(self):
        return self.textframe.Overflowing

    @property
    def Parent(self):
        return Shape(self.textframe.Parent)

    @property
    def PathFormat(self):
        return self.textframe.PathFormat

    @PathFormat.setter
    def PathFormat(self, value):
        self.textframe.PathFormat = value

    @property
    def Previous(self):
        return TextFrame(self.textframe.Previous)

    @property
    def TextRange(self):
        return Range(self.textframe.TextRange)

    @property
    def ThreeD(self):
        return self.textframe.ThreeD

    @property
    def VerticalAnchor(self):
        return self.textframe.VerticalAnchor

    @VerticalAnchor.setter
    def VerticalAnchor(self, value):
        self.textframe.VerticalAnchor = value

    @property
    def WarpFormat(self):
        return self.textframe.WarpFormat

    @WarpFormat.setter
    def WarpFormat(self, value):
        self.textframe.WarpFormat = value

    @property
    def WordWrap(self):
        return self.textframe.WordWrap

    def BreakForwardLink(self):
        self.textframe.BreakForwardLink()

    def DeleteText(self):
        self.textframe.DeleteText()

    def ValidLinkTarget(self, *args, TargetTextFrame=None):
        arguments = {"TargetTextFrame": TargetTextFrame}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.textframe.ValidLinkTarget(*args, **arguments)

class TextInput:

    def __init__(self, textinput=None):
        self.textinput = textinput

    @property
    def Application(self):
        return Application(self.textinput.Application)

    @property
    def Creator(self):
        return self.textinput.Creator

    @property
    def Default(self):
        return self.textinput.Default

    @Default.setter
    def Default(self, value):
        self.textinput.Default = value

    @property
    def Format(self):
        return self.textinput.Format

    @property
    def Parent(self):
        return self.textinput.Parent

    @property
    def Type(self):
        return WdTextFormFieldType(self.textinput.Type)

    @property
    def Valid(self):
        return self.textinput.Valid

    @property
    def Width(self):
        return self.textinput.Width

    @Width.setter
    def Width(self, value):
        self.textinput.Width = value

    def Clear(self):
        self.textinput.Clear()

    def EditType(self, *args, Type=None, Default=None, Format=None, Enabled=None):
        arguments = {"Type": Type, "Default": Default, "Format": Format, "Enabled": Enabled}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.textinput.EditType(*args, **arguments)

class TextRetrievalMode:

    def __init__(self, textretrievalmode=None):
        self.textretrievalmode = textretrievalmode

    @property
    def Application(self):
        return Application(self.textretrievalmode.Application)

    @property
    def Creator(self):
        return self.textretrievalmode.Creator

    @property
    def Duplicate(self):
        return TextRetrievalMode(self.textretrievalmode.Duplicate)

    @property
    def IncludeFieldCodes(self):
        return self.textretrievalmode.IncludeFieldCodes

    @property
    def IncludeHiddenText(self):
        return self.textretrievalmode.IncludeHiddenText

    @property
    def Parent(self):
        return self.textretrievalmode.Parent

    @property
    def ViewType(self):
        return TextRetrievalMode(self.textretrievalmode.ViewType)

    @ViewType.setter
    def ViewType(self, value):
        self.textretrievalmode.ViewType = value

class ThreeDFormat:

    def __init__(self, threedformat=None):
        self.threedformat = threedformat

    @property
    def Application(self):
        return Application(self.threedformat.Application)

    @property
    def BevelBottomDepth(self):
        return self.threedformat.BevelBottomDepth

    @BevelBottomDepth.setter
    def BevelBottomDepth(self, value):
        self.threedformat.BevelBottomDepth = value

    @property
    def BevelBottomInset(self):
        return self.threedformat.BevelBottomInset

    @BevelBottomInset.setter
    def BevelBottomInset(self, value):
        self.threedformat.BevelBottomInset = value

    @property
    def BevelBottomType(self):
        return self.threedformat.BevelBottomType

    @BevelBottomType.setter
    def BevelBottomType(self, value):
        self.threedformat.BevelBottomType = value

    @property
    def BevelTopDepth(self):
        return self.threedformat.BevelTopDepth

    @BevelTopDepth.setter
    def BevelTopDepth(self, value):
        self.threedformat.BevelTopDepth = value

    @property
    def BevelTopInset(self):
        return self.threedformat.BevelTopInset

    @BevelTopInset.setter
    def BevelTopInset(self, value):
        self.threedformat.BevelTopInset = value

    @property
    def BevelTopType(self):
        return self.threedformat.BevelTopType

    @BevelTopType.setter
    def BevelTopType(self, value):
        self.threedformat.BevelTopType = value

    @property
    def ContourColor(self):
        return ColorFormat(self.threedformat.ContourColor)

    @ContourColor.setter
    def ContourColor(self, value):
        self.threedformat.ContourColor = value

    @property
    def ContourWidth(self):
        return self.threedformat.ContourWidth

    @ContourWidth.setter
    def ContourWidth(self, value):
        self.threedformat.ContourWidth = value

    @property
    def Creator(self):
        return self.threedformat.Creator

    @property
    def Depth(self):
        return self.threedformat.Depth

    @Depth.setter
    def Depth(self, value):
        self.threedformat.Depth = value

    @property
    def ExtrusionColor(self):
        return ColorFormat(self.threedformat.ExtrusionColor)

    @property
    def ExtrusionColorType(self):
        return self.threedformat.ExtrusionColorType

    @ExtrusionColorType.setter
    def ExtrusionColorType(self, value):
        self.threedformat.ExtrusionColorType = value

    @property
    def FieldOfView(self):
        return self.threedformat.FieldOfView

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.threedformat.FieldOfView = value

    @property
    def LightAngle(self):
        return self.threedformat.LightAngle

    @LightAngle.setter
    def LightAngle(self, value):
        self.threedformat.LightAngle = value

    @property
    def Parent(self):
        return self.threedformat.Parent

    @property
    def Perspective(self):
        return self.threedformat.Perspective

    @property
    def PresetCamera(self):
        return self.threedformat.PresetCamera

    @property
    def PresetExtrusionDirection(self):
        return self.threedformat.PresetExtrusionDirection

    @property
    def PresetLighting(self):
        return self.threedformat.PresetLighting

    @PresetLighting.setter
    def PresetLighting(self, value):
        self.threedformat.PresetLighting = value

    @property
    def PresetLightingDirection(self):
        return self.threedformat.PresetLightingDirection

    @PresetLightingDirection.setter
    def PresetLightingDirection(self, value):
        self.threedformat.PresetLightingDirection = value

    @property
    def PresetLightingSoftness(self):
        return self.threedformat.PresetLightingSoftness

    @PresetLightingSoftness.setter
    def PresetLightingSoftness(self, value):
        self.threedformat.PresetLightingSoftness = value

    @property
    def PresetMaterial(self):
        return self.threedformat.PresetMaterial

    @PresetMaterial.setter
    def PresetMaterial(self, value):
        self.threedformat.PresetMaterial = value

    @property
    def PresetThreeDFormat(self):
        return self.threedformat.PresetThreeDFormat

    @property
    def ProjectText(self):
        return self.threedformat.ProjectText

    @ProjectText.setter
    def ProjectText(self, value):
        self.threedformat.ProjectText = value

    @property
    def RotationX(self):
        return self.threedformat.RotationX

    @RotationX.setter
    def RotationX(self, value):
        self.threedformat.RotationX = value

    @property
    def RotationY(self):
        return self.threedformat.RotationY

    @RotationY.setter
    def RotationY(self, value):
        self.threedformat.RotationY = value

    @property
    def RotationZ(self):
        return self.threedformat.RotationZ

    @RotationZ.setter
    def RotationZ(self, value):
        self.threedformat.RotationZ = value

    @property
    def Visible(self):
        return self.threedformat.Visible

    @property
    def Z(self):
        return self.threedformat.Z

    @Z.setter
    def Z(self, value):
        self.threedformat.Z = value

    def IncrementRotationHorizontal(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationHorizontal(*args, **arguments)

    def IncrementRotationVertical(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationVertical(*args, **arguments)

    def IncrementRotationX(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationX(*args, **arguments)

    def IncrementRotationY(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationY(*args, **arguments)

    def IncrementRotationZ(self, *args, Increment=None):
        arguments = {"Increment": Increment}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.IncrementRotationZ(*args, **arguments)

    def ResetRotation(self):
        self.threedformat.ResetRotation()

    def SetExtrusionDirection(self, *args, PresetExtrusionDirection=None):
        arguments = {"PresetExtrusionDirection": PresetExtrusionDirection}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.SetExtrusionDirection(*args, **arguments)

    def SetPresetCamera(self, *args, PresetCamera=None):
        arguments = {"PresetCamera": PresetCamera}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.SetPresetCamera(*args, **arguments)

    def SetThreeDFormat(self, *args, PresetThreeDFormat=None):
        arguments = {"PresetThreeDFormat": PresetThreeDFormat}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.threedformat.SetThreeDFormat(*args, **arguments)

class TickLabels:

    def __init__(self, ticklabels=None):
        self.ticklabels = ticklabels

    @property
    def Alignment(self):
        return self.ticklabels.Alignment

    @Alignment.setter
    def Alignment(self, value):
        self.ticklabels.Alignment = value

    @property
    def Application(self):
        return self.ticklabels.Application

    @property
    def Creator(self):
        return self.ticklabels.Creator

    @property
    def Depth(self):
        return self.ticklabels.Depth

    @property
    def Font(self):
        return ChartFont(self.ticklabels.Font)

    @property
    def Format(self):
        return ChartFormat(self.ticklabels.Format)

    @property
    def MultiLevel(self):
        return self.ticklabels.MultiLevel

    @MultiLevel.setter
    def MultiLevel(self, value):
        self.ticklabels.MultiLevel = value

    @property
    def Name(self):
        return self.ticklabels.Name

    @property
    def NumberFormat(self):
        return self.ticklabels.NumberFormat

    @NumberFormat.setter
    def NumberFormat(self, value):
        self.ticklabels.NumberFormat = value

    @property
    def NumberFormatLinked(self):
        return self.ticklabels.NumberFormatLinked

    @property
    def NumberFormatLocal(self):
        return self.ticklabels.NumberFormatLocal

    @NumberFormatLocal.setter
    def NumberFormatLocal(self, value):
        self.ticklabels.NumberFormatLocal = value

    @property
    def Offset(self):
        return self.ticklabels.Offset

    @Offset.setter
    def Offset(self, value):
        self.ticklabels.Offset = value

    @property
    def Orientation(self):
        return self.ticklabels.Orientation

    @Orientation.setter
    def Orientation(self, value):
        self.ticklabels.Orientation = value

    @property
    def Parent(self):
        return self.ticklabels.Parent

    @property
    def ReadingOrder(self):
        return XlReadingOrder(self.ticklabels.ReadingOrder)

    @ReadingOrder.setter
    def ReadingOrder(self, value):
        self.ticklabels.ReadingOrder = value

    def Delete(self):
        self.ticklabels.Delete()

    def Select(self):
        self.ticklabels.Select()

class Trendline:

    def __init__(self, trendline=None):
        self.trendline = trendline

    @property
    def Application(self):
        return self.trendline.Application

    @property
    def Backward2(self):
        return self.trendline.Backward2

    @Backward2.setter
    def Backward2(self, value):
        self.trendline.Backward2 = value

    @property
    def Border(self):
        return ChartBorder(self.trendline.Border)

    @property
    def Creator(self):
        return self.trendline.Creator

    @property
    def DataLabel(self):
        return DataLabel(self.trendline.DataLabel)

    @property
    def DisplayEquation(self):
        return self.trendline.DisplayEquation

    @property
    def DisplayRSquared(self):
        return self.trendline.DisplayRSquared

    @property
    def Format(self):
        return ChartFormat(self.trendline.Format)

    @property
    def Forward2(self):
        return self.trendline.Forward2

    @Forward2.setter
    def Forward2(self, value):
        self.trendline.Forward2 = value

    @property
    def Index(self):
        return self.trendline.Index

    @property
    def Intercept(self):
        return self.trendline.Intercept

    @Intercept.setter
    def Intercept(self, value):
        self.trendline.Intercept = value

    @property
    def InterceptIsAuto(self):
        return self.trendline.InterceptIsAuto

    @property
    def Name(self):
        return self.trendline.Name

    @Name.setter
    def Name(self, value):
        self.trendline.Name = value

    @property
    def NameIsAuto(self):
        return self.trendline.NameIsAuto

    @property
    def Order(self):
        return self.trendline.Order

    @Order.setter
    def Order(self, value):
        self.trendline.Order = value

    @property
    def Parent(self):
        return self.trendline.Parent

    @property
    def Period(self):
        return self.trendline.Period

    @Period.setter
    def Period(self, value):
        self.trendline.Period = value

    @property
    def Type(self):
        return self.trendline.Type

    @Type.setter
    def Type(self, value):
        self.trendline.Type = value

    def ClearFormats(self):
        self.trendline.ClearFormats()

    def Delete(self):
        self.trendline.Delete()

    def Select(self):
        self.trendline.Select()

class Trendlines:

    def __init__(self, trendlines=None):
        self.trendlines = trendlines

    @property
    def Application(self):
        return self.trendlines.Application

    @property
    def Count(self):
        return self.trendlines.Count

    @property
    def Creator(self):
        return self.trendlines.Creator

    @property
    def Parent(self):
        return self.trendlines.Parent

    def Add(self, *args, Type=None, Order=None, Period=None, Forward=None, Backward=None, Intercept=None, DisplayEquation=None, DisplayRSquared=None, Name=None):
        arguments = {"Type": Type, "Order": Order, "Period": Period, "Forward": Forward, "Backward": Backward, "Intercept": Intercept, "DisplayEquation": DisplayEquation, "DisplayRSquared": DisplayRSquared, "Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Trendline(self.trendlines.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return Trendline(self.trendlines.Item(*args, **arguments))

class TwoInitialCapsException:

    def __init__(self, twoinitialcapsexception=None):
        self.twoinitialcapsexception = twoinitialcapsexception

    @property
    def Application(self):
        return Application(self.twoinitialcapsexception.Application)

    @property
    def Creator(self):
        return self.twoinitialcapsexception.Creator

    @property
    def Index(self):
        return self.twoinitialcapsexception.Index

    @property
    def Name(self):
        return self.twoinitialcapsexception.Name

    @property
    def Parent(self):
        return self.twoinitialcapsexception.Parent

    def Delete(self):
        self.twoinitialcapsexception.Delete()

class UndoRecord:

    def __init__(self, undorecord=None):
        self.undorecord = undorecord

    @property
    def Application(self):
        return self.undorecord.Application

    @property
    def Creator(self):
        return self.undorecord.Creator

    @property
    def CustomRecordLevel(self):
        return self.undorecord.CustomRecordLevel

    @property
    def CustomRecordName(self):
        return self.undorecord.CustomRecordName

    @property
    def IsRecordingCustomRecord(self):
        return self.undorecord.IsRecordingCustomRecord

    @property
    def Parent(self):
        return self.undorecord.Parent

    def EndCustomRecord(self):
        self.undorecord.EndCustomRecord()

    def StartCustomRecord(self, *args, Name=None):
        arguments = {"Name": Name}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.undorecord.StartCustomRecord(*args, **arguments)

class UpBars:

    def __init__(self, upbars=None):
        self.upbars = upbars

    @property
    def Application(self):
        return self.upbars.Application

    @property
    def Border(self):
        return ChartBorder(self.upbars.Border)

    @property
    def Creator(self):
        return self.upbars.Creator

    @property
    def Fill(self):
        return FillFormat(self.upbars.Fill)

    @property
    def Format(self):
        return ChartFormat(self.upbars.Format)

    @property
    def Interior(self):
        return Interior(self.upbars.Interior)

    @property
    def Name(self):
        return self.upbars.Name

    @property
    def Parent(self):
        return self.upbars.Parent

    def Delete(self):
        self.upbars.Delete()

    def Select(self):
        self.upbars.Select()

class Variable:

    def __init__(self, variable=None):
        self.variable = variable

    @property
    def Application(self):
        return Application(self.variable.Application)

    @property
    def Creator(self):
        return self.variable.Creator

    @property
    def Index(self):
        return self.variable.Index

    @property
    def Name(self):
        return self.variable.Name

    @property
    def Parent(self):
        return self.variable.Parent

    @property
    def Value(self):
        return self.variable.Value

    @Value.setter
    def Value(self, value):
        self.variable.Value = value

    def Delete(self):
        self.variable.Delete()

class Version:

    def __init__(self, version=None):
        self.version = version

    @property
    def Application(self):
        return Application(self.version.Application)

    @property
    def Comment(self):
        return self.version.Comment

    @property
    def Creator(self):
        return self.version.Creator

    @property
    def Date(self):
        return self.version.Date

    @property
    def Index(self):
        return self.version.Index

    @property
    def Parent(self):
        return self.version.Parent

    @property
    def SavedBy(self):
        return self.version.SavedBy

    def Delete(self):
        self.version.Delete()

    def Open(self):
        return self.version.Open()

class View:

    def __init__(self, view=None):
        self.view = view

    @property
    def Application(self):
        return Application(self.view.Application)

    @property
    def ConflictMode(self):
        return self.view.ConflictMode

    @property
    def Creator(self):
        return self.view.Creator

    @property
    def DisplayBackgrounds(self):
        return self.view.DisplayBackgrounds

    @DisplayBackgrounds.setter
    def DisplayBackgrounds(self, value):
        self.view.DisplayBackgrounds = value

    @property
    def DisplayPageBoundaries(self):
        return self.view.DisplayPageBoundaries

    @property
    def Draft(self):
        return self.view.Draft

    @property
    def FieldShading(self):
        return WdFieldShading(self.view.FieldShading)

    @FieldShading.setter
    def FieldShading(self, value):
        self.view.FieldShading = value

    @property
    def FullScreen(self):
        return self.view.FullScreen

    @property
    def Magnifier(self):
        return self.view.Magnifier

    @property
    def MailMergeDataView(self):
        return self.view.MailMergeDataView

    @property
    def MarkupMode(self):
        return WdRevisionsMode(self.view.MarkupMode)

    @MarkupMode.setter
    def MarkupMode(self, value):
        self.view.MarkupMode = value

    @property
    def PageMovementType(self):
        return WdPageMovementType(self.view.PageMovementType)

    @PageMovementType.setter
    def PageMovementType(self, value):
        self.view.PageMovementType = value

    @property
    def Panning(self):
        return self.view.Panning

    @Panning.setter
    def Panning(self, value):
        self.view.Panning = value

    @property
    def Parent(self):
        return self.view.Parent

    @property
    def ReadingLayout(self):
        return self.view.ReadingLayout

    @property
    def ReadingLayoutActualView(self):
        return self.view.ReadingLayoutActualView

    @property
    def ReadingLayoutTruncateMargins(self):
        return WdReadingLayoutMargin(self.view.ReadingLayoutTruncateMargins)

    @ReadingLayoutTruncateMargins.setter
    def ReadingLayoutTruncateMargins(self, value):
        self.view.ReadingLayoutTruncateMargins = value

    @property
    def RevisionsBalloonShowConnectingLines(self):
        return self.view.RevisionsBalloonShowConnectingLines

    @property
    def RevisionsBalloonSide(self):
        return self.view.RevisionsBalloonSide

    @property
    def RevisionsBalloonWidth(self):
        return self.view.RevisionsBalloonWidth

    @property
    def RevisionsBalloonWidthType(self):
        return self.view.RevisionsBalloonWidthType

    @property
    def SeekView(self):
        return WdSeekView(self.view.SeekView)

    @SeekView.setter
    def SeekView(self, value):
        self.view.SeekView = value

    @property
    def ShadeEditableRanges(self):
        return self.view.ShadeEditableRanges

    @ShadeEditableRanges.setter
    def ShadeEditableRanges(self, value):
        self.view.ShadeEditableRanges = value

    @property
    def ShowAll(self):
        return self.view.ShowAll

    @property
    def ShowBookmarks(self):
        return self.view.ShowBookmarks

    @property
    def ShowComments(self):
        return self.view.ShowComments

    @property
    def ShowCropMarks(self):
        return self.view.ShowCropMarks

    @ShowCropMarks.setter
    def ShowCropMarks(self, value):
        self.view.ShowCropMarks = value

    @property
    def ShowDrawings(self):
        return self.view.ShowDrawings

    @property
    def ShowFieldCodes(self):
        return self.view.ShowFieldCodes

    @property
    def ShowFirstLineOnly(self):
        return self.view.ShowFirstLineOnly

    @property
    def ShowFormat(self):
        return self.view.ShowFormat

    @property
    def ShowFormatChanges(self):
        return self.view.ShowFormatChanges

    @property
    def ShowHiddenText(self):
        return self.view.ShowHiddenText

    @property
    def ShowHighlight(self):
        return self.view.ShowHighlight

    @property
    def ShowHyphens(self):
        return self.view.ShowHyphens

    @property
    def ShowInkAnnotations(self):
        return self.view.ShowInkAnnotations

    @ShowInkAnnotations.setter
    def ShowInkAnnotations(self, value):
        self.view.ShowInkAnnotations = value

    @property
    def ShowInsertionsAndDeletions(self):
        return self.view.ShowInsertionsAndDeletions

    @property
    def ShowMainTextLayer(self):
        return self.view.ShowMainTextLayer

    @property
    def ShowMarkupAreaHighlight(self):
        return self.view.ShowMarkupAreaHighlight

    @ShowMarkupAreaHighlight.setter
    def ShowMarkupAreaHighlight(self, value):
        self.view.ShowMarkupAreaHighlight = value

    @property
    def ShowObjectAnchors(self):
        return self.view.ShowObjectAnchors

    @property
    def ShowOptionalBreaks(self):
        return self.view.ShowOptionalBreaks

    @property
    def ShowOtherAuthors(self):
        return self.view.ShowOtherAuthors

    @property
    def ShowParagraphs(self):
        return self.view.ShowParagraphs

    @property
    def ShowPicturePlaceHolders(self):
        return self.view.ShowPicturePlaceHolders

    @property
    def ShowRevisionsAndComments(self):
        return self.view.ShowRevisionsAndComments

    @property
    def ShowSpaces(self):
        return self.view.ShowSpaces

    @property
    def ShowTabs(self):
        return self.view.ShowTabs

    @property
    def ShowTextBoundaries(self):
        return self.view.ShowTextBoundaries

    @property
    def ShowXMLMarkup(self):
        return self.view.ShowXMLMarkup

    @property
    def SplitSpecial(self):
        return WdSpecialPane(self.view.SplitSpecial)

    @SplitSpecial.setter
    def SplitSpecial(self, value):
        self.view.SplitSpecial = value

    @property
    def TableGridlines(self):
        return self.view.TableGridlines

    @property
    def Type(self):
        return WdViewType(self.view.Type)

    @Type.setter
    def Type(self, value):
        self.view.Type = value

    @property
    def WrapToWindow(self):
        return self.view.WrapToWindow

    @property
    def Zoom(self):
        return Zoom(self.view.Zoom)

    def CollapseOutline(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.view.CollapseOutline(*args, **arguments)

    def ExpandOutline(self, *args, Range=None):
        arguments = {"Range": Range}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.view.ExpandOutline(*args, **arguments)

    def NextHeaderFooter(self):
        self.view.NextHeaderFooter()

    def PreviousHeaderFooter(self):
        self.view.PreviousHeaderFooter()

    def ShowAllHeadings(self):
        self.view.ShowAllHeadings()

    def ShowHeading(self, *args, Level=None):
        arguments = {"Level": Level}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.view.ShowHeading(*args, **arguments)

class Walls:

    def __init__(self, walls=None):
        self.walls = walls

    @property
    def Application(self):
        return self.walls.Application

    @property
    def Creator(self):
        return self.walls.Creator

    @property
    def Format(self):
        return ChartFormat(self.walls.Format)

    @property
    def Name(self):
        return self.walls.Name

    @property
    def Parent(self):
        return self.walls.Parent

    @property
    def PictureType(self):
        return self.walls.PictureType

    @PictureType.setter
    def PictureType(self, value):
        self.walls.PictureType = value

    @property
    def PictureUnit(self):
        return self.walls.PictureUnit

    @PictureUnit.setter
    def PictureUnit(self, value):
        self.walls.PictureUnit = value

    @property
    def Thickness(self):
        return self.walls.Thickness

    @Thickness.setter
    def Thickness(self, value):
        self.walls.Thickness = value

    def ClearFormats(self):
        self.walls.ClearFormats()

    def Paste(self):
        self.walls.Paste()

    def Select(self):
        self.walls.Select()

class WebOptions:

    def __init__(self, weboptions=None):
        self.weboptions = weboptions

    @property
    def AllowPNG(self):
        return self.weboptions.AllowPNG

    @property
    def Application(self):
        return Application(self.weboptions.Application)

    @property
    def BrowserLevel(self):
        return WdBrowserLevel(self.weboptions.BrowserLevel)

    @BrowserLevel.setter
    def BrowserLevel(self, value):
        self.weboptions.BrowserLevel = value

    @property
    def Creator(self):
        return self.weboptions.Creator

    @property
    def Encoding(self):
        return self.weboptions.Encoding

    @Encoding.setter
    def Encoding(self, value):
        self.weboptions.Encoding = value

    @property
    def FolderSuffix(self):
        return self.weboptions.FolderSuffix

    @property
    def OptimizeForBrowser(self):
        return self.weboptions.OptimizeForBrowser

    @property
    def OrganizeInFolder(self):
        return self.weboptions.OrganizeInFolder

    @property
    def Parent(self):
        return self.weboptions.Parent

    @property
    def PixelsPerInch(self):
        return self.weboptions.PixelsPerInch

    @PixelsPerInch.setter
    def PixelsPerInch(self, value):
        self.weboptions.PixelsPerInch = value

    @property
    def RelyOnCSS(self):
        return self.weboptions.RelyOnCSS

    @property
    def RelyOnVML(self):
        return self.weboptions.RelyOnVML

    @property
    def ScreenSize(self):
        return self.weboptions.ScreenSize

    @ScreenSize.setter
    def ScreenSize(self, value):
        self.weboptions.ScreenSize = value

    @property
    def TargetBrowser(self):
        return self.weboptions.TargetBrowser

    @property
    def UseLongFileNames(self):
        return self.weboptions.UseLongFileNames

    def UseDefaultFolderSuffix(self):
        self.weboptions.UseDefaultFolderSuffix()

class Window:

    def __init__(self, window=None):
        self.window = window

    @property
    def Active(self):
        return self.window.Active

    @property
    def ActivePane(self):
        return Pane(self.window.ActivePane)

    @property
    def Application(self):
        return Application(self.window.Application)

    @property
    def Caption(self):
        return self.window.Caption

    @Caption.setter
    def Caption(self, value):
        self.window.Caption = value

    @property
    def Creator(self):
        return self.window.Creator

    @property
    def DisplayHorizontalScrollBar(self):
        return self.window.DisplayHorizontalScrollBar

    @property
    def DisplayLeftScrollBar(self):
        return self.window.DisplayLeftScrollBar

    @property
    def DisplayRightRuler(self):
        return self.window.DisplayRightRuler

    @property
    def DisplayRulers(self):
        return self.window.DisplayRulers

    @property
    def DisplayScreenTips(self):
        return self.window.DisplayScreenTips

    @property
    def DisplayVerticalRuler(self):
        return self.window.DisplayVerticalRuler

    @property
    def DisplayVerticalScrollBar(self):
        return self.window.DisplayVerticalScrollBar

    @property
    def Document(self):
        return Document(self.window.Document)

    @property
    def DocumentMap(self):
        return self.window.DocumentMap

    @property
    def EnvelopeVisible(self):
        return self.window.EnvelopeVisible

    @property
    def Height(self):
        return self.window.Height

    @Height.setter
    def Height(self, value):
        self.window.Height = value

    @property
    def HorizontalPercentScrolled(self):
        return self.window.HorizontalPercentScrolled

    @HorizontalPercentScrolled.setter
    def HorizontalPercentScrolled(self, value):
        self.window.HorizontalPercentScrolled = value

    @property
    def IMEMode(self):
        return WdIMEMode(self.window.IMEMode)

    @IMEMode.setter
    def IMEMode(self, value):
        self.window.IMEMode = value

    @property
    def Index(self):
        return self.window.Index

    @property
    def Left(self):
        return self.window.Left

    @Left.setter
    def Left(self, value):
        self.window.Left = value

    @property
    def Next(self):
        return self.window.Next

    @property
    def Panes(self):
        return self.window.Panes

    @property
    def Parent(self):
        return self.window.Parent

    @property
    def Previous(self):
        return self.window.Previous

    @property
    def Selection(self):
        return Selection(self.window.Selection)

    @property
    def ShowSourceDocuments(self):
        return WdShowSourceDocuments(self.window.ShowSourceDocuments)

    @ShowSourceDocuments.setter
    def ShowSourceDocuments(self, value):
        self.window.ShowSourceDocuments = value

    @property
    def Split(self):
        return self.window.Split

    @property
    def SplitVertical(self):
        return self.window.SplitVertical

    @SplitVertical.setter
    def SplitVertical(self, value):
        self.window.SplitVertical = value

    @property
    def StyleAreaWidth(self):
        return self.window.StyleAreaWidth

    @StyleAreaWidth.setter
    def StyleAreaWidth(self, value):
        self.window.StyleAreaWidth = value

    @property
    def Thumbnails(self):
        return self.window.Thumbnails

    @property
    def Top(self):
        return self.window.Top

    @Top.setter
    def Top(self, value):
        self.window.Top = value

    @property
    def Type(self):
        return WdWindowType(self.window.Type)

    @property
    def UsableHeight(self):
        return self.window.UsableHeight

    @property
    def UsableWidth(self):
        return self.window.UsableWidth

    @property
    def VerticalPercentScrolled(self):
        return self.window.VerticalPercentScrolled

    @VerticalPercentScrolled.setter
    def VerticalPercentScrolled(self, value):
        self.window.VerticalPercentScrolled = value

    @property
    def View(self):
        return View(self.window.View)

    @property
    def Visible(self):
        return self.window.Visible

    @property
    def Width(self):
        return self.window.Width

    @Width.setter
    def Width(self, value):
        self.window.Width = value

    @property
    def WindowState(self):
        return WdWindowState(self.window.WindowState)

    @WindowState.setter
    def WindowState(self, value):
        self.window.WindowState = value

    def Activate(self):
        self.window.Activate()

    def Close(self, *args, SaveChanges=None, RouteDocument=None):
        arguments = {"SaveChanges": SaveChanges, "RouteDocument": RouteDocument}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.window.Close(*args, **arguments)

    def GetPoint(self, *args, ScreenPixelsLeft=None, ScreenPixelsTop=None, ScreenPixelsWidth=None, ScreenPixelsHeight=None, obj=None):
        arguments = {"ScreenPixelsLeft": ScreenPixelsLeft, "ScreenPixelsTop": ScreenPixelsTop, "ScreenPixelsWidth": ScreenPixelsWidth, "ScreenPixelsHeight": ScreenPixelsHeight, "obj": obj}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.window.GetPoint(*args, **arguments)

    def LargeScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.window.LargeScroll(*args, **arguments)

    def NewWindow(self):
        return self.window.NewWindow()

    def PageScroll(self, *args, Down=None, Up=None):
        arguments = {"Down": Down, "Up": Up}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.window.PageScroll(*args, **arguments)

    def PrintOut(self, *args, Background=None, Append=None, Range=None, OutputFileName=None, From=None, To=None, Item=None, Copies=None, Pages=None, PageType=None, PrintToFile=None, Collate=None, FileName=None, ActivePrinterMacGX=None, ManualDuplexPrint=None, PrintZoomColumn=None, PrintZoomRow=None, PrintZoomPaperWidth=None, PrintZoomPaperHeight=None):
        arguments = {"Background": Background, "Append": Append, "Range": Range, "OutputFileName": OutputFileName, "From": From, "To": To, "Item": Item, "Copies": Copies, "Pages": Pages, "PageType": PageType, "PrintToFile": PrintToFile, "Collate": Collate, "FileName": FileName, "ActivePrinterMacGX": ActivePrinterMacGX, "ManualDuplexPrint": ManualDuplexPrint, "PrintZoomColumn": PrintZoomColumn, "PrintZoomRow": PrintZoomRow, "PrintZoomPaperWidth": PrintZoomPaperWidth, "PrintZoomPaperHeight": PrintZoomPaperHeight}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.window.PrintOut(*args, **arguments)

    def RangeFromPoint(self, *args, x=None, y=None):
        arguments = {"x": x, "y": y}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.window.RangeFromPoint(*args, **arguments)

    def ScrollIntoView(self, *args, Obj=None, Start=None):
        arguments = {"Obj": Obj, "Start": Start}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.window.ScrollIntoView(*args, **arguments)

    def SetFocus(self):
        self.window.SetFocus()

    def SmallScroll(self, *args, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = {"Down": Down, "Up": Up, "ToRight": ToRight, "ToLeft": ToLeft}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.window.SmallScroll(*args, **arguments)

    def ToggleRibbon(self):
        self.window.ToggleRibbon()

class WrapFormat:

    def __init__(self, wrapformat=None):
        self.wrapformat = wrapformat

    @property
    def AllowOverlap(self):
        return self.wrapformat.AllowOverlap

    @AllowOverlap.setter
    def AllowOverlap(self, value):
        self.wrapformat.AllowOverlap = value

    @property
    def Application(self):
        return Application(self.wrapformat.Application)

    @property
    def Creator(self):
        return self.wrapformat.Creator

    @property
    def DistanceBottom(self):
        return self.wrapformat.DistanceBottom

    @DistanceBottom.setter
    def DistanceBottom(self, value):
        self.wrapformat.DistanceBottom = value

    @property
    def DistanceLeft(self):
        return self.wrapformat.DistanceLeft

    @DistanceLeft.setter
    def DistanceLeft(self, value):
        self.wrapformat.DistanceLeft = value

    @property
    def DistanceRight(self):
        return self.wrapformat.DistanceRight

    @DistanceRight.setter
    def DistanceRight(self, value):
        self.wrapformat.DistanceRight = value

    @property
    def DistanceTop(self):
        return self.wrapformat.DistanceTop

    @DistanceTop.setter
    def DistanceTop(self, value):
        self.wrapformat.DistanceTop = value

    @property
    def Parent(self):
        return self.wrapformat.Parent

    @property
    def Side(self):
        return WdWrapSideType(self.wrapformat.Side)

    @Side.setter
    def Side(self, value):
        self.wrapformat.Side = value

    @property
    def Type(self):
        return WdWrapType(self.wrapformat.Type)

class XMLMapping:

    def __init__(self, xmlmapping=None):
        self.xmlmapping = xmlmapping

    @property
    def Application(self):
        return Application(self.xmlmapping.Application)

    @property
    def Creator(self):
        return self.xmlmapping.Creator

    @property
    def CustomXMLNode(self):
        return self.xmlmapping.CustomXMLNode

    @property
    def CustomXMLPart(self):
        return self.xmlmapping.CustomXMLPart

    @property
    def IsMapped(self):
        return self.xmlmapping.IsMapped

    @property
    def Parent(self):
        return self.xmlmapping.Parent

    @property
    def PrefixMappings(self):
        return self.xmlmapping.PrefixMappings

    @property
    def XPath(self):
        return self.xmlmapping.XPath

    def Delete(self):
        self.xmlmapping.Delete()

    def SetMapping(self, *args, XPath=None, PrefixMapping=None, Source=None):
        arguments = {"XPath": XPath, "PrefixMapping": PrefixMapping, "Source": Source}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.xmlmapping.SetMapping(*args, **arguments)

    def SetMappingByNode(self, *args, Node=None):
        arguments = {"Node": Node}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.xmlmapping.SetMappingByNode(*args, **arguments)

class XMLNamespace:

    def __init__(self, xmlnamespace=None):
        self.xmlnamespace = xmlnamespace

    @property
    def Alias(self):
        return self.xmlnamespace.Alias

    @property
    def Application(self):
        return Application(self.xmlnamespace.Application)

    @property
    def Creator(self):
        return self.xmlnamespace.Creator

    @property
    def DefaultTransform(self):
        return XSLTransform(self.xmlnamespace.DefaultTransform)

    @property
    def Location(self):
        return self.xmlnamespace.Location

    @Location.setter
    def Location(self, value):
        self.xmlnamespace.Location = value

    @property
    def Parent(self):
        return self.xmlnamespace.Parent

    @property
    def XSLTransforms(self):
        return self.xmlnamespace.XSLTransforms

    def AttachToDocument(self, *args, Document=None):
        arguments = {"Document": Document}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.xmlnamespace.AttachToDocument(*args, **arguments)

    def Delete(self):
        self.xmlnamespace.Delete()

class XMLNode:

    def __init__(self, xmlnode=None):
        self.xmlnode = xmlnode

    @property
    def Application(self):
        return Application(self.xmlnode.Application)

    @property
    def Attributes(self):
        return XMLNodes(self.xmlnode.Attributes)

    @property
    def BaseName(self):
        return self.xmlnode.BaseName

    @property
    def ChildNodes(self):
        return XMLNodes(self.xmlnode.ChildNodes)

    @property
    def Creator(self):
        return self.xmlnode.Creator

    @property
    def FirstChild(self):
        return self.xmlnode.FirstChild

    @property
    def HasChildNodes(self):
        return self.xmlnode.HasChildNodes

    @property
    def LastChild(self):
        return XMLNode(self.xmlnode.LastChild)

    @property
    def Level(self):
        return self.xmlnode.Level

    @property
    def NamespaceURI(self):
        return self.xmlnode.NamespaceURI

    @property
    def NextSibling(self):
        return XMLNode(self.xmlnode.NextSibling)

    @property
    def NodeType(self):
        return self.xmlnode.NodeType

    @property
    def NodeValue(self):
        return self.xmlnode.NodeValue

    @NodeValue.setter
    def NodeValue(self, value):
        self.xmlnode.NodeValue = value

    @property
    def OwnerDocument(self):
        return Document(self.xmlnode.OwnerDocument)

    @property
    def Parent(self):
        return self.xmlnode.Parent

    @property
    def ParentNode(self):
        return XMLNode(self.xmlnode.ParentNode)

    @property
    def PlaceholderText(self):
        return self.xmlnode.PlaceholderText

    @property
    def PreviousSibling(self):
        return XMLNode(self.xmlnode.PreviousSibling)

    @property
    def Range(self):
        return Range(self.xmlnode.Range)

    @property
    def Text(self):
        return self.xmlnode.Text

    @Text.setter
    def Text(self, value):
        self.xmlnode.Text = value

    def ValidationErrorText(self, *args, Advanced=None):
        arguments = {"Advanced": Advanced}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.xmlnode.ValidationErrorText(*args, **arguments)

    @property
    def ValidationStatus(self):
        return self.xmlnode.ValidationStatus

    @property
    def WordOpenXML(self):
        return self.xmlnode.WordOpenXML

    def Copy(self):
        return self.xmlnode.Copy()

    def Cut(self):
        self.xmlnode.Cut()

    def Delete(self):
        self.xmlnode.Delete()

    def RemoveChild(self, *args, ChildElement=None):
        arguments = {"ChildElement": ChildElement}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.xmlnode.RemoveChild(*args, **arguments)

    def SelectNodes(self, *args, XPath=None, PrefixMapping=None, FastSearchSkippingTextNodes=None):
        arguments = {"XPath": XPath, "PrefixMapping": PrefixMapping, "FastSearchSkippingTextNodes": FastSearchSkippingTextNodes}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.xmlnode.SelectNodes(*args, **arguments)

    def SelectSingleNode(self, *args, XPath=None, PrefixMapping=None, FastSearchSkippingTextNodes=None):
        arguments = {"XPath": XPath, "PrefixMapping": PrefixMapping, "FastSearchSkippingTextNodes": FastSearchSkippingTextNodes}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.xmlnode.SelectSingleNode(*args, **arguments)

    def SetValidationError(self, *args, Status=None, ErrorText=None, ClearedAutomatically=None):
        arguments = {"Status": Status, "ErrorText": ErrorText, "ClearedAutomatically": ClearedAutomatically}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        self.xmlnode.SetValidationError(*args, **arguments)

    def Validate(self):
        return self.xmlnode.Validate()

class XMLNodes:

    def __init__(self, xmlnodes=None):
        self.xmlnodes = xmlnodes

    def __call__(self, item):
        return XMLNode(self.xmlnodes(item))

    @property
    def Application(self):
        return Application(self.xmlnodes.Application)

    @property
    def Count(self):
        return self.xmlnodes.Count

    @property
    def Creator(self):
        return self.xmlnodes.Creator

    @property
    def Parent(self):
        return self.xmlnodes.Parent

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.xmlnodes.Item(*args, **arguments)

class XMLSchemaReference:

    def __init__(self, xmlschemareference=None):
        self.xmlschemareference = xmlschemareference

    @property
    def Application(self):
        return Application(self.xmlschemareference.Application)

    @property
    def Creator(self):
        return self.xmlschemareference.Creator

    @property
    def Location(self):
        return self.xmlschemareference.Location

    @property
    def NamespaceURI(self):
        return self.xmlschemareference.NamespaceURI

    @property
    def Parent(self):
        return self.xmlschemareference.Parent

    def Delete(self):
        self.xmlschemareference.Delete()

    def Reload(self):
        return self.xmlschemareference.Reload()

class XMLSchemaReferences:

    def __init__(self, xmlschemareferences=None):
        self.xmlschemareferences = xmlschemareferences

    def __call__(self, item):
        return XMLSchemaReference(self.xmlschemareferences(item))

    @property
    def Application(self):
        return Application(self.xmlschemareferences.Application)

    @property
    def Count(self):
        return self.xmlschemareferences.Count

    @property
    def Creator(self):
        return self.xmlschemareferences.Creator

    @property
    def HideValidationErrors(self):
        return self.xmlschemareferences.HideValidationErrors

    @property
    def IgnoreMixedContent(self):
        return self.xmlschemareferences.IgnoreMixedContent

    @property
    def Parent(self):
        return self.xmlschemareferences.Parent

    @property
    def ShowPlaceholderText(self):
        return self.xmlschemareferences.ShowPlaceholderText

    @ShowPlaceholderText.setter
    def ShowPlaceholderText(self, value):
        self.xmlschemareferences.ShowPlaceholderText = value

    def Add(self, *args, NamespaceURI=None, Alias=None, FileName=None, InstallForAllUsers=None):
        arguments = {"NamespaceURI": NamespaceURI, "Alias": Alias, "FileName": FileName, "InstallForAllUsers": InstallForAllUsers}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return XMLSchemaReference(self.xmlschemareferences.Add(*args, **arguments))

    def Item(self, *args, Index=None):
        arguments = {"Index": Index}
        arguments = {key: value for key, value in arguments.items() if value is not None}
        return self.xmlschemareferences.Item(*args, **arguments)

    def Validate(self):
        return self.xmlschemareferences.Validate()

class XSLTransform:

    def __init__(self, xsltransform=None):
        self.xsltransform = xsltransform

    @property
    def Alias(self):
        return self.xsltransform.Alias

    @property
    def Application(self):
        return Application(self.xsltransform.Application)

    @property
    def Creator(self):
        return self.xsltransform.Creator

    @property
    def ID(self):
        return self.xsltransform.ID

    @property
    def Location(self):
        return self.xsltransform.Location

    @Location.setter
    def Location(self, value):
        self.xsltransform.Location = value

    @property
    def Parent(self):
        return self.xsltransform.Parent

    def Delete(self):
        self.xsltransform.Delete()

class Zoom:

    def __init__(self, zoom=None):
        self.zoom = zoom

    @property
    def Application(self):
        return Application(self.zoom.Application)

    @property
    def Creator(self):
        return self.zoom.Creator

    @property
    def PageColumns(self):
        return self.zoom.PageColumns

    @PageColumns.setter
    def PageColumns(self, value):
        self.zoom.PageColumns = value

    @property
    def PageFit(self):
        return WdPageFit(self.zoom.PageFit)

    @PageFit.setter
    def PageFit(self, value):
        self.zoom.PageFit = value

    @property
    def PageRows(self):
        return self.zoom.PageRows

    @PageRows.setter
    def PageRows(self, value):
        self.zoom.PageRows = value

    @property
    def Parent(self):
        return self.zoom.Parent

    @property
    def Percentage(self):
        return self.zoom.Percentage

    @Percentage.setter
    def Percentage(self, value):
        self.zoom.Percentage = value
