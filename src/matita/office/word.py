from . import com_arguments

import win32com.client
import pythoncom

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

    @Installed.setter
    def Installed(self, value):
        self.addin.Installed = value

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

    @BrowseExtraFileTypes.setter
    def BrowseExtraFileTypes(self, value):
        self.application.BrowseExtraFileTypes = value

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

    @CheckLanguage.setter
    def CheckLanguage(self, value):
        self.application.CheckLanguage = value

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

    @DefaultLegalBlackline.setter
    def DefaultLegalBlackline(self, value):
        self.application.DefaultLegalBlackline = value

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

    @DisplayAutoCompleteTips.setter
    def DisplayAutoCompleteTips(self, value):
        self.application.DisplayAutoCompleteTips = value

    @property
    def DisplayDocumentInformationPanel(self):
        return self.application.DisplayDocumentInformationPanel

    @DisplayDocumentInformationPanel.setter
    def DisplayDocumentInformationPanel(self, value):
        self.application.DisplayDocumentInformationPanel = value

    @property
    def DisplayRecentFiles(self):
        return self.application.DisplayRecentFiles

    @DisplayRecentFiles.setter
    def DisplayRecentFiles(self, value):
        self.application.DisplayRecentFiles = value

    @property
    def DisplayScreenTips(self):
        return self.application.DisplayScreenTips

    @DisplayScreenTips.setter
    def DisplayScreenTips(self, value):
        self.application.DisplayScreenTips = value

    @property
    def DisplayScrollBars(self):
        return self.application.DisplayScrollBars

    @DisplayScrollBars.setter
    def DisplayScrollBars(self, value):
        self.application.DisplayScrollBars = value

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

    def FileDialog(self, FileDialogType=None):
        arguments = com_arguments([FileDialogType])
        if callable(self.application.FileDialog):
            return self.application.FileDialog(*arguments)
        else:
            return self.application.GetFileDialog(*arguments)

    @property
    def FileValidation(self):
        return self.application.FileValidation

    @FileValidation.setter
    def FileValidation(self, value):
        self.application.FileValidation = value

    def FindKey(self, KeyCode=None, KeyCode2=None):
        arguments = com_arguments([KeyCode, KeyCode2])
        if callable(self.application.FindKey):
            return KeyBinding(self.application.FindKey(*arguments))
        else:
            return KeyBinding(self.application.GetFindKey(*arguments))

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

    def International(self, Index=None):
        arguments = com_arguments([Index])
        if callable(self.application.International):
            return self.application.International(*arguments)
        else:
            return self.application.GetInternational(*arguments)

    def IsObjectValid(self, Object=None):
        arguments = com_arguments([Object])
        if callable(self.application.IsObjectValid):
            return self.application.IsObjectValid(*arguments)
        else:
            return self.application.GetIsObjectValid(*arguments)

    @property
    def IsSandboxed(self):
        return self.application.IsSandboxed

    @property
    def KeyBindings(self):
        return self.application.KeyBindings

    def KeysBoundTo(self, KeyCategory=None, Command=None, CommandParameter=None):
        arguments = com_arguments([KeyCategory, Command, CommandParameter])
        if callable(self.application.KeysBoundTo):
            return self.application.KeysBoundTo(*arguments)
        else:
            return self.application.GetKeysBoundTo(*arguments)

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

    @PrintPreview.setter
    def PrintPreview(self, value):
        self.application.PrintPreview = value

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

    @ScreenUpdating.setter
    def ScreenUpdating(self, value):
        self.application.ScreenUpdating = value

    @property
    def Selection(self):
        return Selection(self.application.Selection)

    @property
    def SensitivityLabelPolicy(self):
        return self.application.SensitivityLabelPolicy

    @property
    def ShowStartupDialog(self):
        return self.application.ShowStartupDialog

    @ShowStartupDialog.setter
    def ShowStartupDialog(self, value):
        self.application.ShowStartupDialog = value

    @property
    def ShowStylePreviews(self):
        return self.application.ShowStylePreviews

    @ShowStylePreviews.setter
    def ShowStylePreviews(self, value):
        self.application.ShowStylePreviews = value

    @property
    def ShowVisualBasicEditor(self):
        return self.application.ShowVisualBasicEditor

    @ShowVisualBasicEditor.setter
    def ShowVisualBasicEditor(self, value):
        self.application.ShowVisualBasicEditor = value

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

    def SynonymInfo(self, Word=None, LanguageID=None):
        arguments = com_arguments([Word, LanguageID])
        if callable(self.application.SynonymInfo):
            return SynonymInfo(self.application.SynonymInfo(*arguments))
        else:
            return SynonymInfo(self.application.GetSynonymInfo(*arguments))

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

    @Visible.setter
    def Visible(self, value):
        self.application.Visible = value

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

    def AddAddress(self, TagID=None, Value=None):
        arguments = com_arguments([TagID, Value])
        self.application.AddAddress(*arguments)

    def Application(self):
        self.application.Application()

    def AutomaticChange(self):
        self.application.AutomaticChange()

    def BuildKeyCode(self, Arg1=None, Arg2=None, Arg3=None, Arg4=None):
        arguments = com_arguments([Arg1, Arg2, Arg3, Arg4])
        self.application.BuildKeyCode(*arguments)

    def CentimetersToPoints(self, Centimeters=None):
        arguments = com_arguments([Centimeters])
        self.application.CentimetersToPoints(*arguments)

    def ChangeFileOpenDirectory(self, Path=None):
        arguments = com_arguments([Path])
        self.application.ChangeFileOpenDirectory(*arguments)

    def CheckGrammar(self, String=None):
        arguments = com_arguments([String])
        return self.application.CheckGrammar(*arguments)

    def CheckSpelling(self, Word=None, CustomDictionary=None, IgnoreUppercase=None, MainDictionary=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = com_arguments([Word, CustomDictionary, IgnoreUppercase, MainDictionary, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10])
        return self.application.CheckSpelling(*arguments)

    def CleanString(self, String=None):
        arguments = com_arguments([String])
        self.application.CleanString(*arguments)

    def CompareDocuments(self, OriginalDocument=None, RevisedDocument=None, Destination=None, Granularity=None, CompareFormatting=None, CompareCaseChanges=None, CompareWhitespace=None, CompareTables=None, CompareHeaders=None, CompareFootnotes=None, CompareTextboxes=None, CompareFields=None, CompareComments=None, CompareMoves=None, RevisedAuthor=None, IgnoreAllComparisonWarnings=None):
        arguments = com_arguments([OriginalDocument, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes, CompareFields, CompareComments, CompareMoves, RevisedAuthor, IgnoreAllComparisonWarnings])
        return self.application.CompareDocuments(*arguments)

    def DDEInitiate(self, App=None, Topic=None):
        arguments = com_arguments([App, Topic])
        self.application.DDEInitiate(*arguments)

    def DDEPoke(self, Channel=None, Item=None, Data=None):
        arguments = com_arguments([Channel, Item, Data])
        self.application.DDEPoke(*arguments)

    def DDERequest(self, Channel=None, Item=None):
        arguments = com_arguments([Channel, Item])
        self.application.DDERequest(*arguments)

    def DDETerminate(self, Channel=None):
        arguments = com_arguments([Channel])
        self.application.DDETerminate(*arguments)

    def DDETerminateAll(self):
        self.application.DDETerminateAll()

    def DefaultWebOptions(self):
        return self.application.DefaultWebOptions()

    def GetAddress(self, Name=None, AddressProperties=None, UseAutoText=None, DisplaySelectDialog=None, SelectDialog=None, CheckNamesDialog=None, RecentAddressesChoice=None, UpdateRecentAddresses=None):
        arguments = com_arguments([Name, AddressProperties, UseAutoText, DisplaySelectDialog, SelectDialog, CheckNamesDialog, RecentAddressesChoice, UpdateRecentAddresses])
        return self.application.GetAddress(*arguments)

    def GetDefaultTheme(self, DocumentType=None):
        arguments = com_arguments([DocumentType])
        self.application.GetDefaultTheme(*arguments)

    def GetSpellingSuggestions(self, Word=None, CustomDictionary=None, IgnoreUppercase=None, MainDictionary=None, SuggestionMode=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = com_arguments([Word, CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10])
        self.application.GetSpellingSuggestions(*arguments)

    def GoBack(self):
        self.application.GoBack()

    def GoForward(self):
        self.application.GoForward()

    def Help(self, HelpType=None):
        arguments = com_arguments([HelpType])
        self.application.Help(*arguments)

    def HelpTool(self):
        self.application.HelpTool()

    def InchesToPoints(self, Inches=None):
        arguments = com_arguments([Inches])
        self.application.InchesToPoints(*arguments)

    def Keyboard(self, LangId=None):
        arguments = com_arguments([LangId])
        self.application.Keyboard(*arguments)

    def KeyboardBidi(self):
        self.application.KeyboardBidi()

    def KeyboardLatin(self):
        self.application.KeyboardLatin()

    def KeyString(self, KeyCode=None, KeyCode2=None):
        arguments = com_arguments([KeyCode, KeyCode2])
        return self.application.KeyString(*arguments)

    def LinesToPoints(self, Lines=None):
        arguments = com_arguments([Lines])
        return self.application.LinesToPoints(*arguments)

    def ListCommands(self, ListAllCommands=None):
        arguments = com_arguments([ListAllCommands])
        self.application.ListCommands(*arguments)

    def LoadMasterList(self, FileName=None):
        arguments = com_arguments([FileName])
        self.application.LoadMasterList(*arguments)

    def LookupNameProperties(self, Name=None):
        arguments = com_arguments([Name])
        self.application.LookupNameProperties(*arguments)

    def MergeDocuments(self, OriginalDocument=None, RevisedDocument=None, Destination=None, Granularity=None, CompareFormatting=None, CompareCaseChanges=None, CompareWhitespace=None, CompareTables=None, CompareHeaders=None, CompareFootnotes=None, CompareTextboxes=None, CompareFields=None, CompareComments=None, OriginalAuthor=None, RevisedAuthor=None, FormatFrom=None):
        arguments = com_arguments([OriginalDocument, RevisedDocument, Destination, Granularity, CompareFormatting, CompareCaseChanges, CompareWhitespace, CompareTables, CompareHeaders, CompareFootnotes, CompareTextboxes, CompareFields, CompareComments, OriginalAuthor, RevisedAuthor, FormatFrom])
        return self.application.MergeDocuments(*arguments)

    def MillimetersToPoints(self, Millimeters=None):
        arguments = com_arguments([Millimeters])
        return self.application.MillimetersToPoints(*arguments)

    def Move(self, Left=None, Top=None):
        arguments = com_arguments([Left, Top])
        self.application.Move(*arguments)

    def NewWindow(self):
        return self.application.NewWindow()

    def OnTime(self, When=None, Name=None, Tolerance=None):
        arguments = com_arguments([When, Name, Tolerance])
        self.application.OnTime(*arguments)

    def OrganizerCopy(self, Source=None, Destination=None, Name=None, Object=None):
        arguments = com_arguments([Source, Destination, Name, Object])
        self.application.OrganizerCopy(*arguments)

    def OrganizerDelete(self, Source=None, Name=None, Object=None):
        arguments = com_arguments([Source, Name, Object])
        self.application.OrganizerDelete(*arguments)

    def OrganizerRename(self, Source=None, Name=None, NewName=None, Object=None):
        arguments = com_arguments([Source, Name, NewName, Object])
        self.application.OrganizerRename(*arguments)

    def PicasToPoints(self, Picas=None):
        arguments = com_arguments([Picas])
        return self.application.PicasToPoints(*arguments)

    def PixelsToPoints(self, Pixels=None, fVertical=None):
        arguments = com_arguments([Pixels, fVertical])
        return self.application.PixelsToPoints(*arguments)

    def PointsToCentimeters(self, Points=None):
        arguments = com_arguments([Points])
        return self.application.PointsToCentimeters(*arguments)

    def PointsToInches(self, Points=None):
        arguments = com_arguments([Points])
        return self.application.PointsToInches(*arguments)

    def PointsToLines(self, Points=None):
        arguments = com_arguments([Points])
        return self.application.PointsToLines(*arguments)

    def PointsToMillimeters(self, Points=None):
        arguments = com_arguments([Points])
        return self.application.PointsToMillimeters(*arguments)

    def PointsToPicas(self, Points=None):
        arguments = com_arguments([Points])
        return self.application.PointsToPicas(*arguments)

    def PointsToPixels(self, Points=None, fVertical=None):
        arguments = com_arguments([Points, fVertical])
        return self.application.PointsToPixels(*arguments)

    def PrintOut(self, Background=None, Append=None, Range=None, OutputFileName=None, From=None, To=None, Item=None, Copies=None, Pages=None, PageType=None, PrintToFile=None, Collate=None, FileName=None, ActivePrinterMacGX=None, ManualDuplexPrint=None, PrintZoomColumn=None, PrintZoomRow=None, PrintZoomPaperWidth=None, PrintZoomPaperHeight=None):
        arguments = com_arguments([Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight])
        self.application.PrintOut(*arguments)

    def ProductCode(self):
        return self.application.ProductCode()

    def PutFocusInMailHeader(self):
        self.application.PutFocusInMailHeader()

    def Quit(self, SaveChanges=None, OriginalFormat=None, RouteDocument=None):
        arguments = com_arguments([SaveChanges, OriginalFormat, RouteDocument])
        self.application.Quit(*arguments)

    def Repeat(self, Times=None):
        arguments = com_arguments([Times])
        return self.application.Repeat(*arguments)

    def ResetIgnoreAll(self):
        self.application.ResetIgnoreAll()

    def Resize(self, Width=None, Height=None):
        arguments = com_arguments([Width, Height])
        self.application.Resize(*arguments)

    def Run(self, MacroName=None, varg1=None, varg2=None, varg3=None, varg4=None, varg5=None, varg6=None, varg7=None, varg8=None, varg9=None, varg10=None, varg11=None, varg12=None, varg13=None, varg14=None, varg15=None, varg16=None, varg17=None, varg18=None, varg19=None, varg20=None, varg21=None, varg22=None, varg23=None, varg24=None, varg25=None, varg26=None, varg27=None, varg28=None, varg29=None, varg30=None):
        arguments = com_arguments([MacroName, varg1, varg2, varg3, varg4, varg5, varg6, varg7, varg8, varg9, varg10, varg11, varg12, varg13, varg14, varg15, varg16, varg17, varg18, varg19, varg20, varg21, varg22, varg23, varg24, varg25, varg26, varg27, varg28, varg29, varg30])
        self.application.Run(*arguments)

    def ScreenRefresh(self):
        self.application.ScreenRefresh()

    def SetDefaultTheme(self, Name=None, DocumentType=None):
        arguments = com_arguments([Name, DocumentType])
        self.application.SetDefaultTheme(*arguments)

    def ShowClipboard(self):
        self.application.ShowClipboard()

    def ShowMe(self):
        self.application.ShowMe()

    def SubstituteFont(self, UnavailableFont=None, SubstituteFont=None):
        arguments = com_arguments([UnavailableFont, SubstituteFont])
        self.application.SubstituteFont(*arguments)

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

    @AutoInsert.setter
    def AutoInsert(self, value):
        self.autocaption.AutoInsert = value

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

    @CorrectCapsLock.setter
    def CorrectCapsLock(self, value):
        self.autocorrect.CorrectCapsLock = value

    @property
    def CorrectDays(self):
        return self.autocorrect.CorrectDays

    @CorrectDays.setter
    def CorrectDays(self, value):
        self.autocorrect.CorrectDays = value

    @property
    def CorrectHangulAndAlphabet(self):
        return self.autocorrect.CorrectHangulAndAlphabet

    @CorrectHangulAndAlphabet.setter
    def CorrectHangulAndAlphabet(self, value):
        self.autocorrect.CorrectHangulAndAlphabet = value

    @property
    def CorrectInitialCaps(self):
        return self.autocorrect.CorrectInitialCaps

    @CorrectInitialCaps.setter
    def CorrectInitialCaps(self, value):
        self.autocorrect.CorrectInitialCaps = value

    @property
    def CorrectKeyboardSetting(self):
        return self.autocorrect.CorrectKeyboardSetting

    @CorrectKeyboardSetting.setter
    def CorrectKeyboardSetting(self, value):
        self.autocorrect.CorrectKeyboardSetting = value

    @property
    def CorrectSentenceCaps(self):
        return self.autocorrect.CorrectSentenceCaps

    @CorrectSentenceCaps.setter
    def CorrectSentenceCaps(self, value):
        self.autocorrect.CorrectSentenceCaps = value

    @property
    def CorrectTableCells(self):
        return self.autocorrect.CorrectTableCells

    @CorrectTableCells.setter
    def CorrectTableCells(self, value):
        self.autocorrect.CorrectTableCells = value

    @property
    def Creator(self):
        return self.autocorrect.Creator

    @property
    def DisplayAutoCorrectOptions(self):
        return self.autocorrect.DisplayAutoCorrectOptions

    @DisplayAutoCorrectOptions.setter
    def DisplayAutoCorrectOptions(self, value):
        self.autocorrect.DisplayAutoCorrectOptions = value

    @property
    def Entries(self):
        return self.autocorrect.Entries

    @property
    def FirstLetterAutoAdd(self):
        return self.autocorrect.FirstLetterAutoAdd

    @FirstLetterAutoAdd.setter
    def FirstLetterAutoAdd(self, value):
        self.autocorrect.FirstLetterAutoAdd = value

    @property
    def FirstLetterExceptions(self):
        return self.autocorrect.FirstLetterExceptions

    @property
    def HangulAndAlphabetAutoAdd(self):
        return self.autocorrect.HangulAndAlphabetAutoAdd

    @HangulAndAlphabetAutoAdd.setter
    def HangulAndAlphabetAutoAdd(self, value):
        self.autocorrect.HangulAndAlphabetAutoAdd = value

    @property
    def HangulAndAlphabetExceptions(self):
        return self.autocorrect.HangulAndAlphabetExceptions

    @property
    def OtherCorrectionsAutoAdd(self):
        return self.autocorrect.OtherCorrectionsAutoAdd

    @OtherCorrectionsAutoAdd.setter
    def OtherCorrectionsAutoAdd(self, value):
        self.autocorrect.OtherCorrectionsAutoAdd = value

    @property
    def OtherCorrectionsExceptions(self):
        return self.autocorrect.OtherCorrectionsExceptions

    @property
    def Parent(self):
        return self.autocorrect.Parent

    @property
    def ReplaceText(self):
        return self.autocorrect.ReplaceText

    @ReplaceText.setter
    def ReplaceText(self, value):
        self.autocorrect.ReplaceText = value

    @property
    def ReplaceTextFromSpellingChecker(self):
        return self.autocorrect.ReplaceTextFromSpellingChecker

    @ReplaceTextFromSpellingChecker.setter
    def ReplaceTextFromSpellingChecker(self, value):
        self.autocorrect.ReplaceTextFromSpellingChecker = value

    @property
    def TwoInitialCapsAutoAdd(self):
        return self.autocorrect.TwoInitialCapsAutoAdd

    @TwoInitialCapsAutoAdd.setter
    def TwoInitialCapsAutoAdd(self, value):
        self.autocorrect.TwoInitialCapsAutoAdd = value

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

    def Apply(self, Range=None):
        arguments = com_arguments([Range])
        self.autocorrectentry.Apply(*arguments)

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

    def Insert(self, Where=None, RichText=None):
        arguments = com_arguments([Where, RichText])
        return self.autotextentry.Insert(*arguments)


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

    def Item(self, Type=None, AxisGroup=None):
        arguments = com_arguments([Type, AxisGroup])
        self.axes.Item(*arguments)


class Axis:

    def __init__(self, axis=None):
        self.axis = axis

    @property
    def Application(self):
        return self.axis.Application

    @property
    def AxisBetweenCategories(self):
        return self.axis.AxisBetweenCategories

    @AxisBetweenCategories.setter
    def AxisBetweenCategories(self, value):
        self.axis.AxisBetweenCategories = value

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

    @BaseUnitIsAuto.setter
    def BaseUnitIsAuto(self, value):
        self.axis.BaseUnitIsAuto = value

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

    @DisplayUnitCustom.setter
    def DisplayUnitCustom(self, value):
        self.axis.DisplayUnitCustom = value

    @property
    def DisplayUnitLabel(self):
        return DisplayUnitLabel(self.axis.DisplayUnitLabel)

    @property
    def Format(self):
        return ChartFormat(self.axis.Format)

    @property
    def HasDisplayUnitLabel(self):
        return self.axis.HasDisplayUnitLabel

    @HasDisplayUnitLabel.setter
    def HasDisplayUnitLabel(self, value):
        self.axis.HasDisplayUnitLabel = value

    @property
    def HasMajorGridlines(self):
        return self.axis.HasMajorGridlines

    @HasMajorGridlines.setter
    def HasMajorGridlines(self, value):
        self.axis.HasMajorGridlines = value

    @property
    def HasMinorGridlines(self):
        return self.axis.HasMinorGridlines

    @HasMinorGridlines.setter
    def HasMinorGridlines(self, value):
        self.axis.HasMinorGridlines = value

    @property
    def HasTitle(self):
        return self.axis.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.axis.HasTitle = value

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

    @MajorUnitIsAuto.setter
    def MajorUnitIsAuto(self, value):
        self.axis.MajorUnitIsAuto = value

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

    @MaximumScaleIsAuto.setter
    def MaximumScaleIsAuto(self, value):
        self.axis.MaximumScaleIsAuto = value

    @property
    def MinimumScale(self):
        return self.axis.MinimumScale

    @MinimumScale.setter
    def MinimumScale(self, value):
        self.axis.MinimumScale = value

    @property
    def MinimumScaleIsAuto(self):
        return self.axis.MinimumScaleIsAuto

    @MinimumScaleIsAuto.setter
    def MinimumScaleIsAuto(self, value):
        self.axis.MinimumScaleIsAuto = value

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

    @MinorUnitIsAuto.setter
    def MinorUnitIsAuto(self, value):
        self.axis.MinorUnitIsAuto = value

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

    @ReversePlotOrder.setter
    def ReversePlotOrder(self, value):
        self.axis.ReversePlotOrder = value

    @property
    def ScaleType(self):
        return self.axis.ScaleType

    @ScaleType.setter
    def ScaleType(self, value):
        self.axis.ScaleType = value

    @property
    def TickLabelPosition(self):
        return self.axis.TickLabelPosition

    @TickLabelPosition.setter
    def TickLabelPosition(self, value):
        self.axis.TickLabelPosition = value

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        if callable(self.axistitle.Characters):
            return ChartCharacters(self.axistitle.Characters(*arguments))
        else:
            return ChartCharacters(self.axistitle.GetCharacters(*arguments))

    @property
    def Creator(self):
        return self.axistitle.Creator

    @property
    def Format(self):
        return ChartFormat(self.axistitle.Format)

    @property
    def Formula(self):
        return self.axistitle.Formula

    @Formula.setter
    def Formula(self, value):
        self.axistitle.Formula = value

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

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.axistitle.IncludeInLayout = value

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

    def Copy(self, Name=None):
        arguments = com_arguments([Name])
        self.bookmark.Copy(*arguments)

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

    @Visible.setter
    def Visible(self, value):
        self.border.Visible = value


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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.breaks.Item(*arguments)


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

    def Insert(self, Where=None, RichText=None):
        arguments = com_arguments([Where, RichText])
        return self.buildingblock.Insert(*arguments)


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

    def Add(self, Name=None, Type=None, Category=None, Range=None, Description=None, InsertOptions=None):
        arguments = com_arguments([Name, Type, Category, Range, Description, InsertOptions])
        return self.buildingblockentries.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.buildingblockentries.Item(*arguments)


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

    def Add(self, Name=None, Range=None, Description=None, InsertOptions=None):
        arguments = com_arguments([Name, Range, Description, InsertOptions])
        return self.buildingblocks.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.buildingblocks.Item(*arguments)


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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.buildingblocktypes.Item(*arguments)


class CalloutFormat:

    def __init__(self, calloutformat=None):
        self.calloutformat = calloutformat

    @property
    def Accent(self):
        return self.calloutformat.Accent

    @Accent.setter
    def Accent(self, value):
        self.calloutformat.Accent = value

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

    def CustomDrop(self, Drop=None):
        arguments = com_arguments([Drop])
        self.calloutformat.CustomDrop(*arguments)

    def CustomLength(self, Length=None):
        arguments = com_arguments([Length])
        self.calloutformat.CustomLength(*arguments)

    def PresetDrop(self, DropType=None):
        arguments = com_arguments([DropType])
        self.calloutformat.PresetDrop(*arguments)


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

    def AddCallout(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Type, Left, Top, Width, Height])
        self.canvasshapes.AddCallout(*arguments)

    def AddConnector(self, Type=None, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = com_arguments([Type, BeginX, BeginY, EndX, EndY])
        self.canvasshapes.AddConnector(*arguments)

    def AddCurve(self, SafeArrayOfPoints=None):
        arguments = com_arguments([SafeArrayOfPoints])
        self.canvasshapes.AddCurve(*arguments)

    def AddLabel(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Orientation, Left, Top, Width, Height])
        self.canvasshapes.AddLabel(*arguments)

    def AddLine(self, BeginX=None, BeginY=None, EndX=None, EndY=None):
        arguments = com_arguments([BeginX, BeginY, EndX, EndY])
        self.canvasshapes.AddLine(*arguments)

    def AddPicture(self, FileName=None, LinkToFile=None, SaveWithDocument=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([FileName, LinkToFile, SaveWithDocument, Left, Top, Width, Height])
        self.canvasshapes.AddPicture(*arguments)

    def AddPolyline(self, SafeArrayOfPoints=None):
        arguments = com_arguments([SafeArrayOfPoints])
        self.canvasshapes.AddPolyline(*arguments)

    def AddShape(self, Type=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Type, Left, Top, Width, Height])
        self.canvasshapes.AddShape(*arguments)

    def AddTextbox(self, Orientation=None, Left=None, Top=None, Width=None, Height=None):
        arguments = com_arguments([Orientation, Left, Top, Width, Height])
        self.canvasshapes.AddTextbox(*arguments)

    def AddTextEffect(self, PresetTextEffect=None, Text=None, FontName=None, FontSize=None, FontBold=None, FontItalic=None, Left=None, Top=None):
        arguments = com_arguments([PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top])
        self.canvasshapes.AddTextEffect(*arguments)

    def BuildFreeform(self, EditingType=None, X1=None, Y1=None):
        arguments = com_arguments([EditingType, X1, Y1])
        self.canvasshapes.BuildFreeform(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.canvasshapes.Item(*arguments)

    def Range(self, Index=None):
        arguments = com_arguments([Index])
        return self.canvasshapes.Range(*arguments)

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

    @IncludeChapterNumber.setter
    def IncludeChapterNumber(self, value):
        self.captionlabel.IncludeChapterNumber = value

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.categories.Item(*arguments)


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

    @FitText.setter
    def FitText(self, value):
        self.cell.FitText = value

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

    @WordWrap.setter
    def WordWrap(self, value):
        self.cell.WordWrap = value

    def AutoSum(self):
        self.cell.AutoSum()

    def Delete(self, ShiftCells=None):
        arguments = com_arguments([ShiftCells])
        self.cell.Delete(*arguments)

    def Formula(self, Formula=None, NumFormat=None):
        arguments = com_arguments([Formula, NumFormat])
        self.cell.Formula(*arguments)

    def Merge(self, MergeTo=None):
        arguments = com_arguments([MergeTo])
        self.cell.Merge(*arguments)

    def Select(self):
        self.cell.Select()

    def SetHeight(self, RowHeight=None, HeightRule=None):
        arguments = com_arguments([RowHeight, HeightRule])
        self.cell.SetHeight(*arguments)

    def SetWidth(self, ColumnWidth=None, RulerStyle=None):
        arguments = com_arguments([ColumnWidth, RulerStyle])
        self.cell.SetWidth(*arguments)

    def Split(self, NumRows=None, NumColumns=None):
        arguments = com_arguments([NumRows, NumColumns])
        self.cell.Split(*arguments)


class Chart:

    def __init__(self, chart=None):
        self.chart = chart

    @property
    def Application(self):
        return self.chart.Application

    @property
    def AutoScaling(self):
        return self.chart.AutoScaling

    @AutoScaling.setter
    def AutoScaling(self, value):
        self.chart.AutoScaling = value

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

    def ChartGroups(self, Index=None):
        arguments = com_arguments([Index])
        if callable(self.chart.ChartGroups):
            return self.chart.ChartGroups(*arguments)
        else:
            return self.chart.GetChartGroups(*arguments)

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

    @HasDataTable.setter
    def HasDataTable(self, value):
        self.chart.HasDataTable = value

    @property
    def HasLegend(self):
        return self.chart.HasLegend

    @HasLegend.setter
    def HasLegend(self, value):
        self.chart.HasLegend = value

    @property
    def HasTitle(self):
        return self.chart.HasTitle

    @HasTitle.setter
    def HasTitle(self, value):
        self.chart.HasTitle = value

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

    @PlotVisibleOnly.setter
    def PlotVisibleOnly(self, value):
        self.chart.PlotVisibleOnly = value

    @property
    def RightAngleAxes(self):
        return self.chart.RightAngleAxes

    @RightAngleAxes.setter
    def RightAngleAxes(self, value):
        self.chart.RightAngleAxes = value

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

    def ApplyChartTemplate(self, FileName=None):
        arguments = com_arguments([FileName])
        self.chart.ApplyChartTemplate(*arguments)

    def ApplyDataLabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = com_arguments([Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator])
        self.chart.ApplyDataLabels(*arguments)

    def ApplyLayout(self, Layout=None, ChartType=None):
        arguments = com_arguments([Layout, ChartType])
        self.chart.ApplyLayout(*arguments)

    def Axes(self, Type=None, AxisGroup=None):
        arguments = com_arguments([Type, AxisGroup])
        return self.chart.Axes(*arguments)

    def ChartWizard(self, Source=None, Gallery=None, Format=None, PlotBy=None, CategoryLabels=None, SeriesLabels=None, HasLegend=None, Title=None, CategoryTitle=None, ValueTitle=None, ExtraTitle=None):
        arguments = com_arguments([Source, Gallery, Format, PlotBy, CategoryLabels, SeriesLabels, HasLegend, Title, CategoryTitle, ValueTitle, ExtraTitle])
        self.chart.ChartWizard(*arguments)

    def ClearToMatchStyle(self):
        self.chart.ClearToMatchStyle()

    def Copy(self, Before=None, After=None):
        arguments = com_arguments([Before, After])
        self.chart.Copy(*arguments)

    def CopyPicture(self, Appearance=None, Format=None, Size=None):
        arguments = com_arguments([Appearance, Format, Size])
        self.chart.CopyPicture(*arguments)

    def Delete(self):
        self.chart.Delete()

    def Export(self, FileName=None, FilterName=None, Interactive=None):
        arguments = com_arguments([FileName, FilterName, Interactive])
        return self.chart.Export(*arguments)

    def GetChartElement(self, x=None, y=None, ElementID=None, Arg1=None, Arg2=None):
        arguments = com_arguments([x, y, ElementID, Arg1, Arg2])
        self.chart.GetChartElement(*arguments)

    def Paste(self, Type=None):
        arguments = com_arguments([Type])
        self.chart.Paste(*arguments)

    def Refresh(self):
        self.chart.Refresh()

    def SaveChartTemplate(self, FileName=None):
        arguments = com_arguments([FileName])
        self.chart.SaveChartTemplate(*arguments)

    def Select(self, Replace=None):
        arguments = com_arguments([Replace])
        self.chart.Select(*arguments)

    def SeriesCollection(self, Index=None):
        arguments = com_arguments([Index])
        return SeriesCollection(self.chart.SeriesCollection(*arguments))

    def SetBackgroundPicture(self, FileName=None):
        arguments = com_arguments([FileName])
        self.chart.SetBackgroundPicture(*arguments)

    def SetDefaultChart(self, Name=None):
        arguments = com_arguments([Name])
        self.chart.SetDefaultChart(*arguments)

    def SetElement(self, Element=None):
        arguments = com_arguments([Element])
        self.chart.SetElement(*arguments)

    def SetSourceData(self, Source=None, PlotBy=None):
        arguments = com_arguments([Source, PlotBy])
        self.chart.SetSourceData(*arguments)


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

    def Insert(self, String=None):
        arguments = com_arguments([String])
        self.chartcharacters.Insert(*arguments)


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

    @Bold.setter
    def Bold(self, value):
        self.chartfont.Bold = value

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

    @Italic.setter
    def Italic(self, value):
        self.chartfont.Italic = value

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

    @StrikeThrough.setter
    def StrikeThrough(self, value):
        self.chartfont.StrikeThrough = value

    @property
    def Subscript(self):
        return self.chartfont.Subscript

    @Subscript.setter
    def Subscript(self, value):
        self.chartfont.Subscript = value

    @property
    def Superscript(self):
        return self.chartfont.Superscript

    @Superscript.setter
    def Superscript(self, value):
        self.chartfont.Superscript = value

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

    @AxisGroup.setter
    def AxisGroup(self, value):
        self.chartgroup.AxisGroup = value

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

    @GapWidth.setter
    def GapWidth(self, value):
        self.chartgroup.GapWidth = value

    @property
    def Has3DShading(self):
        return self.chartgroup.Has3DShading

    @Has3DShading.setter
    def Has3DShading(self, value):
        self.chartgroup.Has3DShading = value

    @property
    def HasDropLines(self):
        return self.chartgroup.HasDropLines

    @HasDropLines.setter
    def HasDropLines(self, value):
        self.chartgroup.HasDropLines = value

    @property
    def HasHiLoLines(self):
        return self.chartgroup.HasHiLoLines

    @HasHiLoLines.setter
    def HasHiLoLines(self, value):
        self.chartgroup.HasHiLoLines = value

    @property
    def HasRadarAxisLabels(self):
        return self.chartgroup.HasRadarAxisLabels

    @HasRadarAxisLabels.setter
    def HasRadarAxisLabels(self, value):
        self.chartgroup.HasRadarAxisLabels = value

    @property
    def HasSeriesLines(self):
        return self.chartgroup.HasSeriesLines

    @HasSeriesLines.setter
    def HasSeriesLines(self, value):
        self.chartgroup.HasSeriesLines = value

    @property
    def HasUpDownBars(self):
        return self.chartgroup.HasUpDownBars

    @HasUpDownBars.setter
    def HasUpDownBars(self, value):
        self.chartgroup.HasUpDownBars = value

    @property
    def HiLoLines(self):
        return HiLoLines(self.chartgroup.HiLoLines)

    @property
    def Index(self):
        return self.chartgroup.Index

    @property
    def Overlap(self):
        return self.chartgroup.Overlap

    @Overlap.setter
    def Overlap(self, value):
        self.chartgroup.Overlap = value

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

    @ShowNegativeBubbles.setter
    def ShowNegativeBubbles(self, value):
        self.chartgroup.ShowNegativeBubbles = value

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

    @VaryByCategories.setter
    def VaryByCategories(self, value):
        self.chartgroup.VaryByCategories = value

    def SeriesCollection(self, Index=None):
        arguments = com_arguments([Index])
        return SeriesCollection(self.chartgroup.SeriesCollection(*arguments))


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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return ChartGroup(self.chartgroups.Item(*arguments))


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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        if callable(self.charttitle.Characters):
            return ChartCharacters(self.charttitle.Characters(*arguments))
        else:
            return ChartCharacters(self.charttitle.GetCharacters(*arguments))

    @property
    def Creator(self):
        return self.charttitle.Creator

    @property
    def Format(self):
        return ChartFormat(self.charttitle.Format)

    @property
    def Formula(self):
        return self.charttitle.Formula

    @Formula.setter
    def Formula(self, value):
        self.charttitle.Formula = value

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

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.charttitle.IncludeInLayout = value

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

    @AutoSize.setter
    def AutoSize(self, value):
        self.checkbox.AutoSize = value

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

    @Value.setter
    def Value(self, value):
        self.checkbox.Value = value


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

    def Add(self, Range=None, Type=None):
        arguments = com_arguments([Range, Type])
        return CoAuthLock(self.coauthlocks.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return CoAuthLock(self.coauthlocks.Item(*arguments))

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.coauthors.Item(*arguments)


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

    @Brightness.setter
    def Brightness(self, value):
        self.colorformat.Brightness = value

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

    @TintAndShade.setter
    def TintAndShade(self, value):
        self.colorformat.TintAndShade = value

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

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([RowIndex, ColumnIndex])
        if callable(self.column.Cells):
            return self.column.Cells(*arguments)
        else:
            return self.column.GetCells(*arguments)

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

    def SetWidth(self, ColumnWidth=None, RulerStyle=None):
        arguments = com_arguments([ColumnWidth, RulerStyle])
        self.column.SetWidth(*arguments)

    def Sort(self, ExcludeHeader=None, SortFieldType=None, SortOrder=None, CaseSensitive=None, BidiSort=None, IgnoreThe=None, IgnoreKashida=None, IgnoreDiacritics=None, IgnoreHe=None, LanguageID=None):
        arguments = com_arguments([ExcludeHeader, SortFieldType, SortOrder, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID])
        self.column.Sort(*arguments)


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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.conflicts.Item(*arguments)

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

    @DateDisplayLocale.setter
    def DateDisplayLocale(self, value):
        self.contentcontrol.DateDisplayLocale = value

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

    @MultiLine.setter
    def MultiLine(self, value):
        self.contentcontrol.MultiLine = value

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

    def Delete(self, DeleteContents=None):
        arguments = com_arguments([DeleteContents])
        self.contentcontrol.Delete(*arguments)

    def SetCheckedSymbol(self, CharacterNumber=None, Font=None):
        arguments = com_arguments([CharacterNumber, Font])
        self.contentcontrol.SetCheckedSymbol(*arguments)

    def SetPlaceholderText(self, BuildingBlock=None, Range=None, Text=None):
        arguments = com_arguments([BuildingBlock, Range, Text])
        self.contentcontrol.SetPlaceholderText(*arguments)

    def SetUncheckedSymbol(self, CharacterNumber=None, Font=None):
        arguments = com_arguments([CharacterNumber, Font])
        self.contentcontrol.SetUncheckedSymbol(*arguments)

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

    def Add(self, Text=None, Value=None, Index=None):
        arguments = com_arguments([Text, Value, Index])
        return self.contentcontrollistentries.Add(*arguments)

    def Clear(self):
        self.contentcontrollistentries.Clear()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.contentcontrollistentries.Item(*arguments)


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

    def Add(self, Type=None, Range=None):
        arguments = com_arguments([Type, Range])
        return ContentControl(self.contentcontrols.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.contentcontrols.Item(*arguments)


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

    def Add(self, Name=None, Value=None):
        arguments = com_arguments([Name, Value])
        return CustomPropertie(self.customproperties.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.customproperties.Item(*arguments)


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

    @AutoText.setter
    def AutoText(self, value):
        self.datalabel.AutoText = value

    @property
    def Caption(self):
        return self.datalabel.Caption

    @Caption.setter
    def Caption(self, value):
        self.datalabel.Caption = value

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        if callable(self.datalabel.Characters):
            return ChartCharacters(self.datalabel.Characters(*arguments))
        else:
            return ChartCharacters(self.datalabel.GetCharacters(*arguments))

    @property
    def Creator(self):
        return self.datalabel.Creator

    @property
    def Format(self):
        return ChartFormat(self.datalabel.Format)

    @property
    def Formula(self):
        return self.datalabel.Formula

    @Formula.setter
    def Formula(self, value):
        self.datalabel.Formula = value

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

    @Height.setter
    def Height(self, value):
        self.datalabel.Height = value

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

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.datalabel.NumberFormatLinked = value

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

    @ShowBubbleSize.setter
    def ShowBubbleSize(self, value):
        self.datalabel.ShowBubbleSize = value

    @property
    def ShowCategoryName(self):
        return self.datalabel.ShowCategoryName

    @ShowCategoryName.setter
    def ShowCategoryName(self, value):
        self.datalabel.ShowCategoryName = value

    @property
    def ShowLegendKey(self):
        return self.datalabel.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.datalabel.ShowLegendKey = value

    @property
    def ShowPercentage(self):
        return self.datalabel.ShowPercentage

    @ShowPercentage.setter
    def ShowPercentage(self, value):
        self.datalabel.ShowPercentage = value

    @property
    def ShowSeriesName(self):
        return self.datalabel.ShowSeriesName

    @ShowSeriesName.setter
    def ShowSeriesName(self, value):
        self.datalabel.ShowSeriesName = value

    @property
    def ShowValue(self):
        return self.datalabel.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.datalabel.ShowValue = value

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

    @Width.setter
    def Width(self, value):
        self.datalabel.Width = value

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

    @AutoText.setter
    def AutoText(self, value):
        self.datalabels.AutoText = value

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

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.datalabels.NumberFormatLinked = value

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

    @Separator.setter
    def Separator(self, value):
        self.datalabels.Separator = value

    @property
    def Shadow(self):
        return self.datalabels.Shadow

    @Shadow.setter
    def Shadow(self, value):
        self.datalabels.Shadow = value

    @property
    def ShowBubbleSize(self):
        return self.datalabels.ShowBubbleSize

    @ShowBubbleSize.setter
    def ShowBubbleSize(self, value):
        self.datalabels.ShowBubbleSize = value

    @property
    def ShowCategoryName(self):
        return self.datalabels.ShowCategoryName

    @ShowCategoryName.setter
    def ShowCategoryName(self, value):
        self.datalabels.ShowCategoryName = value

    @property
    def ShowLegendKey(self):
        return self.datalabels.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.datalabels.ShowLegendKey = value

    @property
    def ShowPercentage(self):
        return self.datalabels.ShowPercentage

    @ShowPercentage.setter
    def ShowPercentage(self, value):
        self.datalabels.ShowPercentage = value

    @property
    def ShowSeriesName(self):
        return self.datalabels.ShowSeriesName

    @ShowSeriesName.setter
    def ShowSeriesName(self, value):
        self.datalabels.ShowSeriesName = value

    @property
    def ShowValue(self):
        return self.datalabels.ShowValue

    @ShowValue.setter
    def ShowValue(self, value):
        self.datalabels.ShowValue = value

    @property
    def VerticalAlignment(self):
        return self.datalabels.VerticalAlignment

    @VerticalAlignment.setter
    def VerticalAlignment(self, value):
        self.datalabels.VerticalAlignment = value

    def Delete(self):
        self.datalabels.Delete()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return DataLabel(self.datalabels.Item(*arguments))

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

    @HasBorderHorizontal.setter
    def HasBorderHorizontal(self, value):
        self.datatable.HasBorderHorizontal = value

    @property
    def HasBorderOutline(self):
        return self.datatable.HasBorderOutline

    @HasBorderOutline.setter
    def HasBorderOutline(self, value):
        self.datatable.HasBorderOutline = value

    @property
    def HasBorderVertical(self):
        return self.datatable.HasBorderVertical

    @HasBorderVertical.setter
    def HasBorderVertical(self, value):
        self.datatable.HasBorderVertical = value

    @property
    def Parent(self):
        return self.datatable.Parent

    @property
    def ShowLegendKey(self):
        return self.datatable.ShowLegendKey

    @ShowLegendKey.setter
    def ShowLegendKey(self, value):
        self.datatable.ShowLegendKey = value

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

    @AllowPNG.setter
    def AllowPNG(self, value):
        self.defaultweboptions.AllowPNG = value

    @property
    def AlwaysSaveInDefaultEncoding(self):
        return self.defaultweboptions.AlwaysSaveInDefaultEncoding

    @AlwaysSaveInDefaultEncoding.setter
    def AlwaysSaveInDefaultEncoding(self, value):
        self.defaultweboptions.AlwaysSaveInDefaultEncoding = value

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

    @CheckIfOfficeIsHTMLEditor.setter
    def CheckIfOfficeIsHTMLEditor(self, value):
        self.defaultweboptions.CheckIfOfficeIsHTMLEditor = value

    @property
    def CheckIfWordIsDefaultHTMLEditor(self):
        return self.defaultweboptions.CheckIfWordIsDefaultHTMLEditor

    @CheckIfWordIsDefaultHTMLEditor.setter
    def CheckIfWordIsDefaultHTMLEditor(self, value):
        self.defaultweboptions.CheckIfWordIsDefaultHTMLEditor = value

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

    @OptimizeForBrowser.setter
    def OptimizeForBrowser(self, value):
        self.defaultweboptions.OptimizeForBrowser = value

    @property
    def OrganizeInFolder(self):
        return self.defaultweboptions.OrganizeInFolder

    @OrganizeInFolder.setter
    def OrganizeInFolder(self, value):
        self.defaultweboptions.OrganizeInFolder = value

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

    @RelyOnCSS.setter
    def RelyOnCSS(self, value):
        self.defaultweboptions.RelyOnCSS = value

    @property
    def RelyOnVML(self):
        return self.defaultweboptions.RelyOnVML

    @RelyOnVML.setter
    def RelyOnVML(self, value):
        self.defaultweboptions.RelyOnVML = value

    @property
    def SaveNewWebPagesAsWebArchives(self):
        return self.defaultweboptions.SaveNewWebPagesAsWebArchives

    @SaveNewWebPagesAsWebArchives.setter
    def SaveNewWebPagesAsWebArchives(self, value):
        self.defaultweboptions.SaveNewWebPagesAsWebArchives = value

    @property
    def ScreenSize(self):
        return self.defaultweboptions.ScreenSize

    @ScreenSize.setter
    def ScreenSize(self, value):
        self.defaultweboptions.ScreenSize = value

    @property
    def TargetBrowser(self):
        return self.defaultweboptions.TargetBrowser

    @TargetBrowser.setter
    def TargetBrowser(self, value):
        self.defaultweboptions.TargetBrowser = value

    @property
    def UpdateLinksOnSave(self):
        return self.defaultweboptions.UpdateLinksOnSave

    @UpdateLinksOnSave.setter
    def UpdateLinksOnSave(self, value):
        self.defaultweboptions.UpdateLinksOnSave = value

    @property
    def UseLongFileNames(self):
        return self.defaultweboptions.UseLongFileNames

    @UseLongFileNames.setter
    def UseLongFileNames(self, value):
        self.defaultweboptions.UseLongFileNames = value


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

    def Display(self, TimeOut=None):
        arguments = com_arguments([TimeOut])
        return self.dialog.Display(*arguments)

    def Execute(self):
        self.dialog.Execute()

    def Show(self, TimeOut=None):
        arguments = com_arguments([TimeOut])
        return self.dialog.Show(*arguments)

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

    @LanguageSpecific.setter
    def LanguageSpecific(self, value):
        self.dictionary.LanguageSpecific = value

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

    def Characters(self, Start=None, Length=None):
        arguments = com_arguments([Start, Length])
        if callable(self.displayunitlabel.Characters):
            return ChartCharacters(self.displayunitlabel.Characters(*arguments))
        else:
            return ChartCharacters(self.displayunitlabel.GetCharacters(*arguments))

    @property
    def Creator(self):
        return self.displayunitlabel.Creator

    @property
    def Format(self):
        return ChartFormat(self.displayunitlabel.Format)

    @property
    def Formula(self):
        return self.displayunitlabel.Formula

    @Formula.setter
    def Formula(self, value):
        self.displayunitlabel.Formula = value

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

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.displayunitlabel.IncludeInLayout = value

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

    @AttachedTemplate.setter
    def AttachedTemplate(self, value):
        self.document.AttachedTemplate = value

    @property
    def AutoFormatOverride(self):
        return self.document.AutoFormatOverride

    @AutoFormatOverride.setter
    def AutoFormatOverride(self, value):
        self.document.AutoFormatOverride = value

    @property
    def AutoHyphenation(self):
        return self.document.AutoHyphenation

    @AutoHyphenation.setter
    def AutoHyphenation(self, value):
        self.document.AutoHyphenation = value

    @property
    def AutoSaveOn(self):
        return self.document.AutoSaveOn

    @AutoSaveOn.setter
    def AutoSaveOn(self, value):
        self.document.AutoSaveOn = value

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

    @property
    def Compatibility(self):
        return self.document.Compatibility

    @Compatibility.setter
    def Compatibility(self, value):
        self.document.Compatibility = value

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

    @DisableFeatures.setter
    def DisableFeatures(self, value):
        self.document.DisableFeatures = value

    @property
    def DisableFeaturesIntroducedAfter(self):
        return self.document.DisableFeaturesIntroducedAfter

    @DisableFeaturesIntroducedAfter.setter
    def DisableFeaturesIntroducedAfter(self, value):
        self.document.DisableFeaturesIntroducedAfter = value

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

    @DoNotEmbedSystemFonts.setter
    def DoNotEmbedSystemFonts(self, value):
        self.document.DoNotEmbedSystemFonts = value

    @property
    def Email(self):
        return Email(self.document.Email)

    @property
    def EmbedLinguisticData(self):
        return self.document.EmbedLinguisticData

    @EmbedLinguisticData.setter
    def EmbedLinguisticData(self, value):
        self.document.EmbedLinguisticData = value

    @property
    def EmbedTrueTypeFonts(self):
        return self.document.EmbedTrueTypeFonts

    @EmbedTrueTypeFonts.setter
    def EmbedTrueTypeFonts(self, value):
        self.document.EmbedTrueTypeFonts = value

    @property
    def EncryptionProvider(self):
        return self.document.EncryptionProvider

    @EncryptionProvider.setter
    def EncryptionProvider(self, value):
        self.document.EncryptionProvider = value

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

    @FormattingShowClear.setter
    def FormattingShowClear(self, value):
        self.document.FormattingShowClear = value

    @property
    def FormattingShowFilter(self):
        return self.document.FormattingShowFilter

    @FormattingShowFilter.setter
    def FormattingShowFilter(self, value):
        self.document.FormattingShowFilter = value

    @property
    def FormattingShowFont(self):
        return self.document.FormattingShowFont

    @FormattingShowFont.setter
    def FormattingShowFont(self, value):
        self.document.FormattingShowFont = value

    @property
    def FormattingShowNextLevel(self):
        return self.document.FormattingShowNextLevel

    @FormattingShowNextLevel.setter
    def FormattingShowNextLevel(self, value):
        self.document.FormattingShowNextLevel = value

    @property
    def FormattingShowNumbering(self):
        return self.document.FormattingShowNumbering

    @FormattingShowNumbering.setter
    def FormattingShowNumbering(self, value):
        self.document.FormattingShowNumbering = value

    @property
    def FormattingShowParagraph(self):
        return self.document.FormattingShowParagraph

    @FormattingShowParagraph.setter
    def FormattingShowParagraph(self, value):
        self.document.FormattingShowParagraph = value

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

    @GrammarChecked.setter
    def GrammarChecked(self, value):
        self.document.GrammarChecked = value

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

    @GridOriginFromMargin.setter
    def GridOriginFromMargin(self, value):
        self.document.GridOriginFromMargin = value

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

    @HyphenateCaps.setter
    def HyphenateCaps(self, value):
        self.document.HyphenateCaps = value

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

    @KerningByAlgorithm.setter
    def KerningByAlgorithm(self, value):
        self.document.KerningByAlgorithm = value

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

    @OptimizeForWord97.setter
    def OptimizeForWord97(self, value):
        self.document.OptimizeForWord97 = value

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

    @PrintFormsData.setter
    def PrintFormsData(self, value):
        self.document.PrintFormsData = value

    @property
    def PrintPostScriptOverText(self):
        return self.document.PrintPostScriptOverText

    @PrintPostScriptOverText.setter
    def PrintPostScriptOverText(self, value):
        self.document.PrintPostScriptOverText = value

    @property
    def PrintRevisions(self):
        return self.document.PrintRevisions

    @PrintRevisions.setter
    def PrintRevisions(self, value):
        self.document.PrintRevisions = value

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

    @ReadOnlyRecommended.setter
    def ReadOnlyRecommended(self, value):
        self.document.ReadOnlyRecommended = value

    @property
    def RemoveDateAndTime(self):
        return self.document.RemoveDateAndTime

    @property
    def RemovePersonalInformation(self):
        return self.document.RemovePersonalInformation

    @RemovePersonalInformation.setter
    def RemovePersonalInformation(self, value):
        self.document.RemovePersonalInformation = value

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

    @Saved.setter
    def Saved(self, value):
        self.document.Saved = value

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

    @SaveFormsData.setter
    def SaveFormsData(self, value):
        self.document.SaveFormsData = value

    @property
    def SaveSubsetFonts(self):
        return self.document.SaveSubsetFonts

    @SaveSubsetFonts.setter
    def SaveSubsetFonts(self, value):
        self.document.SaveSubsetFonts = value

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

    @ShowGrammaticalErrors.setter
    def ShowGrammaticalErrors(self, value):
        self.document.ShowGrammaticalErrors = value

    @property
    def ShowSpellingErrors(self):
        return self.document.ShowSpellingErrors

    @ShowSpellingErrors.setter
    def ShowSpellingErrors(self, value):
        self.document.ShowSpellingErrors = value

    @property
    def Signatures(self):
        return self.document.Signatures

    @property
    def SmartDocument(self):
        return self.document.SmartDocument

    @property
    def SnapToGrid(self):
        return self.document.SnapToGrid

    @SnapToGrid.setter
    def SnapToGrid(self, value):
        self.document.SnapToGrid = value

    @property
    def SnapToShapes(self):
        return self.document.SnapToShapes

    @SnapToShapes.setter
    def SnapToShapes(self, value):
        self.document.SnapToShapes = value

    @property
    def SpellingChecked(self):
        return self.document.SpellingChecked

    @SpellingChecked.setter
    def SpellingChecked(self, value):
        self.document.SpellingChecked = value

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

    @TrackRevisions.setter
    def TrackRevisions(self, value):
        self.document.TrackRevisions = value

    @property
    def Type(self):
        return WdDocumentType(self.document.Type)

    @property
    def UpdateStylesOnOpen(self):
        return self.document.UpdateStylesOnOpen

    @UpdateStylesOnOpen.setter
    def UpdateStylesOnOpen(self, value):
        self.document.UpdateStylesOnOpen = value

    @property
    def UseMathDefaults(self):
        return self.document.UseMathDefaults

    @UseMathDefaults.setter
    def UseMathDefaults(self, value):
        self.document.UseMathDefaults = value

    @property
    def UserControl(self):
        return self.document.UserControl

    @UserControl.setter
    def UserControl(self, value):
        self.document.UserControl = value

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

    def ApplyQuickStyleSet2(self, Style=None):
        arguments = com_arguments([Style])
        self.document.ApplyQuickStyleSet2(*arguments)

    def ApplyTheme(self, Name=None):
        arguments = com_arguments([Name])
        self.document.ApplyTheme(*arguments)

    def AutoFormat(self):
        self.document.AutoFormat()

    def CanCheckin(self):
        return self.document.CanCheckin()

    def CheckConsistency(self):
        self.document.CheckConsistency()

    def CheckGrammar(self):
        self.document.CheckGrammar()

    def CheckIn(self, SaveChanges=None, Comments=None, MakePublic=None):
        arguments = com_arguments([SaveChanges, Comments, MakePublic])
        self.document.CheckIn(*arguments)

    def CheckInWithVersion(self, SaveChanges=None, Comments=None, MakePublic=None, VersionType=None):
        arguments = com_arguments([SaveChanges, Comments, MakePublic, VersionType])
        self.document.CheckInWithVersion(*arguments)

    def CheckSpelling(self, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = com_arguments([CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10])
        self.document.CheckSpelling(*arguments)

    def Close(self, SaveChanges=None, OriginalFormat=None, RouteDocument=None):
        arguments = com_arguments([SaveChanges, OriginalFormat, RouteDocument])
        self.document.Close(*arguments)

    def ClosePrintPreview(self):
        self.document.ClosePrintPreview()

    def Compare(self, Name=None, AuthorName=None, CompareTarget=None, DetectFormatChanges=None, IgnoreAllComparisonWarnings=None, AddToRecentFiles=None, RemovePersonalInformation=None, RemoveDateAndTime=None):
        arguments = com_arguments([Name, AuthorName, CompareTarget, DetectFormatChanges, IgnoreAllComparisonWarnings, AddToRecentFiles, RemovePersonalInformation, RemoveDateAndTime])
        self.document.Compare(*arguments)

    def ComputeStatistics(self, Statistic=None, IncludeFootnotesAndEndnotes=None):
        arguments = com_arguments([Statistic, IncludeFootnotesAndEndnotes])
        self.document.ComputeStatistics(*arguments)

    def Convert(self):
        self.document.Convert()

    def ConvertAutoHyphens(self):
        self.document.ConvertAutoHyphens()

    def ConvertNumbersToText(self):
        self.document.ConvertNumbersToText()

    def ConvertVietDoc(self, CodePageOrigin=None):
        arguments = com_arguments([CodePageOrigin])
        self.document.ConvertVietDoc(*arguments)

    def CopyStylesFromTemplate(self, Template=None):
        arguments = com_arguments([Template])
        self.document.CopyStylesFromTemplate(*arguments)

    def CountNumberedItems(self, NumberType=None, Level=None):
        arguments = com_arguments([NumberType, Level])
        self.document.CountNumberedItems(*arguments)

    def CreateLetterContent(self, DateFormat=None, IncludeHeaderFooter=None, PageDesign=None, LetterStyle=None, Letterhead=None, LetterheadLocation=None, LetterheadSize=None, RecipientName=None, RecipientAddress=None, Salutation=None, SalutationType=None, RecipientReference=None, MailingInstructions=None, AttentionLine=None, Subject=None, CCList=None, ReturnAddress=None, SenderName=None, Closing=None, SenderCompany=None, SenderJobTitle=None, SenderInitials=None, EnclosureNumber=None, InfoBlock=None, RecipientCode=None, RecipientGender=None, ReturnAddressShortForm=None, SenderCity=None, SenderCode=None, SenderGender=None, SenderReference=None):
        arguments = com_arguments([DateFormat, IncludeHeaderFooter, PageDesign, LetterStyle, Letterhead, LetterheadLocation, LetterheadSize, RecipientName, RecipientAddress, Salutation, SalutationType, RecipientReference, MailingInstructions, AttentionLine, Subject, CCList, ReturnAddress, SenderName, Closing, SenderCompany, SenderJobTitle, SenderInitials, EnclosureNumber, InfoBlock, RecipientCode, RecipientGender, ReturnAddressShortForm, SenderCity, SenderCode, SenderGender, SenderReference])
        return self.document.CreateLetterContent(*arguments)

    def DataForm(self):
        self.document.DataForm()

    def DeleteAllComments(self):
        self.document.DeleteAllComments()

    def DeleteAllCommentsShown(self):
        self.document.DeleteAllCommentsShown()

    def DeleteAllEditableRanges(self, EditorID=None):
        arguments = com_arguments([EditorID])
        self.document.DeleteAllEditableRanges(*arguments)

    def DeleteAllInkAnnotations(self):
        self.document.DeleteAllInkAnnotations()

    def DetectLanguage(self):
        self.document.DetectLanguage()

    def DowngradeDocument(self):
        self.document.DowngradeDocument()

    def EndReview(self):
        self.document.EndReview()

    def ExportAsFixedFormat(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, Range=None, From=None, To=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr])
        self.document.ExportAsFixedFormat(*arguments)

    def ExportAsFixedFormat2(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, Range=None, From=None, To=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, OptimizeForImageQuality=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, FixedFormatExtClassPtr])
        self.document.ExportAsFixedFormat2(*arguments)

    def ExportAsFixedFormat3(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, Range=None, From=None, To=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, OptimizeForImageQuality=None, ImproveExportTagging=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, Range, From, To, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, ImproveExportTagging, FixedFormatExtClassPtr])
        self.document.ExportAsFixedFormat3(*arguments)

    def FitToPages(self):
        self.document.FitToPages()

    def FollowHyperlink(self, Address=None, SubAddress=None, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = com_arguments([Address, SubAddress, NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo])
        self.document.FollowHyperlink(*arguments)

    def FreezeLayout(self):
        self.document.FreezeLayout()

    def GetCrossReferenceItems(self, ReferenceType=None):
        arguments = com_arguments([ReferenceType])
        self.document.GetCrossReferenceItems(*arguments)

    def GetLetterContent(self):
        return self.document.GetLetterContent()

    def GetWorkflowTasks(self):
        return self.document.GetWorkflowTasks()

    def GetWorkflowTemplates(self):
        return self.document.GetWorkflowTemplates()

    def GoTo(self, What=None, Which=None, Count=None, Name=None):
        arguments = com_arguments([What, Which, Count, Name])
        self.document.GoTo(*arguments)

    def LockServerFile(self):
        self.document.LockServerFile()

    def MakeCompatibilityDefault(self):
        self.document.MakeCompatibilityDefault()

    def ManualHyphenation(self):
        self.document.ManualHyphenation()

    def Merge(self, Name=None, MergeTarget=None, DetectFormatChanges=None, UseFormattingFrom=None, AddToRecentFiles=None):
        arguments = com_arguments([Name, MergeTarget, DetectFormatChanges, UseFormattingFrom, AddToRecentFiles])
        self.document.Merge(*arguments)

    def Post(self):
        self.document.Post()

    def PresentIt(self):
        self.document.PresentIt()

    def PrintOut(self, Background=None, Append=None, Range=None, OutputFileName=None, From=None, To=None, Item=None, Copies=None, Pages=None, PageType=None, PrintToFile=None, Collate=None, FileName=None, ActivePrinterMacGX=None, ManualDuplexPrint=None, PrintZoomColumn=None, PrintZoomRow=None, PrintZoomPaperWidth=None, PrintZoomPaperHeight=None):
        arguments = com_arguments([Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight])
        self.document.PrintOut(*arguments)

    def PrintPreview(self):
        self.document.PrintPreview()

    def Range(self, Start=None, End=None):
        arguments = com_arguments([Start, End])
        return self.document.Range(*arguments)

    def Redo(self, Times=None):
        arguments = com_arguments([Times])
        return self.document.Redo(*arguments)

    def RejectAllRevisions(self):
        self.document.RejectAllRevisions()

    def RejectAllRevisionsShown(self):
        self.document.RejectAllRevisionsShown()

    def Reload(self):
        self.document.Reload()

    def ReloadAs(self, Encoding=None):
        arguments = com_arguments([Encoding])
        self.document.ReloadAs(*arguments)

    def RemoveDocumentInformation(self, RemoveDocInfoType=None):
        arguments = com_arguments([RemoveDocInfoType])
        self.document.RemoveDocumentInformation(*arguments)

    def RemoveLockedStyles(self):
        self.document.RemoveLockedStyles()

    def RemoveNumbers(self, NumberType=None):
        arguments = com_arguments([NumberType])
        self.document.RemoveNumbers(*arguments)

    def RemoveTheme(self):
        self.document.RemoveTheme()

    def Repaginate(self):
        self.document.Repaginate()

    def ReplyWithChanges(self, ShowMessage=None):
        arguments = com_arguments([ShowMessage])
        self.document.ReplyWithChanges(*arguments)

    def ResetFormFields(self):
        self.document.ResetFormFields()

    def RunAutoMacro(self, Which=None):
        arguments = com_arguments([Which])
        self.document.RunAutoMacro(*arguments)

    def RunLetterWizard(self, LetterContent=None, WizardMode=None):
        arguments = com_arguments([LetterContent, WizardMode])
        self.document.RunLetterWizard(*arguments)

    def Save(self):
        self.document.Save()

    def SaveAsQuickStyleSet(self, FileName=None):
        arguments = com_arguments([FileName])
        self.document.SaveAsQuickStyleSet(*arguments)

    def Select(self):
        self.document.Select()

    def SelectAllEditableRanges(self, EditorID=None):
        arguments = com_arguments([EditorID])
        self.document.SelectAllEditableRanges(*arguments)

    def SelectContentControlsByTag(self, Tag=None):
        arguments = com_arguments([Tag])
        return self.document.SelectContentControlsByTag(*arguments)

    def SelectContentControlsByTitle(self, Title=None):
        arguments = com_arguments([Title])
        return self.document.SelectContentControlsByTitle(*arguments)

    def SelectLinkedControls(self, Node=None):
        arguments = com_arguments([Node])
        return self.document.SelectLinkedControls(*arguments)

    def SelectNodes(self, XPath=None, PrefixMapping=None, FastSearchSkippingTextNodes=None):
        arguments = com_arguments([XPath, PrefixMapping, FastSearchSkippingTextNodes])
        return self.document.SelectNodes(*arguments)

    def SelectSingleNode(self, XPath=None, PrefixMapping=None, FastSearchSkippingTextNodes=None):
        arguments = com_arguments([XPath, PrefixMapping, FastSearchSkippingTextNodes])
        return self.document.SelectSingleNode(*arguments)

    def SelectUnlinkedControls(self, Stream=None):
        arguments = com_arguments([Stream])
        return self.document.SelectUnlinkedControls(*arguments)

    def SendFax(self, Address=None, Subject=None):
        arguments = com_arguments([Address, Subject])
        self.document.SendFax(*arguments)

    def SendFaxOverInternet(self, Recipients=None, Subject=None, ShowMessage=None):
        arguments = com_arguments([Recipients, Subject, ShowMessage])
        self.document.SendFaxOverInternet(*arguments)

    def SendForReview(self, Recipients=None, Subject=None, ShowMessage=None, IncludeAttachment=None):
        arguments = com_arguments([Recipients, Subject, ShowMessage, IncludeAttachment])
        self.document.SendForReview(*arguments)

    def SendMail(self):
        self.document.SendMail()

    def SetDefaultTableStyle(self, Style=None, SetInTemplate=None):
        arguments = com_arguments([Style, SetInTemplate])
        self.document.SetDefaultTableStyle(*arguments)

    def SetLetterContent(self, LetterContent=None):
        arguments = com_arguments([LetterContent])
        self.document.SetLetterContent(*arguments)

    def SetPasswordEncryptionOptions(self, PasswordEncryptionProvider=None, PasswordEncryptionAlgorithm=None, PasswordEncryptionKeyLength=None, PasswordEncryptionFileProperties=None):
        arguments = com_arguments([PasswordEncryptionProvider, PasswordEncryptionAlgorithm, PasswordEncryptionKeyLength, PasswordEncryptionFileProperties])
        self.document.SetPasswordEncryptionOptions(*arguments)

    def ToggleFormsDesign(self):
        self.document.ToggleFormsDesign()

    def TransformDocument(self, Path=None, DataOnly=None):
        arguments = com_arguments([Path, DataOnly])
        self.document.TransformDocument(*arguments)

    def Undo(self, Times=None):
        arguments = com_arguments([Times])
        return self.document.Undo(*arguments)

    def UndoClear(self):
        self.document.UndoClear()

    def Unprotect(self, Password=None):
        arguments = com_arguments([Password])
        self.document.Unprotect(*arguments)

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

    def Add(self, EditorID=None):
        arguments = com_arguments([EditorID])
        self.editors.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.editors.Item(*arguments)


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

    @AutoFormatAsYouTypeApplyBorders.setter
    def AutoFormatAsYouTypeApplyBorders(self, value):
        self.emailoptions.AutoFormatAsYouTypeApplyBorders = value

    @property
    def AutoFormatAsYouTypeApplyBulletedLists(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyBulletedLists

    @AutoFormatAsYouTypeApplyBulletedLists.setter
    def AutoFormatAsYouTypeApplyBulletedLists(self, value):
        self.emailoptions.AutoFormatAsYouTypeApplyBulletedLists = value

    @property
    def AutoFormatAsYouTypeApplyClosings(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyClosings

    @AutoFormatAsYouTypeApplyClosings.setter
    def AutoFormatAsYouTypeApplyClosings(self, value):
        self.emailoptions.AutoFormatAsYouTypeApplyClosings = value

    @property
    def AutoFormatAsYouTypeApplyDates(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyDates

    @AutoFormatAsYouTypeApplyDates.setter
    def AutoFormatAsYouTypeApplyDates(self, value):
        self.emailoptions.AutoFormatAsYouTypeApplyDates = value

    @property
    def AutoFormatAsYouTypeApplyFirstIndents(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyFirstIndents

    @AutoFormatAsYouTypeApplyFirstIndents.setter
    def AutoFormatAsYouTypeApplyFirstIndents(self, value):
        self.emailoptions.AutoFormatAsYouTypeApplyFirstIndents = value

    @property
    def AutoFormatAsYouTypeApplyHeadings(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyHeadings

    @AutoFormatAsYouTypeApplyHeadings.setter
    def AutoFormatAsYouTypeApplyHeadings(self, value):
        self.emailoptions.AutoFormatAsYouTypeApplyHeadings = value

    @property
    def AutoFormatAsYouTypeApplyNumberedLists(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyNumberedLists

    @AutoFormatAsYouTypeApplyNumberedLists.setter
    def AutoFormatAsYouTypeApplyNumberedLists(self, value):
        self.emailoptions.AutoFormatAsYouTypeApplyNumberedLists = value

    @property
    def AutoFormatAsYouTypeApplyTables(self):
        return self.emailoptions.AutoFormatAsYouTypeApplyTables

    @AutoFormatAsYouTypeApplyTables.setter
    def AutoFormatAsYouTypeApplyTables(self, value):
        self.emailoptions.AutoFormatAsYouTypeApplyTables = value

    @property
    def AutoFormatAsYouTypeAutoLetterWizard(self):
        return self.emailoptions.AutoFormatAsYouTypeAutoLetterWizard

    @AutoFormatAsYouTypeAutoLetterWizard.setter
    def AutoFormatAsYouTypeAutoLetterWizard(self, value):
        self.emailoptions.AutoFormatAsYouTypeAutoLetterWizard = value

    @property
    def AutoFormatAsYouTypeDefineStyles(self):
        return self.emailoptions.AutoFormatAsYouTypeDefineStyles

    @AutoFormatAsYouTypeDefineStyles.setter
    def AutoFormatAsYouTypeDefineStyles(self, value):
        self.emailoptions.AutoFormatAsYouTypeDefineStyles = value

    @property
    def AutoFormatAsYouTypeDeleteAutoSpaces(self):
        return self.emailoptions.AutoFormatAsYouTypeDeleteAutoSpaces

    @AutoFormatAsYouTypeDeleteAutoSpaces.setter
    def AutoFormatAsYouTypeDeleteAutoSpaces(self, value):
        self.emailoptions.AutoFormatAsYouTypeDeleteAutoSpaces = value

    @property
    def AutoFormatAsYouTypeFormatListItemBeginning(self):
        return self.emailoptions.AutoFormatAsYouTypeFormatListItemBeginning

    @AutoFormatAsYouTypeFormatListItemBeginning.setter
    def AutoFormatAsYouTypeFormatListItemBeginning(self, value):
        self.emailoptions.AutoFormatAsYouTypeFormatListItemBeginning = value

    @property
    def AutoFormatAsYouTypeInsertClosings(self):
        return self.emailoptions.AutoFormatAsYouTypeInsertClosings

    @AutoFormatAsYouTypeInsertClosings.setter
    def AutoFormatAsYouTypeInsertClosings(self, value):
        self.emailoptions.AutoFormatAsYouTypeInsertClosings = value

    @property
    def AutoFormatAsYouTypeInsertOvers(self):
        return self.emailoptions.AutoFormatAsYouTypeInsertOvers

    @AutoFormatAsYouTypeInsertOvers.setter
    def AutoFormatAsYouTypeInsertOvers(self, value):
        self.emailoptions.AutoFormatAsYouTypeInsertOvers = value

    @property
    def AutoFormatAsYouTypeMatchParentheses(self):
        return self.emailoptions.AutoFormatAsYouTypeMatchParentheses

    @AutoFormatAsYouTypeMatchParentheses.setter
    def AutoFormatAsYouTypeMatchParentheses(self, value):
        self.emailoptions.AutoFormatAsYouTypeMatchParentheses = value

    @property
    def AutoFormatAsYouTypeReplaceFarEastDashes(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceFarEastDashes

    @AutoFormatAsYouTypeReplaceFarEastDashes.setter
    def AutoFormatAsYouTypeReplaceFarEastDashes(self, value):
        self.emailoptions.AutoFormatAsYouTypeReplaceFarEastDashes = value

    @property
    def AutoFormatAsYouTypeReplaceFractions(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceFractions

    @AutoFormatAsYouTypeReplaceFractions.setter
    def AutoFormatAsYouTypeReplaceFractions(self, value):
        self.emailoptions.AutoFormatAsYouTypeReplaceFractions = value

    @property
    def AutoFormatAsYouTypeReplaceHyperlinks(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceHyperlinks

    @AutoFormatAsYouTypeReplaceHyperlinks.setter
    def AutoFormatAsYouTypeReplaceHyperlinks(self, value):
        self.emailoptions.AutoFormatAsYouTypeReplaceHyperlinks = value

    @property
    def AutoFormatAsYouTypeReplaceOrdinals(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceOrdinals

    @AutoFormatAsYouTypeReplaceOrdinals.setter
    def AutoFormatAsYouTypeReplaceOrdinals(self, value):
        self.emailoptions.AutoFormatAsYouTypeReplaceOrdinals = value

    @property
    def AutoFormatAsYouTypeReplacePlainTextEmphasis(self):
        return self.emailoptions.AutoFormatAsYouTypeReplacePlainTextEmphasis

    @AutoFormatAsYouTypeReplacePlainTextEmphasis.setter
    def AutoFormatAsYouTypeReplacePlainTextEmphasis(self, value):
        self.emailoptions.AutoFormatAsYouTypeReplacePlainTextEmphasis = value

    @property
    def AutoFormatAsYouTypeReplaceQuotes(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceQuotes

    @AutoFormatAsYouTypeReplaceQuotes.setter
    def AutoFormatAsYouTypeReplaceQuotes(self, value):
        self.emailoptions.AutoFormatAsYouTypeReplaceQuotes = value

    @property
    def AutoFormatAsYouTypeReplaceSymbols(self):
        return self.emailoptions.AutoFormatAsYouTypeReplaceSymbols

    @AutoFormatAsYouTypeReplaceSymbols.setter
    def AutoFormatAsYouTypeReplaceSymbols(self, value):
        self.emailoptions.AutoFormatAsYouTypeReplaceSymbols = value

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

    @HTMLFidelity.setter
    def HTMLFidelity(self, value):
        self.emailoptions.HTMLFidelity = value

    @property
    def MarkComments(self):
        return self.emailoptions.MarkComments

    @MarkComments.setter
    def MarkComments(self, value):
        self.emailoptions.MarkComments = value

    @property
    def MarkCommentsWith(self):
        return self.emailoptions.MarkCommentsWith

    @MarkCommentsWith.setter
    def MarkCommentsWith(self, value):
        self.emailoptions.MarkCommentsWith = value

    @property
    def NewColorOnReply(self):
        return self.emailoptions.NewColorOnReply

    @NewColorOnReply.setter
    def NewColorOnReply(self, value):
        self.emailoptions.NewColorOnReply = value

    @property
    def Parent(self):
        return self.emailoptions.Parent

    @property
    def PlainTextStyle(self):
        return Style(self.emailoptions.PlainTextStyle)

    @property
    def RelyOnCSS(self):
        return self.emailoptions.RelyOnCSS

    @RelyOnCSS.setter
    def RelyOnCSS(self, value):
        self.emailoptions.RelyOnCSS = value

    @property
    def ReplyStyle(self):
        return Style(self.emailoptions.ReplyStyle)

    @property
    def TabIndentKey(self):
        return self.emailoptions.TabIndentKey

    @TabIndentKey.setter
    def TabIndentKey(self, value):
        self.emailoptions.TabIndentKey = value

    @property
    def ThemeName(self):
        return self.emailoptions.ThemeName

    @ThemeName.setter
    def ThemeName(self, value):
        self.emailoptions.ThemeName = value

    @property
    def UseThemeStyle(self):
        return self.emailoptions.UseThemeStyle

    @UseThemeStyle.setter
    def UseThemeStyle(self, value):
        self.emailoptions.UseThemeStyle = value

    @property
    def UseThemeStyleOnReply(self):
        return self.emailoptions.UseThemeStyleOnReply

    @UseThemeStyleOnReply.setter
    def UseThemeStyleOnReply(self, value):
        self.emailoptions.UseThemeStyleOnReply = value


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

    def Add(self, Name=None, Range=None):
        arguments = com_arguments([Name, Range])
        return EmailSignatureEntrie(self.emailsignatureentries.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.emailsignatureentries.Item(*arguments)


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

    @DefaultFaceUp.setter
    def DefaultFaceUp(self, value):
        self.envelope.DefaultFaceUp = value

    @property
    def DefaultHeight(self):
        return self.envelope.DefaultHeight

    @DefaultHeight.setter
    def DefaultHeight(self, value):
        self.envelope.DefaultHeight = value

    @property
    def DefaultOmitReturnAddress(self):
        return self.envelope.DefaultOmitReturnAddress

    @DefaultOmitReturnAddress.setter
    def DefaultOmitReturnAddress(self, value):
        self.envelope.DefaultOmitReturnAddress = value

    @property
    def DefaultOrientation(self):
        return WdEnvelopeOrientation(self.envelope.DefaultOrientation)

    @DefaultOrientation.setter
    def DefaultOrientation(self, value):
        self.envelope.DefaultOrientation = value

    @property
    def DefaultPrintFIMA(self):
        return self.envelope.DefaultPrintFIMA

    @DefaultPrintFIMA.setter
    def DefaultPrintFIMA(self, value):
        self.envelope.DefaultPrintFIMA = value

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

    @Vertical.setter
    def Vertical(self, value):
        self.envelope.Vertical = value

    def Insert(self, ExtractAddress=None, Address=None, AutoText=None, OmitReturnAddress=None, ReturnAddress=None, ReturnAutoText=None, PrintBarCode=None, PrintFIMA=None, Size=None, Height=None, Width=None, FeedSource=None, AddressFromLeft=None, AddressFromTop=None, ReturnAddressFromLeft=None, ReturnAddressFromTop=None, DefaultFaceUp=None, DefaultOrientation=None, PrintEPostage=None, Vertical=None, RecipientNamefromLeft=None, RecipientNamefromTop=None, RecipientPostalfromLeft=None, RecipientPostalfromTop=None, SenderNamefromLeft=None, SenderNamefromTop=None, SenderPostalfromLeft=None, SenderPostalfromTop=None):
        arguments = com_arguments([ExtractAddress, Address, AutoText, OmitReturnAddress, ReturnAddress, ReturnAutoText, PrintBarCode, PrintFIMA, Size, Height, Width, FeedSource, AddressFromLeft, AddressFromTop, ReturnAddressFromLeft, ReturnAddressFromTop, DefaultFaceUp, DefaultOrientation, PrintEPostage, Vertical, RecipientNamefromLeft, RecipientNamefromTop, RecipientPostalfromLeft, RecipientPostalfromTop, SenderNamefromLeft, SenderNamefromTop, SenderPostalfromLeft, SenderPostalfromTop])
        self.envelope.Insert(*arguments)

    def Options(self):
        self.envelope.Options()

    def PrintOut(self, ExtractAddress=None, Address=None, AutoText=None, OmitReturnAddress=None, ReturnAddress=None, ReturnAutoText=None, PrintBarCode=None, PrintFIMA=None, Size=None, Height=None, Width=None, FeedSource=None, AddressFromLeft=None, AddressFromTop=None, ReturnAddressFromLeft=None, ReturnAddressFromTop=None, DefaultFaceUp=None, DefaultOrientation=None, PrintEPostage=None, Vertical=None, RecipientNamefromLeft=None, RecipientNamefromTop=None, RecipientPostalfromLeft=None, RecipientPostalfromTop=None, SenderNamefromLeft=None, SenderNamefromTop=None, SenderPostalfromLeft=None, SenderPostalfromTop=None):
        arguments = com_arguments([ExtractAddress, Address, AutoText, OmitReturnAddress, ReturnAddress, ReturnAutoText, PrintBarCode, PrintFIMA, Size, Height, Width, FeedSource, AddressFromLeft, AddressFromTop, ReturnAddressFromLeft, ReturnAddressFromTop, DefaultFaceUp, DefaultOrientation, PrintEPostage, Vertical, RecipientNamefromLeft, RecipientNamefromTop, RecipientPostalfromLeft, RecipientPostalfromTop, SenderNamefromLeft, SenderNamefromTop, SenderPostalfromLeft, SenderPostalfromTop])
        self.envelope.PrintOut(*arguments)

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

    @Code.setter
    def Code(self, value):
        self.field.Code = value

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

    @Locked.setter
    def Locked(self, value):
        self.field.Locked = value

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

    @Result.setter
    def Result(self, value):
        self.field.Result = value

    @property
    def ShowCodes(self):
        return self.field.ShowCodes

    @ShowCodes.setter
    def ShowCodes(self, value):
        self.field.ShowCodes = value

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

    @Visible.setter
    def Visible(self, value):
        self.fillformat.Visible = value

    def OneColorGradient(self, Style=None, Variant=None, Degree=None):
        arguments = com_arguments([Style, Variant, Degree])
        self.fillformat.OneColorGradient(*arguments)

    def Patterned(self, Pattern=None):
        arguments = com_arguments([Pattern])
        self.fillformat.Patterned(*arguments)

    def PresetGradient(self, Style=None, Variant=None, PresetGradientType=None):
        arguments = com_arguments([Style, Variant, PresetGradientType])
        self.fillformat.PresetGradient(*arguments)

    def PresetTextured(self, PresetTexture=None):
        arguments = com_arguments([PresetTexture])
        self.fillformat.PresetTextured(*arguments)

    def Solid(self):
        self.fillformat.Solid()

    def TwoColorGradient(self, Style=None, Variant=None):
        arguments = com_arguments([Style, Variant])
        self.fillformat.TwoColorGradient(*arguments)

    def UserPicture(self, PictureFile=None):
        arguments = com_arguments([PictureFile])
        self.fillformat.UserPicture(*arguments)

    def UserTextured(self, TextureFile=None):
        arguments = com_arguments([TextureFile])
        self.fillformat.UserTextured(*arguments)


class Find:

    def __init__(self, find=None):
        self.find = find

    @property
    def Application(self):
        return Application(self.find.Application)

    @property
    def CorrectHangulEndings(self):
        return self.find.CorrectHangulEndings

    @CorrectHangulEndings.setter
    def CorrectHangulEndings(self, value):
        self.find.CorrectHangulEndings = value

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

    @Format.setter
    def Format(self, value):
        self.find.Format = value

    @property
    def Forward(self):
        return self.find.Forward

    @Forward.setter
    def Forward(self, value):
        self.find.Forward = value

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

    @Highlight.setter
    def Highlight(self, value):
        self.find.Highlight = value

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

    @MatchAlefHamza.setter
    def MatchAlefHamza(self, value):
        self.find.MatchAlefHamza = value

    @property
    def MatchAllWordForms(self):
        return self.find.MatchAllWordForms

    @MatchAllWordForms.setter
    def MatchAllWordForms(self, value):
        self.find.MatchAllWordForms = value

    @property
    def MatchByte(self):
        return self.find.MatchByte

    @MatchByte.setter
    def MatchByte(self, value):
        self.find.MatchByte = value

    @property
    def MatchCase(self):
        return self.find.MatchCase

    @MatchCase.setter
    def MatchCase(self, value):
        self.find.MatchCase = value

    @property
    def MatchControl(self):
        return self.find.MatchControl

    @MatchControl.setter
    def MatchControl(self, value):
        self.find.MatchControl = value

    @property
    def MatchDiacritics(self):
        return self.find.MatchDiacritics

    @MatchDiacritics.setter
    def MatchDiacritics(self, value):
        self.find.MatchDiacritics = value

    @property
    def MatchFuzzy(self):
        return self.find.MatchFuzzy

    @MatchFuzzy.setter
    def MatchFuzzy(self, value):
        self.find.MatchFuzzy = value

    @property
    def MatchKashida(self):
        return self.find.MatchKashida

    @MatchKashida.setter
    def MatchKashida(self, value):
        self.find.MatchKashida = value

    @property
    def MatchPhrase(self):
        return self.find.MatchPhrase

    @MatchPhrase.setter
    def MatchPhrase(self, value):
        self.find.MatchPhrase = value

    @property
    def MatchPrefix(self):
        return self.find.MatchPrefix

    @MatchPrefix.setter
    def MatchPrefix(self, value):
        self.find.MatchPrefix = value

    @property
    def MatchSoundsLike(self):
        return self.find.MatchSoundsLike

    @MatchSoundsLike.setter
    def MatchSoundsLike(self, value):
        self.find.MatchSoundsLike = value

    @property
    def MatchSuffix(self):
        return self.find.MatchSuffix

    @MatchSuffix.setter
    def MatchSuffix(self, value):
        self.find.MatchSuffix = value

    @property
    def MatchWholeWord(self):
        return self.find.MatchWholeWord

    @MatchWholeWord.setter
    def MatchWholeWord(self, value):
        self.find.MatchWholeWord = value

    @property
    def MatchWildcards(self):
        return self.find.MatchWildcards

    @MatchWildcards.setter
    def MatchWildcards(self, value):
        self.find.MatchWildcards = value

    @property
    def NoProofing(self):
        return self.find.NoProofing

    @NoProofing.setter
    def NoProofing(self, value):
        self.find.NoProofing = value

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

    def Execute(self, FindText=None, MatchCase=None, MatchWholeWord=None, MatchWildcards=None, MatchSoundsLike=None, MatchAllWordForms=None, Forward=None, Wrap=None, Format=None, ReplaceWith=None, Replace=None, MatchKashida=None, MatchDiacritics=None, MatchAlefHamza=None, MatchControl=None):
        arguments = com_arguments([FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl])
        return self.find.Execute(*arguments)

    def Execute2007(self, FindText=None, MatchCase=None, MatchWholeWord=None, MatchWildcards=None, MatchSoundsLike=None, MatchAllWordForms=None, Forward=None, Wrap=None, Format=None, ReplaceWith=None, Replace=None, MatchKashida=None, MatchDiacritics=None, MatchAlefHamza=None, MatchControl=None, MatchPrefix=None, MatchSuffix=None, MatchPhrase=None, IgnoreSpace=None, IgnorePunct=None):
        arguments = com_arguments([FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms, Forward, Wrap, Format, ReplaceWith, Replace, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl, MatchPrefix, MatchSuffix, MatchPhrase, IgnoreSpace, IgnorePunct])
        return self.find.Execute2007(*arguments)

    def HitHighlight(self, FindText=None, HighlightColor=None, TextColor=None, MatchCase=None, MatchWholeWord=None, MatchPrefix=None, MatchSuffix=None, MatchPhrase=None, MatchWildcards=None, MatchSoundsLike=None, MatchAllWordForms=None, MatchByte=None, MatchFuzzy=None, MatchKashida=None, MatchDiacritics=None, MatchAlefHamza=None, MatchControl=None, IgnoreSpace=None, IgnorePunct=None, HanjaPhoneticHangul=None):
        arguments = com_arguments([FindText, HighlightColor, TextColor, MatchCase, MatchWholeWord, MatchPrefix, MatchSuffix, MatchPhrase, MatchWildcards, MatchSoundsLike, MatchAllWordForms, MatchByte, MatchFuzzy, MatchKashida, MatchDiacritics, MatchAlefHamza, MatchControl, IgnoreSpace, IgnorePunct, HanjaPhoneticHangul])
        return self.find.HitHighlight(*arguments)

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

    @AllCaps.setter
    def AllCaps(self, value):
        self.font.AllCaps = value

    @property
    def Application(self):
        return Application(self.font.Application)

    @property
    def Bold(self):
        return self.font.Bold

    @Bold.setter
    def Bold(self, value):
        self.font.Bold = value

    @property
    def BoldBi(self):
        return self.font.BoldBi

    @BoldBi.setter
    def BoldBi(self, value):
        self.font.BoldBi = value

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

    @ContextualAlternates.setter
    def ContextualAlternates(self, value):
        self.font.ContextualAlternates = value

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

    @DisableCharacterSpaceGrid.setter
    def DisableCharacterSpaceGrid(self, value):
        self.font.DisableCharacterSpaceGrid = value

    @property
    def DoubleStrikeThrough(self):
        return self.font.DoubleStrikeThrough

    @property
    def Duplicate(self):
        return Font(self.font.Duplicate)

    @property
    def Emboss(self):
        return self.font.Emboss

    @Emboss.setter
    def Emboss(self, value):
        self.font.Emboss = value

    @property
    def EmphasisMark(self):
        return WdEmphasisMark(self.font.EmphasisMark)

    @EmphasisMark.setter
    def EmphasisMark(self, value):
        self.font.EmphasisMark = value

    @property
    def Engrave(self):
        return self.font.Engrave

    @Engrave.setter
    def Engrave(self, value):
        self.font.Engrave = value

    @property
    def Fill(self):
        return self.font.Fill

    @property
    def Glow(self):
        return GlowFormat(self.font.Glow)

    @property
    def Hidden(self):
        return self.font.Hidden

    @Hidden.setter
    def Hidden(self, value):
        self.font.Hidden = value

    @property
    def Italic(self):
        return self.font.Italic

    @Italic.setter
    def Italic(self, value):
        self.font.Italic = value

    @property
    def ItalicBi(self):
        return self.font.ItalicBi

    @ItalicBi.setter
    def ItalicBi(self, value):
        self.font.ItalicBi = value

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

    @Line.setter
    def Line(self, value):
        self.font.Line = value

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

    @Outline.setter
    def Outline(self, value):
        self.font.Outline = value

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

    @Shadow.setter
    def Shadow(self, value):
        self.font.Shadow = value

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

    @SmallCaps.setter
    def SmallCaps(self, value):
        self.font.SmallCaps = value

    @property
    def Spacing(self):
        return self.font.Spacing

    @Spacing.setter
    def Spacing(self, value):
        self.font.Spacing = value

    @property
    def StrikeThrough(self):
        return self.font.StrikeThrough

    @StrikeThrough.setter
    def StrikeThrough(self, value):
        self.font.StrikeThrough = value

    @property
    def StylisticSet(self):
        return self.font.StylisticSet

    @StylisticSet.setter
    def StylisticSet(self, value):
        self.font.StylisticSet = value

    @property
    def Subscript(self):
        return self.font.Subscript

    @Subscript.setter
    def Subscript(self, value):
        self.font.Subscript = value

    @property
    def Superscript(self):
        return self.font.Superscript

    @Superscript.setter
    def Superscript(self, value):
        self.font.Superscript = value

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.fontnames.Item(*arguments)


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

    @CalculateOnExit.setter
    def CalculateOnExit(self, value):
        self.formfield.CalculateOnExit = value

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

    @Enabled.setter
    def Enabled(self, value):
        self.formfield.Enabled = value

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

    @OwnHelp.setter
    def OwnHelp(self, value):
        self.formfield.OwnHelp = value

    @property
    def OwnStatus(self):
        return self.formfield.OwnStatus

    @OwnStatus.setter
    def OwnStatus(self, value):
        self.formfield.OwnStatus = value

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

    @Result.setter
    def Result(self, value):
        self.formfield.Result = value

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

    @LockAnchor.setter
    def LockAnchor(self, value):
        self.frame.LockAnchor = value

    @property
    def Parent(self):
        return self.frame.Parent

    @property
    def Range(self):
        return Range(self.frame.Range)

    @property
    def RelativeHorizontalPosition(self):
        return self.frame.RelativeHorizontalPosition

    @RelativeHorizontalPosition.setter
    def RelativeHorizontalPosition(self, value):
        self.frame.RelativeHorizontalPosition = value

    @property
    def RelativeVerticalPosition(self):
        return self.frame.RelativeVerticalPosition

    @RelativeVerticalPosition.setter
    def RelativeVerticalPosition(self, value):
        self.frame.RelativeVerticalPosition = value

    @property
    def Shading(self):
        return Shading(self.frame.Shading)

    @property
    def TextWrap(self):
        return self.frame.TextWrap

    @TextWrap.setter
    def TextWrap(self, value):
        self.frame.TextWrap = value

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

    def ChildFramesetItem(self, Index=None):
        arguments = com_arguments([Index])
        if callable(self.frameset.ChildFramesetItem):
            return Frameset(self.frameset.ChildFramesetItem(*arguments))
        else:
            return Frameset(self.frameset.GetChildFramesetItem(*arguments))

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

    @FrameDisplayBorders.setter
    def FrameDisplayBorders(self, value):
        self.frameset.FrameDisplayBorders = value

    @property
    def FrameLinkToFile(self):
        return self.frameset.FrameLinkToFile

    @FrameLinkToFile.setter
    def FrameLinkToFile(self, value):
        self.frameset.FrameLinkToFile = value

    @property
    def FrameName(self):
        return self.frameset.FrameName

    @FrameName.setter
    def FrameName(self, value):
        self.frameset.FrameName = value

    @property
    def FrameResizable(self):
        return self.frameset.FrameResizable

    @FrameResizable.setter
    def FrameResizable(self, value):
        self.frameset.FrameResizable = value

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

    def AddNewFrame(self, Where=None):
        arguments = com_arguments([Where])
        self.frameset.AddNewFrame(*arguments)

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

    def AddNodes(self, SegmentType=None, EditingType=None, X1=None, Y1=None, X2=None, Y2=None, X3=None, Y3=None):
        arguments = com_arguments([SegmentType, EditingType, X1, Y1, X2, Y2, X3, Y3])
        self.freeformbuilder.AddNodes(*arguments)

    def ConvertToShape(self, Anchor=None):
        arguments = com_arguments([Anchor])
        self.freeformbuilder.ConvertToShape(*arguments)


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

    @Exists.setter
    def Exists(self, value):
        self.headerfooter.Exists = value

    @property
    def Index(self):
        return WdHeaderFooterIndex(self.headerfooter.Index)

    @property
    def IsHeader(self):
        return self.headerfooter.IsHeader

    @property
    def LinkToPrevious(self):
        return self.headerfooter.LinkToPrevious

    @LinkToPrevious.setter
    def LinkToPrevious(self, value):
        self.headerfooter.LinkToPrevious = value

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

    @NoShade.setter
    def NoShade(self, value):
        self.horizontallineformat.NoShade = value

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

    def HTMLDivisionParent(self, LevelsUp=None):
        arguments = com_arguments([LevelsUp])
        return self.htmldivision.HTMLDivisionParent(*arguments)


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

    def Add(self, Range=None):
        arguments = com_arguments([Range])
        return HTMLDivision(self.htmldivisions.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.htmldivisions.Item(*arguments)


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

    def CreateNewDocument(self, FileName=None, EditNow=None, Overwrite=None):
        arguments = com_arguments([FileName, EditNow, Overwrite])
        self.hyperlink.CreateNewDocument(*arguments)

    def Delete(self):
        self.hyperlink.Delete()

    def Follow(self, NewWindow=None, AddHistory=None, ExtraInfo=None, Method=None, HeaderInfo=None):
        arguments = com_arguments([NewWindow, AddHistory, ExtraInfo, Method, HeaderInfo])
        self.hyperlink.Follow(*arguments)


class Index:

    def __init__(self, index=None):
        self.index = index

    @property
    def AccentedLetters(self):
        return self.index.AccentedLetters

    @AccentedLetters.setter
    def AccentedLetters(self, value):
        self.index.AccentedLetters = value

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

    @NumberOfColumns.setter
    def NumberOfColumns(self, value):
        self.index.NumberOfColumns = value

    @property
    def Parent(self):
        return self.index.Parent

    @property
    def Range(self):
        return Range(self.index.Range)

    @property
    def RightAlignPageNumbers(self):
        return self.index.RightAlignPageNumbers

    @RightAlignPageNumbers.setter
    def RightAlignPageNumbers(self, value):
        self.index.RightAlignPageNumbers = value

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

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.inlineshape.LockAspectRatio = value

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

    @ScaleHeight.setter
    def ScaleHeight(self, value):
        self.inlineshape.ScaleHeight = value

    @property
    def ScaleWidth(self):
        return self.inlineshape.ScaleWidth

    @ScaleWidth.setter
    def ScaleWidth(self, value):
        self.inlineshape.ScaleWidth = value

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

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.interior.InvertIfNegative = value

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

    def Rebind(self, KeyCategory=None, Command=None, CommandParameter=None):
        arguments = com_arguments([KeyCategory, Command, CommandParameter])
        self.keybinding.Rebind(*arguments)


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

    @IncludeInLayout.setter
    def IncludeInLayout(self, value):
        self.legend.IncludeInLayout = value

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return LegendEntry(self.legendentries.Item(*arguments))


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

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.legendkey.InvertIfNegative = value

    @property
    def Left(self):
        return self.legendkey.Left

    @property
    def MarkerBackgroundColor(self):
        return self.legendkey.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.legendkey.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return self.legendkey.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.legendkey.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.legendkey.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.legendkey.MarkerForegroundColor = value

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

    @Smooth.setter
    def Smooth(self, value):
        self.legendkey.Smooth = value

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

    @IncludeHeaderFooter.setter
    def IncludeHeaderFooter(self, value):
        self.lettercontent.IncludeHeaderFooter = value

    @property
    def InfoBlock(self):
        return self.lettercontent.InfoBlock

    @property
    def Letterhead(self):
        return self.lettercontent.Letterhead

    @Letterhead.setter
    def Letterhead(self, value):
        self.lettercontent.Letterhead = value

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

    @SenderReference.setter
    def SenderReference(self, value):
        self.lettercontent.SenderReference = value

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

    @InsetPen.setter
    def InsetPen(self, value):
        self.lineformat.InsetPen = value

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

    @Visible.setter
    def Visible(self, value):
        self.lineformat.Visible = value

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

    @Active.setter
    def Active(self, value):
        self.linenumbering.Active = value

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.lines.Item(*arguments)


class LinkFormat:

    def __init__(self, linkformat=None):
        self.linkformat = linkformat

    @property
    def Application(self):
        return Application(self.linkformat.Application)

    @property
    def AutoUpdate(self):
        return self.linkformat.AutoUpdate

    @AutoUpdate.setter
    def AutoUpdate(self, value):
        self.linkformat.AutoUpdate = value

    @property
    def Creator(self):
        return self.linkformat.Creator

    @property
    def Locked(self):
        return self.linkformat.Locked

    @Locked.setter
    def Locked(self, value):
        self.linkformat.Locked = value

    @property
    def Parent(self):
        return self.linkformat.Parent

    @property
    def SavePictureWithDocument(self):
        return self.linkformat.SavePictureWithDocument

    @SavePictureWithDocument.setter
    def SavePictureWithDocument(self, value):
        self.linkformat.SavePictureWithDocument = value

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

    def ApplyListTemplate(self, ListTemplate=None, ContinuePreviousList=None, ApplyTo=None, DefaultListBehavior=None):
        arguments = com_arguments([ListTemplate, ContinuePreviousList, ApplyTo, DefaultListBehavior])
        self.list.ApplyListTemplate(*arguments)

    def ApplyListTemplateWithLevel(self, ListTemplate=None, ContinuePreviousList=None, DefaultListBehavior=None, ApplyLevel=None):
        arguments = com_arguments([ListTemplate, ContinuePreviousList, DefaultListBehavior, ApplyLevel])
        self.list.ApplyListTemplateWithLevel(*arguments)

    def CanContinuePreviousList(self, ListTemplate=None):
        arguments = com_arguments([ListTemplate])
        self.list.CanContinuePreviousList(*arguments)

    def ConvertNumbersToText(self):
        self.list.ConvertNumbersToText()

    def CountNumberedItems(self):
        self.list.CountNumberedItems()

    def RemoveNumbers(self, NumberType=None):
        arguments = com_arguments([NumberType])
        self.list.RemoveNumbers(*arguments)


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

    def ApplyBulletDefault(self, DefaultListBehavior=None):
        arguments = com_arguments([DefaultListBehavior])
        self.listformat.ApplyBulletDefault(*arguments)

    def ApplyListTemplate(self, ListTemplate=None, ContinuePreviousList=None, ApplyTo=None, DefaultListBehavior=None):
        arguments = com_arguments([ListTemplate, ContinuePreviousList, ApplyTo, DefaultListBehavior])
        self.listformat.ApplyListTemplate(*arguments)

    def ApplyListTemplateWithLevel(self, ListTemplate=None, ContinuePreviousList=None, ApplyTo=None, DefaultListBehavior=None, ApplyLevel=None):
        arguments = com_arguments([ListTemplate, ContinuePreviousList, ApplyTo, DefaultListBehavior, ApplyLevel])
        self.listformat.ApplyListTemplateWithLevel(*arguments)

    def ApplyNumberDefault(self, DefaultListBehavior=None):
        arguments = com_arguments([DefaultListBehavior])
        self.listformat.ApplyNumberDefault(*arguments)

    def ApplyOutlineNumberDefault(self, DefaultListBehavior=None):
        arguments = com_arguments([DefaultListBehavior])
        self.listformat.ApplyOutlineNumberDefault(*arguments)

    def CanContinuePreviousList(self, ListTemplate=None):
        arguments = com_arguments([ListTemplate])
        self.listformat.CanContinuePreviousList(*arguments)

    def ConvertNumbersToText(self):
        self.listformat.ConvertNumbersToText()

    def CountNumberedItems(self):
        self.listformat.CountNumberedItems()

    def ListIndent(self):
        self.listformat.ListIndent()

    def ListOutdent(self):
        self.listformat.ListOutdent()

    def RemoveNumbers(self, NumberType=None):
        arguments = com_arguments([NumberType])
        self.listformat.RemoveNumbers(*arguments)


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

    def Modified(self, Index=None):
        arguments = com_arguments([Index])
        if callable(self.listgallery.Modified):
            return self.listgallery.Modified(*arguments)
        else:
            return self.listgallery.GetModified(*arguments)

    @property
    def Parent(self):
        return self.listgallery.Parent

    def Reset(self, Index=None):
        arguments = com_arguments([Index])
        self.listgallery.Reset(*arguments)


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

    @ResetOnHigher.setter
    def ResetOnHigher(self, value):
        self.listlevel.ResetOnHigher = value

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

    def ApplyPictureBullet(self, FileName=None):
        arguments = com_arguments([FileName])
        self.listlevel.ApplyPictureBullet(*arguments)


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

    @OutlineNumbered.setter
    def OutlineNumbered(self, value):
        self.listtemplate.OutlineNumbered = value

    @property
    def Parent(self):
        return self.listtemplate.Parent

    def Convert(self, Level=None):
        arguments = com_arguments([Level])
        self.listtemplate.Convert(*arguments)


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

    @Vertical.setter
    def Vertical(self, value):
        self.mailinglabel.Vertical = value

    def CreateNewDocument(self, Name=None, Address=None, AutoText=None, ExtractAddress=None, LaserTray=None, PrintEPostageLabel=None, Vertical=None):
        arguments = com_arguments([Name, Address, AutoText, ExtractAddress, LaserTray, PrintEPostageLabel, Vertical])
        return self.mailinglabel.CreateNewDocument(*arguments)

    def CreateNewDocumentByID(self, LabelID=None, Address=None, AutoText=None, ExtractAddress=None, LaserTray=None, PrintEPostageLabel=None, Vertical=None):
        arguments = com_arguments([LabelID, Address, AutoText, ExtractAddress, LaserTray, PrintEPostageLabel, Vertical])
        return self.mailinglabel.CreateNewDocumentByID(*arguments)

    def LabelOptions(self):
        self.mailinglabel.LabelOptions()

    def PrintOut(self, Name=None, Address=None, ExtractAddress=None, LaserTray=None, SingleLabel=None, Row=None, Column=None, PrintEPostageLabel=None, Vertical=None):
        arguments = com_arguments([Name, Address, ExtractAddress, LaserTray, SingleLabel, Row, Column, PrintEPostageLabel, Vertical])
        self.mailinglabel.PrintOut(*arguments)

    def PrintOutByID(self, LabelID=None, Address=None, ExtractAddress=None, LaserTray=None, SingleLabel=None, Row=None, Column=None, PrintEPostageLabel=None, Vertical=None):
        arguments = com_arguments([LabelID, Address, ExtractAddress, LaserTray, SingleLabel, Row, Column, PrintEPostageLabel, Vertical])
        self.mailinglabel.PrintOutByID(*arguments)


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

    @HighlightMergeFields.setter
    def HighlightMergeFields(self, value):
        self.mailmerge.HighlightMergeFields = value

    @property
    def MailAddressFieldName(self):
        return self.mailmerge.MailAddressFieldName

    @MailAddressFieldName.setter
    def MailAddressFieldName(self, value):
        self.mailmerge.MailAddressFieldName = value

    @property
    def MailAsAttachment(self):
        return self.mailmerge.MailAsAttachment

    @MailAsAttachment.setter
    def MailAsAttachment(self, value):
        self.mailmerge.MailAsAttachment = value

    @property
    def MailFormat(self):
        return WdMailMergeMailFormat(self.mailmerge.MailFormat)

    @MailFormat.setter
    def MailFormat(self, value):
        self.mailmerge.MailFormat = value

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

    @SuppressBlankLines.setter
    def SuppressBlankLines(self, value):
        self.mailmerge.SuppressBlankLines = value

    @property
    def ViewMailMergeFieldCodes(self):
        return self.mailmerge.ViewMailMergeFieldCodes

    @ViewMailMergeFieldCodes.setter
    def ViewMailMergeFieldCodes(self, value):
        self.mailmerge.ViewMailMergeFieldCodes = value

    @property
    def WizardState(self):
        return self.mailmerge.WizardState

    @WizardState.setter
    def WizardState(self, value):
        self.mailmerge.WizardState = value

    def Check(self):
        self.mailmerge.Check()

    def CreateDataSource(self, Name=None, PasswordDocument=None, WritePasswordDocument=None, HeaderRecord=None, MSQuery=None, SQLStatement=None, SQLStatement1=None, Connection=None, LinkToSource=None):
        arguments = com_arguments([Name, PasswordDocument, WritePasswordDocument, HeaderRecord, MSQuery, SQLStatement, SQLStatement1, Connection, LinkToSource])
        self.mailmerge.CreateDataSource(*arguments)

    def CreateHeaderSource(self, Name=None, PasswordDocument=None, WritePasswordDocument=None, HeaderRecord=None):
        arguments = com_arguments([Name, PasswordDocument, WritePasswordDocument, HeaderRecord])
        self.mailmerge.CreateHeaderSource(*arguments)

    def EditDataSource(self):
        self.mailmerge.EditDataSource()

    def EditHeaderSource(self):
        self.mailmerge.EditHeaderSource()

    def EditMainDocument(self):
        self.mailmerge.EditMainDocument()

    def Execute(self, Pause=None):
        arguments = com_arguments([Pause])
        self.mailmerge.Execute(*arguments)

    def OpenDataSource(self, Name=None, Format=None, ConfirmConversions=None, ReadOnly=None, LinkToSource=None, AddToRecentFiles=None, PasswordDocument=None, PasswordTemplate=None, Revert=None, WritePasswordDocument=None, WritePasswordTemplate=None, Connection=None, SQLStatement=None, SQLStatement1=None, OpenExclusive=None, SubType=None):
        arguments = com_arguments([Name, Format, ConfirmConversions, ReadOnly, LinkToSource, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, Connection, SQLStatement, SQLStatement1, OpenExclusive, SubType])
        self.mailmerge.OpenDataSource(*arguments)

    def OpenHeaderSource(self, Name=None, Format=None, ConfirmConversions=None, ReadOnly=None, AddToRecentFiles=None, PasswordDocument=None, PasswordTemplate=None, Revert=None, WritePasswordDocument=None, WritePasswordTemplate=None, OpenExclusive=None):
        arguments = com_arguments([Name, Format, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert, WritePasswordDocument, WritePasswordTemplate, OpenExclusive])
        self.mailmerge.OpenHeaderSource(*arguments)

    def ShowWizard(self, InitialState=None, ShowDocumentStep=None, ShowTemplateStep=None, ShowDataStep=None, ShowWriteStep=None, ShowPreviewStep=None, ShowMergeStep=None):
        arguments = com_arguments([InitialState, ShowDocumentStep, ShowTemplateStep, ShowDataStep, ShowWriteStep, ShowPreviewStep, ShowMergeStep])
        self.mailmerge.ShowWizard(*arguments)


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

    @Included.setter
    def Included(self, value):
        self.mailmergedatasource.Included = value

    @property
    def InvalidAddress(self):
        return self.mailmergedatasource.InvalidAddress

    @InvalidAddress.setter
    def InvalidAddress(self, value):
        self.mailmergedatasource.InvalidAddress = value

    @property
    def InvalidComments(self):
        return self.mailmergedatasource.InvalidComments

    @InvalidComments.setter
    def InvalidComments(self, value):
        self.mailmergedatasource.InvalidComments = value

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

    def FindRecord(self, FindText=None, Field=None):
        arguments = com_arguments([FindText, Field])
        return self.mailmergedatasource.FindRecord(*arguments)

    def SetAllErrorFlags(self, Invalid=None, InvalidComment=None):
        arguments = com_arguments([Invalid, InvalidComment])
        self.mailmergedatasource.SetAllErrorFlags(*arguments)

    def SetAllIncludedFlags(self, Included=None):
        arguments = com_arguments([Included])
        self.mailmergedatasource.SetAllIncludedFlags(*arguments)


class MailMergeField:

    def __init__(self, mailmergefield=None):
        self.mailmergefield = mailmergefield

    @property
    def Application(self):
        return Application(self.mailmergefield.Application)

    @property
    def Code(self):
        return Range(self.mailmergefield.Code)

    @Code.setter
    def Code(self, value):
        self.mailmergefield.Code = value

    @property
    def Creator(self):
        return self.mailmergefield.Creator

    @property
    def Locked(self):
        return self.mailmergefield.Locked

    @Locked.setter
    def Locked(self, value):
        self.mailmergefield.Locked = value

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

    @DataFieldName.setter
    def DataFieldName(self, value):
        self.mappeddatafield.DataFieldName = value

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.mappeddatafields.Item(*arguments)


class Model3DFormat:

    def __init__(self, model3dformat=None):
        self.model3dformat = model3dformat

    @property
    def Application(self):
        return Application(self.model3dformat.Application)

    @property
    def AutoFit(self):
        return self.model3dformat.AutoFit

    @AutoFit.setter
    def AutoFit(self, value):
        self.model3dformat.AutoFit = value

    @property
    def CameraPositionX(self):
        return self.model3dformat.CameraPositionX

    @CameraPositionX.setter
    def CameraPositionX(self, value):
        self.model3dformat.CameraPositionX = value

    @property
    def CameraPositionY(self):
        return self.model3dformat.CameraPositionY

    @CameraPositionY.setter
    def CameraPositionY(self, value):
        self.model3dformat.CameraPositionY = value

    @property
    def CameraPositionZ(self):
        return self.model3dformat.CameraPositionZ

    @CameraPositionZ.setter
    def CameraPositionZ(self, value):
        self.model3dformat.CameraPositionZ = value

    @property
    def Creator(self):
        return self.model3dformat.Creator

    @property
    def FieldOfView(self):
        return self.model3dformat.FieldOfView

    @FieldOfView.setter
    def FieldOfView(self, value):
        self.model3dformat.FieldOfView = value

    @property
    def LookAtPointX(self):
        return self.model3dformat.LookAtPointX

    @LookAtPointX.setter
    def LookAtPointX(self, value):
        self.model3dformat.LookAtPointX = value

    @property
    def LookAtPointY(self):
        return self.model3dformat.LookAtPointY

    @LookAtPointY.setter
    def LookAtPointY(self, value):
        self.model3dformat.LookAtPointY = value

    @property
    def LookAtPointZ(self):
        return self.model3dformat.LookAtPointZ

    @LookAtPointZ.setter
    def LookAtPointZ(self, value):
        self.model3dformat.LookAtPointZ = value

    @property
    def Parent(self):
        return self.model3dformat.Parent

    @property
    def RotationX(self):
        return self.model3dformat.RotationX

    @RotationX.setter
    def RotationX(self, value):
        self.model3dformat.RotationX = value

    @property
    def RotationY(self):
        return self.model3dformat.RotationY

    @RotationY.setter
    def RotationY(self, value):
        self.model3dformat.RotationY = value

    @property
    def RotationZ(self):
        return self.model3dformat.RotationZ

    @RotationZ.setter
    def RotationZ(self, value):
        self.model3dformat.RotationZ = value

    def IncrementRotationX(self, Increment=None):
        arguments = com_arguments([Increment])
        self.model3dformat.IncrementRotationX(*arguments)

    def IncrementRotationY(self, Increment=None):
        arguments = com_arguments([Increment])
        self.model3dformat.IncrementRotationY(*arguments)

    def IncrementRotationZ(self, Increment=None):
        arguments = com_arguments([Increment])
        self.model3dformat.IncrementRotationZ(*arguments)

    def ResetModel(self, ResetSize=None):
        arguments = com_arguments([ResetSize])
        self.model3dformat.ResetModel(*arguments)


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

    @DisplayAsIcon.setter
    def DisplayAsIcon(self, value):
        self.oleformat.DisplayAsIcon = value

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

    @PreserveFormattingOnUpdate.setter
    def PreserveFormattingOnUpdate(self, value):
        self.oleformat.PreserveFormattingOnUpdate = value

    @property
    def ProgID(self):
        return self.oleformat.ProgID

    def Activate(self):
        self.oleformat.Activate()

    def ActivateAs(self, ClassType=None):
        arguments = com_arguments([ClassType])
        self.oleformat.ActivateAs(*arguments)

    def ConvertTo(self, ClassType=None, DisplayAsIcon=None, IconFileName=None, IconIndex=None, IconLabel=None):
        arguments = com_arguments([ClassType, DisplayAsIcon, IconFileName, IconIndex, IconLabel])
        self.oleformat.ConvertTo(*arguments)

    def DoVerb(self, VerbIndex=None):
        arguments = com_arguments([VerbIndex])
        self.oleformat.DoVerb(*arguments)

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

    def Add(self, BeforeArg=None):
        arguments = com_arguments([BeforeArg])
        return self.omathargs.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.omathargs.Item(*arguments)


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

    def Add(self, Name=None, Value=None):
        arguments = com_arguments([Name, Value])
        return self.omathautocorrectentries.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.omathautocorrectentries.Item(*arguments)


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

    def Add(self, Range=None):
        arguments = com_arguments([Range])
        return self.omathbreaks.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.omathbreaks.Item(*arguments)


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

    def Add(self, Range=None, Type=None, NumArgs=None, NumCols=None):
        arguments = com_arguments([Range, Type, NumArgs, NumCols])
        return self.omathfunctions.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.omathfunctions.Item(*arguments)


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

    def Cell(self, Row=None, Col=None):
        arguments = com_arguments([Row, Col])
        if callable(self.omathmat.Cell):
            return OMath(self.omathmat.Cell(*arguments))
        else:
            return OMath(self.omathmat.GetCell(*arguments))

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

    def Add(self, BeforeCol=None):
        arguments = com_arguments([BeforeCol])
        return self.omathmatcols.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.omathmatcols.Item(*arguments)


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

    def Add(self, BeforeRow=None):
        arguments = com_arguments([BeforeRow])
        return self.omathmatrows.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.omathmatrows.Item(*arguments)


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

    def Add(self, Name=None):
        arguments = com_arguments([Name])
        return self.omathrecognizedfunctions.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.omathrecognizedfunctions.Item(*arguments)


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

    def Add(self, Range=None):
        arguments = com_arguments([Range])
        return OMath(self.omaths.Add(*arguments))

    def BuildUp(self):
        return self.omaths.BuildUp()

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.omaths.Item(*arguments)

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

    @AddBiDirectionalMarksWhenSavingTextFile.setter
    def AddBiDirectionalMarksWhenSavingTextFile(self, value):
        self.options.AddBiDirectionalMarksWhenSavingTextFile = value

    @property
    def AddControlCharacters(self):
        return self.options.AddControlCharacters

    @AddControlCharacters.setter
    def AddControlCharacters(self, value):
        self.options.AddControlCharacters = value

    @property
    def AddHebDoubleQuote(self):
        return self.options.AddHebDoubleQuote

    @AddHebDoubleQuote.setter
    def AddHebDoubleQuote(self, value):
        self.options.AddHebDoubleQuote = value

    @property
    def AllowAccentedUppercase(self):
        return self.options.AllowAccentedUppercase

    @AllowAccentedUppercase.setter
    def AllowAccentedUppercase(self, value):
        self.options.AllowAccentedUppercase = value

    @property
    def AllowClickAndTypeMouse(self):
        return self.options.AllowClickAndTypeMouse

    @AllowClickAndTypeMouse.setter
    def AllowClickAndTypeMouse(self, value):
        self.options.AllowClickAndTypeMouse = value

    @property
    def AllowCombinedAuxiliaryForms(self):
        return self.options.AllowCombinedAuxiliaryForms

    @AllowCombinedAuxiliaryForms.setter
    def AllowCombinedAuxiliaryForms(self, value):
        self.options.AllowCombinedAuxiliaryForms = value

    @property
    def AllowCompoundNounProcessing(self):
        return self.options.AllowCompoundNounProcessing

    @AllowCompoundNounProcessing.setter
    def AllowCompoundNounProcessing(self, value):
        self.options.AllowCompoundNounProcessing = value

    @property
    def AllowDragAndDrop(self):
        return self.options.AllowDragAndDrop

    @AllowDragAndDrop.setter
    def AllowDragAndDrop(self, value):
        self.options.AllowDragAndDrop = value

    @property
    def AllowOpenInDraftView(self):
        return self.options.AllowOpenInDraftView

    @AllowOpenInDraftView.setter
    def AllowOpenInDraftView(self, value):
        self.options.AllowOpenInDraftView = value

    @property
    def AllowPixelUnits(self):
        return self.options.AllowPixelUnits

    @AllowPixelUnits.setter
    def AllowPixelUnits(self, value):
        self.options.AllowPixelUnits = value

    @property
    def AllowReadingMode(self):
        return self.options.AllowReadingMode

    @AllowReadingMode.setter
    def AllowReadingMode(self, value):
        self.options.AllowReadingMode = value

    @property
    def AnimateScreenMovements(self):
        return self.options.AnimateScreenMovements

    @AnimateScreenMovements.setter
    def AnimateScreenMovements(self, value):
        self.options.AnimateScreenMovements = value

    @property
    def Application(self):
        return Application(self.options.Application)

    @property
    def ApplyFarEastFontsToAscii(self):
        return self.options.ApplyFarEastFontsToAscii

    @ApplyFarEastFontsToAscii.setter
    def ApplyFarEastFontsToAscii(self, value):
        self.options.ApplyFarEastFontsToAscii = value

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

    @AutoCreateNewDrawings.setter
    def AutoCreateNewDrawings(self, value):
        self.options.AutoCreateNewDrawings = value

    @property
    def AutoFormatApplyBulletedLists(self):
        return self.options.AutoFormatApplyBulletedLists

    @AutoFormatApplyBulletedLists.setter
    def AutoFormatApplyBulletedLists(self, value):
        self.options.AutoFormatApplyBulletedLists = value

    @property
    def AutoFormatApplyFirstIndents(self):
        return self.options.AutoFormatApplyFirstIndents

    @AutoFormatApplyFirstIndents.setter
    def AutoFormatApplyFirstIndents(self, value):
        self.options.AutoFormatApplyFirstIndents = value

    @property
    def AutoFormatApplyHeadings(self):
        return self.options.AutoFormatApplyHeadings

    @AutoFormatApplyHeadings.setter
    def AutoFormatApplyHeadings(self, value):
        self.options.AutoFormatApplyHeadings = value

    @property
    def AutoFormatApplyLists(self):
        return self.options.AutoFormatApplyLists

    @AutoFormatApplyLists.setter
    def AutoFormatApplyLists(self, value):
        self.options.AutoFormatApplyLists = value

    @property
    def AutoFormatApplyOtherParas(self):
        return self.options.AutoFormatApplyOtherParas

    @AutoFormatApplyOtherParas.setter
    def AutoFormatApplyOtherParas(self, value):
        self.options.AutoFormatApplyOtherParas = value

    @property
    def AutoFormatAsYouTypeApplyBorders(self):
        return self.options.AutoFormatAsYouTypeApplyBorders

    @AutoFormatAsYouTypeApplyBorders.setter
    def AutoFormatAsYouTypeApplyBorders(self, value):
        self.options.AutoFormatAsYouTypeApplyBorders = value

    @property
    def AutoFormatAsYouTypeApplyBulletedLists(self):
        return self.options.AutoFormatAsYouTypeApplyBulletedLists

    @AutoFormatAsYouTypeApplyBulletedLists.setter
    def AutoFormatAsYouTypeApplyBulletedLists(self, value):
        self.options.AutoFormatAsYouTypeApplyBulletedLists = value

    @property
    def AutoFormatAsYouTypeApplyClosings(self):
        return self.options.AutoFormatAsYouTypeApplyClosings

    @AutoFormatAsYouTypeApplyClosings.setter
    def AutoFormatAsYouTypeApplyClosings(self, value):
        self.options.AutoFormatAsYouTypeApplyClosings = value

    @property
    def AutoFormatAsYouTypeApplyDates(self):
        return self.options.AutoFormatAsYouTypeApplyDates

    @AutoFormatAsYouTypeApplyDates.setter
    def AutoFormatAsYouTypeApplyDates(self, value):
        self.options.AutoFormatAsYouTypeApplyDates = value

    @property
    def AutoFormatAsYouTypeApplyFirstIndents(self):
        return self.options.AutoFormatAsYouTypeApplyFirstIndents

    @AutoFormatAsYouTypeApplyFirstIndents.setter
    def AutoFormatAsYouTypeApplyFirstIndents(self, value):
        self.options.AutoFormatAsYouTypeApplyFirstIndents = value

    @property
    def AutoFormatAsYouTypeApplyHeadings(self):
        return self.options.AutoFormatAsYouTypeApplyHeadings

    @AutoFormatAsYouTypeApplyHeadings.setter
    def AutoFormatAsYouTypeApplyHeadings(self, value):
        self.options.AutoFormatAsYouTypeApplyHeadings = value

    @property
    def AutoFormatAsYouTypeApplyNumberedLists(self):
        return self.options.AutoFormatAsYouTypeApplyNumberedLists

    @AutoFormatAsYouTypeApplyNumberedLists.setter
    def AutoFormatAsYouTypeApplyNumberedLists(self, value):
        self.options.AutoFormatAsYouTypeApplyNumberedLists = value

    @property
    def AutoFormatAsYouTypeApplyTables(self):
        return self.options.AutoFormatAsYouTypeApplyTables

    @AutoFormatAsYouTypeApplyTables.setter
    def AutoFormatAsYouTypeApplyTables(self, value):
        self.options.AutoFormatAsYouTypeApplyTables = value

    @property
    def AutoFormatAsYouTypeAutoLetterWizard(self):
        return self.options.AutoFormatAsYouTypeAutoLetterWizard

    @AutoFormatAsYouTypeAutoLetterWizard.setter
    def AutoFormatAsYouTypeAutoLetterWizard(self, value):
        self.options.AutoFormatAsYouTypeAutoLetterWizard = value

    @property
    def AutoFormatAsYouTypeDefineStyles(self):
        return self.options.AutoFormatAsYouTypeDefineStyles

    @AutoFormatAsYouTypeDefineStyles.setter
    def AutoFormatAsYouTypeDefineStyles(self, value):
        self.options.AutoFormatAsYouTypeDefineStyles = value

    @property
    def AutoFormatAsYouTypeDeleteAutoSpaces(self):
        return self.options.AutoFormatAsYouTypeDeleteAutoSpaces

    @AutoFormatAsYouTypeDeleteAutoSpaces.setter
    def AutoFormatAsYouTypeDeleteAutoSpaces(self, value):
        self.options.AutoFormatAsYouTypeDeleteAutoSpaces = value

    @property
    def AutoFormatAsYouTypeFormatListItemBeginning(self):
        return self.options.AutoFormatAsYouTypeFormatListItemBeginning

    @AutoFormatAsYouTypeFormatListItemBeginning.setter
    def AutoFormatAsYouTypeFormatListItemBeginning(self, value):
        self.options.AutoFormatAsYouTypeFormatListItemBeginning = value

    @property
    def AutoFormatAsYouTypeInsertClosings(self):
        return self.options.AutoFormatAsYouTypeInsertClosings

    @AutoFormatAsYouTypeInsertClosings.setter
    def AutoFormatAsYouTypeInsertClosings(self, value):
        self.options.AutoFormatAsYouTypeInsertClosings = value

    @property
    def AutoFormatAsYouTypeInsertOvers(self):
        return self.options.AutoFormatAsYouTypeInsertOvers

    @AutoFormatAsYouTypeInsertOvers.setter
    def AutoFormatAsYouTypeInsertOvers(self, value):
        self.options.AutoFormatAsYouTypeInsertOvers = value

    @property
    def AutoFormatAsYouTypeMatchParentheses(self):
        return self.options.AutoFormatAsYouTypeMatchParentheses

    @AutoFormatAsYouTypeMatchParentheses.setter
    def AutoFormatAsYouTypeMatchParentheses(self, value):
        self.options.AutoFormatAsYouTypeMatchParentheses = value

    @property
    def AutoFormatAsYouTypeReplaceFarEastDashes(self):
        return self.options.AutoFormatAsYouTypeReplaceFarEastDashes

    @AutoFormatAsYouTypeReplaceFarEastDashes.setter
    def AutoFormatAsYouTypeReplaceFarEastDashes(self, value):
        self.options.AutoFormatAsYouTypeReplaceFarEastDashes = value

    @property
    def AutoFormatAsYouTypeReplaceFractions(self):
        return self.options.AutoFormatAsYouTypeReplaceFractions

    @AutoFormatAsYouTypeReplaceFractions.setter
    def AutoFormatAsYouTypeReplaceFractions(self, value):
        self.options.AutoFormatAsYouTypeReplaceFractions = value

    @property
    def AutoFormatAsYouTypeReplaceHyperlinks(self):
        return self.options.AutoFormatAsYouTypeReplaceHyperlinks

    @AutoFormatAsYouTypeReplaceHyperlinks.setter
    def AutoFormatAsYouTypeReplaceHyperlinks(self, value):
        self.options.AutoFormatAsYouTypeReplaceHyperlinks = value

    @property
    def AutoFormatAsYouTypeReplaceOrdinals(self):
        return self.options.AutoFormatAsYouTypeReplaceOrdinals

    @AutoFormatAsYouTypeReplaceOrdinals.setter
    def AutoFormatAsYouTypeReplaceOrdinals(self, value):
        self.options.AutoFormatAsYouTypeReplaceOrdinals = value

    @property
    def AutoFormatAsYouTypeReplacePlainTextEmphasis(self):
        return self.options.AutoFormatAsYouTypeReplacePlainTextEmphasis

    @AutoFormatAsYouTypeReplacePlainTextEmphasis.setter
    def AutoFormatAsYouTypeReplacePlainTextEmphasis(self, value):
        self.options.AutoFormatAsYouTypeReplacePlainTextEmphasis = value

    @property
    def AutoFormatAsYouTypeReplaceQuotes(self):
        return self.options.AutoFormatAsYouTypeReplaceQuotes

    @AutoFormatAsYouTypeReplaceQuotes.setter
    def AutoFormatAsYouTypeReplaceQuotes(self, value):
        self.options.AutoFormatAsYouTypeReplaceQuotes = value

    @property
    def AutoFormatAsYouTypeReplaceSymbols(self):
        return self.options.AutoFormatAsYouTypeReplaceSymbols

    @AutoFormatAsYouTypeReplaceSymbols.setter
    def AutoFormatAsYouTypeReplaceSymbols(self, value):
        self.options.AutoFormatAsYouTypeReplaceSymbols = value

    @property
    def AutoFormatDeleteAutoSpaces(self):
        return self.options.AutoFormatDeleteAutoSpaces

    @AutoFormatDeleteAutoSpaces.setter
    def AutoFormatDeleteAutoSpaces(self, value):
        self.options.AutoFormatDeleteAutoSpaces = value

    @property
    def AutoFormatMatchParentheses(self):
        return self.options.AutoFormatMatchParentheses

    @AutoFormatMatchParentheses.setter
    def AutoFormatMatchParentheses(self, value):
        self.options.AutoFormatMatchParentheses = value

    @property
    def AutoFormatPlainTextWordMail(self):
        return self.options.AutoFormatPlainTextWordMail

    @AutoFormatPlainTextWordMail.setter
    def AutoFormatPlainTextWordMail(self, value):
        self.options.AutoFormatPlainTextWordMail = value

    @property
    def AutoFormatPreserveStyles(self):
        return self.options.AutoFormatPreserveStyles

    @AutoFormatPreserveStyles.setter
    def AutoFormatPreserveStyles(self, value):
        self.options.AutoFormatPreserveStyles = value

    @property
    def AutoFormatReplaceFarEastDashes(self):
        return self.options.AutoFormatReplaceFarEastDashes

    @AutoFormatReplaceFarEastDashes.setter
    def AutoFormatReplaceFarEastDashes(self, value):
        self.options.AutoFormatReplaceFarEastDashes = value

    @property
    def AutoFormatReplaceFractions(self):
        return self.options.AutoFormatReplaceFractions

    @AutoFormatReplaceFractions.setter
    def AutoFormatReplaceFractions(self, value):
        self.options.AutoFormatReplaceFractions = value

    @property
    def AutoFormatReplaceHyperlinks(self):
        return self.options.AutoFormatReplaceHyperlinks

    @AutoFormatReplaceHyperlinks.setter
    def AutoFormatReplaceHyperlinks(self, value):
        self.options.AutoFormatReplaceHyperlinks = value

    @property
    def AutoFormatReplaceOrdinals(self):
        return self.options.AutoFormatReplaceOrdinals

    @AutoFormatReplaceOrdinals.setter
    def AutoFormatReplaceOrdinals(self, value):
        self.options.AutoFormatReplaceOrdinals = value

    @property
    def AutoFormatReplacePlainTextEmphasis(self):
        return self.options.AutoFormatReplacePlainTextEmphasis

    @AutoFormatReplacePlainTextEmphasis.setter
    def AutoFormatReplacePlainTextEmphasis(self, value):
        self.options.AutoFormatReplacePlainTextEmphasis = value

    @property
    def AutoFormatReplaceQuotes(self):
        return self.options.AutoFormatReplaceQuotes

    @AutoFormatReplaceQuotes.setter
    def AutoFormatReplaceQuotes(self, value):
        self.options.AutoFormatReplaceQuotes = value

    @property
    def AutoFormatReplaceSymbols(self):
        return self.options.AutoFormatReplaceSymbols

    @AutoFormatReplaceSymbols.setter
    def AutoFormatReplaceSymbols(self, value):
        self.options.AutoFormatReplaceSymbols = value

    @property
    def AutoKeyboardSwitching(self):
        return self.options.AutoKeyboardSwitching

    @AutoKeyboardSwitching.setter
    def AutoKeyboardSwitching(self, value):
        self.options.AutoKeyboardSwitching = value

    @property
    def AutoWordSelection(self):
        return self.options.AutoWordSelection

    @AutoWordSelection.setter
    def AutoWordSelection(self, value):
        self.options.AutoWordSelection = value

    @property
    def BackgroundSave(self):
        return self.options.BackgroundSave

    @BackgroundSave.setter
    def BackgroundSave(self, value):
        self.options.BackgroundSave = value

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

    @CheckGrammarAsYouType.setter
    def CheckGrammarAsYouType(self, value):
        self.options.CheckGrammarAsYouType = value

    @property
    def CheckGrammarWithSpelling(self):
        return self.options.CheckGrammarWithSpelling

    @CheckGrammarWithSpelling.setter
    def CheckGrammarWithSpelling(self, value):
        self.options.CheckGrammarWithSpelling = value

    @property
    def CheckHangulEndings(self):
        return self.options.CheckHangulEndings

    @CheckHangulEndings.setter
    def CheckHangulEndings(self, value):
        self.options.CheckHangulEndings = value

    @property
    def CheckSpellingAsYouType(self):
        return self.options.CheckSpellingAsYouType

    @CheckSpellingAsYouType.setter
    def CheckSpellingAsYouType(self, value):
        self.options.CheckSpellingAsYouType = value

    @property
    def CommentsColor(self):
        return WdColorIndex(self.options.CommentsColor)

    @CommentsColor.setter
    def CommentsColor(self, value):
        self.options.CommentsColor = value

    @property
    def ConfirmConversions(self):
        return self.options.ConfirmConversions

    @ConfirmConversions.setter
    def ConfirmConversions(self, value):
        self.options.ConfirmConversions = value

    @property
    def ContextualSpeller(self):
        return self.options.ContextualSpeller

    @ContextualSpeller.setter
    def ContextualSpeller(self, value):
        self.options.ContextualSpeller = value

    @property
    def ConvertHighAnsiToFarEast(self):
        return self.options.ConvertHighAnsiToFarEast

    @ConvertHighAnsiToFarEast.setter
    def ConvertHighAnsiToFarEast(self, value):
        self.options.ConvertHighAnsiToFarEast = value

    @property
    def CreateBackup(self):
        return self.options.CreateBackup

    @CreateBackup.setter
    def CreateBackup(self, value):
        self.options.CreateBackup = value

    @property
    def Creator(self):
        return self.options.Creator

    @property
    def CtrlClickHyperlinkToOpen(self):
        return self.options.CtrlClickHyperlinkToOpen

    @CtrlClickHyperlinkToOpen.setter
    def CtrlClickHyperlinkToOpen(self, value):
        self.options.CtrlClickHyperlinkToOpen = value

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

    @DefaultEPostageApp.setter
    def DefaultEPostageApp(self, value):
        self.options.DefaultEPostageApp = value

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

    @DisableFeaturesbyDefault.setter
    def DisableFeaturesbyDefault(self, value):
        self.options.DisableFeaturesbyDefault = value

    @property
    def DisableFeaturesIntroducedAfterbyDefault(self):
        return self.options.DisableFeaturesIntroducedAfterbyDefault

    @DisableFeaturesIntroducedAfterbyDefault.setter
    def DisableFeaturesIntroducedAfterbyDefault(self, value):
        self.options.DisableFeaturesIntroducedAfterbyDefault = value

    @property
    def DisplayGridLines(self):
        return self.options.DisplayGridLines

    @DisplayGridLines.setter
    def DisplayGridLines(self, value):
        self.options.DisplayGridLines = value

    @property
    def DisplayPasteOptions(self):
        return self.options.DisplayPasteOptions

    @DisplayPasteOptions.setter
    def DisplayPasteOptions(self, value):
        self.options.DisplayPasteOptions = value

    @property
    def DocumentViewDirection(self):
        return WdDocumentViewDirection(self.options.DocumentViewDirection)

    @DocumentViewDirection.setter
    def DocumentViewDirection(self, value):
        self.options.DocumentViewDirection = value

    @property
    def DoNotPromptForConvert(self):
        return self.options.DoNotPromptForConvert

    @DoNotPromptForConvert.setter
    def DoNotPromptForConvert(self, value):
        self.options.DoNotPromptForConvert = value

    @property
    def EnableHangulHanjaRecentOrdering(self):
        return self.options.EnableHangulHanjaRecentOrdering

    @EnableHangulHanjaRecentOrdering.setter
    def EnableHangulHanjaRecentOrdering(self, value):
        self.options.EnableHangulHanjaRecentOrdering = value

    @property
    def EnableLegacyIMEMode(self):
        return self.options.EnableLegacyIMEMode

    @EnableLegacyIMEMode.setter
    def EnableLegacyIMEMode(self, value):
        self.options.EnableLegacyIMEMode = value

    @property
    def EnableLivePreview(self):
        return self.options.EnableLivePreview

    @EnableLivePreview.setter
    def EnableLivePreview(self, value):
        self.options.EnableLivePreview = value

    @property
    def EnableMisusedWordsDictionary(self):
        return self.options.EnableMisusedWordsDictionary

    @EnableMisusedWordsDictionary.setter
    def EnableMisusedWordsDictionary(self, value):
        self.options.EnableMisusedWordsDictionary = value

    @property
    def EnableSound(self):
        return self.options.EnableSound

    @EnableSound.setter
    def EnableSound(self, value):
        self.options.EnableSound = value

    @property
    def EnvelopeFeederInstalled(self):
        return self.options.EnvelopeFeederInstalled

    @property
    def FormatScanning(self):
        return self.options.FormatScanning

    @FormatScanning.setter
    def FormatScanning(self, value):
        self.options.FormatScanning = value

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

    @HangulHanjaFastConversion.setter
    def HangulHanjaFastConversion(self, value):
        self.options.HangulHanjaFastConversion = value

    @property
    def HebrewMode(self):
        return WdHebSpellStart(self.options.HebrewMode)

    @HebrewMode.setter
    def HebrewMode(self, value):
        self.options.HebrewMode = value

    @property
    def IgnoreInternetAndFileAddresses(self):
        return self.options.IgnoreInternetAndFileAddresses

    @IgnoreInternetAndFileAddresses.setter
    def IgnoreInternetAndFileAddresses(self, value):
        self.options.IgnoreInternetAndFileAddresses = value

    @property
    def IgnoreMixedDigits(self):
        return self.options.IgnoreMixedDigits

    @IgnoreMixedDigits.setter
    def IgnoreMixedDigits(self, value):
        self.options.IgnoreMixedDigits = value

    @property
    def IgnoreUppercase(self):
        return self.options.IgnoreUppercase

    @IgnoreUppercase.setter
    def IgnoreUppercase(self, value):
        self.options.IgnoreUppercase = value

    @property
    def IMEAutomaticControl(self):
        return self.options.IMEAutomaticControl

    @IMEAutomaticControl.setter
    def IMEAutomaticControl(self, value):
        self.options.IMEAutomaticControl = value

    @property
    def InlineConversion(self):
        return self.options.InlineConversion

    @InlineConversion.setter
    def InlineConversion(self, value):
        self.options.InlineConversion = value

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

    @INSKeyForOvertype.setter
    def INSKeyForOvertype(self, value):
        self.options.INSKeyForOvertype = value

    @property
    def INSKeyForPaste(self):
        return self.options.INSKeyForPaste

    @INSKeyForPaste.setter
    def INSKeyForPaste(self, value):
        self.options.INSKeyForPaste = value

    @property
    def InterpretHighAnsi(self):
        return WdHighAnsiText(self.options.InterpretHighAnsi)

    @InterpretHighAnsi.setter
    def InterpretHighAnsi(self, value):
        self.options.InterpretHighAnsi = value

    @property
    def LocalNetworkFile(self):
        return self.options.LocalNetworkFile

    @LocalNetworkFile.setter
    def LocalNetworkFile(self, value):
        self.options.LocalNetworkFile = value

    @property
    def MapPaperSize(self):
        return self.options.MapPaperSize

    @MapPaperSize.setter
    def MapPaperSize(self, value):
        self.options.MapPaperSize = value

    @property
    def MatchFuzzyAY(self):
        return self.options.MatchFuzzyAY

    @MatchFuzzyAY.setter
    def MatchFuzzyAY(self, value):
        self.options.MatchFuzzyAY = value

    @property
    def MatchFuzzyBV(self):
        return self.options.MatchFuzzyBV

    @MatchFuzzyBV.setter
    def MatchFuzzyBV(self, value):
        self.options.MatchFuzzyBV = value

    @property
    def MatchFuzzyByte(self):
        return self.options.MatchFuzzyByte

    @MatchFuzzyByte.setter
    def MatchFuzzyByte(self, value):
        self.options.MatchFuzzyByte = value

    @property
    def MatchFuzzyCase(self):
        return self.options.MatchFuzzyCase

    @MatchFuzzyCase.setter
    def MatchFuzzyCase(self, value):
        self.options.MatchFuzzyCase = value

    @property
    def MatchFuzzyDash(self):
        return self.options.MatchFuzzyDash

    @MatchFuzzyDash.setter
    def MatchFuzzyDash(self, value):
        self.options.MatchFuzzyDash = value

    @property
    def MatchFuzzyDZ(self):
        return self.options.MatchFuzzyDZ

    @MatchFuzzyDZ.setter
    def MatchFuzzyDZ(self, value):
        self.options.MatchFuzzyDZ = value

    @property
    def MatchFuzzyHF(self):
        return self.options.MatchFuzzyHF

    @MatchFuzzyHF.setter
    def MatchFuzzyHF(self, value):
        self.options.MatchFuzzyHF = value

    @property
    def MatchFuzzyHiragana(self):
        return self.options.MatchFuzzyHiragana

    @MatchFuzzyHiragana.setter
    def MatchFuzzyHiragana(self, value):
        self.options.MatchFuzzyHiragana = value

    @property
    def MatchFuzzyIterationMark(self):
        return self.options.MatchFuzzyIterationMark

    @MatchFuzzyIterationMark.setter
    def MatchFuzzyIterationMark(self, value):
        self.options.MatchFuzzyIterationMark = value

    @property
    def MatchFuzzyKanji(self):
        return self.options.MatchFuzzyKanji

    @MatchFuzzyKanji.setter
    def MatchFuzzyKanji(self, value):
        self.options.MatchFuzzyKanji = value

    @property
    def MatchFuzzyKiKu(self):
        return self.options.MatchFuzzyKiKu

    @MatchFuzzyKiKu.setter
    def MatchFuzzyKiKu(self, value):
        self.options.MatchFuzzyKiKu = value

    @property
    def MatchFuzzyOldKana(self):
        return self.options.MatchFuzzyOldKana

    @MatchFuzzyOldKana.setter
    def MatchFuzzyOldKana(self, value):
        self.options.MatchFuzzyOldKana = value

    @property
    def MatchFuzzyProlongedSoundMark(self):
        return self.options.MatchFuzzyProlongedSoundMark

    @MatchFuzzyProlongedSoundMark.setter
    def MatchFuzzyProlongedSoundMark(self, value):
        self.options.MatchFuzzyProlongedSoundMark = value

    @property
    def MatchFuzzyPunctuation(self):
        return self.options.MatchFuzzyPunctuation

    @MatchFuzzyPunctuation.setter
    def MatchFuzzyPunctuation(self, value):
        self.options.MatchFuzzyPunctuation = value

    @property
    def MatchFuzzySmallKana(self):
        return self.options.MatchFuzzySmallKana

    @MatchFuzzySmallKana.setter
    def MatchFuzzySmallKana(self, value):
        self.options.MatchFuzzySmallKana = value

    @property
    def MatchFuzzySpace(self):
        return self.options.MatchFuzzySpace

    @MatchFuzzySpace.setter
    def MatchFuzzySpace(self, value):
        self.options.MatchFuzzySpace = value

    @property
    def MatchFuzzyTC(self):
        return self.options.MatchFuzzyTC

    @MatchFuzzyTC.setter
    def MatchFuzzyTC(self, value):
        self.options.MatchFuzzyTC = value

    @property
    def MatchFuzzyZJ(self):
        return self.options.MatchFuzzyZJ

    @MatchFuzzyZJ.setter
    def MatchFuzzyZJ(self, value):
        self.options.MatchFuzzyZJ = value

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

    @OptimizeForWord97byDefault.setter
    def OptimizeForWord97byDefault(self, value):
        self.options.OptimizeForWord97byDefault = value

    @property
    def Options(self):
        return self.options.Options

    @property
    def Overtype(self):
        return self.options.Overtype

    @Overtype.setter
    def Overtype(self, value):
        self.options.Overtype = value

    @property
    def Pagination(self):
        return self.options.Pagination

    @Pagination.setter
    def Pagination(self, value):
        self.options.Pagination = value

    @property
    def Parent(self):
        return self.options.Parent

    @property
    def PasteAdjustParagraphSpacing(self):
        return self.options.PasteAdjustParagraphSpacing

    @PasteAdjustParagraphSpacing.setter
    def PasteAdjustParagraphSpacing(self, value):
        self.options.PasteAdjustParagraphSpacing = value

    @property
    def PasteAdjustTableFormatting(self):
        return self.options.PasteAdjustTableFormatting

    @PasteAdjustTableFormatting.setter
    def PasteAdjustTableFormatting(self, value):
        self.options.PasteAdjustTableFormatting = value

    @property
    def PasteAdjustWordSpacing(self):
        return self.options.PasteAdjustWordSpacing

    @PasteAdjustWordSpacing.setter
    def PasteAdjustWordSpacing(self, value):
        self.options.PasteAdjustWordSpacing = value

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

    @PasteMergeFromPPT.setter
    def PasteMergeFromPPT(self, value):
        self.options.PasteMergeFromPPT = value

    @property
    def PasteMergeFromXL(self):
        return self.options.PasteMergeFromXL

    @PasteMergeFromXL.setter
    def PasteMergeFromXL(self, value):
        self.options.PasteMergeFromXL = value

    @property
    def PasteMergeLists(self):
        return self.options.PasteMergeLists

    @PasteMergeLists.setter
    def PasteMergeLists(self, value):
        self.options.PasteMergeLists = value

    @property
    def PasteOptionKeepBulletsAndNumbers(self):
        return self.options.PasteOptionKeepBulletsAndNumbers

    @PasteOptionKeepBulletsAndNumbers.setter
    def PasteOptionKeepBulletsAndNumbers(self, value):
        self.options.PasteOptionKeepBulletsAndNumbers = value

    @property
    def PasteSmartCutPaste(self):
        return self.options.PasteSmartCutPaste

    @PasteSmartCutPaste.setter
    def PasteSmartCutPaste(self, value):
        self.options.PasteSmartCutPaste = value

    @property
    def PasteSmartStyleBehavior(self):
        return self.options.PasteSmartStyleBehavior

    @PasteSmartStyleBehavior.setter
    def PasteSmartStyleBehavior(self, value):
        self.options.PasteSmartStyleBehavior = value

    @property
    def PictureEditor(self):
        return self.options.PictureEditor

    @PictureEditor.setter
    def PictureEditor(self, value):
        self.options.PictureEditor = value

    @property
    def PictureWrapType(self):
        return self.options.PictureWrapType

    @PictureWrapType.setter
    def PictureWrapType(self, value):
        self.options.PictureWrapType = value

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

    @PrintBackground.setter
    def PrintBackground(self, value):
        self.options.PrintBackground = value

    @property
    def PrintBackgrounds(self):
        return self.options.PrintBackgrounds

    @property
    def PrintComments(self):
        return self.options.PrintComments

    @PrintComments.setter
    def PrintComments(self, value):
        self.options.PrintComments = value

    @property
    def PrintDraft(self):
        return self.options.PrintDraft

    @PrintDraft.setter
    def PrintDraft(self, value):
        self.options.PrintDraft = value

    @property
    def PrintDrawingObjects(self):
        return self.options.PrintDrawingObjects

    @PrintDrawingObjects.setter
    def PrintDrawingObjects(self, value):
        self.options.PrintDrawingObjects = value

    @property
    def PrintEvenPagesInAscendingOrder(self):
        return self.options.PrintEvenPagesInAscendingOrder

    @PrintEvenPagesInAscendingOrder.setter
    def PrintEvenPagesInAscendingOrder(self, value):
        self.options.PrintEvenPagesInAscendingOrder = value

    @property
    def PrintFieldCodes(self):
        return self.options.PrintFieldCodes

    @PrintFieldCodes.setter
    def PrintFieldCodes(self, value):
        self.options.PrintFieldCodes = value

    @property
    def PrintHiddenText(self):
        return self.options.PrintHiddenText

    @PrintHiddenText.setter
    def PrintHiddenText(self, value):
        self.options.PrintHiddenText = value

    @property
    def PrintOddPagesInAscendingOrder(self):
        return self.options.PrintOddPagesInAscendingOrder

    @PrintOddPagesInAscendingOrder.setter
    def PrintOddPagesInAscendingOrder(self, value):
        self.options.PrintOddPagesInAscendingOrder = value

    @property
    def PrintProperties(self):
        return self.options.PrintProperties

    @PrintProperties.setter
    def PrintProperties(self, value):
        self.options.PrintProperties = value

    @property
    def PrintReverse(self):
        return self.options.PrintReverse

    @PrintReverse.setter
    def PrintReverse(self, value):
        self.options.PrintReverse = value

    @property
    def PrintXMLTag(self):
        return self.options.PrintXMLTag

    @property
    def PromptUpdateStyle(self):
        return self.options.PromptUpdateStyle

    @PromptUpdateStyle.setter
    def PromptUpdateStyle(self, value):
        self.options.PromptUpdateStyle = value

    @property
    def RepeatWord(self):
        return self.options.RepeatWord

    @RepeatWord.setter
    def RepeatWord(self, value):
        self.options.RepeatWord = value

    @property
    def ReplaceSelection(self):
        return self.options.ReplaceSelection

    @ReplaceSelection.setter
    def ReplaceSelection(self, value):
        self.options.ReplaceSelection = value

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

    @SaveNormalPrompt.setter
    def SaveNormalPrompt(self, value):
        self.options.SaveNormalPrompt = value

    @property
    def SavePropertiesPrompt(self):
        return self.options.SavePropertiesPrompt

    @SavePropertiesPrompt.setter
    def SavePropertiesPrompt(self, value):
        self.options.SavePropertiesPrompt = value

    @property
    def SendMailAttach(self):
        return self.options.SendMailAttach

    @SendMailAttach.setter
    def SendMailAttach(self, value):
        self.options.SendMailAttach = value

    @property
    def SequenceCheck(self):
        return self.options.SequenceCheck

    @SequenceCheck.setter
    def SequenceCheck(self, value):
        self.options.SequenceCheck = value

    @property
    def ShowControlCharacters(self):
        return self.options.ShowControlCharacters

    @ShowControlCharacters.setter
    def ShowControlCharacters(self, value):
        self.options.ShowControlCharacters = value

    @property
    def ShowDevTools(self):
        return self.options.ShowDevTools

    @ShowDevTools.setter
    def ShowDevTools(self, value):
        self.options.ShowDevTools = value

    @property
    def ShowDiacritics(self):
        return self.options.ShowDiacritics

    @ShowDiacritics.setter
    def ShowDiacritics(self, value):
        self.options.ShowDiacritics = value

    @property
    def ShowFormatError(self):
        return self.options.ShowFormatError

    @ShowFormatError.setter
    def ShowFormatError(self, value):
        self.options.ShowFormatError = value

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

    @ShowReadabilityStatistics.setter
    def ShowReadabilityStatistics(self, value):
        self.options.ShowReadabilityStatistics = value

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

    @SmartCutPaste.setter
    def SmartCutPaste(self, value):
        self.options.SmartCutPaste = value

    @property
    def SmartParaSelection(self):
        return self.options.SmartParaSelection

    @SmartParaSelection.setter
    def SmartParaSelection(self, value):
        self.options.SmartParaSelection = value

    @property
    def SnapToGrid(self):
        return self.options.SnapToGrid

    @SnapToGrid.setter
    def SnapToGrid(self, value):
        self.options.SnapToGrid = value

    @property
    def SnapToShapes(self):
        return self.options.SnapToShapes

    @SnapToShapes.setter
    def SnapToShapes(self, value):
        self.options.SnapToShapes = value

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

    @StoreRSIDOnSave.setter
    def StoreRSIDOnSave(self, value):
        self.options.StoreRSIDOnSave = value

    @property
    def StrictFinalYaa(self):
        return self.options.StrictFinalYaa

    @StrictFinalYaa.setter
    def StrictFinalYaa(self, value):
        self.options.StrictFinalYaa = value

    @property
    def StrictInitialAlefHamza(self):
        return self.options.StrictInitialAlefHamza

    @StrictInitialAlefHamza.setter
    def StrictInitialAlefHamza(self, value):
        self.options.StrictInitialAlefHamza = value

    @property
    def StrictRussianE(self):
        return self.options.StrictRussianE

    @StrictRussianE.setter
    def StrictRussianE(self, value):
        self.options.StrictRussianE = value

    @property
    def StrictTaaMarboota(self):
        return self.options.StrictTaaMarboota

    @StrictTaaMarboota.setter
    def StrictTaaMarboota(self, value):
        self.options.StrictTaaMarboota = value

    @property
    def SuggestFromMainDictionaryOnly(self):
        return self.options.SuggestFromMainDictionaryOnly

    @SuggestFromMainDictionaryOnly.setter
    def SuggestFromMainDictionaryOnly(self, value):
        self.options.SuggestFromMainDictionaryOnly = value

    @property
    def SuggestSpellingCorrections(self):
        return self.options.SuggestSpellingCorrections

    @SuggestSpellingCorrections.setter
    def SuggestSpellingCorrections(self, value):
        self.options.SuggestSpellingCorrections = value

    @property
    def TabIndentKey(self):
        return self.options.TabIndentKey

    @TabIndentKey.setter
    def TabIndentKey(self, value):
        self.options.TabIndentKey = value

    @property
    def TypeNReplace(self):
        return self.options.TypeNReplace

    @TypeNReplace.setter
    def TypeNReplace(self, value):
        self.options.TypeNReplace = value

    @property
    def UpdateFieldsAtPrint(self):
        return self.options.UpdateFieldsAtPrint

    @UpdateFieldsAtPrint.setter
    def UpdateFieldsAtPrint(self, value):
        self.options.UpdateFieldsAtPrint = value

    @property
    def UpdateFieldsWithTrackedChangesAtPrint(self):
        return self.options.UpdateFieldsWithTrackedChangesAtPrint

    @UpdateFieldsWithTrackedChangesAtPrint.setter
    def UpdateFieldsWithTrackedChangesAtPrint(self, value):
        self.options.UpdateFieldsWithTrackedChangesAtPrint = value

    @property
    def UpdateLinksAtOpen(self):
        return self.options.UpdateLinksAtOpen

    @UpdateLinksAtOpen.setter
    def UpdateLinksAtOpen(self, value):
        self.options.UpdateLinksAtOpen = value

    @property
    def UpdateLinksAtPrint(self):
        return self.options.UpdateLinksAtPrint

    @UpdateLinksAtPrint.setter
    def UpdateLinksAtPrint(self, value):
        self.options.UpdateLinksAtPrint = value

    @property
    def UpdateStyleListBehavior(self):
        return self.options.UpdateStyleListBehavior

    @UpdateStyleListBehavior.setter
    def UpdateStyleListBehavior(self, value):
        self.options.UpdateStyleListBehavior = value

    @property
    def UseCharacterUnit(self):
        return self.options.UseCharacterUnit

    @UseCharacterUnit.setter
    def UseCharacterUnit(self, value):
        self.options.UseCharacterUnit = value

    @property
    def UseDiffDiacColor(self):
        return self.options.UseDiffDiacColor

    @UseDiffDiacColor.setter
    def UseDiffDiacColor(self, value):
        self.options.UseDiffDiacColor = value

    @property
    def UseGermanSpellingReform(self):
        return self.options.UseGermanSpellingReform

    @UseGermanSpellingReform.setter
    def UseGermanSpellingReform(self, value):
        self.options.UseGermanSpellingReform = value

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

    @WarnBeforeSavingPrintingSendingMarkup.setter
    def WarnBeforeSavingPrintingSendingMarkup(self, value):
        self.options.WarnBeforeSavingPrintingSendingMarkup = value


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

    @BookFoldPrinting.setter
    def BookFoldPrinting(self, value):
        self.pagesetup.BookFoldPrinting = value

    @property
    def BookFoldPrintingSheets(self):
        return self.pagesetup.BookFoldPrintingSheets

    @BookFoldPrintingSheets.setter
    def BookFoldPrintingSheets(self, value):
        self.pagesetup.BookFoldPrintingSheets = value

    @property
    def BookFoldRevPrinting(self):
        return self.pagesetup.BookFoldRevPrinting

    @BookFoldRevPrinting.setter
    def BookFoldRevPrinting(self, value):
        self.pagesetup.BookFoldRevPrinting = value

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

    @DifferentFirstPageHeaderFooter.setter
    def DifferentFirstPageHeaderFooter(self, value):
        self.pagesetup.DifferentFirstPageHeaderFooter = value

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

    @MirrorMargins.setter
    def MirrorMargins(self, value):
        self.pagesetup.MirrorMargins = value

    @property
    def OddAndEvenPagesHeaderFooter(self):
        return self.pagesetup.OddAndEvenPagesHeaderFooter

    @OddAndEvenPagesHeaderFooter.setter
    def OddAndEvenPagesHeaderFooter(self, value):
        self.pagesetup.OddAndEvenPagesHeaderFooter = value

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

    @SuppressEndnotes.setter
    def SuppressEndnotes(self, value):
        self.pagesetup.SuppressEndnotes = value

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

    @TwoPagesOnOne.setter
    def TwoPagesOnOne(self, value):
        self.pagesetup.TwoPagesOnOne = value

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

    @DisplayRulers.setter
    def DisplayRulers(self, value):
        self.pane.DisplayRulers = value

    @property
    def DisplayVerticalRuler(self):
        return self.pane.DisplayVerticalRuler

    @DisplayVerticalRuler.setter
    def DisplayVerticalRuler(self, value):
        self.pane.DisplayVerticalRuler = value

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

    def AutoScroll(self, Velocity=None):
        arguments = com_arguments([Velocity])
        self.pane.AutoScroll(*arguments)

    def Close(self):
        self.pane.Close()

    def LargeScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = com_arguments([Down, Up, ToRight, ToLeft])
        self.pane.LargeScroll(*arguments)

    def NewFrameset(self):
        self.pane.NewFrameset()

    def PageScroll(self, Down=None, Up=None):
        arguments = com_arguments([Down, Up])
        self.pane.PageScroll(*arguments)

    def SmallScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = com_arguments([Down, Up, ToRight, ToLeft])
        self.pane.SmallScroll(*arguments)

    def TOCInFrameset(self):
        self.pane.TOCInFrameset()


class Paragraph:

    def __init__(self, paragraph=None):
        self.paragraph = paragraph

    @property
    def AddSpaceBetweenFarEastAndAlpha(self):
        return self.paragraph.AddSpaceBetweenFarEastAndAlpha

    @AddSpaceBetweenFarEastAndAlpha.setter
    def AddSpaceBetweenFarEastAndAlpha(self, value):
        self.paragraph.AddSpaceBetweenFarEastAndAlpha = value

    @property
    def AddSpaceBetweenFarEastAndDigit(self):
        return self.paragraph.AddSpaceBetweenFarEastAndDigit

    @AddSpaceBetweenFarEastAndDigit.setter
    def AddSpaceBetweenFarEastAndDigit(self, value):
        self.paragraph.AddSpaceBetweenFarEastAndDigit = value

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

    @AutoAdjustRightIndent.setter
    def AutoAdjustRightIndent(self, value):
        self.paragraph.AutoAdjustRightIndent = value

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

    @DisableLineHeightGrid.setter
    def DisableLineHeightGrid(self, value):
        self.paragraph.DisableLineHeightGrid = value

    @property
    def DropCap(self):
        return DropCap(self.paragraph.DropCap)

    @property
    def FarEastLineBreakControl(self):
        return self.paragraph.FarEastLineBreakControl

    @FarEastLineBreakControl.setter
    def FarEastLineBreakControl(self, value):
        self.paragraph.FarEastLineBreakControl = value

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

    @HalfWidthPunctuationOnTopOfLine.setter
    def HalfWidthPunctuationOnTopOfLine(self, value):
        self.paragraph.HalfWidthPunctuationOnTopOfLine = value

    @property
    def HangingPunctuation(self):
        return self.paragraph.HangingPunctuation

    @HangingPunctuation.setter
    def HangingPunctuation(self, value):
        self.paragraph.HangingPunctuation = value

    @property
    def Hyphenation(self):
        return self.paragraph.Hyphenation

    @Hyphenation.setter
    def Hyphenation(self, value):
        self.paragraph.Hyphenation = value

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

    @KeepTogether.setter
    def KeepTogether(self, value):
        self.paragraph.KeepTogether = value

    @property
    def KeepWithNext(self):
        return self.paragraph.KeepWithNext

    @KeepWithNext.setter
    def KeepWithNext(self, value):
        self.paragraph.KeepWithNext = value

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

    def ListNumberOriginal(self, Level=None):
        arguments = com_arguments([Level])
        if callable(self.paragraph.ListNumberOriginal):
            return self.paragraph.ListNumberOriginal(*arguments)
        else:
            return self.paragraph.GetListNumberOriginal(*arguments)

    @property
    def MirrorIndents(self):
        return self.paragraph.MirrorIndents

    @MirrorIndents.setter
    def MirrorIndents(self, value):
        self.paragraph.MirrorIndents = value

    @property
    def NoLineNumber(self):
        return self.paragraph.NoLineNumber

    @NoLineNumber.setter
    def NoLineNumber(self, value):
        self.paragraph.NoLineNumber = value

    @property
    def OutlineLevel(self):
        return WdOutlineLevel(self.paragraph.OutlineLevel)

    @OutlineLevel.setter
    def OutlineLevel(self, value):
        self.paragraph.OutlineLevel = value

    @property
    def PageBreakBefore(self):
        return self.paragraph.PageBreakBefore

    @PageBreakBefore.setter
    def PageBreakBefore(self, value):
        self.paragraph.PageBreakBefore = value

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

    @SpaceAfterAuto.setter
    def SpaceAfterAuto(self, value):
        self.paragraph.SpaceAfterAuto = value

    @property
    def SpaceBefore(self):
        return self.paragraph.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.paragraph.SpaceBefore = value

    @property
    def SpaceBeforeAuto(self):
        return self.paragraph.SpaceBeforeAuto

    @SpaceBeforeAuto.setter
    def SpaceBeforeAuto(self, value):
        self.paragraph.SpaceBeforeAuto = value

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

    @WidowControl.setter
    def WidowControl(self, value):
        self.paragraph.WidowControl = value

    @property
    def WordWrap(self):
        return self.paragraph.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.paragraph.WordWrap = value

    def CloseUp(self):
        self.paragraph.CloseUp()

    def Indent(self):
        self.paragraph.Indent()

    def IndentCharWidth(self, Count=None):
        arguments = com_arguments([Count])
        self.paragraph.IndentCharWidth(*arguments)

    def IndentFirstLineCharWidth(self, Count=None):
        arguments = com_arguments([Count])
        self.paragraph.IndentFirstLineCharWidth(*arguments)

    def JoinList(self):
        self.paragraph.JoinList()

    def ListAdvanceTo(self, Level1=None, Level2=None, Level3=None, Level4=None, Level5=None, Level6=None, Level7=None, Level8=None, Level9=None):
        arguments = com_arguments([Level1, Level2, Level3, Level4, Level5, Level6, Level7, Level8, Level9])
        self.paragraph.ListAdvanceTo(*arguments)

    def Next(self, Count=None):
        arguments = com_arguments([Count])
        return self.paragraph.Next(*arguments)

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

    def Previous(self, Count=None):
        arguments = com_arguments([Count])
        return self.paragraph.Previous(*arguments)

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

    def TabHangingIndent(self, Count=None):
        arguments = com_arguments([Count])
        self.paragraph.TabHangingIndent(*arguments)

    def TabIndent(self, Count=None):
        arguments = com_arguments([Count])
        self.paragraph.TabIndent(*arguments)


class ParagraphFormat:

    def __init__(self, paragraphformat=None):
        self.paragraphformat = paragraphformat

    @property
    def AddSpaceBetweenFarEastAndAlpha(self):
        return self.paragraphformat.AddSpaceBetweenFarEastAndAlpha

    @AddSpaceBetweenFarEastAndAlpha.setter
    def AddSpaceBetweenFarEastAndAlpha(self, value):
        self.paragraphformat.AddSpaceBetweenFarEastAndAlpha = value

    @property
    def AddSpaceBetweenFarEastAndDigit(self):
        return self.paragraphformat.AddSpaceBetweenFarEastAndDigit

    @AddSpaceBetweenFarEastAndDigit.setter
    def AddSpaceBetweenFarEastAndDigit(self, value):
        self.paragraphformat.AddSpaceBetweenFarEastAndDigit = value

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

    @AutoAdjustRightIndent.setter
    def AutoAdjustRightIndent(self, value):
        self.paragraphformat.AutoAdjustRightIndent = value

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

    @DisableLineHeightGrid.setter
    def DisableLineHeightGrid(self, value):
        self.paragraphformat.DisableLineHeightGrid = value

    @property
    def Duplicate(self):
        return ParagraphFormat(self.paragraphformat.Duplicate)

    @property
    def FarEastLineBreakControl(self):
        return self.paragraphformat.FarEastLineBreakControl

    @FarEastLineBreakControl.setter
    def FarEastLineBreakControl(self, value):
        self.paragraphformat.FarEastLineBreakControl = value

    @property
    def FirstLineIndent(self):
        return self.paragraphformat.FirstLineIndent

    @FirstLineIndent.setter
    def FirstLineIndent(self, value):
        self.paragraphformat.FirstLineIndent = value

    @property
    def HalfWidthPunctuationOnTopOfLine(self):
        return self.paragraphformat.HalfWidthPunctuationOnTopOfLine

    @HalfWidthPunctuationOnTopOfLine.setter
    def HalfWidthPunctuationOnTopOfLine(self, value):
        self.paragraphformat.HalfWidthPunctuationOnTopOfLine = value

    @property
    def HangingPunctuation(self):
        return self.paragraphformat.HangingPunctuation

    @HangingPunctuation.setter
    def HangingPunctuation(self, value):
        self.paragraphformat.HangingPunctuation = value

    @property
    def Hyphenation(self):
        return self.paragraphformat.Hyphenation

    @Hyphenation.setter
    def Hyphenation(self, value):
        self.paragraphformat.Hyphenation = value

    @property
    def KeepTogether(self):
        return self.paragraphformat.KeepTogether

    @KeepTogether.setter
    def KeepTogether(self, value):
        self.paragraphformat.KeepTogether = value

    @property
    def KeepWithNext(self):
        return self.paragraphformat.KeepWithNext

    @KeepWithNext.setter
    def KeepWithNext(self, value):
        self.paragraphformat.KeepWithNext = value

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

    @NoLineNumber.setter
    def NoLineNumber(self, value):
        self.paragraphformat.NoLineNumber = value

    @property
    def OutlineLevel(self):
        return WdOutlineLevel(self.paragraphformat.OutlineLevel)

    @OutlineLevel.setter
    def OutlineLevel(self, value):
        self.paragraphformat.OutlineLevel = value

    @property
    def PageBreakBefore(self):
        return self.paragraphformat.PageBreakBefore

    @PageBreakBefore.setter
    def PageBreakBefore(self, value):
        self.paragraphformat.PageBreakBefore = value

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

    @SpaceAfterAuto.setter
    def SpaceAfterAuto(self, value):
        self.paragraphformat.SpaceAfterAuto = value

    @property
    def SpaceBefore(self):
        return self.paragraphformat.SpaceBefore

    @SpaceBefore.setter
    def SpaceBefore(self, value):
        self.paragraphformat.SpaceBefore = value

    @property
    def SpaceBeforeAuto(self):
        return self.paragraphformat.SpaceBeforeAuto

    @SpaceBeforeAuto.setter
    def SpaceBeforeAuto(self, value):
        self.paragraphformat.SpaceBeforeAuto = value

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

    @WidowControl.setter
    def WidowControl(self, value):
        self.paragraphformat.WidowControl = value

    @property
    def WordWrap(self):
        return self.paragraphformat.WordWrap

    @WordWrap.setter
    def WordWrap(self, value):
        self.paragraphformat.WordWrap = value

    def CloseUp(self):
        self.paragraphformat.CloseUp()

    def IndentCharWidth(self, Count=None):
        arguments = com_arguments([Count])
        self.paragraphformat.IndentCharWidth(*arguments)

    def IndentFirstLineCharWidth(self, Count=None):
        arguments = com_arguments([Count])
        self.paragraphformat.IndentFirstLineCharWidth(*arguments)

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

    def TabHangingIndent(self, Count=None):
        arguments = com_arguments([Count])
        self.paragraphformat.TabHangingIndent(*arguments)

    def TabIndent(self, Count=None):
        arguments = com_arguments([Count])
        self.paragraphformat.TabIndent(*arguments)


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

    @TransparentBackground.setter
    def TransparentBackground(self, value):
        self.pictureformat.TransparentBackground = value

    def IncrementBrightness(self, Increment=None):
        arguments = com_arguments([Increment])
        self.pictureformat.IncrementBrightness(*arguments)

    def IncrementContrast(self, Increment=None):
        arguments = com_arguments([Increment])
        self.pictureformat.IncrementContrast(*arguments)


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

    @ApplyPictToEnd.setter
    def ApplyPictToEnd(self, value):
        self.point.ApplyPictToEnd = value

    @property
    def ApplyPictToFront(self):
        return self.point.ApplyPictToFront

    @ApplyPictToFront.setter
    def ApplyPictToFront(self, value):
        self.point.ApplyPictToFront = value

    @property
    def ApplyPictToSides(self):
        return self.point.ApplyPictToSides

    @ApplyPictToSides.setter
    def ApplyPictToSides(self, value):
        self.point.ApplyPictToSides = value

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

    @Has3DEffect.setter
    def Has3DEffect(self, value):
        self.point.Has3DEffect = value

    @property
    def HasDataLabel(self):
        return self.point.HasDataLabel

    @HasDataLabel.setter
    def HasDataLabel(self, value):
        self.point.HasDataLabel = value

    @property
    def Height(self):
        return self.point.Height

    @property
    def InvertIfNegative(self):
        return self.point.InvertIfNegative

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.point.InvertIfNegative = value

    @property
    def Left(self):
        return self.point.Left

    @property
    def MarkerBackgroundColor(self):
        return self.point.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.point.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return self.point.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.point.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.point.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.point.MarkerForegroundColor = value

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

    @SecondaryPlot.setter
    def SecondaryPlot(self, value):
        self.point.SecondaryPlot = value

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

    def ApplyDataLabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = com_arguments([Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator])
        self.point.ApplyDataLabels(*arguments)

    def ClearFormats(self):
        self.point.ClearFormats()

    def Copy(self):
        self.point.Copy()

    def Delete(self):
        self.point.Delete()

    def Paste(self):
        self.point.Paste()

    def PieSliceLocation(self, loc=None, Index=None):
        arguments = com_arguments([loc, Index])
        return self.point.PieSliceLocation(*arguments)

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Point(self.points.Item(*arguments))


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

    @Visible.setter
    def Visible(self, value):
        self.protectedviewwindow.Visible = value

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.protectedviewwindows.Item(*arguments)

    def Open(self, FileName=None, AddToRecentFiles=None, PasswordDocument=None, Visible=None, OpenAndRepair=None):
        arguments = com_arguments([FileName, AddToRecentFiles, PasswordDocument, Visible, OpenAndRepair])
        return ProtectedViewWindow(self.protectedviewwindows.Open(*arguments))


class Range:

    def __init__(self, range=None):
        self.range = range

    @property
    def Application(self):
        return Application(self.range.Application)

    @property
    def Bold(self):
        return self.range.Bold

    @Bold.setter
    def Bold(self, value):
        self.range.Bold = value

    @property
    def BoldBi(self):
        return self.range.BoldBi

    @BoldBi.setter
    def BoldBi(self, value):
        self.range.BoldBi = value

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

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([RowIndex, ColumnIndex])
        if callable(self.range.Cells):
            return self.range.Cells(*arguments)
        else:
            return self.range.GetCells(*arguments)

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

    @CombineCharacters.setter
    def CombineCharacters(self, value):
        self.range.CombineCharacters = value

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

    @DisableCharacterSpaceGrid.setter
    def DisableCharacterSpaceGrid(self, value):
        self.range.DisableCharacterSpaceGrid = value

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

    @GrammarChecked.setter
    def GrammarChecked(self, value):
        self.range.GrammarChecked = value

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

    def Information(self, Type=None):
        arguments = com_arguments([Type])
        if callable(self.range.Information):
            return self.range.Information(*arguments)
        else:
            return self.range.GetInformation(*arguments)

    @property
    def InlineShapes(self):
        return self.range.InlineShapes

    @property
    def IsEndOfRowMark(self):
        return self.range.IsEndOfRowMark

    @property
    def Italic(self):
        return self.range.Italic

    @Italic.setter
    def Italic(self, value):
        self.range.Italic = value

    @property
    def ItalicBi(self):
        return self.range.ItalicBi

    @ItalicBi.setter
    def ItalicBi(self, value):
        self.range.ItalicBi = value

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

    @NoProofing.setter
    def NoProofing(self, value):
        self.range.NoProofing = value

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

    @ShowAll.setter
    def ShowAll(self, value):
        self.range.ShowAll = value

    @property
    def SpellingChecked(self):
        return self.range.SpellingChecked

    @SpellingChecked.setter
    def SpellingChecked(self, value):
        self.range.SpellingChecked = value

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

    @TextRetrievalMode.setter
    def TextRetrievalMode(self, value):
        self.range.TextRetrievalMode = value

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

    def XML(self, DataOnly=None):
        arguments = com_arguments([DataOnly])
        if callable(self.range.XML):
            return self.range.XML(*arguments)
        else:
            return self.range.GetXML(*arguments)

    def AutoFormat(self):
        self.range.AutoFormat()

    def Calculate(self):
        self.range.Calculate()

    def CheckGrammar(self):
        self.range.CheckGrammar()

    def CheckSpelling(self, CustomDictionary=None, IgnoreUppercase=None, AlwaysSuggest=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = com_arguments([CustomDictionary, IgnoreUppercase, AlwaysSuggest, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10])
        self.range.CheckSpelling(*arguments)

    def CheckSynonyms(self):
        self.range.CheckSynonyms()

    def Collapse(self, Direction=None):
        arguments = com_arguments([Direction])
        self.range.Collapse(*arguments)

    def ComputeStatistics(self, Statistic=None):
        arguments = com_arguments([Statistic])
        self.range.ComputeStatistics(*arguments)

    def ConvertHangulAndHanja(self, ConversionsMode=None, FastConversion=None, CheckHangulEnding=None, EnableRecentOrdering=None, CustomDictionary=None):
        arguments = com_arguments([ConversionsMode, FastConversion, CheckHangulEnding, EnableRecentOrdering, CustomDictionary])
        self.range.ConvertHangulAndHanja(*arguments)

    def ConvertToTable(self, Separator=None, NumRows=None, NumColumns=None, InitialColumnWidth=None, Format=None, ApplyBorders=None, ApplyShading=None, ApplyFont=None, ApplyColor=None, ApplyHeadingRows=None, ApplyLastRow=None, ApplyFirstColumn=None, ApplyLastColumn=None, AutoFit=None, AutoFitBehavior=None, DefaultTableBehavior=None):
        arguments = com_arguments([Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior])
        return self.range.ConvertToTable(*arguments)

    def Copy(self):
        self.range.Copy()

    def CopyAsPicture(self):
        self.range.CopyAsPicture()

    def Cut(self):
        self.range.Cut()

    def Delete(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.range.Delete(*arguments)

    def DetectLanguage(self):
        self.range.DetectLanguage()

    def EndOf(self, Unit=None, Extend=None):
        arguments = com_arguments([Unit, Extend])
        self.range.EndOf(*arguments)

    def Expand(self, Unit=None):
        arguments = com_arguments([Unit])
        self.range.Expand(*arguments)

    def ExportAsFixedFormat(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr])
        self.range.ExportAsFixedFormat(*arguments)

    def ExportAsFixedFormat2(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, OptimizeForImageQuality=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, FixedFormatExtClassPtr])
        self.range.ExportAsFixedFormat2(*arguments)

    def ExportAsFixedFormat3(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, OptimizeForImageQuality=None, ImproveExportTagging=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, ImproveExportTagging, FixedFormatExtClassPtr])
        self.range.ExportAsFixedFormat3(*arguments)

    def ExportFragment(self, FileName=None, Format=None):
        arguments = com_arguments([FileName, Format])
        return self.range.ExportFragment(*arguments)

    def GetSpellingSuggestions(self, CustomDictionary=None, IgnoreUppercase=None, MainDictionary=None, SuggestionMode=None, CustomDictionary2=None, CustomDictionary3=None, CustomDictionary4=None, CustomDictionary5=None, CustomDictionary6=None, CustomDictionary7=None, CustomDictionary8=None, CustomDictionary9=None, CustomDictionary10=None):
        arguments = com_arguments([CustomDictionary, IgnoreUppercase, MainDictionary, SuggestionMode, CustomDictionary2, CustomDictionary3, CustomDictionary4, CustomDictionary5, CustomDictionary6, CustomDictionary7, CustomDictionary8, CustomDictionary9, CustomDictionary10])
        return self.range.GetSpellingSuggestions(*arguments)

    def GoTo(self, What=None, Which=None, Count=None, Name=None):
        arguments = com_arguments([What, Which, Count, Name])
        self.range.GoTo(*arguments)

    def GoToEditableRange(self, EditorID=None):
        arguments = com_arguments([EditorID])
        self.range.GoToEditableRange(*arguments)

    def GoToNext(self, What=None):
        arguments = com_arguments([What])
        self.range.GoToNext(*arguments)

    def GoToPrevious(self, What=None):
        arguments = com_arguments([What])
        self.range.GoToPrevious(*arguments)

    def ImportFragment(self, FileName=None, MatchDestination=None):
        arguments = com_arguments([FileName, MatchDestination])
        return self.range.ImportFragment(*arguments)

    def InRange(self, Range=None):
        arguments = com_arguments([Range])
        return self.range.InRange(*arguments)

    def InsertAfter(self, Text=None):
        arguments = com_arguments([Text])
        self.range.InsertAfter(*arguments)

    def InsertAlignmentTab(self, Alignment=None, RelativeTo=None):
        arguments = com_arguments([Alignment, RelativeTo])
        self.range.InsertAlignmentTab(*arguments)

    def InsertAutoText(self):
        self.range.InsertAutoText()

    def InsertBefore(self, Text=None):
        arguments = com_arguments([Text])
        self.range.InsertBefore(*arguments)

    def InsertBreak(self, Type=None):
        arguments = com_arguments([Type])
        self.range.InsertBreak(*arguments)

    def InsertCaption(self, Label=None, Title=None, TitleAutoText=None, Position=None, ExcludeLabel=None):
        arguments = com_arguments([Label, Title, TitleAutoText, Position, ExcludeLabel])
        self.range.InsertCaption(*arguments)

    def InsertCrossReference(self, ReferenceType=None, ReferenceKind=None, ReferenceItem=None, InsertAsHyperlink=None, IncludePosition=None, SeparateNumbers=None, SeparatorString=None):
        arguments = com_arguments([ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString])
        self.range.InsertCrossReference(*arguments)

    def InsertDatabase(self, Format=None, Style=None, LinkToSource=None, Connection=None, SQLStatement=None, SQLStatement1=None, PasswordDocument=None, PasswordTemplate=None, WritePasswordDocument=None, WritePasswordTemplate=None, DataSource=None, From=None, To=None, IncludeFields=None):
        arguments = com_arguments([Format, Style, LinkToSource, Connection, SQLStatement, SQLStatement1, PasswordDocument, PasswordTemplate, WritePasswordDocument, WritePasswordTemplate, DataSource, From, To, IncludeFields])
        self.range.InsertDatabase(*arguments)

    def InsertDateTime(self, DateTimeFormat=None, InsertAsField=None, InsertAsFullWidth=None, DateLanguage=None, CalendarType=None):
        arguments = com_arguments([DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType])
        self.range.InsertDateTime(*arguments)

    def InsertFile(self, FileName=None, Range=None, ConfirmConversions=None, Link=None, Attachment=None):
        arguments = com_arguments([FileName, Range, ConfirmConversions, Link, Attachment])
        self.range.InsertFile(*arguments)

    def InsertParagraph(self):
        self.range.InsertParagraph()

    def InsertParagraphAfter(self):
        self.range.InsertParagraphAfter()

    def InsertParagraphBefore(self):
        self.range.InsertParagraphBefore()

    def InsertSymbol(self, CharacterNumber=None, Font=None, Unicode=None, Bias=None):
        arguments = com_arguments([CharacterNumber, Font, Unicode, Bias])
        self.range.InsertSymbol(*arguments)

    def InsertXML(self, XML=None, Transform=None):
        arguments = com_arguments([XML, Transform])
        return self.range.InsertXML(*arguments)

    def InStory(self, Range=None):
        arguments = com_arguments([Range])
        return self.range.InStory(*arguments)

    def IsEqual(self, Range=None):
        arguments = com_arguments([Range])
        return self.range.IsEqual(*arguments)

    def LookupNameProperties(self):
        self.range.LookupNameProperties()

    def ModifyEnclosure(self, Style=None, Symbol=None, EnclosedText=None):
        arguments = com_arguments([Style, Symbol, EnclosedText])
        self.range.ModifyEnclosure(*arguments)

    def Move(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.range.Move(*arguments)

    def MoveEnd(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        self.range.MoveEnd(*arguments)

    def MoveEndUntil(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.range.MoveEndUntil(*arguments)

    def MoveEndWhile(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.range.MoveEndWhile(*arguments)

    def MoveStart(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.range.MoveStart(*arguments)

    def MoveStartUntil(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.range.MoveStartUntil(*arguments)

    def MoveStartWhile(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.range.MoveStartWhile(*arguments)

    def MoveUntil(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.range.MoveUntil(*arguments)

    def MoveWhile(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.range.MoveWhile(*arguments)

    def Next(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.range.Next(*arguments)

    def NextSubdocument(self):
        self.range.NextSubdocument()

    def Paste(self):
        self.range.Paste()

    def PasteAndFormat(self, Type=None):
        arguments = com_arguments([Type])
        self.range.PasteAndFormat(*arguments)

    def PasteAppendTable(self):
        self.range.PasteAppendTable()

    def PasteAsNestedTable(self):
        self.range.PasteAsNestedTable()

    def PasteExcelTable(self, LinkedToExcel=None, WordFormatting=None, RTF=None):
        arguments = com_arguments([LinkedToExcel, WordFormatting, RTF])
        self.range.PasteExcelTable(*arguments)

    def PasteSpecial(self, IconIndex=None, Link=None, Placement=None, DisplayAsIcon=None, DataType=None, IconFileName=None, IconLabel=None):
        arguments = com_arguments([IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel])
        self.range.PasteSpecial(*arguments)

    def PhoneticGuide(self, Text=None, Alignment=None, Raise=None, FontSize=None, FontName=None):
        arguments = com_arguments([Text, Alignment, Raise, FontSize, FontName])
        self.range.PhoneticGuide(*arguments)

    def Previous(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.range.Previous(*arguments)

    def PreviousSubdocument(self):
        self.range.PreviousSubdocument()

    def Relocate(self, Direction=None):
        arguments = com_arguments([Direction])
        self.range.Relocate(*arguments)

    def Select(self):
        self.range.Select()

    def SetListLevel(self, Level=None):
        arguments = com_arguments([Level])
        self.range.SetListLevel(*arguments)

    def SetRange(self, Start=None, End=None):
        arguments = com_arguments([Start, End])
        self.range.SetRange(*arguments)

    def Sort(self, ExcludeHeader=None, FieldNumber=None, SortFieldType=None, SortOrder=None, FieldNumber2=None, SortFieldType2=None, SortOrder2=None, FieldNumber3=None, SortFieldType3=None, SortOrder3=None, SortColumn=None, Separator=None, CaseSensitive=None, BidiSort=None, IgnoreThe=None, IgnoreKashida=None, IgnoreDiacritics=None, IgnoreHe=None, LanguageID=None):
        arguments = com_arguments([ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID])
        self.range.Sort(*arguments)

    def SortAscending(self):
        self.range.SortAscending()

    def SortDescending(self):
        self.range.SortDescending()

    def StartOf(self, Unit=None, Extend=None):
        arguments = com_arguments([Unit, Extend])
        self.range.StartOf(*arguments)

    def TCSCConverter(self, WdTCSCConverterDirection=None, CommonTerms=None, UseVariants=None):
        arguments = com_arguments([WdTCSCConverterDirection, CommonTerms, UseVariants])
        self.range.TCSCConverter(*arguments)

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

    @ReadOnly.setter
    def ReadOnly(self, value):
        self.recentfile.ReadOnly = value

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.rectangles.Item(*arguments)


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

    @Highlight.setter
    def Highlight(self, value):
        self.replacement.Highlight = value

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

    @NoProofing.setter
    def NoProofing(self, value):
        self.replacement.NoProofing = value

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

    def IsResearchService(self, ServiceID=None):
        arguments = com_arguments([ServiceID])
        return self.research.IsResearchService(*arguments)

    def Query(self, ServiceID=None, QueryString=None, QueryLanguage=None, UseSelection=None, RequeryContextXML=None, NewQueryContextXML=None, LaunchQuery=None):
        arguments = com_arguments([ServiceID, QueryString, QueryLanguage, UseSelection, RequeryContextXML, NewQueryContextXML, LaunchQuery])
        return self.research.Query(*arguments)

    def SetLanguagePair(self, LanguageFrom=None, LanguageTo=None):
        arguments = com_arguments([LanguageFrom, LanguageTo])
        return self.research.SetLanguagePair(*arguments)


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

    @Visible.setter
    def Visible(self, value):
        self.reviewer.Visible = value


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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.reviewers.Item(*arguments)


class Revision:

    def __init__(self, revision=None):
        self.revision = revision

    @property
    def Application(self):
        return Application(self.revision.Application)

    @property
    def Author(self):
        return self.revision.Author

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([RowIndex, ColumnIndex])
        if callable(self.revision.Cells):
            return self.revision.Cells(*arguments)
        else:
            return self.revision.GetCells(*arguments)

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

    @AllowBreakAcrossPages.setter
    def AllowBreakAcrossPages(self, value):
        self.row.AllowBreakAcrossPages = value

    @property
    def Application(self):
        return Application(self.row.Application)

    @property
    def Borders(self):
        return self.row.Borders

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([RowIndex, ColumnIndex])
        if callable(self.row.Cells):
            return self.row.Cells(*arguments)
        else:
            return self.row.GetCells(*arguments)

    @property
    def Creator(self):
        return self.row.Creator

    @property
    def HeadingFormat(self):
        return self.row.HeadingFormat

    @HeadingFormat.setter
    def HeadingFormat(self, value):
        self.row.HeadingFormat = value

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

    def ConvertToText(self, Separator=None, NestedTables=None):
        arguments = com_arguments([Separator, NestedTables])
        self.row.ConvertToText(*arguments)

    def Delete(self):
        self.row.Delete()

    def Select(self):
        self.row.Select()

    def SetHeight(self, RowHeight=None, HeightRule=None):
        arguments = com_arguments([RowHeight, HeightRule])
        self.row.SetHeight(*arguments)

    def SetLeftIndent(self, LeftIndent=None, RulerStyle=None):
        arguments = com_arguments([LeftIndent, RulerStyle])
        self.row.SetLeftIndent(*arguments)


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

    @ProtectedForForms.setter
    def ProtectedForForms(self, value):
        self.section.ProtectedForForms = value

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

    def Cells(self, RowIndex=None, ColumnIndex=None):
        arguments = com_arguments([RowIndex, ColumnIndex])
        if callable(self.selection.Cells):
            return self.selection.Cells(*arguments)
        else:
            return self.selection.GetCells(*arguments)

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

    @ColumnSelectMode.setter
    def ColumnSelectMode(self, value):
        self.selection.ColumnSelectMode = value

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

    @ExtendMode.setter
    def ExtendMode(self, value):
        self.selection.ExtendMode = value

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

    def Information(self, Type=None):
        arguments = com_arguments([Type])
        if callable(self.selection.Information):
            return self.selection.Information(*arguments)
        else:
            return self.selection.GetInformation(*arguments)

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

    @NoProofing.setter
    def NoProofing(self, value):
        self.selection.NoProofing = value

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

    @StartIsActive.setter
    def StartIsActive(self, value):
        self.selection.StartIsActive = value

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

    def XML(self, DataOnly=None):
        arguments = com_arguments([DataOnly])
        if callable(self.selection.XML):
            return self.selection.XML(*arguments)
        else:
            return self.selection.GetXML(*arguments)

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

    def Collapse(self, Direction=None):
        arguments = com_arguments([Direction])
        self.selection.Collapse(*arguments)

    def ConvertToTable(self, Separator=None, NumRows=None, NumColumns=None, InitialColumnWidth=None, Format=None, ApplyBorders=None, ApplyShading=None, ApplyFont=None, ApplyColor=None, ApplyHeadingRows=None, ApplyLastRow=None, ApplyFirstColumn=None, ApplyLastColumn=None, AutoFit=None, AutoFitBehavior=None, DefaultTableBehavior=None):
        arguments = com_arguments([Separator, NumRows, NumColumns, InitialColumnWidth, Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit, AutoFitBehavior, DefaultTableBehavior])
        return self.selection.ConvertToTable(*arguments)

    def Copy(self):
        self.selection.Copy()

    def CopyAsPicture(self):
        self.selection.CopyAsPicture()

    def CopyFormat(self):
        self.selection.CopyFormat()

    def CreateAutoTextEntry(self, Name=None, StyleName=None):
        arguments = com_arguments([Name, StyleName])
        self.selection.CreateAutoTextEntry(*arguments)

    def CreateTextbox(self):
        self.selection.CreateTextbox()

    def Cut(self):
        self.selection.Cut()

    def Delete(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.selection.Delete(*arguments)

    def DetectLanguage(self):
        self.selection.DetectLanguage()

    def EndKey(self, Unit=None, Extend=None):
        arguments = com_arguments([Unit, Extend])
        self.selection.EndKey(*arguments)

    def EndOf(self, Unit=None, Extend=None):
        arguments = com_arguments([Unit, Extend])
        self.selection.EndOf(*arguments)

    def EscapeKey(self):
        self.selection.EscapeKey()

    def Expand(self, Unit=None):
        arguments = com_arguments([Unit])
        self.selection.Expand(*arguments)

    def ExportAsFixedFormat(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, FixedFormatExtClassPtr])
        self.selection.ExportAsFixedFormat(*arguments)

    def ExportAsFixedFormat2(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, OptimizeForImageQuality=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, FixedFormatExtClassPtr])
        self.selection.ExportAsFixedFormat2(*arguments)

    def ExportAsFixedFormat3(self, OutputFileName=None, ExportFormat=None, OpenAfterExport=None, OptimizeFor=None, ExportCurrentPage=None, Item=None, IncludeDocProps=None, KeepIRM=None, CreateBookmarks=None, DocStructureTags=None, BitmapMissingFonts=None, UseISO19005_1=None, OptimizeForImageQuality=None, ImproveExportTagging=None, FixedFormatExtClassPtr=None):
        arguments = com_arguments([OutputFileName, ExportFormat, OpenAfterExport, OptimizeFor, ExportCurrentPage, Item, IncludeDocProps, KeepIRM, CreateBookmarks, DocStructureTags, BitmapMissingFonts, UseISO19005_1, OptimizeForImageQuality, ImproveExportTagging, FixedFormatExtClassPtr])
        self.selection.ExportAsFixedFormat3(*arguments)

    def Extend(self, Character=None):
        arguments = com_arguments([Character])
        self.selection.Extend(*arguments)

    def GoTo(self, What=None, Which=None, Count=None, Name=None):
        arguments = com_arguments([What, Which, Count, Name])
        return self.selection.GoTo(*arguments)

    def GoToEditableRange(self, EditorID=None):
        arguments = com_arguments([EditorID])
        return self.selection.GoToEditableRange(*arguments)

    def GoToNext(self, What=None):
        arguments = com_arguments([What])
        self.selection.GoToNext(*arguments)

    def GoToPrevious(self, What=None):
        arguments = com_arguments([What])
        self.selection.GoToPrevious(*arguments)

    def HomeKey(self, Unit=None, Extend=None):
        arguments = com_arguments([Unit, Extend])
        self.selection.HomeKey(*arguments)

    def InRange(self, Range=None):
        arguments = com_arguments([Range])
        return self.selection.InRange(*arguments)

    def InsertAfter(self, Text=None):
        arguments = com_arguments([Text])
        self.selection.InsertAfter(*arguments)

    def InsertBefore(self, Text=None):
        arguments = com_arguments([Text])
        self.selection.InsertBefore(*arguments)

    def InsertBreak(self, Type=None):
        arguments = com_arguments([Type])
        self.selection.InsertBreak(*arguments)

    def InsertCaption(self, Label=None, Title=None, TitleAutoText=None, Position=None, ExcludeLabel=None):
        arguments = com_arguments([Label, Title, TitleAutoText, Position, ExcludeLabel])
        self.selection.InsertCaption(*arguments)

    def InsertCells(self, ShiftCells=None):
        arguments = com_arguments([ShiftCells])
        self.selection.InsertCells(*arguments)

    def InsertColumns(self):
        self.selection.InsertColumns()

    def InsertColumnsRight(self):
        self.selection.InsertColumnsRight()

    def InsertCrossReference(self, ReferenceType=None, ReferenceKind=None, ReferenceItem=None, InsertAsHyperlink=None, IncludePosition=None, SeparateNumbers=None, SeparatorString=None):
        arguments = com_arguments([ReferenceType, ReferenceKind, ReferenceItem, InsertAsHyperlink, IncludePosition, SeparateNumbers, SeparatorString])
        self.selection.InsertCrossReference(*arguments)

    def InsertDateTime(self, DateTimeFormat=None, InsertAsField=None, InsertAsFullWidth=None, DateLanguage=None, CalendarType=None):
        arguments = com_arguments([DateTimeFormat, InsertAsField, InsertAsFullWidth, DateLanguage, CalendarType])
        self.selection.InsertDateTime(*arguments)

    def InsertFile(self, FileName=None, Range=None, ConfirmConversions=None, Link=None, Attachment=None):
        arguments = com_arguments([FileName, Range, ConfirmConversions, Link, Attachment])
        self.selection.InsertFile(*arguments)

    def InsertFormula(self, Formula=None, NumberFormat=None):
        arguments = com_arguments([Formula, NumberFormat])
        self.selection.InsertFormula(*arguments)

    def InsertNewPage(self):
        self.selection.InsertNewPage()

    def InsertParagraph(self):
        self.selection.InsertParagraph()

    def InsertParagraphAfter(self):
        self.selection.InsertParagraphAfter()

    def InsertParagraphBefore(self):
        self.selection.InsertParagraphBefore()

    def InsertRows(self, NumRows=None):
        arguments = com_arguments([NumRows])
        self.selection.InsertRows(*arguments)

    def InsertRowsAbove(self):
        self.selection.InsertRowsAbove()

    def InsertRowsBelow(self):
        self.selection.InsertRowsBelow()

    def InsertStyleSeparator(self):
        self.selection.InsertStyleSeparator()

    def InsertSymbol(self, CharacterNumber=None, Font=None, Unicode=None, Bias=None):
        arguments = com_arguments([CharacterNumber, Font, Unicode, Bias])
        self.selection.InsertSymbol(*arguments)

    def InsertXML(self, XML=None, Transform=None):
        arguments = com_arguments([XML, Transform])
        return self.selection.InsertXML(*arguments)

    def InStory(self, Range=None):
        arguments = com_arguments([Range])
        return self.selection.InStory(*arguments)

    def IsEqual(self, Range=None):
        arguments = com_arguments([Range])
        return self.selection.IsEqual(*arguments)

    def ItalicRun(self):
        self.selection.ItalicRun()

    def LtrPara(self):
        self.selection.LtrPara()

    def LtrRun(self):
        self.selection.LtrRun()

    def Move(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.selection.Move(*arguments)

    def MoveDown(self, Unit=None, Count=None, Extend=None):
        arguments = com_arguments([Unit, Count, Extend])
        self.selection.MoveDown(*arguments)

    def MoveEnd(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.selection.MoveEnd(*arguments)

    def MoveEndUntil(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        return self.selection.MoveEndUntil(*arguments)

    def MoveEndWhile(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.selection.MoveEndWhile(*arguments)

    def MoveLeft(self, Unit=None, Count=None, Extend=None):
        arguments = com_arguments([Unit, Count, Extend])
        self.selection.MoveLeft(*arguments)

    def MoveRight(self, Unit=None, Count=None, Extend=None):
        arguments = com_arguments([Unit, Count, Extend])
        return self.selection.MoveRight(*arguments)

    def MoveStart(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.selection.MoveStart(*arguments)

    def MoveStartUntil(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.selection.MoveStartUntil(*arguments)

    def MoveStartWhile(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.selection.MoveStartWhile(*arguments)

    def MoveUntil(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.selection.MoveUntil(*arguments)

    def MoveUp(self, Unit=None, Count=None, Extend=None):
        arguments = com_arguments([Unit, Count, Extend])
        return self.selection.MoveUp(*arguments)

    def MoveWhile(self, Cset=None, Count=None):
        arguments = com_arguments([Cset, Count])
        self.selection.MoveWhile(*arguments)

    def Next(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.selection.Next(*arguments)

    def NextField(self):
        return self.selection.NextField()

    def NextRevision(self, Wrap=None):
        arguments = com_arguments([Wrap])
        return self.selection.NextRevision(*arguments)

    def NextSubdocument(self):
        self.selection.NextSubdocument()

    def Paste(self):
        self.selection.Paste()

    def PasteAndFormat(self, Type=None):
        arguments = com_arguments([Type])
        self.selection.PasteAndFormat(*arguments)

    def PasteAppendTable(self):
        self.selection.PasteAppendTable()

    def PasteAsNestedTable(self):
        self.selection.PasteAsNestedTable()

    def PasteExcelTable(self, LinkedToExcel=None, WordFormatting=None, RTF=None):
        arguments = com_arguments([LinkedToExcel, WordFormatting, RTF])
        self.selection.PasteExcelTable(*arguments)

    def PasteFormat(self):
        self.selection.PasteFormat()

    def PasteSpecial(self, IconIndex=None, Link=None, Placement=None, DisplayAsIcon=None, DataType=None, IconFileName=None, IconLabel=None):
        arguments = com_arguments([IconIndex, Link, Placement, DisplayAsIcon, DataType, IconFileName, IconLabel])
        self.selection.PasteSpecial(*arguments)

    def Previous(self, Unit=None, Count=None):
        arguments = com_arguments([Unit, Count])
        return self.selection.Previous(*arguments)

    def PreviousField(self):
        return self.selection.PreviousField()

    def PreviousRevision(self, Wrap=None):
        arguments = com_arguments([Wrap])
        return self.selection.PreviousRevision(*arguments)

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

    def SetRange(self, Start=None, End=None):
        arguments = com_arguments([Start, End])
        self.selection.SetRange(*arguments)

    def Shrink(self):
        self.selection.Shrink()

    def ShrinkDiscontiguousSelection(self):
        self.selection.ShrinkDiscontiguousSelection()

    def Sort(self, ExcludeHeader=None, FieldNumber=None, SortFieldType=None, SortOrder=None, FieldNumber2=None, SortFieldType2=None, SortOrder2=None, FieldNumber3=None, SortFieldType3=None, SortOrder3=None, SortColumn=None, Separator=None, CaseSensitive=None, BidiSort=None, IgnoreThe=None, IgnoreKashida=None, IgnoreDiacritics=None, IgnoreHe=None, LanguageID=None, SubFieldNumber=None, SubFieldNumber2=None, SubFieldNumber3=None):
        arguments = com_arguments([ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, SortColumn, Separator, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID, SubFieldNumber, SubFieldNumber2, SubFieldNumber3])
        self.selection.Sort(*arguments)

    def SortAscending(self):
        self.selection.SortAscending()

    def SortDescending(self):
        self.selection.SortDescending()

    def SplitTable(self):
        self.selection.SplitTable()

    def StartOf(self, Unit=None, Extend=None):
        arguments = com_arguments([Unit, Extend])
        self.selection.StartOf(*arguments)

    def ToggleCharacterCode(self):
        self.selection.ToggleCharacterCode()

    def TypeBackspace(self):
        self.selection.TypeBackspace()

    def TypeParagraph(self):
        self.selection.TypeParagraph()

    def TypeText(self, Text=None):
        arguments = com_arguments([Text])
        self.selection.TypeText(*arguments)

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

    @ApplyPictToEnd.setter
    def ApplyPictToEnd(self, value):
        self.series.ApplyPictToEnd = value

    @property
    def ApplyPictToFront(self):
        return self.series.ApplyPictToFront

    @ApplyPictToFront.setter
    def ApplyPictToFront(self, value):
        self.series.ApplyPictToFront = value

    @property
    def ApplyPictToSides(self):
        return self.series.ApplyPictToSides

    @ApplyPictToSides.setter
    def ApplyPictToSides(self, value):
        self.series.ApplyPictToSides = value

    @property
    def AxisGroup(self):
        return self.series.AxisGroup

    @AxisGroup.setter
    def AxisGroup(self, value):
        self.series.AxisGroup = value

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

    @Has3DEffect.setter
    def Has3DEffect(self, value):
        self.series.Has3DEffect = value

    @property
    def HasDataLabels(self):
        return self.series.HasDataLabels

    @HasDataLabels.setter
    def HasDataLabels(self, value):
        self.series.HasDataLabels = value

    @property
    def HasErrorBars(self):
        return self.series.HasErrorBars

    @HasErrorBars.setter
    def HasErrorBars(self, value):
        self.series.HasErrorBars = value

    @property
    def HasLeaderLines(self):
        return self.series.HasLeaderLines

    @HasLeaderLines.setter
    def HasLeaderLines(self, value):
        self.series.HasLeaderLines = value

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

    @InvertIfNegative.setter
    def InvertIfNegative(self, value):
        self.series.InvertIfNegative = value

    @property
    def LeaderLines(self):
        return LeaderLines(self.series.LeaderLines)

    @property
    def MarkerBackgroundColor(self):
        return self.series.MarkerBackgroundColor

    @MarkerBackgroundColor.setter
    def MarkerBackgroundColor(self, value):
        self.series.MarkerBackgroundColor = value

    @property
    def MarkerBackgroundColorIndex(self):
        return self.series.MarkerBackgroundColorIndex

    @MarkerBackgroundColorIndex.setter
    def MarkerBackgroundColorIndex(self, value):
        self.series.MarkerBackgroundColorIndex = value

    @property
    def MarkerForegroundColor(self):
        return self.series.MarkerForegroundColor

    @MarkerForegroundColor.setter
    def MarkerForegroundColor(self, value):
        self.series.MarkerForegroundColor = value

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

    @Smooth.setter
    def Smooth(self, value):
        self.series.Smooth = value

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

    def ApplyDataLabels(self, Type=None, LegendKey=None, AutoText=None, HasLeaderLines=None, ShowSeriesName=None, ShowCategoryName=None, ShowValue=None, ShowPercentage=None, ShowBubbleSize=None, Separator=None):
        arguments = com_arguments([Type, LegendKey, AutoText, HasLeaderLines, ShowSeriesName, ShowCategoryName, ShowValue, ShowPercentage, ShowBubbleSize, Separator])
        self.series.ApplyDataLabels(*arguments)

    def ClearFormats(self):
        self.series.ClearFormats()

    def Copy(self):
        self.series.Copy()

    def DataLabels(self, Index=None):
        arguments = com_arguments([Index])
        return self.series.DataLabels(*arguments)

    def Delete(self):
        self.series.Delete()

    def ErrorBar(self, Direction=None, Include=None, Type=None, Amount=None, MinusValues=None):
        arguments = com_arguments([Direction, Include, Type, Amount, MinusValues])
        self.series.ErrorBar(*arguments)

    def Paste(self):
        self.series.Paste()

    def Points(self, Index=None):
        arguments = com_arguments([Index])
        return Points(self.series.Points(*arguments))

    def Select(self):
        self.series.Select()

    def Trendlines(self, Index=None):
        arguments = com_arguments([Index])
        return Trendlines(self.series.Trendlines(*arguments))


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

    def Add(self, Source=None, Rowcol=None, SeriesLabels=None, CategoryLabels=None, Replace=None):
        arguments = com_arguments([Source, Rowcol, SeriesLabels, CategoryLabels, Replace])
        return Series(self.seriescollection.Add(*arguments))

    def Extend(self, Source=None, Rowcol=None, CategoryLabels=None):
        arguments = com_arguments([Source, Rowcol, CategoryLabels])
        self.seriescollection.Extend(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Series(self.seriescollection.Item(*arguments))

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

    @Obscured.setter
    def Obscured(self, value):
        self.shadowformat.Obscured = value

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

    @Visible.setter
    def Visible(self, value):
        self.shadowformat.Visible = value

    def IncrementOffsetX(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shadowformat.IncrementOffsetX(*arguments)

    def IncrementOffsetY(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shadowformat.IncrementOffsetY(*arguments)


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

    @BackgroundStyle.setter
    def BackgroundStyle(self, value):
        self.shape.BackgroundStyle = value

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

    @Decorative.setter
    def Decorative(self, value):
        self.shape.Decorative = value

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

    @LockAnchor.setter
    def LockAnchor(self, value):
        self.shape.LockAnchor = value

    @property
    def LockAspectRatio(self):
        return self.shape.LockAspectRatio

    @LockAspectRatio.setter
    def LockAspectRatio(self, value):
        self.shape.LockAspectRatio = value

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

    @RelativeHorizontalPosition.setter
    def RelativeHorizontalPosition(self, value):
        self.shape.RelativeHorizontalPosition = value

    @property
    def RelativeHorizontalSize(self):
        return WdRelativeVerticalSize(self.shape.RelativeHorizontalSize)

    @RelativeHorizontalSize.setter
    def RelativeHorizontalSize(self, value):
        self.shape.RelativeHorizontalSize = value

    @property
    def RelativeVerticalPosition(self):
        return self.shape.RelativeVerticalPosition

    @RelativeVerticalPosition.setter
    def RelativeVerticalPosition(self, value):
        self.shape.RelativeVerticalPosition = value

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

    @Visible.setter
    def Visible(self, value):
        self.shape.Visible = value

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

    def CanvasCropBottom(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shape.CanvasCropBottom(*arguments)

    def CanvasCropLeft(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shape.CanvasCropLeft(*arguments)

    def CanvasCropRight(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shape.CanvasCropRight(*arguments)

    def CanvasCropTop(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shape.CanvasCropTop(*arguments)

    def ConvertToInlineShape(self):
        self.shape.ConvertToInlineShape()

    def Delete(self, Index=None):
        arguments = com_arguments([Index])
        self.shape.Delete(*arguments)

    def Duplicate(self):
        self.shape.Duplicate()

    def Flip(self, FlipCmd=None):
        arguments = com_arguments([FlipCmd])
        self.shape.Flip(*arguments)

    def IncrementLeft(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shape.IncrementLeft(*arguments)

    def IncrementRotation(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shape.IncrementRotation(*arguments)

    def IncrementTop(self, Increment=None):
        arguments = com_arguments([Increment])
        self.shape.IncrementTop(*arguments)

    def PickUp(self):
        self.shape.PickUp()

    def ScaleHeight(self, Factor=None, RelativeToOriginalSize=None, Scale=None):
        arguments = com_arguments([Factor, RelativeToOriginalSize, Scale])
        self.shape.ScaleHeight(*arguments)

    def ScaleWidth(self, Factor=None, RelativeToOriginalSize=None, Scale=None):
        arguments = com_arguments([Factor, RelativeToOriginalSize, Scale])
        self.shape.ScaleWidth(*arguments)

    def Select(self, Replace=None):
        arguments = com_arguments([Replace])
        self.shape.Select(*arguments)

    def SetShapesDefaultProperties(self):
        self.shape.SetShapesDefaultProperties()

    def Ungroup(self):
        return self.shape.Ungroup()

    def ZOrder(self, ZOrderCmd=None):
        arguments = com_arguments([ZOrderCmd])
        return self.shape.ZOrder(*arguments)


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

    def Field(self, Name=None):
        arguments = com_arguments([Name])
        if callable(self.source.Field):
            return self.source.Field(*arguments)
        else:
            return self.source.GetField(*arguments)

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

    def Add(self, Data=None):
        arguments = com_arguments([Data])
        self.sources.Add(*arguments)

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.sources.Item(*arguments)


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

    @AutomaticallyUpdate.setter
    def AutomaticallyUpdate(self, value):
        self.style.AutomaticallyUpdate = value

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

    @Hidden.setter
    def Hidden(self, value):
        self.style.Hidden = value

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

    @LinkStyle.setter
    def LinkStyle(self, value):
        self.style.LinkStyle = value

    @property
    def ListLevelNumber(self):
        return self.style.ListLevelNumber

    @property
    def ListTemplate(self):
        return ListTemplate(self.style.ListTemplate)

    @property
    def Locked(self):
        return self.style.Locked

    @Locked.setter
    def Locked(self, value):
        self.style.Locked = value

    @property
    def NameLocal(self):
        return self.style.NameLocal

    @NameLocal.setter
    def NameLocal(self, value):
        self.style.NameLocal = value

    @property
    def NextParagraphStyle(self):
        return self.style.NextParagraphStyle

    @NextParagraphStyle.setter
    def NextParagraphStyle(self, value):
        self.style.NextParagraphStyle = value

    @property
    def NoProofing(self):
        return self.style.NoProofing

    @NoProofing.setter
    def NoProofing(self, value):
        self.style.NoProofing = value

    @property
    def NoSpaceBetweenParagraphsOfSameStyle(self):
        return self.style.NoSpaceBetweenParagraphsOfSameStyle

    @NoSpaceBetweenParagraphsOfSameStyle.setter
    def NoSpaceBetweenParagraphsOfSameStyle(self, value):
        self.style.NoSpaceBetweenParagraphsOfSameStyle = value

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

    @UnhideWhenUsed.setter
    def UnhideWhenUsed(self, value):
        self.style.UnhideWhenUsed = value

    @property
    def Visibility(self):
        return self.style.Visibility

    @Visibility.setter
    def Visibility(self, value):
        self.style.Visibility = value

    def Delete(self):
        self.style.Delete()

    def LinkToListTemplate(self, ListTemplate=None, ListLevelNumber=None):
        arguments = com_arguments([ListTemplate, ListLevelNumber])
        self.style.LinkToListTemplate(*arguments)


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

    def Move(self, Precedence=None):
        arguments = com_arguments([Precedence])
        self.stylesheet.Move(*arguments)


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

    def Add(self, FileName=None, LinkType=None, Title=None, Precedence=None):
        arguments = com_arguments([FileName, LinkType, Title, Precedence])
        return StyleSheet(self.stylesheets.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.stylesheets.Item(*arguments)


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

    @Locked.setter
    def Locked(self, value):
        self.subdocument.Locked = value

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

    def Split(self, Range=None):
        arguments = com_arguments([Range])
        self.subdocument.Split(*arguments)


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

    def SynonymList(self, Meaning=None):
        arguments = com_arguments([Meaning])
        if callable(self.synonyminfo.SynonymList):
            return self.synonyminfo.SynonymList(*arguments)
        else:
            return self.synonyminfo.GetSynonymList(*arguments)

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

    def Connect(self, Path=None, Drive=None, Password=None):
        arguments = com_arguments([Path, Drive, Password])
        self.system.Connect(*arguments)

    def MSInfo(self):
        self.system.MSInfo()


class Table:

    def __init__(self, table=None):
        self.table = table

    @property
    def AllowAutoFit(self):
        return self.table.AllowAutoFit

    @AllowAutoFit.setter
    def AllowAutoFit(self, value):
        self.table.AllowAutoFit = value

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

    @ApplyStyleFirstColumn.setter
    def ApplyStyleFirstColumn(self, value):
        self.table.ApplyStyleFirstColumn = value

    @property
    def ApplyStyleHeadingRows(self):
        return self.table.ApplyStyleHeadingRows

    @ApplyStyleHeadingRows.setter
    def ApplyStyleHeadingRows(self, value):
        self.table.ApplyStyleHeadingRows = value

    @property
    def ApplyStyleLastColumn(self):
        return self.table.ApplyStyleLastColumn

    @ApplyStyleLastColumn.setter
    def ApplyStyleLastColumn(self, value):
        self.table.ApplyStyleLastColumn = value

    @property
    def ApplyStyleLastRow(self):
        return self.table.ApplyStyleLastRow

    @ApplyStyleLastRow.setter
    def ApplyStyleLastRow(self, value):
        self.table.ApplyStyleLastRow = value

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

    def ApplyStyleDirectFormatting(self, StyleName=None):
        arguments = com_arguments([StyleName])
        self.table.ApplyStyleDirectFormatting(*arguments)

    def AutoFitBehavior(self, Behavior=None):
        arguments = com_arguments([Behavior])
        self.table.AutoFitBehavior(*arguments)

    def AutoFormat(self, Format=None, ApplyBorders=None, ApplyShading=None, ApplyFont=None, ApplyColor=None, ApplyHeadingRows=None, ApplyLastRow=None, ApplyFirstColumn=None, ApplyLastColumn=None, AutoFit=None):
        arguments = com_arguments([Format, ApplyBorders, ApplyShading, ApplyFont, ApplyColor, ApplyHeadingRows, ApplyLastRow, ApplyFirstColumn, ApplyLastColumn, AutoFit])
        self.table.AutoFormat(*arguments)

    def Cell(self, Row=None, Column=None):
        arguments = com_arguments([Row, Column])
        return self.table.Cell(*arguments)

    def ConvertToText(self, Separator=None, NestedTables=None):
        arguments = com_arguments([Separator, NestedTables])
        self.table.ConvertToText(*arguments)

    def Delete(self):
        self.table.Delete()

    def Select(self):
        self.table.Select()

    def Sort(self, ExcludeHeader=None, FieldNumber=None, SortFieldType=None, SortOrder=None, FieldNumber2=None, SortFieldType2=None, SortOrder2=None, FieldNumber3=None, SortFieldType3=None, SortOrder3=None, CaseSensitive=None, BidiSort=None, IgnoreThe=None, IgnoreKashida=None, IgnoreDiacritics=None, IgnoreHe=None, LanguageID=None):
        arguments = com_arguments([ExcludeHeader, FieldNumber, SortFieldType, SortOrder, FieldNumber2, SortFieldType2, SortOrder2, FieldNumber3, SortFieldType3, SortOrder3, CaseSensitive, BidiSort, IgnoreThe, IgnoreKashida, IgnoreDiacritics, IgnoreHe, LanguageID])
        self.table.Sort(*arguments)

    def SortAscending(self):
        self.table.SortAscending()

    def SortDescending(self):
        self.table.SortDescending()

    def Split(self, BeforeRow=None):
        arguments = com_arguments([BeforeRow])
        return self.table.Split(*arguments)

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

    @IncludeCategoryHeader.setter
    def IncludeCategoryHeader(self, value):
        self.tableofauthorities.IncludeCategoryHeader = value

    @property
    def IncludeSequenceName(self):
        return self.tableofauthorities.IncludeSequenceName

    @IncludeSequenceName.setter
    def IncludeSequenceName(self, value):
        self.tableofauthorities.IncludeSequenceName = value

    @property
    def KeepEntryFormatting(self):
        return self.tableofauthorities.KeepEntryFormatting

    @KeepEntryFormatting.setter
    def KeepEntryFormatting(self, value):
        self.tableofauthorities.KeepEntryFormatting = value

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

    @Passim.setter
    def Passim(self, value):
        self.tableofauthorities.Passim = value

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

    @IncludePageNumbers.setter
    def IncludePageNumbers(self, value):
        self.tableofcontents.IncludePageNumbers = value

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

    @RightAlignPageNumbers.setter
    def RightAlignPageNumbers(self, value):
        self.tableofcontents.RightAlignPageNumbers = value

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

    @UseFields.setter
    def UseFields(self, value):
        self.tableofcontents.UseFields = value

    @property
    def UseHeadingStyles(self):
        return self.tableofcontents.UseHeadingStyles

    @UseHeadingStyles.setter
    def UseHeadingStyles(self, value):
        self.tableofcontents.UseHeadingStyles = value

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

    @IncludeLabel.setter
    def IncludeLabel(self, value):
        self.tableoffigures.IncludeLabel = value

    @property
    def IncludePageNumbers(self):
        return self.tableoffigures.IncludePageNumbers

    @IncludePageNumbers.setter
    def IncludePageNumbers(self, value):
        self.tableoffigures.IncludePageNumbers = value

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

    @RightAlignPageNumbers.setter
    def RightAlignPageNumbers(self, value):
        self.tableoffigures.RightAlignPageNumbers = value

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

    @UseFields.setter
    def UseFields(self, value):
        self.tableoffigures.UseFields = value

    @property
    def UseHeadingStyles(self):
        return self.tableoffigures.UseHeadingStyles

    @UseHeadingStyles.setter
    def UseHeadingStyles(self, value):
        self.tableoffigures.UseHeadingStyles = value

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

    @AllowBreakAcrossPage.setter
    def AllowBreakAcrossPage(self, value):
        self.tablestyle.AllowBreakAcrossPage = value

    @property
    def AllowPageBreaks(self):
        return self.tablestyle.AllowPageBreaks

    @AllowPageBreaks.setter
    def AllowPageBreaks(self, value):
        self.tablestyle.AllowPageBreaks = value

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

    def Condition(self, ConditionCode=None):
        arguments = com_arguments([ConditionCode])
        self.tablestyle.Condition(*arguments)


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

    @Visible.setter
    def Visible(self, value):
        self.task.Visible = value

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

    def Activate(self, Wait=None):
        arguments = com_arguments([Wait])
        self.task.Activate(*arguments)

    def Close(self):
        self.task.Close()

    def Move(self, Left=None, Top=None):
        arguments = com_arguments([Left, Top])
        self.task.Move(*arguments)

    def Resize(self, Width=None, Height=None):
        arguments = com_arguments([Width, Height])
        self.task.Resize(*arguments)

    def SendWindowMessage(self, Message=None, wParam=None, IParam=None):
        arguments = com_arguments([Message, wParam, IParam])
        self.task.SendWindowMessage(*arguments)


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

    @Visible.setter
    def Visible(self, value):
        self.taskpane.Visible = value


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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.taskpanes.Item(*arguments)


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

    @KerningByAlgorithm.setter
    def KerningByAlgorithm(self, value):
        self.template.KerningByAlgorithm = value

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

    @NoProofing.setter
    def NoProofing(self, value):
        self.template.NoProofing = value

    @property
    def Parent(self):
        return self.template.Parent

    @property
    def Path(self):
        return self.template.Path

    @property
    def Saved(self):
        return self.template.Saved

    @Saved.setter
    def Saved(self, value):
        self.template.Saved = value

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

    @FontBold.setter
    def FontBold(self, value):
        self.texteffectformat.FontBold = value

    @property
    def FontItalic(self):
        return self.texteffectformat.FontItalic

    @FontItalic.setter
    def FontItalic(self, value):
        self.texteffectformat.FontItalic = value

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

    @KernedPairs.setter
    def KernedPairs(self, value):
        self.texteffectformat.KernedPairs = value

    @property
    def NormalizedHeight(self):
        return self.texteffectformat.NormalizedHeight

    @NormalizedHeight.setter
    def NormalizedHeight(self, value):
        self.texteffectformat.NormalizedHeight = value

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

    @RotatedChars.setter
    def RotatedChars(self, value):
        self.texteffectformat.RotatedChars = value

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

    @NoTextRotation.setter
    def NoTextRotation(self, value):
        self.textframe.NoTextRotation = value

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

    @WordWrap.setter
    def WordWrap(self, value):
        self.textframe.WordWrap = value

    def BreakForwardLink(self):
        self.textframe.BreakForwardLink()

    def DeleteText(self):
        self.textframe.DeleteText()

    def ValidLinkTarget(self, TargetTextFrame=None):
        arguments = com_arguments([TargetTextFrame])
        return self.textframe.ValidLinkTarget(*arguments)


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

    def EditType(self, Type=None, Default=None, Format=None, Enabled=None):
        arguments = com_arguments([Type, Default, Format, Enabled])
        self.textinput.EditType(*arguments)


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

    @IncludeFieldCodes.setter
    def IncludeFieldCodes(self, value):
        self.textretrievalmode.IncludeFieldCodes = value

    @property
    def IncludeHiddenText(self):
        return self.textretrievalmode.IncludeHiddenText

    @IncludeHiddenText.setter
    def IncludeHiddenText(self, value):
        self.textretrievalmode.IncludeHiddenText = value

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

    @Perspective.setter
    def Perspective(self, value):
        self.threedformat.Perspective = value

    @property
    def PresetCamera(self):
        return self.threedformat.PresetCamera

    @property
    def PresetExtrusionDirection(self):
        return self.threedformat.PresetExtrusionDirection

    @PresetExtrusionDirection.setter
    def PresetExtrusionDirection(self, value):
        self.threedformat.PresetExtrusionDirection = value

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

    @Visible.setter
    def Visible(self, value):
        self.threedformat.Visible = value

    @property
    def Z(self):
        return self.threedformat.Z

    @Z.setter
    def Z(self, value):
        self.threedformat.Z = value

    def IncrementRotationHorizontal(self, Increment=None):
        arguments = com_arguments([Increment])
        self.threedformat.IncrementRotationHorizontal(*arguments)

    def IncrementRotationVertical(self, Increment=None):
        arguments = com_arguments([Increment])
        self.threedformat.IncrementRotationVertical(*arguments)

    def IncrementRotationX(self, Increment=None):
        arguments = com_arguments([Increment])
        self.threedformat.IncrementRotationX(*arguments)

    def IncrementRotationY(self, Increment=None):
        arguments = com_arguments([Increment])
        self.threedformat.IncrementRotationY(*arguments)

    def IncrementRotationZ(self, Increment=None):
        arguments = com_arguments([Increment])
        self.threedformat.IncrementRotationZ(*arguments)

    def ResetRotation(self):
        self.threedformat.ResetRotation()

    def SetExtrusionDirection(self, PresetExtrusionDirection=None):
        arguments = com_arguments([PresetExtrusionDirection])
        self.threedformat.SetExtrusionDirection(*arguments)

    def SetPresetCamera(self, PresetCamera=None):
        arguments = com_arguments([PresetCamera])
        self.threedformat.SetPresetCamera(*arguments)

    def SetThreeDFormat(self, PresetThreeDFormat=None):
        arguments = com_arguments([PresetThreeDFormat])
        self.threedformat.SetThreeDFormat(*arguments)


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

    @NumberFormatLinked.setter
    def NumberFormatLinked(self, value):
        self.ticklabels.NumberFormatLinked = value

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

    @DisplayEquation.setter
    def DisplayEquation(self, value):
        self.trendline.DisplayEquation = value

    @property
    def DisplayRSquared(self):
        return self.trendline.DisplayRSquared

    @DisplayRSquared.setter
    def DisplayRSquared(self, value):
        self.trendline.DisplayRSquared = value

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

    @InterceptIsAuto.setter
    def InterceptIsAuto(self, value):
        self.trendline.InterceptIsAuto = value

    @property
    def Name(self):
        return self.trendline.Name

    @Name.setter
    def Name(self, value):
        self.trendline.Name = value

    @property
    def NameIsAuto(self):
        return self.trendline.NameIsAuto

    @NameIsAuto.setter
    def NameIsAuto(self, value):
        self.trendline.NameIsAuto = value

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

    def Add(self, Type=None, Order=None, Period=None, Forward=None, Backward=None, Intercept=None, DisplayEquation=None, DisplayRSquared=None, Name=None):
        arguments = com_arguments([Type, Order, Period, Forward, Backward, Intercept, DisplayEquation, DisplayRSquared, Name])
        return Trendline(self.trendlines.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return Trendline(self.trendlines.Item(*arguments))


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

    def StartCustomRecord(self, Name=None):
        arguments = com_arguments([Name])
        self.undorecord.StartCustomRecord(*arguments)


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

    @ConflictMode.setter
    def ConflictMode(self, value):
        self.view.ConflictMode = value

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

    @DisplayPageBoundaries.setter
    def DisplayPageBoundaries(self, value):
        self.view.DisplayPageBoundaries = value

    @property
    def Draft(self):
        return self.view.Draft

    @Draft.setter
    def Draft(self, value):
        self.view.Draft = value

    @property
    def FieldShading(self):
        return WdFieldShading(self.view.FieldShading)

    @FieldShading.setter
    def FieldShading(self, value):
        self.view.FieldShading = value

    @property
    def FullScreen(self):
        return self.view.FullScreen

    @FullScreen.setter
    def FullScreen(self, value):
        self.view.FullScreen = value

    @property
    def Magnifier(self):
        return self.view.Magnifier

    @Magnifier.setter
    def Magnifier(self, value):
        self.view.Magnifier = value

    @property
    def MailMergeDataView(self):
        return self.view.MailMergeDataView

    @MailMergeDataView.setter
    def MailMergeDataView(self, value):
        self.view.MailMergeDataView = value

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

    @RevisionsBalloonShowConnectingLines.setter
    def RevisionsBalloonShowConnectingLines(self, value):
        self.view.RevisionsBalloonShowConnectingLines = value

    @property
    def RevisionsBalloonSide(self):
        return self.view.RevisionsBalloonSide

    @property
    def RevisionsBalloonWidth(self):
        return self.view.RevisionsBalloonWidth

    @RevisionsBalloonWidth.setter
    def RevisionsBalloonWidth(self, value):
        self.view.RevisionsBalloonWidth = value

    @property
    def RevisionsBalloonWidthType(self):
        return self.view.RevisionsBalloonWidthType

    @RevisionsBalloonWidthType.setter
    def RevisionsBalloonWidthType(self, value):
        self.view.RevisionsBalloonWidthType = value

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

    @ShowAll.setter
    def ShowAll(self, value):
        self.view.ShowAll = value

    @property
    def ShowBookmarks(self):
        return self.view.ShowBookmarks

    @ShowBookmarks.setter
    def ShowBookmarks(self, value):
        self.view.ShowBookmarks = value

    @property
    def ShowComments(self):
        return self.view.ShowComments

    @ShowComments.setter
    def ShowComments(self, value):
        self.view.ShowComments = value

    @property
    def ShowCropMarks(self):
        return self.view.ShowCropMarks

    @ShowCropMarks.setter
    def ShowCropMarks(self, value):
        self.view.ShowCropMarks = value

    @property
    def ShowDrawings(self):
        return self.view.ShowDrawings

    @ShowDrawings.setter
    def ShowDrawings(self, value):
        self.view.ShowDrawings = value

    @property
    def ShowFieldCodes(self):
        return self.view.ShowFieldCodes

    @ShowFieldCodes.setter
    def ShowFieldCodes(self, value):
        self.view.ShowFieldCodes = value

    @property
    def ShowFirstLineOnly(self):
        return self.view.ShowFirstLineOnly

    @ShowFirstLineOnly.setter
    def ShowFirstLineOnly(self, value):
        self.view.ShowFirstLineOnly = value

    @property
    def ShowFormat(self):
        return self.view.ShowFormat

    @ShowFormat.setter
    def ShowFormat(self, value):
        self.view.ShowFormat = value

    @property
    def ShowFormatChanges(self):
        return self.view.ShowFormatChanges

    @ShowFormatChanges.setter
    def ShowFormatChanges(self, value):
        self.view.ShowFormatChanges = value

    @property
    def ShowHiddenText(self):
        return self.view.ShowHiddenText

    @ShowHiddenText.setter
    def ShowHiddenText(self, value):
        self.view.ShowHiddenText = value

    @property
    def ShowHighlight(self):
        return self.view.ShowHighlight

    @ShowHighlight.setter
    def ShowHighlight(self, value):
        self.view.ShowHighlight = value

    @property
    def ShowHyphens(self):
        return self.view.ShowHyphens

    @ShowHyphens.setter
    def ShowHyphens(self, value):
        self.view.ShowHyphens = value

    @property
    def ShowInkAnnotations(self):
        return self.view.ShowInkAnnotations

    @ShowInkAnnotations.setter
    def ShowInkAnnotations(self, value):
        self.view.ShowInkAnnotations = value

    @property
    def ShowInsertionsAndDeletions(self):
        return self.view.ShowInsertionsAndDeletions

    @ShowInsertionsAndDeletions.setter
    def ShowInsertionsAndDeletions(self, value):
        self.view.ShowInsertionsAndDeletions = value

    @property
    def ShowMainTextLayer(self):
        return self.view.ShowMainTextLayer

    @ShowMainTextLayer.setter
    def ShowMainTextLayer(self, value):
        self.view.ShowMainTextLayer = value

    @property
    def ShowMarkupAreaHighlight(self):
        return self.view.ShowMarkupAreaHighlight

    @ShowMarkupAreaHighlight.setter
    def ShowMarkupAreaHighlight(self, value):
        self.view.ShowMarkupAreaHighlight = value

    @property
    def ShowObjectAnchors(self):
        return self.view.ShowObjectAnchors

    @ShowObjectAnchors.setter
    def ShowObjectAnchors(self, value):
        self.view.ShowObjectAnchors = value

    @property
    def ShowOptionalBreaks(self):
        return self.view.ShowOptionalBreaks

    @ShowOptionalBreaks.setter
    def ShowOptionalBreaks(self, value):
        self.view.ShowOptionalBreaks = value

    @property
    def ShowOtherAuthors(self):
        return self.view.ShowOtherAuthors

    @ShowOtherAuthors.setter
    def ShowOtherAuthors(self, value):
        self.view.ShowOtherAuthors = value

    @property
    def ShowParagraphs(self):
        return self.view.ShowParagraphs

    @ShowParagraphs.setter
    def ShowParagraphs(self, value):
        self.view.ShowParagraphs = value

    @property
    def ShowPicturePlaceHolders(self):
        return self.view.ShowPicturePlaceHolders

    @ShowPicturePlaceHolders.setter
    def ShowPicturePlaceHolders(self, value):
        self.view.ShowPicturePlaceHolders = value

    @property
    def ShowRevisionsAndComments(self):
        return self.view.ShowRevisionsAndComments

    @ShowRevisionsAndComments.setter
    def ShowRevisionsAndComments(self, value):
        self.view.ShowRevisionsAndComments = value

    @property
    def ShowSpaces(self):
        return self.view.ShowSpaces

    @ShowSpaces.setter
    def ShowSpaces(self, value):
        self.view.ShowSpaces = value

    @property
    def ShowTabs(self):
        return self.view.ShowTabs

    @ShowTabs.setter
    def ShowTabs(self, value):
        self.view.ShowTabs = value

    @property
    def ShowTextBoundaries(self):
        return self.view.ShowTextBoundaries

    @ShowTextBoundaries.setter
    def ShowTextBoundaries(self, value):
        self.view.ShowTextBoundaries = value

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

    @TableGridlines.setter
    def TableGridlines(self, value):
        self.view.TableGridlines = value

    @property
    def Type(self):
        return WdViewType(self.view.Type)

    @Type.setter
    def Type(self, value):
        self.view.Type = value

    @property
    def WrapToWindow(self):
        return self.view.WrapToWindow

    @WrapToWindow.setter
    def WrapToWindow(self, value):
        self.view.WrapToWindow = value

    @property
    def Zoom(self):
        return Zoom(self.view.Zoom)

    def CollapseOutline(self, Range=None):
        arguments = com_arguments([Range])
        self.view.CollapseOutline(*arguments)

    def ExpandOutline(self, Range=None):
        arguments = com_arguments([Range])
        self.view.ExpandOutline(*arguments)

    def NextHeaderFooter(self):
        self.view.NextHeaderFooter()

    def PreviousHeaderFooter(self):
        self.view.PreviousHeaderFooter()

    def ShowAllHeadings(self):
        self.view.ShowAllHeadings()

    def ShowHeading(self, Level=None):
        arguments = com_arguments([Level])
        self.view.ShowHeading(*arguments)


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


# WdAlertLevel enumeration
wdAlertsAll = -1
wdAlertsMessageBox = -2
wdAlertsNone = 0

# WdAlignmentTabAlignment enumeration
wdCenter = 1
wdLeft = 0
wdRight = 2

# WdAlignmentTabRelative enumeration
wdIndent = 1
wdMargin = 0

# WdApplyQuickStyleSets enumeration
wdSessionStartSet = 1
wdTemplateSet = 2

# WdArabicNumeral enumeration
wdNumeralArabic = 0
wdNumeralContext = 2
wdNumeralHindi = 1
wdNumeralSystem = 3

# WdAraSpeller enumeration
wdBoth = 3
wdFinalYaa = 2
wdInitialAlef = 1
wdNone = 0

# WdArrangeStyle enumeration
wdIcons = 1
wdTiled = 0

# WdAutoFitBehavior enumeration
wdAutoFitContent = 1
wdAutoFitFixed = 0
wdAutoFitWindow = 2

# WdAutoMacros enumeration
wdAutoClose = 3
wdAutoExec = 0
wdAutoExit = 4
wdAutoNew = 1
wdAutoOpen = 2
wdAutoSync = 5

# WdAutoVersions enumeration
wdAutoVersionOff = 0
wdAutoVersionOnClose = 1

# WdBaselineAlignment enumeration
wdBaselineAlignAuto = 4
wdBaselineAlignBaseline = 2
wdBaselineAlignCenter = 1
wdBaselineAlignFarEast50 = 3
wdBaselineAlignTop = 0

# WdBookmarkSortBy enumeration
wdSortByLocation = 1
wdSortByName = 0

# WdBorderDistanceFrom enumeration
wdBorderDistanceFromPageEdge = 1
wdBorderDistanceFromText = 0

# WdBorderType enumeration
wdBorderBottom = -3
wdBorderDiagonalDown = -7
wdBorderDiagonalUp = -8
wdBorderHorizontal = -5
wdBorderLeft = -2
wdBorderRight = -4
wdBorderTop = -1
wdBorderVertical = -6

# WdBreakType enumeration
wdColumnBreak = 8
wdLineBreak = 6
wdLineBreakClearLeft = 9
wdLineBreakClearRight = 10
wdPageBreak = 7
wdSectionBreakContinuous = 3
wdSectionBreakEvenPage = 4
wdSectionBreakNextPage = 2
wdSectionBreakOddPage = 5
wdTextWrappingBreak = 11

# WdBrowserLevel enumeration
wdBrowserLevelMicrosoftInternetExplorer5 = 1
wdBrowserLevelMicrosoftInternetExplorer6 = 2
wdBrowserLevelV4 = 0

# WdBrowseTarget enumeration
wdBrowseComment = 3
wdBrowseEdit = 10
wdBrowseEndnote = 5
wdBrowseField = 6
wdBrowseFind = 11
wdBrowseFootnote = 4
wdBrowseGoTo = 12
wdBrowseGraphic = 8
wdBrowseHeading = 9
wdBrowsePage = 1
wdBrowseSection = 2
wdBrowseTable = 7

# WdBuildingBlockTypes enumeration
wdTypeAutoText = 9
wdTypeBibliography = 34
wdTypeCoverPage = 2
wdTypeCustom1 = 29
wdTypeCustom2 = 30
wdTypeCustom3 = 31
wdTypeCustom4 = 32
wdTypeCustom5 = 33
wdTypeCustomAutoText = 23
wdTypeCustomBibliography = 35
wdTypeCustomCoverPage = 16
wdTypeCustomEquations = 17
wdTypeCustomFooters = 18
wdTypeCustomHeaders = 19
wdTypeCustomPageNumber = 20
wdTypeCustomPageNumberBottom = 26
wdTypeCustomPageNumberPage = 27
wdTypeCustomPageNumberTop = 25
wdTypeCustomQuickParts = 15
wdTypeCustomTableOfContents = 28
wdTypeCustomTables = 21
wdTypeCustomTextBox = 24
wdTypeCustomWatermarks = 22
wdTypeEquations = 3
wdTypeFooters = 4
wdTypeHeaders = 5
wdTypePageNumber = 6
wdTypePageNumberBottom = 12
wdTypePageNumberPage = 13
wdTypePageNumberTop = 11
wdTypeQuickParts = 1
wdTypeTableOfContents = 14
wdTypeTables = 7
wdTypeTextBox = 10
wdTypeWatermarks = 8

# WdBuiltInProperty enumeration
wdPropertyAppName = 9
wdPropertyAuthor = 3
wdPropertyBytes = 22
wdPropertyCategory = 18
wdPropertyCharacters = 16
wdPropertyCharsWSpaces = 30
wdPropertyComments = 5
wdPropertyCompany = 21
wdPropertyFormat = 19
wdPropertyHiddenSlides = 27
wdPropertyHyperlinkBase = 29
wdPropertyKeywords = 4
wdPropertyLastAuthor = 7
wdPropertyLines = 23
wdPropertyManager = 20
wdPropertyMMClips = 28
wdPropertyNotes = 26
wdPropertyPages = 14
wdPropertyParas = 24
wdPropertyRevision = 8
wdPropertySecurity = 17
wdPropertySlides = 25
wdPropertySubject = 2
wdPropertyTemplate = 6
wdPropertyTimeCreated = 11
wdPropertyTimeLastPrinted = 10
wdPropertyTimeLastSaved = 12
wdPropertyTitle = 1
wdPropertyVBATotalEdit = 13
wdPropertyWords = 15

# WdBuiltinStyle enumeration
wdStyleBlockQuotation = -85
wdStyleBodyText = -67
wdStyleBodyText2 = -81
wdStyleBodyText3 = -82
wdStyleBodyTextFirstIndent = -78
wdStyleBodyTextFirstIndent2 = -79
wdStyleBodyTextIndent = -68
wdStyleBodyTextIndent2 = -83
wdStyleBodyTextIndent3 = -84
wdStyleBookTitle = -265
wdStyleCaption = -35
wdStyleClosing = -64
wdStyleCommentReference = -40
wdStyleCommentText = -31
wdStyleDate = -77
wdStyleDefaultParagraphFont = -66
wdStyleEmphasis = -89
wdStyleEndnoteReference = -43
wdStyleEndnoteText = -44
wdStyleEnvelopeAddress = -37
wdStyleEnvelopeReturn = -38
wdStyleFooter = -33
wdStyleFootnoteReference = -39
wdStyleFootnoteText = -30
wdStyleHeader = -32
wdStyleHeading1 = -2
wdStyleHeading2 = -3
wdStyleHeading3 = -4
wdStyleHeading4 = -5
wdStyleHeading5 = -6
wdStyleHeading6 = -7
wdStyleHeading7 = -8
wdStyleHeading8 = -9
wdStyleHeading9 = -10
wdStyleHtmlAcronym = -96
wdStyleHtmlAddress = -97
wdStyleHtmlCite = -98
wdStyleHtmlCode = -99
wdStyleHtmlDfn = -100
wdStyleHtmlKbd = -101
wdStyleHtmlNormal = -95
wdStyleHtmlPre = -102
wdStyleHtmlSamp = -103
wdStyleHtmlTt = -104
wdStyleHtmlVar = -105
wdStyleHyperlink = -86
wdStyleHyperlinkFollowed = -87
wdStyleIndex1 = -11
wdStyleIndex2 = -12
wdStyleIndex3 = -13
wdStyleIndex4 = -14
wdStyleIndex5 = -15
wdStyleIndex6 = -16
wdStyleIndex7 = -17
wdStyleIndex8 = -18
wdStyleIndex9 = -19
wdStyleIndexHeading = -34
wdStyleIntenseEmphasis = -262
wdStyleIntenseQuote = -182
wdStyleIntenseReference = -264
wdStyleLineNumber = -41
wdStyleList = -48
wdStyleList2 = -51
wdStyleList3 = -52
wdStyleList4 = -53
wdStyleList5 = -54
wdStyleListBullet = -49
wdStyleListBullet2 = -55
wdStyleListBullet3 = -56
wdStyleListBullet4 = -57
wdStyleListBullet5 = -58
wdStyleListContinue = -69
wdStyleListContinue2 = -70
wdStyleListContinue3 = -71
wdStyleListContinue4 = -72
wdStyleListContinue5 = -73
wdStyleListNumber = -50
wdStyleListNumber2 = -59
wdStyleListNumber3 = -60
wdStyleListNumber4 = -61
wdStyleListNumber5 = -62
wdStyleListParagraph = -180
wdStyleMacroText = -46
wdStyleMessageHeader = -74
wdStyleNavPane = -90
wdStyleNormal = -1
wdStyleNormalIndent = -29
wdStyleNormalObject = -158
wdStyleNormalTable = -106
wdStyleNoteHeading = -80
wdStylePageNumber = -42
wdStylePlainText = -91
wdStyleQuote = -181
wdStyleSalutation = -76
wdStyleSignature = -65
wdStyleStrong = -88
wdStyleSubtitle = -75
wdStyleSubtleEmphasis = -261
wdStyleSubtleReference = -263
wdStyleTableColorfulGrid = -172
wdStyleTableColorfulList = -171
wdStyleTableColorfulShading = -170
wdStyleTableDarkList = -169
wdStyleTableLightGrid = -161
wdStyleTableLightGridAccent1 = -175
wdStyleTableLightList = -160
wdStyleTableLightListAccent1 = -174
wdStyleTableLightShading = -159
wdStyleTableLightShadingAccent1 = -173
wdStyleTableMediumGrid1 = -166
wdStyleTableMediumGrid2 = -167
wdStyleTableMediumGrid3 = -168
wdStyleTableMediumList1 = -164
wdStyleTableMediumList1Accent1 = -178
wdStyleTableMediumList2 = -165
wdStyleTableMediumShading1 = -162
wdStyleTableMediumShading1Accent1 = -176
wdStyleTableMediumShading2 = -163
wdStyleTableMediumShading2Accent1 = -177
wdStyleTableOfAuthorities = -45
wdStyleTableOfFigures = -36
wdStyleTitle = -63
wdStyleTOAHeading = -47
wdStyleTOC1 = -20
wdStyleTOC2 = -21
wdStyleTOC3 = -22
wdStyleTOC4 = -23
wdStyleTOC5 = -24
wdStyleTOC6 = -25
wdStyleTOC7 = -26
wdStyleTOC8 = -27
wdStyleTOC9 = -28

# WdCalendarType enumeration
wdCalendarArabic = 1
wdCalendarHebrew = 2
wdCalendarJapan = 4
wdCalendarKorean = 6
wdCalendarSakaEra = 7
wdCalendarTaiwan = 3
wdCalendarThai = 5
wdCalendarTranslitEnglish = 8
wdCalendarTranslitFrench = 9
wdCalendarUmalqura = 13
wdCalendarWestern = 0

# WdCalendarTypeBi enumeration
wdCalendarTypeBidi = 99
wdCalendarTypeGregorian = 100

# WdCaptionLabelID enumeration
wdCaptionEquation = -3
wdCaptionFigure = -1
wdCaptionTable = -2

# WdCaptionNumberStyle enumeration
wdCaptionNumberStyleArabic = 0
wdCaptionNumberStyleArabicFullWidth = 14
wdCaptionNumberStyleArabicLetter1 = 46
wdCaptionNumberStyleArabicLetter2 = 48
wdCaptionNumberStyleChosung = 25
wdCaptionNumberStyleGanada = 24
wdCaptionNumberStyleHanjaRead = 41
wdCaptionNumberStyleHanjaReadDigit = 42
wdCaptionNumberStyleHebrewLetter1 = 45
wdCaptionNumberStyleHebrewLetter2 = 47
wdCaptionNumberStyleHindiArabic = 51
wdCaptionNumberStyleHindiCardinalText = 52
wdCaptionNumberStyleHindiLetter1 = 49
wdCaptionNumberStyleHindiLetter2 = 50
wdCaptionNumberStyleKanji = 10
wdCaptionNumberStyleKanjiDigit = 11
wdCaptionNumberStyleKanjiTraditional = 16
wdCaptionNumberStyleLowercaseLetter = 4
wdCaptionNumberStyleLowercaseRoman = 2
wdCaptionNumberStyleNumberInCircle = 18
wdCaptionNumberStyleSimpChinNum2 = 38
wdCaptionNumberStyleSimpChinNum3 = 39
wdCaptionNumberStyleThaiArabic = 54
wdCaptionNumberStyleThaiCardinalText = 55
wdCaptionNumberStyleThaiLetter = 53
wdCaptionNumberStyleTradChinNum2 = 34
wdCaptionNumberStyleTradChinNum3 = 35
wdCaptionNumberStyleUppercaseLetter = 3
wdCaptionNumberStyleUppercaseRoman = 1
wdCaptionNumberStyleVietCardinalText = 56
wdCaptionNumberStyleZodiac1 = 30
wdCaptionNumberStyleZodiac2 = 31

# WdCaptionPosition enumeration
wdCaptionPositionAbove = 0
wdCaptionPositionBelow = 1

# WdCellColor enumeration
wdCellColorByAuthor = -1
wdCellColorLightBlue = 2
wdCellColorLightGray = 7
wdCellColorLightGreen = 6
wdCellColorLightOrange = 5
wdCellColorLightPurple = 4
wdCellColorLightYellow = 3
wdCellColorNoHighlight = 0
wdCellColorPink = 1

# WdCellVerticalAlignment enumeration
wdCellAlignVerticalBottom = 3
wdCellAlignVerticalCenter = 1
wdCellAlignVerticalTop = 0

# WdCharacterCase enumeration
wdFullWidth = 7
wdHalfWidth = 6
wdHiragana = 9
wdKatakana = 8
wdLowerCase = 0
wdNextCase = -1
wdTitleSentence = 4
wdTitleWord = 2
wdToggleCase = 5
wdUpperCase = 1

# WdCharacterWidth enumeration
wdWidthFullWidth = 7
wdWidthHalfWidth = 6

# WdCheckInVersionType enumeration
wdCheckInMajorVersion = 1
wdCheckInMinorVersion = 0
wdCheckInOverwriteVersion = 2

# WdChevronConvertRule enumeration
wdAlwaysConvert = 1
wdAskToConvert = 3
wdAskToNotConvert = 2
wdNeverConvert = 0

# WdCollapseDirection enumeration
wdCollapseEnd = 0
wdCollapseStart = 1

# WdColorIndex enumeration
wdAuto = 0
wdBlack = 1
wdBlue = 2
wdBrightGreen = 4
wdByAuthor = -1
wdDarkBlue = 9
wdDarkRed = 13
wdDarkYellow = 14
wdGray25 = 16
wdGray50 = 15
wdGreen = 11
wdNoHighlight = 0
wdPink = 5
wdRed = 6
wdTeal = 10
wdTurquoise = 3
wdViolet = 12
wdWhite = 8
wdYellow = 7

# WdCompareDestination enumeration
wdCompareDestinationNew = 2
wdCompareDestinationOriginal = 0
wdCompareDestinationRevised = 1

# WdCompareTarget enumeration
wdCompareTargetCurrent = 1
wdCompareTargetNew = 2
wdCompareTargetSelected = 0

# WdCompatibility enumeration
wdAlignTablesRowByRow = 39
wdApplyBreakingRules = 46
wdAutospaceLikeWW7 = 38
wdConvMailMergeEsc = 6
wdDontAdjustLineHeightInTable = 36
wdDontBalanceSingleByteDoubleByteWidth = 16
wdDontBreakWrappedTables = 43
wdDontSnapTextToGridInTableWithObjects = 44
wdDontULTrailSpace = 15
wdDontUseAsianBreakRulesInGrid = 48
wdDontUseHTMLParagraphAutoSpacing = 35
wdDontWrapTextWithPunctuation = 47
wdExactOnTop = 28
wdExpandShiftReturn = 14
wdFootnoteLayoutLikeWW8 = 34
wdForgetLastTabAlignment = 37
wdGrowAutofit = 50
wdLayoutRawTableWidth = 40
wdLayoutTableRowsApart = 41
wdLeaveBackslashAlone = 13
wdLineWrapLikeWord6 = 32
wdMWSmallCaps = 22
wdNoColumnBalance = 5
wdNoExtraLineSpacing = 23
wdNoLeading = 20
wdNoSpaceForUL = 21
wdNoSpaceRaiseLower = 2
wdNoTabHangIndent = 1
wdOrigWordTableRules = 9
wdPrintBodyTextBeforeHeader = 19
wdPrintColBlack = 3
wdSelectFieldWithFirstOrLastCharacter = 45
wdShapeLayoutLikeWW8 = 33
wdShowBreaksInFrames = 11
wdSpacingInWholePoints = 18
wdSubFontBySize = 25
wdSuppressBottomSpacing = 29
wdSuppressSpBfAfterPgBrk = 7
wdSuppressTopSpacing = 8
wdSuppressTopSpacingMac5 = 17
wdSwapBordersFacingPages = 12
wdTransparentMetafiles = 10
wdTruncateFontHeight = 24
wdUsePrinterMetrics = 26
wdUseWord2002TableStyleRules = 49
wdUseWord2010TableStyleRules = 69
wdUseWord97LineBreakingRules = 42
wdWPJustification = 31
wdWPSpaceWidth = 30
wdWrapTrailSpaces = 4
wdWW6BorderRules = 27
wdAllowSpaceOfSameStyleInTable = 54
wdAutofitLikeWW11 = 57
wdDontAutofitConstrainedTables = 56
wdDontUseIndentAsNumberingTabStop = 52
wdFELineBreak11 = 53
wdHangulWidthLikeWW11 = 59
wdSplitPgBreakAndParaMark = 60
wdUnderlineTabInNumList = 58
wdUseNormalStyleForList = 51
wdWW11IndentRules = 55

# WdCompatibilityMode enumeration
wdCurrent = 65535
wdWord2003 = 11
wdWord2007 = 12
wdWord2010 = 14
wdWord2013 = 15

# WdConditionCode enumeration
wdEvenColumnBanding = 7
wdEvenRowBanding = 3
wdFirstColumn = 4
wdFirstRow = 0
wdLastColumn = 5
wdLastRow = 1
wdNECell = 8
wdNWCell = 9
wdOddColumnBanding = 6
wdOddRowBanding = 2
wdSECell = 10
wdSWCell = 11

# WdConstants enumeration
wdAutoPosition = 0
wdBackward = -1073741823
wdCreatorCode = 1297307460
wdFirst = 1
wdForward = 1073741823
wdToggle = 9999998
wdUndefined = 9999999

# WdContentControlDateStorageFormat enumeration
wdContentControlDateStorageDate = 1
wdContentControlDateStorageDateTime = 2
wdContentControlDateStorageText = 0

# WdContentControlType enumeration
wdContentControlBuildingBlockGallery = 5
wdContentControlCheckbox = 8
wdContentControlComboBox = 3
wdContentControlDate = 6
wdContentControlGroup = 7
wdContentControlDropdownList = 4
wdContentControlPicture = 2
wdContentControlRepeatingSection = 9
wdContentControlRichText = 0
wdContentControlText = 1

# WdContinue enumeration
wdContinueDisabled = 0
wdContinueList = 2
wdResetList = 1

# WdCountry enumeration
wdArgentina = 54
wdBrazil = 55
wdCanada = 2
wdChile = 56
wdChina = 86
wdDenmark = 45
wdFinland = 358
wdFrance = 33
wdGermany = 49
wdIceland = 354
wdItaly = 39
wdJapan = 81
wdKorea = 82
wdLatinAmerica = 3
wdMexico = 52
wdNetherlands = 31
wdNorway = 47
wdPeru = 51
wdSpain = 34
wdSweden = 46
wdTaiwan = 886
wdUK = 44
wdUS = 1
wdVenezuela = 58

# WdCursorMovement enumeration
wdCursorMovementLogical = 0
wdCursorMovementVisual = 1

# WdCursorType enumeration
wdCursorIBeam = 1
wdCursorNormal = 2
wdCursorNorthwestArrow = 3
wdCursorWait = 0

# WdCustomLabelPageSize enumeration
wdCustomLabelA4 = 2
wdCustomLabelA4LS = 3
wdCustomLabelA5 = 4
wdCustomLabelA5LS = 5
wdCustomLabelB4JIS = 13
wdCustomLabelB5 = 6
wdCustomLabelFanfold = 8
wdCustomLabelHigaki = 11
wdCustomLabelHigakiLS = 12
wdCustomLabelLetter = 0
wdCustomLabelLetterLS = 1
wdCustomLabelMini = 7
wdCustomLabelVertHalfSheet = 9
wdCustomLabelVertHalfSheetLS = 10

# WdDateLanguage enumeration
wdDateLanguageBidi = 10
wdDateLanguageLatin = 1033

# WdDefaultFilePath enumeration
wdAutoRecoverPath = 5
wdBorderArtPath = 19
wdCurrentFolderPath = 14
wdDocumentsPath = 0
wdGraphicsFiltersPath = 10
wdPicturesPath = 1
wdProgramPath = 9
wdProofingToolsPath = 12
wdStartupPath = 8
wdStyleGalleryPath = 15
wdTempFilePath = 13
wdTextConvertersPath = 11
wdToolsPath = 6
wdTutorialPath = 7
wdUserOptionsPath = 4
wdUserTemplatesPath = 2
wdWorkgroupTemplatesPath = 3

# WdDefaultListBehavior enumeration
wdWord10ListBehavior = 2
wdWord8ListBehavior = 0
wdWord9ListBehavior = 1

# WdDefaultTableBehavior enumeration
wdWord8TableBehavior = 0
wdWord9TableBehavior = 1

# WdDeleteCells enumeration
wdDeleteCellsEntireColumn = 3
wdDeleteCellsEntireRow = 2
wdDeleteCellsShiftLeft = 0
wdDeleteCellsShiftUp = 1

# WdDeletedTextMark enumeration
wdDeletedTextMarkBold = 5
wdDeletedTextMarkCaret = 2
wdDeletedTextMarkColorOnly = 9
wdDeletedTextMarkDoubleUnderline = 8
wdDeletedTextMarkHidden = 0
wdDeletedTextMarkItalic = 6
wdDeletedTextMarkNone = 4
wdDeletedTextMarkPound = 3
wdDeletedTextMarkStrikeThrough = 1
wdDeletedTextMarkUnderline = 7
wdDeletedTextMarkDoubleStrikeThrough = 10

# WdDiacriticColor enumeration
wdDiacriticColorBidi = 0
wdDiacriticColorLatin = 1

# WdDictionaryType enumeration
wdGrammar = 1
wdHangulHanjaConversion = 8
wdHangulHanjaConversionCustom = 9
wdHyphenation = 3
wdSpelling = 0
wdSpellingComplete = 4
wdSpellingCustom = 5
wdSpellingLegal = 6
wdSpellingMedical = 7
wdThesaurus = 2

# WdDisableFeaturesIntroducedAfter enumeration
wd70 = 0
wd70FE = 1
wd80 = 2

# WdDocPartInsertOptions enumeration
wdInsertContent = 0
wdInsertPage = 2
wdInsertParagraph = 1

# WdDocumentDirection enumeration
wdLeftToRight = 0
wdRightToLeft = 1

# WdDocumentKind enumeration
wdDocumentEmail = 2
wdDocumentLetter = 1
wdDocumentNotSpecified = 0

# WdDocumentMedium enumeration
wdDocument = 1
wdEmailMessage = 0
wdWebPage = 2

# WdDocumentType enumeration
wdTypeDocument = 0
wdTypeFrameset = 2
wdTypeTemplate = 1

# WdDocumentViewDirection enumeration
wdDocumentViewLtr = 1
wdDocumentViewRtl = 0

# WdDropPosition enumeration
wdDropMargin = 2
wdDropNone = 0
wdDropNormal = 1

# WdEditionOption enumeration
wdAutomaticUpdate = 3
wdCancelPublisher = 0
wdChangeAttributes = 5
wdManualUpdate = 4
wdOpenSource = 7
wdSelectPublisher = 2
wdSendPublisher = 1
wdUpdateSubscriber = 6

# WdEditionType enumeration
wdPublisher = 0
wdSubscriber = 1

# WdEditorType enumeration
wdEditorCurrent = -6
wdEditorEditors = -5
wdEditorEveryone = -1
wdEditorOwners = -4

# WdEmailHTMLFidelity enumeration
wdEmailHTMLFidelityHigh = 3
wdEmailHTMLFidelityLow = 1
wdEmailHTMLFidelityMedium = 2

# WdEmphasisMark enumeration
wdEmphasisMarkNone = 0
wdEmphasisMarkOverComma = 2
wdEmphasisMarkOverSolidCircle = 1
wdEmphasisMarkOverWhiteCircle = 3
wdEmphasisMarkUnderSolidCircle = 4

# WdEnableCancelKey enumeration
wdCancelDisabled = 0
wdCancelInterrupt = 1

# WdEncloseStyle enumeration
wdEncloseStyleLarge = 2
wdEncloseStyleNone = 0
wdEncloseStyleSmall = 1

# WdEnclosureType enumeration
wdEnclosureCircle = 0
wdEnclosureDiamond = 3
wdEnclosureSquare = 1
wdEnclosureTriangle = 2

# WdEndnoteLocation enumeration
wdEndOfDocument = 1
wdEndOfSection = 0

# WdEnvelopeOrientation enumeration
wdCenterClockwise = 7
wdCenterLandscape = 4
wdCenterPortrait = 1
wdLeftClockwise = 6
wdLeftLandscape = 3
wdLeftPortrait = 0
wdRightClockwise = 8
wdRightLandscape = 5
wdRightPortrait = 2

# WdExportCreateBookmarks enumeration
wdExportCreateHeadingBookmarks = 1
wdExportCreateNoBookmarks = 0
wdExportCreateWordBookmarks = 2

# WdExportFormat enumeration
wdExportFormatPDF = 17
wdExportFormatXPS = 18

# WdExportItem enumeration
wdExportDocumentContent = 0
wdExportDocumentWithMarkup = 7

# WdExportOptimizeFor enumeration
wdExportOptimizeForOnScreen = 1
wdExportOptimizeForPrint = 0

# WdExportRange enumeration
wdExportAllDocument = 0
wdExportCurrentPage = 2
wdExportFromTo = 3
wdExportSelection = 1

# WdFarEastLineBreakLanguageID enumeration
wdLineBreakJapanese = 1041
wdLineBreakKorean = 1042
wdLineBreakSimplifiedChinese = 2052
wdLineBreakTraditionalChinese = 1028

# WdFarEastLineBreakLevel enumeration
wdFarEastLineBreakLevelCustom = 2
wdFarEastLineBreakLevelNormal = 0
wdFarEastLineBreakLevelStrict = 1

# WdFieldKind enumeration
wdFieldKindCold = 3
wdFieldKindHot = 1
wdFieldKindNone = 0
wdFieldKindWarm = 2

# WdFieldShading enumeration
wdFieldShadingAlways = 1
wdFieldShadingNever = 0
wdFieldShadingWhenSelected = 2

# WdFieldType enumeration
wdFieldAddin = 81
wdFieldAddressBlock = 93
wdFieldAdvance = 84
wdFieldAsk = 38
wdFieldAuthor = 17
wdFieldAutoNum = 54
wdFieldAutoNumLegal = 53
wdFieldAutoNumOutline = 52
wdFieldAutoText = 79
wdFieldAutoTextList = 89
wdFieldBarCode = 63
wdFieldBidiOutline = 92
wdFieldComments = 19
wdFieldCompare = 80
wdFieldCreateDate = 21
wdFieldData = 40
wdFieldDatabase = 78
wdFieldDate = 31
wdFieldDDE = 45
wdFieldDDEAuto = 46
wdFieldDisplayBarcode = 99
wdFieldDocProperty = 85
wdFieldDocVariable = 64
wdFieldEditTime = 25
wdFieldEmbed = 58
wdFieldEmpty = -1
wdFieldExpression = 34
wdFieldFileName = 29
wdFieldFileSize = 69
wdFieldFillIn = 39
wdFieldFootnoteRef = 5
wdFieldFormCheckBox = 71
wdFieldFormDropDown = 83
wdFieldFormTextInput = 70
wdFieldFormula = 49
wdFieldGlossary = 47
wdFieldGoToButton = 50
wdFieldGreetingLine = 94
wdFieldHTMLActiveX = 91
wdFieldHyperlink = 88
wdFieldIf = 7
wdFieldImport = 55
wdFieldInclude = 36
wdFieldIncludePicture = 67
wdFieldIncludeText = 68
wdFieldIndex = 8
wdFieldIndexEntry = 4
wdFieldInfo = 14
wdFieldKeyWord = 18
wdFieldLastSavedBy = 20
wdFieldLink = 56
wdFieldListNum = 90
wdFieldMacroButton = 51
wdFieldMergeBarcode = 98
wdFieldMergeField = 59
wdFieldMergeRec = 44
wdFieldMergeSeq = 75
wdFieldNext = 41
wdFieldNextIf = 42
wdFieldNoteRef = 72
wdFieldNumChars = 28
wdFieldNumPages = 26
wdFieldNumWords = 27
wdFieldOCX = 87
wdFieldPage = 33
wdFieldPageRef = 37
wdFieldPrint = 48
wdFieldPrintDate = 23
wdFieldPrivate = 77
wdFieldQuote = 35
wdFieldRef = 3
wdFieldRefDoc = 11
wdFieldRevisionNum = 24
wdFieldSaveDate = 22
wdFieldSection = 65
wdFieldSectionPages = 66
wdFieldSequence = 12
wdFieldSet = 6
wdFieldShape = 95
wdFieldSkipIf = 43
wdFieldStyleRef = 10
wdFieldSubject = 16
wdFieldSubscriber = 82
wdFieldSymbol = 57
wdFieldTemplate = 30
wdFieldTime = 32
wdFieldTitle = 15
wdFieldTOA = 73
wdFieldTOAEntry = 74
wdFieldTOC = 13
wdFieldTOCEntry = 9
wdFieldUserAddress = 62
wdFieldUserInitials = 61
wdFieldUserName = 60
wdFieldBibliography = 97
wdFieldCitation = 96

# WdFindMatch enumeration
wdMatchAnyCharacter = 65599
wdMatchAnyDigit = 65567
wdMatchAnyLetter = 65583
wdMatchCaretCharacter = 11
wdMatchColumnBreak = 14
wdMatchCommentMark = 5
wdMatchEmDash = 8212
wdMatchEnDash = 8211
wdMatchEndnoteMark = 65555
wdMatchField = 19
wdMatchFootnoteMark = 65554
wdMatchGraphic = 1
wdMatchManualLineBreak = 65551
wdMatchManualPageBreak = 65564
wdMatchNonbreakingHyphen = 30
wdMatchNonbreakingSpace = 160
wdMatchOptionalHyphen = 31
wdMatchParagraphMark = 65551
wdMatchSectionBreak = 65580
wdMatchTabCharacter = 9
wdMatchWhiteSpace = 65655

# WdFindWrap enumeration
wdFindAsk = 2
wdFindContinue = 1
wdFindStop = 0

# WdFlowDirection enumeration
wdFlowLtr = 0
wdFlowRtl = 1

# WdFontBias enumeration
wdFontBiasDefault = 0
wdFontBiasDontCare = 255
wdFontBiasFareast = 1

# WdFootnoteLocation enumeration
wdBeneathText = 1
wdBottomOfPage = 0

# WdFramePosition enumeration
wdFrameBottom = -999997
wdFrameCenter = -999995
wdFrameInside = -999994
wdFrameLeft = -999998
wdFrameOutside = -999993
wdFrameRight = -999996
wdFrameTop = -999999

# WdFramesetNewFrameLocation enumeration
wdFramesetNewFrameAbove = 0
wdFramesetNewFrameBelow = 1
wdFramesetNewFrameLeft = 3
wdFramesetNewFrameRight = 2

# WdFramesetSizeType enumeration
wdFramesetSizeTypeFixed = 1
wdFramesetSizeTypePercent = 0
wdFramesetSizeTypeRelative = 2

# WdFramesetType enumeration
wdFramesetTypeFrame = 1
wdFramesetTypeFrameset = 0

# WdFrameSizeRule enumeration
wdFrameAtLeast = 1
wdFrameAuto = 0
wdFrameExact = 2

# WdFrenchSpeller enumeration
wdFrenchBoth = 0
wdFrenchPostReform = 2
wdFrenchPreReform = 1

# WdGoToDirection enumeration
wdGoToAbsolute = 1
wdGoToFirst = 1
wdGoToLast = -1
wdGoToNext = 2
wdGoToPrevious = 3
wdGoToRelative = 2

# WdGoToItem enumeration
wdGoToBookmark = -1
wdGoToComment = 6
wdGoToEndnote = 5
wdGoToEquation = 10
wdGoToField = 7
wdGoToFootnote = 4
wdGoToGrammaticalError = 14
wdGoToGraphic = 8
wdGoToHeading = 11
wdGoToLine = 3
wdGoToObject = 9
wdGoToPage = 1
wdGoToPercent = 12
wdGoToProofreadingError = 15
wdGoToSection = 0
wdGoToSpellingError = 13
wdGoToTable = 2

# WdGranularity enumeration
wdGranularityCharLevel = 0
wdGranularityWordLevel = 1

# WdGutterStyle enumeration
wdGutterPosLeft = 0
wdGutterPosRight = 2
wdGutterPosTop = 1

# WdGutterStyleOld enumeration
wdGutterStyleBidi = 2
wdGutterStyleLatin = -10

# WdHeaderFooterIndex enumeration
wdHeaderFooterEvenPages = 3
wdHeaderFooterFirstPage = 2
wdHeaderFooterPrimary = 1

# WdHeadingSeparator enumeration
wdHeadingSeparatorBlankLine = 1
wdHeadingSeparatorLetter = 2
wdHeadingSeparatorLetterFull = 4
wdHeadingSeparatorLetterLow = 3
wdHeadingSeparatorNone = 0

# WdHebSpellStart enumeration
wdFullScript = 0
wdMixedAuthorizedScript = 3
wdMixedScript = 2
wdPartialScript = 1

# WdHelpType enumeration
wdHelp = 0
wdHelpAbout = 1
wdHelpActiveWindow = 2
wdHelpContents = 3
wdHelpExamplesAndDemos = 4
wdHelpHWP = 13
wdHelpIchitaro = 11
wdHelpIndex = 5
wdHelpKeyboard = 6
wdHelpPE2 = 12
wdHelpPSSHelp = 7
wdHelpQuickPreview = 8
wdHelpSearch = 9
wdHelpUsingHelp = 10

# WdHighAnsiText enumeration
wdAutoDetectHighAnsiFarEast = 2
wdHighAnsiIsFarEast = 0
wdHighAnsiIsHighAnsi = 1

# WdHorizontalInVerticalType enumeration
wdHorizontalInVerticalFitInLine = 1
wdHorizontalInVerticalNone = 0
wdHorizontalInVerticalResizeLine = 2

# WdHorizontalLineAlignment enumeration
wdHorizontalLineAlignCenter = 1
wdHorizontalLineAlignLeft = 0
wdHorizontalLineAlignRight = 2

# WdHorizontalLineWidthType enumeration
wdHorizontalLineFixedWidth = -2
wdHorizontalLinePercentWidth = -1

# WdIMEMode enumeration
wdIMEModeAlpha = 8
wdIMEModeAlphaFull = 7
wdIMEModeHangul = 10
wdIMEModeHangulFull = 9
wdIMEModeHiragana = 4
wdIMEModeKatakana = 5
wdIMEModeKatakanaHalf = 6
wdIMEModeNoControl = 0
wdIMEModeOff = 2
wdIMEModeOn = 1

# WdIndexFilter enumeration
wdIndexFilterAiueo = 1
wdIndexFilterAkasatana = 2
wdIndexFilterChosung = 3
wdIndexFilterFull = 6
wdIndexFilterLow = 4
wdIndexFilterMedium = 5
wdIndexFilterNone = 0

# WdIndexFormat enumeration
wdIndexBulleted = 4
wdIndexClassic = 1
wdIndexFancy = 2
wdIndexFormal = 5
wdIndexModern = 3
wdIndexSimple = 6
wdIndexTemplate = 0

# WdIndexSortBy enumeration
wdIndexSortByStroke = 0
wdIndexSortBySyllable = 1

# WdIndexType enumeration
wdIndexIndent = 0
wdIndexRunin = 1

# WdInformation enumeration
wdActiveEndAdjustedPageNumber = 1
wdActiveEndPageNumber = 3
wdActiveEndSectionNumber = 2
wdAtEndOfRowMarker = 31
wdCapsLock = 21
wdEndOfRangeColumnNumber = 17
wdEndOfRangeRowNumber = 14
wdFirstCharacterColumnNumber = 9
wdFirstCharacterLineNumber = 10
wdFrameIsSelected = 11
wdHeaderFooterType = 33
wdHorizontalPositionRelativeToPage = 5
wdHorizontalPositionRelativeToTextBoundary = 7
wdInBibliography = 42
wdInCitation = 43
wdInClipboard = 38
wdInCommentPane = 26
wdInContentControl = 46
wdInCoverPage = 41
wdInEndnote = 36
wdInFieldCode = 44
wdInFieldResult = 45
wdInFootnote = 35
wdInFootnoteEndnotePane = 25
wdInHeaderFooter = 28
wdInMasterDocument = 34
wdInWordMail = 37
wdMaximumNumberOfColumns = 18
wdMaximumNumberOfRows = 15
wdNumberOfPagesInDocument = 4
wdNumLock = 22
wdOverType = 23
wdReferenceOfType = 32
wdRevisionMarking = 24
wdSelectionMode = 20
wdStartOfRangeColumnNumber = 16
wdStartOfRangeRowNumber = 13
wdVerticalPositionRelativeToPage = 6
wdVerticalPositionRelativeToTextBoundary = 8
wdWithInTable = 12
wdZoomPercentage = 19

# WdInlineShapeType enumeration
wdInlineShape3DModel = 19
wdInlineShapeChart = 12
wdInlineShapeDiagram = 13
wdInlineShapeEmbeddedOLEObject = 1
wdInlineShapeHorizontalLine = 6
wdInlineShapeLinked3DModel = 20
wdInlineShapeLinkedOLEObject = 2
wdInlineShapeLinkedPicture = 4
wdInlineShapeLinkedPictureHorizontalLine = 8
wdInlineShapeLockedCanvas = 14
wdInlineShapeOLEControlObject = 5
wdInlineShapeOWSAnchor = 11
wdInlineShapePicture = 3
wdInlineShapePictureBullet = 9
wdInlineShapePictureHorizontalLine = 7
wdInlineShapeScriptAnchor = 10
wdInlineShapeSmartArt = 15
wdInlineShapeWebVideo = 16

# WdInsertCells enumeration
wdInsertCellsEntireColumn = 3
wdInsertCellsEntireRow = 2
wdInsertCellsShiftDown = 1
wdInsertCellsShiftRight = 0

# WdInsertedTextMark enumeration
wdInsertedTextMarkBold = 1
wdInsertedTextMarkColorOnly = 5
wdInsertedTextMarkDoubleUnderline = 4
wdInsertedTextMarkItalic = 2
wdInsertedTextMarkNone = 0
wdInsertedTextMarkStrikeThrough = 6
wdInsertedTextMarkUnderline = 3
wdInsertedTextMarkDoubleStrikeThrough = 7

# WdInternationalIndex enumeration
wd24HourClock = 21
wdCurrencyCode = 20
wdDateSeparator = 25
wdDecimalSeparator = 18
wdInternationalAM = 22
wdInternationalPM = 23
wdListSeparator = 17
wdProductLanguageID = 26
wdThousandsSeparator = 19
wdTimeSeparator = 24

# WdJustificationMode enumeration
wdJustificationModeCompress = 1
wdJustificationModeCompressKana = 2
wdJustificationModeExpand = 0

# WdKana enumeration
wdKanaHiragana = 9
wdKanaKatakana = 8

# WdKey enumeration
wdKey0 = 48
wdKey1 = 49
wdKey2 = 50
wdKey3 = 51
wdKey4 = 52
wdKey5 = 53
wdKey6 = 54
wdKey7 = 55
wdKey8 = 56
wdKey9 = 57
wdKeyA = 65
wdKeyAlt = 1024
wdKeyB = 66
wdKeyBackSingleQuote = 192
wdKeyBackSlash = 220
wdKeyBackspace = 8
wdKeyC = 67
wdKeyCloseSquareBrace = 221
wdKeyComma = 188
wdKeyCommand = 512
wdKeyControl = 512
wdKeyD = 68
wdKeyDelete = 46
wdKeyE = 69
wdKeyEnd = 35
wdKeyEquals = 187
wdKeyEsc = 27
wdKeyF = 70
wdKeyF1 = 112
wdKeyF10 = 121
wdKeyF11 = 122
wdKeyF12 = 123
wdKeyF13 = 124
wdKeyF14 = 125
wdKeyF15 = 126
wdKeyF16 = 127
wdKeyF2 = 113
wdKeyF3 = 114
wdKeyF4 = 115
wdKeyF5 = 116
wdKeyF6 = 117
wdKeyF7 = 118
wdKeyF8 = 119
wdKeyF9 = 120
wdKeyG = 71
wdKeyH = 72
wdKeyHome = 36
wdKeyHyphen = 189
wdKeyI = 73
wdKeyInsert = 45
wdKeyJ = 74
wdKeyK = 75
wdKeyL = 76
wdKeyM = 77
wdKeyN = 78
wdKeyNumeric0 = 96
wdKeyNumeric1 = 97
wdKeyNumeric2 = 98
wdKeyNumeric3 = 99
wdKeyNumeric4 = 100
wdKeyNumeric5 = 101
wdKeyNumeric5Special = 12
wdKeyNumeric6 = 102
wdKeyNumeric7 = 103
wdKeyNumeric8 = 104
wdKeyNumeric9 = 105
wdKeyNumericAdd = 107
wdKeyNumericDecimal = 110
wdKeyNumericDivide = 111
wdKeyNumericMultiply = 106
wdKeyNumericSubtract = 109
wdKeyO = 79
wdKeyOpenSquareBrace = 219
wdKeyOption = 1024
wdKeyP = 80
wdKeyPageDown = 34
wdKeyPageUp = 33
wdKeyPause = 19
wdKeyPeriod = 190
wdKeyQ = 81
wdKeyR = 82
wdKeyReturn = 13
wdKeyS = 83
wdKeyScrollLock = 145
wdKeySemiColon = 186
wdKeyShift = 256
wdKeySingleQuote = 222
wdKeySlash = 191
wdKeySpacebar = 32
wdKeyT = 84
wdKeyTab = 9
wdKeyU = 85
wdKeyV = 86
wdKeyW = 87
wdKeyX = 88
wdKeyY = 89
wdKeyZ = 90
wdNoKey = 255

# WdKeyCategory enumeration
wdKeyCategoryAutoText = 4
wdKeyCategoryCommand = 1
wdKeyCategoryDisable = 0
wdKeyCategoryFont = 3
wdKeyCategoryMacro = 2
wdKeyCategoryNil = -1
wdKeyCategoryPrefix = 7
wdKeyCategoryStyle = 5
wdKeyCategorySymbol = 6

# WdLanguageID enumeration
wdAfrikaans = 1078
wdAlbanian = 1052
wdAmharic = 1118
wdArabic = 1025
wdArabicAlgeria = 5121
wdArabicBahrain = 15361
wdArabicEgypt = 3073
wdArabicIraq = 2049
wdArabicJordan = 11265
wdArabicKuwait = 13313
wdArabicLebanon = 12289
wdArabicLibya = 4097
wdArabicMorocco = 6145
wdArabicOman = 8193
wdArabicQatar = 16385
wdArabicSyria = 10241
wdArabicTunisia = 7169
wdArabicUAE = 14337
wdArabicYemen = 9217
wdArmenian = 1067
wdAssamese = 1101
wdAzeriCyrillic = 2092
wdAzeriLatin = 1068
wdBasque = 1069
wdBelgianDutch = 2067
wdBelgianFrench = 2060
wdBengali = 1093
wdBulgarian = 1026
wdBurmese = 1109
wdByelorussian = 1059
wdCatalan = 1027
wdCherokee = 1116
wdChineseHongKongSAR = 3076
wdChineseMacaoSAR = 5124
wdChineseSingapore = 4100
wdCroatian = 1050
wdCzech = 1029
wdDanish = 1030
wdDivehi = 1125
wdDutch = 1043
wdEdo = 1126
wdEnglishAUS = 3081
wdEnglishBelize = 10249
wdEnglishCanadian = 4105
wdEnglishCaribbean = 9225
wdEnglishIndonesia = 14345
wdEnglishIreland = 6153
wdEnglishJamaica = 8201
wdEnglishNewZealand = 5129
wdEnglishPhilippines = 13321
wdEnglishSouthAfrica = 7177
wdEnglishTrinidadTobago = 11273
wdEnglishUK = 2057
wdEnglishUS = 1033
wdEnglishZimbabwe = 12297
wdEstonian = 1061
wdFaeroese = 1080
wdFilipino = 1124
wdFinnish = 1035
wdFrench = 1036
wdFrenchCameroon = 11276
wdFrenchCanadian = 3084
wdFrenchCongoDRC = 9228
wdFrenchCotedIvoire = 12300
wdFrenchHaiti = 15372
wdFrenchLuxembourg = 5132
wdFrenchMali = 13324
wdFrenchMonaco = 6156
wdFrenchMorocco = 14348
wdFrenchReunion = 8204
wdFrenchSenegal = 10252
wdFrenchWestIndies = 7180
wdFrisianNetherlands = 1122
wdFulfulde = 1127
wdGaelicIreland = 2108
wdGaelicScotland = 1084
wdGalician = 1110
wdGeorgian = 1079
wdGerman = 1031
wdGermanAustria = 3079
wdGermanLiechtenstein = 5127
wdGermanLuxembourg = 4103
wdGreek = 1032
wdGuarani = 1140
wdGujarati = 1095
wdHausa = 1128
wdHawaiian = 1141
wdHebrew = 1037
wdHindi = 1081
wdHungarian = 1038
wdIbibio = 1129
wdIcelandic = 1039
wdIgbo = 1136
wdIndonesian = 1057
wdInuktitut = 1117
wdItalian = 1040
wdJapanese = 1041
wdKannada = 1099
wdKanuri = 1137
wdKashmiri = 1120
wdKazakh = 1087
wdKhmer = 1107
wdKirghiz = 1088
wdKonkani = 1111
wdKorean = 1042
wdKyrgyz = 1088
wdLanguageNone = 0
wdLao = 1108
wdLatin = 1142
wdLatvian = 1062
wdLithuanian = 1063
wdMacedonianFYROM = 1071
wdMalayalam = 1100
wdMalayBruneiDarussalam = 2110
wdMalaysian = 1086
wdMaltese = 1082
wdManipuri = 1112
wdMarathi = 1102
wdMexicanSpanish = 2058
wdMongolian = 1104
wdNepali = 1121
wdNoProofing = 1024
wdNorwegianBokmol = 1044
wdNorwegianNynorsk = 2068
wdOriya = 1096
wdOromo = 1138
wdPashto = 1123
wdPersian = 1065
wdPolish = 1045
wdPortuguese = 2070
wdPortugueseBrazil = 1046
wdPunjabi = 1094
wdRhaetoRomanic = 1047
wdRomanian = 1048
wdRomanianMoldova = 2072
wdRussian = 1049
wdRussianMoldova = 2073
wdSamiLappish = 1083
wdSanskrit = 1103
wdSerbianCyrillic = 3098
wdSerbianLatin = 2074
wdSesotho = 1072
wdSimplifiedChinese = 2052
wdSindhi = 1113
wdSindhiPakistan = 2137
wdSinhalese = 1115
wdSlovak = 1051
wdSlovenian = 1060
wdSomali = 1143
wdSorbian = 1070
wdSpanish = 1034
wdSpanishArgentina = 11274
wdSpanishBolivia = 16394
wdSpanishChile = 13322
wdSpanishColombia = 9226
wdSpanishCostaRica = 5130
wdSpanishDominicanRepublic = 7178
wdSpanishEcuador = 12298
wdSpanishElSalvador = 17418
wdSpanishGuatemala = 4106
wdSpanishHonduras = 18442
wdSpanishModernSort = 3082
wdSpanishNicaragua = 19466
wdSpanishPanama = 6154
wdSpanishParaguay = 15370
wdSpanishPeru = 10250
wdSpanishPuertoRico = 20490
wdSpanishUruguay = 14346
wdSpanishVenezuela = 8202
wdSutu = 1072
wdSwahili = 1089
wdSwedish = 1053
wdSwedishFinland = 2077
wdSwissFrench = 4108
wdSwissGerman = 2055
wdSwissItalian = 2064
wdSyriac = 1114
wdTajik = 1064
wdTamazight = 1119
wdTamazightLatin = 2143
wdTamil = 1097
wdTatar = 1092
wdTelugu = 1098
wdThai = 1054
wdTibetan = 1105
wdTigrignaEritrea = 2163
wdTigrignaEthiopic = 1139
wdTraditionalChinese = 1028
wdTsonga = 1073
wdTswana = 1074
wdTurkish = 1055
wdTurkmen = 1090
wdUkrainian = 1058
wdUrdu = 1056
wdUzbekCyrillic = 2115
wdUzbekLatin = 1091
wdVenda = 1075
wdVietnamese = 1066
wdWelsh = 1106
wdXhosa = 1076
wdYi = 1144
wdYiddish = 1085
wdYoruba = 1130
wdZulu = 1077

# WdLayoutMode enumeration
wdLayoutModeDefault = 0
wdLayoutModeGenko = 3
wdLayoutModeGrid = 1
wdLayoutModeLineGrid = 2

# WdLetterheadLocation enumeration
wdLetterBottom = 1
wdLetterLeft = 2
wdLetterRight = 3
wdLetterTop = 0

# WdLetterStyle enumeration
wdFullBlock = 0
wdModifiedBlock = 1
wdSemiBlock = 2

# WdLigatures enumeration
wdLigaturesAll = 15
wdLigaturesContextual = 2
wdLigaturesContextualDiscretional = 10
wdLigaturesContextualHistorical = 6
wdLigaturesContextualHistoricalDiscretional = 14
wdLigaturesDiscretional = 8
wdLigaturesHistorical = 4
wdLigaturesHistoricalDiscretional = 12
wdLigaturesNone = 0
wdLigaturesStandard = 1
wdLigaturesStandardContextual = 3
wdLigaturesStandardContextualDiscretional = 11
wdLigaturesStandardContextualHistorical = 7
wdLigaturesStandardDiscretional = 9
wdLigaturesStandardHistorical = 5
wdLigaturesStandardHistoricalDiscretional = 13

# WdLineEndingType enumeration
wdCRLF = 0
wdCROnly = 1
wdLFCR = 3
wdLFOnly = 2
wdLSPS = 4

# WdLineSpacing enumeration
wdLineSpace1pt5 = 1
wdLineSpaceAtLeast = 3
wdLineSpaceDouble = 2
wdLineSpaceExactly = 4
wdLineSpaceMultiple = 5
wdLineSpaceSingle = 0

# WdLineStyle enumeration
wdLineStyleDashDot = 5
wdLineStyleDashDotDot = 6
wdLineStyleDashDotStroked = 20
wdLineStyleDashLargeGap = 4
wdLineStyleDashSmallGap = 3
wdLineStyleDot = 2
wdLineStyleDouble = 7
wdLineStyleDoubleWavy = 19
wdLineStyleEmboss3D = 21
wdLineStyleEngrave3D = 22
wdLineStyleInset = 24
wdLineStyleNone = 0
wdLineStyleOutset = 23
wdLineStyleSingle = 1
wdLineStyleSingleWavy = 18
wdLineStyleThickThinLargeGap = 16
wdLineStyleThickThinMedGap = 13
wdLineStyleThickThinSmallGap = 10
wdLineStyleThinThickLargeGap = 15
wdLineStyleThinThickMedGap = 12
wdLineStyleThinThickSmallGap = 9
wdLineStyleThinThickThinLargeGap = 17
wdLineStyleThinThickThinMedGap = 14
wdLineStyleThinThickThinSmallGap = 11
wdLineStyleTriple = 8

# WdLineType enumeration
wdTableRow = 1
wdTextLine = 0

# WdLineWidth enumeration
wdLineWidth025pt = 2
wdLineWidth050pt = 4
wdLineWidth075pt = 6
wdLineWidth100pt = 8
wdLineWidth150pt = 12
wdLineWidth225pt = 18
wdLineWidth300pt = 24
wdLineWidth450pt = 36
wdLineWidth600pt = 48

# WdLinkType enumeration
wdLinkTypeChart = 8
wdLinkTypeDDE = 6
wdLinkTypeDDEAuto = 7
wdLinkTypeImport = 5
wdLinkTypeInclude = 4
wdLinkTypeOLE = 0
wdLinkTypePicture = 1
wdLinkTypeReference = 3
wdLinkTypeText = 2

# WdListApplyTo enumeration
wdListApplyToSelection = 2
wdListApplyToThisPointForward = 1
wdListApplyToWholeList = 0

# WdListGalleryType enumeration
wdBulletGallery = 1
wdNumberGallery = 2
wdOutlineNumberGallery = 3

# WdListLevelAlignment enumeration
wdListLevelAlignCenter = 1
wdListLevelAlignLeft = 0
wdListLevelAlignRight = 2

# WdListNumberStyle enumeration
wdListNumberStyleAiueo = 20
wdListNumberStyleAiueoHalfWidth = 12
wdListNumberStyleArabic = 0
wdListNumberStyleArabic1 = 46
wdListNumberStyleArabic2 = 48
wdListNumberStyleArabicFullWidth = 14
wdListNumberStyleArabicLZ = 22
wdListNumberStyleArabicLZ2 = 62
wdListNumberStyleArabicLZ3 = 63
wdListNumberStyleArabicLZ4 = 64
wdListNumberStyleBullet = 23
wdListNumberStyleCardinalText = 6
wdListNumberStyleChosung = 25
wdListNumberStyleGanada = 24
wdListNumberStyleGBNum1 = 26
wdListNumberStyleGBNum2 = 27
wdListNumberStyleGBNum3 = 28
wdListNumberStyleGBNum4 = 29
wdListNumberStyleHangul = 43
wdListNumberStyleHanja = 44
wdListNumberStyleHanjaRead = 41
wdListNumberStyleHanjaReadDigit = 42
wdListNumberStyleHebrew1 = 45
wdListNumberStyleHebrew2 = 47
wdListNumberStyleHindiArabic = 51
wdListNumberStyleHindiCardinalText = 52
wdListNumberStyleHindiLetter1 = 49
wdListNumberStyleHindiLetter2 = 50
wdListNumberStyleIroha = 21
wdListNumberStyleIrohaHalfWidth = 13
wdListNumberStyleKanji = 10
wdListNumberStyleKanjiDigit = 11
wdListNumberStyleKanjiTraditional = 16
wdListNumberStyleKanjiTraditional2 = 17
wdListNumberStyleLegal = 253
wdListNumberStyleLegalLZ = 254
wdListNumberStyleLowercaseBulgarian = 67
wdListNumberStyleLowercaseGreek = 60
wdListNumberStyleLowercaseLetter = 4
wdListNumberStyleLowercaseRoman = 2
wdListNumberStyleLowercaseRussian = 58
wdListNumberStyleLowercaseTurkish = 65
wdListNumberStyleNone = 255
wdListNumberStyleNumberInCircle = 18
wdListNumberStyleOrdinal = 5
wdListNumberStyleOrdinalText = 7
wdListNumberStylePictureBullet = 249
wdListNumberStyleSimpChinNum1 = 37
wdListNumberStyleSimpChinNum2 = 38
wdListNumberStyleSimpChinNum3 = 39
wdListNumberStyleSimpChinNum4 = 40
wdListNumberStyleThaiArabic = 54
wdListNumberStyleThaiCardinalText = 55
wdListNumberStyleThaiLetter = 53
wdListNumberStyleTradChinNum1 = 33
wdListNumberStyleTradChinNum2 = 34
wdListNumberStyleTradChinNum3 = 35
wdListNumberStyleTradChinNum4 = 36
wdListNumberStyleUppercaseBulgarian = 68
wdListNumberStyleUppercaseGreek = 61
wdListNumberStyleUppercaseLetter = 3
wdListNumberStyleUppercaseRoman = 1
wdListNumberStyleUppercaseRussian = 59
wdListNumberStyleUppercaseTurkish = 66
wdListNumberStyleVietCardinalText = 56
wdListNumberStyleZodiac1 = 30
wdListNumberStyleZodiac2 = 31
wdListNumberStyleZodiac3 = 32

# WdListType enumeration
wdListBullet = 2
wdListListNumOnly = 1
wdListMixedNumbering = 5
wdListNoNumbering = 0
wdListOutlineNumbering = 4
wdListPictureBullet = 6
wdListSimpleNumbering = 3

# WdLockType enumeration
wdLockChanged = 3
wdLockEphemeral = 2
wdLockNone = 0
wdLockReservation = 1

# WdMailerPriority enumeration
wdPriorityHigh = 3
wdPriorityLow = 2
wdPriorityNormal = 1

# WdMailMergeActiveRecord enumeration
wdFirstDataSourceRecord = -6
wdFirstRecord = -4
wdLastDataSourceRecord = -7
wdLastRecord = -5
wdNextDataSourceRecord = -8
wdNextRecord = -2
wdNoActiveRecord = -1
wdPreviousDataSourceRecord = -9
wdPreviousRecord = -3

# WdMailMergeComparison enumeration
wdMergeIfEqual = 0
wdMergeIfGreaterThan = 3
wdMergeIfGreaterThanOrEqual = 5
wdMergeIfIsBlank = 6
wdMergeIfIsNotBlank = 7
wdMergeIfLessThan = 2
wdMergeIfLessThanOrEqual = 4
wdMergeIfNotEqual = 1

# WdMailMergeDataSource enumeration
wdMergeInfoFromAccessDDE = 1
wdMergeInfoFromExcelDDE = 2
wdMergeInfoFromMSQueryDDE = 3
wdMergeInfoFromODBC = 4
wdMergeInfoFromODSO = 5
wdMergeInfoFromWord = 0
wdNoMergeInfo = -1

# WdMailMergeDefaultRecord enumeration
wdDefaultFirstRecord = 1
wdDefaultLastRecord = -16

# WdMailMergeDestination enumeration
wdSendToEmail = 2
wdSendToFax = 3
wdSendToNewDocument = 0
wdSendToPrinter = 1

# WdMailMergeMailFormat enumeration
wdMailFormatHTML = 1
wdMailFormatPlainText = 0

# WdMailMergeMainDocType enumeration
wdCatalog = 3
wdDirectory = 3
wdEMail = 4
wdEnvelopes = 2
wdFax = 5
wdFormLetters = 0
wdMailingLabels = 1
wdNotAMergeDocument = -1

# WdMailMergeState enumeration
wdDataSource = 5
wdMainAndDataSource = 2
wdMainAndHeader = 3
wdMainAndSourceAndHeader = 4
wdMainDocumentOnly = 1
wdNormalDocument = 0

# WdMailSystem enumeration
wdMAPI = 1
wdMAPIandPowerTalk = 3
wdNoMailSystem = 0
wdPowerTalk = 2

# WdMappedDataFields enumeration
wdAddress1 = 10
wdAddress2 = 11
wdAddress3 = 29
wdBusinessFax = 17
wdBusinessPhone = 16
wdCity = 12
wdCompany = 9
wdCountryRegion = 15
wdCourtesyTitle = 2
wdDepartment = 30
wdEmailAddress = 20
wdFirstName = 3
wdHomeFax = 19
wdHomePhone = 18
wdJobTitle = 8
wdLastName = 5
wdMiddleName = 4
wdNickname = 7
wdPostalCode = 14
wdRubyFirstName = 27
wdRubyLastName = 28
wdSpouseCourtesyTitle = 22
wdSpouseFirstName = 23
wdSpouseLastName = 25
wdSpouseMiddleName = 24
wdSpouseNickname = 26
wdState = 13
wdSuffix = 6
wdUniqueIdentifier = 1
wdWebPageURL = 21

# WdMeasurementUnits enumeration
wdCentimeters = 1
wdInches = 0
wdMillimeters = 2
wdPicas = 4
wdPoints = 3

# WdMergeFormatFrom enumeration
wdMergeFormatFromOriginal = 0
wdMergeFormatFromPrompt = 2
wdMergeFormatFromRevised = 1

# WdMergeSubType enumeration
wdMergeSubTypeAccess = 1
wdMergeSubTypeOAL = 2
wdMergeSubTypeOLEDBText = 5
wdMergeSubTypeOLEDBWord = 3
wdMergeSubTypeOther = 0
wdMergeSubTypeOutlook = 6
wdMergeSubTypeWord = 7
wdMergeSubTypeWord2000 = 8
wdMergeSubTypeWorks = 4

# WdMergeTarget enumeration
wdMergeTargetCurrent = 1
wdMergeTargetNew = 2
wdMergeTargetSelected = 0

# WdMonthNames enumeration
wdMonthNamesArabic = 0
wdMonthNamesEnglish = 1
wdMonthNamesFrench = 2

# WdMoveFromTextMark enumeration
wdMoveFromTextMarkBold = 6
wdMoveFromTextMarkCaret = 3
wdMoveFromTextMarkColorOnly = 10
wdMoveFromTextMarkDoubleStrikeThrough = 1
wdMoveFromTextMarkDoubleUnderline = 9
wdMoveFromTextMarkHidden = 0
wdMoveFromTextMarkItalic = 7
wdMoveFromTextMarkNone = 5
wdMoveFromTextMarkPound = 4
wdMoveFromTextMarkStrikeThrough = 2
wdMoveFromTextMarkUnderline = 8

# WdMovementType enumeration
wdExtend = 1
wdMove = 0

# WdMoveToTextMark enumeration
wdMoveToTextMarkBold = 1
wdMoveToTextMarkColorOnly = 5
wdMoveToTextMarkDoubleStrikeThrough = 7
wdMoveToTextMarkDoubleUnderline = 4
wdMoveToTextMarkItalic = 2
wdMoveToTextMarkNone = 0
wdMoveToTextMarkStrikeThrough = 6
wdMoveToTextMarkUnderline = 3

# WdMultipleWordConversionsMode enumeration
wdHangulToHanja = 0
wdHanjaToHangul = 1

# WdNewDocumentType enumeration
wdNewBlankDocument = 0
wdNewEmailMessage = 2
wdNewFrameset = 3
wdNewWebPage = 1
wdNewXMLDocument = 4

# WdNoteNumberStyle enumeration
wdNoteNumberStyleArabic = 0
wdNoteNumberStyleArabicFullWidth = 14
wdNoteNumberStyleArabicLetter1 = 46
wdNoteNumberStyleArabicLetter2 = 48
wdNoteNumberStyleHanjaRead = 41
wdNoteNumberStyleHanjaReadDigit = 42
wdNoteNumberStyleHebrewLetter1 = 45
wdNoteNumberStyleHebrewLetter2 = 47
wdNoteNumberStyleHindiArabic = 51
wdNoteNumberStyleHindiCardinalText = 52
wdNoteNumberStyleHindiLetter1 = 49
wdNoteNumberStyleHindiLetter2 = 50
wdNoteNumberStyleKanji = 10
wdNoteNumberStyleKanjiDigit = 11
wdNoteNumberStyleKanjiTraditional = 16
wdNoteNumberStyleLowercaseLetter = 4
wdNoteNumberStyleLowercaseRoman = 2
wdNoteNumberStyleNumberInCircle = 18
wdNoteNumberStyleSimpChinNum1 = 37
wdNoteNumberStyleSimpChinNum2 = 38
wdNoteNumberStyleSymbol = 9
wdNoteNumberStyleThaiArabic = 54
wdNoteNumberStyleThaiCardinalText = 55
wdNoteNumberStyleThaiLetter = 53
wdNoteNumberStyleTradChinNum1 = 33
wdNoteNumberStyleTradChinNum2 = 34
wdNoteNumberStyleUppercaseLetter = 3
wdNoteNumberStyleUppercaseRoman = 1
wdNoteNumberStyleVietCardinalText = 56

# WdNumberForm enumeration
wdNumberFormDefault = 0
wdNumberFormLining = 1
wdNumberFormOldstyle = 2

# WdNumberingRule enumeration
wdRestartContinuous = 0
wdRestartPage = 2
wdRestartSection = 1

# WdNumberSpacing enumeration
wdNumberSpacingDefault = 0
wdNumberSpacingProportional = 1
wdNumberSpacingTabular = 2

# WdNumberStyleWordBasicBiDi enumeration
wdCaptionNumberStyleBidiLetter1 = 49
wdCaptionNumberStyleBidiLetter2 = 50
wdListNumberStyleBidi1 = 49
wdListNumberStyleBidi2 = 50
wdNoteNumberStyleBidiLetter1 = 49
wdNoteNumberStyleBidiLetter2 = 50
wdPageNumberStyleBidiLetter1 = 49
wdPageNumberStyleBidiLetter2 = 50

# WdNumberType enumeration
wdNumberAllNumbers = 3
wdNumberListNum = 2
wdNumberParagraph = 1

# WdOLEPlacement enumeration
wdFloatOverText = 1
wdInLine = 0

# WdOLEType enumeration
wdOLEControl = 2
wdOLEEmbed = 1
wdOLELink = 0

# WdOLEVerb enumeration
wdOLEVerbDiscardUndoState = -6
wdOLEVerbHide = -3
wdOLEVerbInPlaceActivate = -5
wdOLEVerbOpen = -2
wdOLEVerbPrimary = 0
wdOLEVerbShow = -1
wdOLEVerbUIActivate = -4

# WdOMathBreakBin enumeration
wdOMathBreakBinAfter = 1
wdOMathBreakBinBefore = 0
wdOMathBreakBinRepeat = 2

# WdOMathBreakSub enumeration
wdOMathBreakSubMinusMinus = 0
wdOMathBreakSubMinusPlus = 2
wdOMathBreakSubPlusMinus = 1

# WdOMathFracType enumeration
wdOMathFracBar = 0
wdOMathFracLin = 3
wdOMathFracNoBar = 1
wdOMathFracSkw = 2

# WdOMathFunctionType enumeration
wdOMathFunctionAcc = 1
wdOMathFunctionBar = 2
wdOMathFunctionBorderBox = 4
wdOMathFunctionBox = 3
wdOMathFunctionDelim = 5
wdOMathFunctionEqArray = 6
wdOMathFunctionFrac = 7
wdOMathFunctionFunc = 8
wdOMathFunctionGroupChar = 9
wdOMathFunctionLimLow = 10
wdOMathFunctionLimUpp = 11
wdOMathFunctionMat = 12
wdOMathFunctionNary = 13
wdOMathFunctionNormalText = 21
wdOMathFunctionPhantom = 14
wdOMathFunctionRad = 16
wdOMathFunctionScrPre = 15
wdOMathFunctionScrSub = 17
wdOMathFunctionScrSubSup = 18
wdOMathFunctionScrSup = 19
wdOMathFunctionText = 20

# WdOMathHorizAlignType enumeration
wdOMathHorizAlignCenter = 0
wdOMathHorizAlignLeft = 1
wdOMathHorizAlignRight = 2

# WdOMathJc enumeration
wdOMathJcCenter = 2
wdOMathJcCenterGroup = 1
wdOMathJcInline = 7
wdOMathJcLeft = 3
wdOMathJcRight = 4

# WdOMathShapeType enumeration
wdOMathShapeCentered = 0
wdOMathShapeMatch = 1

# WdOMathSpacingRule enumeration
wdOMathSpacing1pt5 = 1
wdOMathSpacingDouble = 2
wdOMathSpacingExactly = 3
wdOMathSpacingMultiple = 4
wdOMathSpacingSingle = 0

# WdOMathType enumeration
wdOMathDisplay = 0
wdOMathInline = 1

# WdOMathVertAlignType enumeration
wdOMathVertAlignBottom = 2
wdOMathVertAlignCenter = 0
wdOMathVertAlignTop = 1

# WdOpenFormat enumeration
wdOpenFormatAllWord = 6
wdOpenFormatAuto = 0
wdOpenFormatDocument = 1
wdOpenFormatEncodedText = 5
wdOpenFormatRTF = 3
wdOpenFormatTemplate = 2
wdOpenFormatText = 4
wdOpenFormatOpenDocumentText = "18 (&H12)"
wdOpenFormatUnicodeText = 5
wdOpenFormatWebPages = 7
wdOpenFormatXML = 8
wdOpenFormatAllWordTemplates = 13
wdOpenFormatDocument97 = 1
wdOpenFormatTemplate97 = 2
wdOpenFormatXMLDocument = 9
wdOpenFormatXMLDocumentSerialized = 14
wdOpenFormatXMLDocumentMacroEnabled = 10
wdOpenFormatXMLDocumentMacroEnabledSerialized = 15
wdOpenFormatXMLTemplate = 11
wdOpenFormatXMLTemplateSerialized = "16 (&H10)"
wdOpenFormatXMLTemplateMacroEnabled = 12
wdOpenFormatXMLTemplateMacroEnabledSerialized = "17 (&H11)"

# WdOrganizerObject enumeration
wdOrganizerObjectAutoText = 1
wdOrganizerObjectCommandBars = 2
wdOrganizerObjectProjectItems = 3
wdOrganizerObjectStyles = 0

# WdOrientation enumeration
wdOrientLandscape = 1
wdOrientPortrait = 0

# WdOriginalFormat enumeration
wdOriginalDocumentFormat = 1
wdPromptUser = 2
wdWordDocument = 0

# WdOutlineLevel enumeration
wdOutlineLevel1 = 1
wdOutlineLevel2 = 2
wdOutlineLevel3 = 3
wdOutlineLevel4 = 4
wdOutlineLevel5 = 5
wdOutlineLevel6 = 6
wdOutlineLevel7 = 7
wdOutlineLevel8 = 8
wdOutlineLevel9 = 9
wdOutlineLevelBodyText = 10

# WdPageBorderArt enumeration
wdArtApples = 1
wdArtArchedScallops = 97
wdArtBabyPacifier = 70
wdArtBabyRattle = 71
wdArtBalloons3Colors = 11
wdArtBalloonsHotAir = 12
wdArtBasicBlackDashes = 155
wdArtBasicBlackDots = 156
wdArtBasicBlackSquares = 154
wdArtBasicThinLines = 151
wdArtBasicWhiteDashes = 152
wdArtBasicWhiteDots = 147
wdArtBasicWhiteSquares = 153
wdArtBasicWideInline = 150
wdArtBasicWideMidline = 148
wdArtBasicWideOutline = 149
wdArtBats = 37
wdArtBirds = 102
wdArtBirdsFlight = 35
wdArtCabins = 72
wdArtCakeSlice = 3
wdArtCandyCorn = 4
wdArtCelticKnotwork = 99
wdArtCertificateBanner = 158
wdArtChainLink = 128
wdArtChampagneBottle = 6
wdArtCheckedBarBlack = 145
wdArtCheckedBarColor = 61
wdArtCheckered = 144
wdArtChristmasTree = 8
wdArtCirclesLines = 91
wdArtCirclesRectangles = 140
wdArtClassicalWave = 56
wdArtClocks = 27
wdArtCompass = 54
wdArtConfetti = 31
wdArtConfettiGrays = 115
wdArtConfettiOutline = 116
wdArtConfettiStreamers = 14
wdArtConfettiWhite = 117
wdArtCornerTriangles = 141
wdArtCouponCutoutDashes = 163
wdArtCouponCutoutDots = 164
wdArtCrazyMaze = 100
wdArtCreaturesButterfly = 32
wdArtCreaturesFish = 34
wdArtCreaturesInsects = 142
wdArtCreaturesLadyBug = 33
wdArtCrossStitch = 138
wdArtCup = 67
wdArtDecoArch = 89
wdArtDecoArchColor = 50
wdArtDecoBlocks = 90
wdArtDiamondsGray = 88
wdArtDoubleD = 55
wdArtDoubleDiamonds = 127
wdArtEarth1 = 22
wdArtEarth2 = 21
wdArtEclipsingSquares1 = 101
wdArtEclipsingSquares2 = 86
wdArtEggsBlack = 66
wdArtFans = 51
wdArtFilm = 52
wdArtFirecrackers = 28
wdArtFlowersBlockPrint = 49
wdArtFlowersDaisies = 48
wdArtFlowersModern1 = 45
wdArtFlowersModern2 = 44
wdArtFlowersPansy = 43
wdArtFlowersRedRose = 39
wdArtFlowersRoses = 38
wdArtFlowersTeacup = 103
wdArtFlowersTiny = 42
wdArtGems = 139
wdArtGingerbreadMan = 69
wdArtGradient = 122
wdArtHandmade1 = 159
wdArtHandmade2 = 160
wdArtHeartBalloon = 16
wdArtHeartGray = 68
wdArtHearts = 15
wdArtHeebieJeebies = 120
wdArtHolly = 41
wdArtHouseFunky = 73
wdArtHypnotic = 87
wdArtIceCreamCones = 5
wdArtLightBulb = 121
wdArtLightning1 = 53
wdArtLightning2 = 119
wdArtMapleLeaf = 81
wdArtMapleMuffins = 2
wdArtMapPins = 30
wdArtMarquee = 146
wdArtMarqueeToothed = 131
wdArtMoons = 125
wdArtMosaic = 118
wdArtMusicNotes = 79
wdArtNorthwest = 104
wdArtOvals = 126
wdArtPackages = 26
wdArtPalmsBlack = 80
wdArtPalmsColor = 10
wdArtPaperClips = 82
wdArtPapyrus = 92
wdArtPartyFavor = 13
wdArtPartyGlass = 7
wdArtPencils = 25
wdArtPeople = 84
wdArtPeopleHats = 23
wdArtPeopleWaving = 85
wdArtPoinsettias = 40
wdArtPostageStamp = 135
wdArtPumpkin1 = 65
wdArtPushPinNote1 = 63
wdArtPushPinNote2 = 64
wdArtPyramids = 113
wdArtPyramidsAbove = 114
wdArtQuadrants = 60
wdArtRings = 29
wdArtSafari = 98
wdArtSawtooth = 133
wdArtSawtoothGray = 134
wdArtScaredCat = 36
wdArtSeattle = 78
wdArtShadowedSquares = 57
wdArtSharksTeeth = 132
wdArtShorebirdTracks = 83
wdArtSkyrocket = 77
wdArtSnowflakeFancy = 76
wdArtSnowflakes = 75
wdArtSombrero = 24
wdArtSouthwest = 105
wdArtStars = 19
wdArtStars3D = 17
wdArtStarsBlack = 74
wdArtStarsShadowed = 18
wdArtStarsTop = 157
wdArtSun = 20
wdArtSwirligig = 62
wdArtTornPaper = 161
wdArtTornPaperBlack = 162
wdArtTrees = 9
wdArtTriangleParty = 123
wdArtTriangles = 129
wdArtTribal1 = 130
wdArtTribal2 = 109
wdArtTribal3 = 108
wdArtTribal4 = 107
wdArtTribal5 = 110
wdArtTribal6 = 106
wdArtTwistedLines1 = 58
wdArtTwistedLines2 = 124
wdArtVine = 47
wdArtWaveline = 59
wdArtWeavingAngles = 96
wdArtWeavingBraid = 94
wdArtWeavingRibbon = 95
wdArtWeavingStrips = 136
wdArtWhiteFlowers = 46
wdArtWoodwork = 93
wdArtXIllusions = 111
wdArtZanyTriangles = 112
wdArtZigZag = 137
wdArtZigZagStitch = 143

# WdPageFit enumeration
wdPageFitBestFit = 2
wdPageFitFullPage = 1
wdPageFitNone = 0
wdPageFitTextFit = 3

# WdPageMovementType enumeration
wdVertical = 1
wdSideToSide = 2

# WdPageNumberAlignment enumeration
wdAlignPageNumberCenter = 1
wdAlignPageNumberInside = 3
wdAlignPageNumberLeft = 0
wdAlignPageNumberOutside = 4
wdAlignPageNumberRight = 2

# WdPageNumberStyle enumeration
wdPageNumberStyleArabic = 0
wdPageNumberStyleArabicFullWidth = 14
wdPageNumberStyleArabicLetter1 = 46
wdPageNumberStyleArabicLetter2 = 48
wdPageNumberStyleHanjaRead = 41
wdPageNumberStyleHanjaReadDigit = 42
wdPageNumberStyleHebrewLetter1 = 45
wdPageNumberStyleHebrewLetter2 = 47
wdPageNumberStyleHindiArabic = 51
wdPageNumberStyleHindiCardinalText = 52
wdPageNumberStyleHindiLetter1 = 49
wdPageNumberStyleHindiLetter2 = 50
wdPageNumberStyleKanji = 10
wdPageNumberStyleKanjiDigit = 11
wdPageNumberStyleKanjiTraditional = 16
wdPageNumberStyleLowercaseLetter = 4
wdPageNumberStyleLowercaseRoman = 2
wdPageNumberStyleNumberInCircle = 18
wdPageNumberStyleNumberInDash = 57
wdPageNumberStyleSimpChinNum1 = 37
wdPageNumberStyleSimpChinNum2 = 38
wdPageNumberStyleThaiArabic = 54
wdPageNumberStyleThaiCardinalText = 55
wdPageNumberStyleThaiLetter = 53
wdPageNumberStyleTradChinNum1 = 33
wdPageNumberStyleTradChinNum2 = 34
wdPageNumberStyleUppercaseLetter = 3
wdPageNumberStyleUppercaseRoman = 1
wdPageNumberStyleVietCardinalText = 56

# WdPaperSize enumeration
wdPaper10x14 = 0
wdPaper11x17 = 1
wdPaperA3 = 6
wdPaperA4 = 7
wdPaperA4Small = 8
wdPaperA5 = 9
wdPaperB4 = 10
wdPaperB5 = 11
wdPaperCSheet = 12
wdPaperCustom = 41
wdPaperDSheet = 13
wdPaperEnvelope10 = 25
wdPaperEnvelope11 = 26
wdPaperEnvelope12 = 27
wdPaperEnvelope14 = 28
wdPaperEnvelope9 = 24
wdPaperEnvelopeB4 = 29
wdPaperEnvelopeB5 = 30
wdPaperEnvelopeB6 = 31
wdPaperEnvelopeC3 = 32
wdPaperEnvelopeC4 = 33
wdPaperEnvelopeC5 = 34
wdPaperEnvelopeC6 = 35
wdPaperEnvelopeC65 = 36
wdPaperEnvelopeDL = 37
wdPaperEnvelopeItaly = 38
wdPaperEnvelopeMonarch = 39
wdPaperEnvelopePersonal = 40
wdPaperESheet = 14
wdPaperExecutive = 5
wdPaperFanfoldLegalGerman = 15
wdPaperFanfoldStdGerman = 16
wdPaperFanfoldUS = 17
wdPaperFolio = 18
wdPaperLedger = 19
wdPaperLegal = 4
wdPaperLetter = 2
wdPaperLetterSmall = 3
wdPaperNote = 20
wdPaperQuarto = 21
wdPaperStatement = 22
wdPaperTabloid = 23

# WdPaperTray enumeration
wdPrinterAutomaticSheetFeed = 7
wdPrinterDefaultBin = 0
wdPrinterEnvelopeFeed = 5
wdPrinterFormSource = 15
wdPrinterLargeCapacityBin = 11
wdPrinterLargeFormatBin = 10
wdPrinterLowerBin = 2
wdPrinterManualEnvelopeFeed = 6
wdPrinterManualFeed = 4
wdPrinterMiddleBin = 3
wdPrinterOnlyBin = 1
wdPrinterPaperCassette = 14
wdPrinterSmallFormatBin = 9
wdPrinterTractorFeed = 8
wdPrinterUpperBin = 1

# WdParagraphAlignment enumeration
wdAlignParagraphCenter = 1
wdAlignParagraphDistribute = 4
wdAlignParagraphJustify = 3
wdAlignParagraphJustifyHi = 7
wdAlignParagraphJustifyLow = 8
wdAlignParagraphJustifyMed = 5
wdAlignParagraphLeft = 0
wdAlignParagraphRight = 2
wdAlignParagraphThaiJustify = 9

# WdPartOfSpeech enumeration
wdAdjective = 0
wdAdverb = 2
wdConjunction = 5
wdIdiom = 8
wdInterjection = 7
wdNoun = 1
wdOther = 9
wdPreposition = 6
wdPronoun = 4
wdVerb = 3

# WdPasteDataType enumeration
wdPasteBitmap = 4
wdPasteDeviceIndependentBitmap = 5
wdPasteEnhancedMetafile = 9
wdPasteHTML = 10
wdPasteHyperlink = 7
wdPasteMetafilePicture = 3
wdPasteOLEObject = 0
wdPasteRTF = 1
wdPasteShape = 8
wdPasteText = 2

# WdPasteOptions enumeration
wdKeepSourceFormatting = 0
wdKeepTextOnly = 2
wdMatchDestinationFormatting = 1
wdUseDestinationStyles = 3

# WdPhoneticGuideAlignmentType enumeration
wdPhoneticGuideAlignmentCenter = 0
wdPhoneticGuideAlignmentLeft = 3
wdPhoneticGuideAlignmentOneTwoOne = 2
wdPhoneticGuideAlignmentRight = 4
wdPhoneticGuideAlignmentRightVertical = 5
wdPhoneticGuideAlignmentZeroOneZero = 1

# WdPictureLinkType enumeration
wdLinkDataInDoc = 1
wdLinkDataOnDisk = 2
wdLinkNone = 0

# WdPortugueseReform enumeration
wdPortugueseBoth = 3
wdPortuguesePostReform = 2
wdPortuguesePreReform = 1

# WdPreferredWidthType enumeration
wdPreferredWidthAuto = 1
wdPreferredWidthPercent = 2
wdPreferredWidthPoints = 3

# WdPrintOutItem enumeration
wdPrintAutoTextEntries = 4
wdPrintComments = 2
wdPrintDocumentContent = 0
wdPrintDocumentWithMarkup = 7
wdPrintEnvelope = 6
wdPrintKeyAssignments = 5
wdPrintMarkup = 2
wdPrintProperties = 1
wdPrintStyles = 3

# WdPrintOutPages enumeration
wdPrintAllPages = 0
wdPrintEvenPagesOnly = 2
wdPrintOddPagesOnly = 1

# WdPrintOutRange enumeration
wdPrintAllDocument = 0
wdPrintCurrentPage = 2
wdPrintFromTo = 3
wdPrintRangeOfPages = 4
wdPrintSelection = 1

# WdProofreadingErrorType enumeration
wdGrammaticalError = 1
wdSpellingError = 0

# WdProtectedViewCloseReason enumeration
wdProtectedViewCloseEdit = 1
wdProtectedViewCloseForced = 2
wdProtectedViewCloseNormal = 0

# WdProtectionType enumeration
wdAllowOnlyComments = 1
wdAllowOnlyFormFields = 2
wdAllowOnlyReading = 3
wdAllowOnlyRevisions = 0
wdNoProtection = -1

# WdReadingLayoutMargin enumeration
wdAutomaticMargin = 0
wdFullMargin = 2
wdSuppressMargin = 1

# WdReadingOrder enumeration
wdReadingOrderLtr = 1
wdReadingOrderRtl = 0

# WdRecoveryType enumeration
wdChart = 14
wdChartLinked = 15
wdChartPicture = 13
wdFormatOriginalFormatting = 16
wdFormatPlainText = 22
wdFormatSurroundingFormattingWithEmphasis = 20
wdListCombineWithExistingList = 24
wdListContinueNumbering = 7
wdListDontMerge = 25
wdListRestartNumbering = 8
wdPasteDefault = 0
wdSingleCellTable = 6
wdSingleCellText = 5
wdTableAppendTable = 10
wdTableInsertAsRows = 11
wdTableOriginalFormatting = 12
wdTableOverwriteCells = 23
wdUseDestinationStylesRecovery = 19

# WdRectangleType enumeration
wdLineBetweenColumnRectangle = 5
wdMarkupRectangle = 2
wdMarkupRectangleButton = 3
wdPageBorderRectangle = 4
wdSelection = 6
wdShapeRectangle = 1
wdSystem = 7
wdTextRectangle = 0
wdDocumentControlRectangle = 13
wdMailNavArea = 12
wdMarkupRectangleArea = 8
wdMarkupRectangleMoveMatch = 10
wdReadingModeNavigation = 9
wdReadingModePanningArea = 11

# WdReferenceKind enumeration
wdContentText = -1
wdEndnoteNumber = 6
wdEndnoteNumberFormatted = 17
wdEntireCaption = 2
wdFootnoteNumber = 5
wdFootnoteNumberFormatted = 16
wdNumberFullContext = -4
wdNumberNoContext = -3
wdNumberRelativeContext = -2
wdOnlyCaptionText = 4
wdOnlyLabelAndNumber = 3
wdPageNumber = 7
wdPosition = 15

# WdReferenceType enumeration
wdRefTypeBookmark = 2
wdRefTypeEndnote = 4
wdRefTypeFootnote = 3
wdRefTypeHeading = 1
wdRefTypeNumberedItem = 0

# WdRelativeHorizontalPosition enumeration
wdRelativeHorizontalPositionCharacter = 3
wdRelativeHorizontalPositionColumn = 2
wdRelativeHorizontalPositionMargin = 0
wdRelativeHorizontalPositionPage = 1
wdRelativeHorizontalPositionInnerMarginArea = 6
wdRelativeHorizontalPositionLeftMarginArea = 4
wdRelativeHorizontalPositionOuterMarginArea = 7
wdRelativeHorizontalPositionRightMarginArea = 5

# WdRelativeHorizontalSize enumeration
wdRelativeHorizontalSizeInnerMarginArea = 4
wdRelativeHorizontalSizeLeftMarginArea = 2
wdRelativeHorizontalSizeMargin = 0
wdRelativeHorizontalSizeOuterMarginArea = 5
wdRelativeHorizontalSizePage = 1
wdRelativeHorizontalSizeRightMarginArea = 3

# WdRelativeVerticalPosition enumeration
wdRelativeVerticalPositionLine = 3
wdRelativeVerticalPositionMargin = 0
wdRelativeVerticalPositionPage = 1
wdRelativeVerticalPositionParagraph = 2
wdRelativeVerticalPositionBottomMarginArea = 5
wdRelativeVerticalPositionInnerMarginArea = 6
wdRelativeVerticalPositionOuterMarginArea = 7
wdRelativeVerticalPositionTopMarginArea = 4

# WdRelativeVerticalSize enumeration
wdRelativeVerticalSizeBottomMarginArea = 3
wdRelativeVerticalSizeInnerMarginArea = 4
wdRelativeVerticalSizeMargin = 0
wdRelativeVerticalSizeOuterMarginArea = 5
wdRelativeVerticalSizePage = 1
wdRelativeVerticalSizeTopMarginArea = 2

# WdRelocate enumeration
wdRelocateDown = 1
wdRelocateUp = 0

# WdRemoveDocInfoType enumeration
wdRDIAll = 99
wdRDIComments = 1
wdRDIContentType = 16
wdRDIDocumentManagementPolicy = 15
wdRDIDocumentProperties = 8
wdRDIDocumentServerProperties = 14
wdRDIDocumentWorkspace = 10
wdRDIEmailHeader = 5
wdRDIInkAnnotations = 11
wdRDIRemovePersonalInformation = 4
wdRDIRevisions = 2
wdRDIRoutingSlip = 6
wdRDISendForReview = 7
wdRDITemplate = 9
wdRDITaskpaneWebExtensions = 17
wdRDIVersions = 3

# WdReplace enumeration
wdReplaceAll = 2
wdReplaceNone = 0
wdReplaceOne = 1

# WdRevisedLinesMark enumeration
wdRevisedLinesMarkLeftBorder = 1
wdRevisedLinesMarkNone = 0
wdRevisedLinesMarkOutsideBorder = 3
wdRevisedLinesMarkRightBorder = 2

# WdRevisedPropertiesMark enumeration
wdRevisedPropertiesMarkBold = 1
wdRevisedPropertiesMarkColorOnly = 5
wdRevisedPropertiesMarkDoubleStrikeThrough = 7
wdRevisedPropertiesMarkDoubleUnderline = 4
wdRevisedPropertiesMarkItalic = 2
wdRevisedPropertiesMarkNone = 0
wdRevisedPropertiesMarkStrikeThrough = 6
wdRevisedPropertiesMarkUnderline = 3

# WdRevisionsBalloonMargin enumeration
wdLeftMargin = 0
wdRightMargin = 1

# WdRevisionsBalloonPrintOrientation enumeration
wdBalloonPrintOrientationAuto = 0
wdBalloonPrintOrientationForceLandscape = 2
wdBalloonPrintOrientationPreserve = 1

# WdRevisionsBalloonWidthType enumeration
wdBalloonWidthPercent = 0
wdBalloonWidthPoints = 1

# WdRevisionsMode enumeration
wdBalloonRevisions = 0
wdInLineRevisions = 1
wdMixedRevisions = 2

# WdRevisionsView enumeration
wdRevisionsViewFinal = 0
wdRevisionsViewOriginal = 1

# WdRevisionsWrap enumeration
wdWrapAlways = 1
wdWrapAsk = 2
wdWrapNever = 0

# WdRevisionType enumeration
wdNoRevision = 0
wdRevisionCellDeletion = 17
wdRevisionCellInsertion = 16
wdRevisionCellMerge = 18
wdRevisionCellSplit = 19
wdRevisionConflict = 7
wdRevisionConflictDelete = 21
wdRevisionConflictInsert = 20
wdRevisionDelete = 2
wdRevisionDisplayField = 5
wdRevisionInsert = 1
wdRevisionMovedFrom = 14
wdRevisionMovedTo = 15
wdRevisionParagraphNumber = 4
wdRevisionParagraphProperty = 10
wdRevisionProperty = 3
wdRevisionReconcile = 6
wdRevisionReplace = 9
wdRevisionSectionProperty = 12
wdRevisionStyle = 8
wdRevisionStyleDefinition = 13
wdRevisionTableProperty = 11

# WdRowAlignment enumeration
wdAlignRowCenter = 1
wdAlignRowLeft = 0
wdAlignRowRight = 2

# WdRowHeightRule enumeration
wdRowHeightAtLeast = 1
wdRowHeightAuto = 0
wdRowHeightExactly = 2

# WdRulerStyle enumeration
wdAdjustFirstColumn = 2
wdAdjustNone = 0
wdAdjustProportional = 1
wdAdjustSameWidth = 3

# WdSalutationGender enumeration
wdGenderFemale = 0
wdGenderMale = 1
wdGenderNeutral = 2
wdGenderUnknown = 3

# WdSalutationType enumeration
wdSalutationBusiness = 2
wdSalutationFormal = 1
wdSalutationInformal = 0
wdSalutationOther = 3

# WdSaveFormat enumeration
wdFormatDocument = 0
wdFormatDOSText = 4
wdFormatDOSTextLineBreaks = 5
wdFormatEncodedText = 7
wdFormatFilteredHTML = 10
wdFormatFlatXML = 19
wdFormatFlatXMLMacroEnabled = 20
wdFormatFlatXMLTemplate = 21
wdFormatFlatXMLTemplateMacroEnabled = 22
wdFormatOpenDocumentText = 23
wdFormatHTML = 8
wdFormatRTF = 6
wdFormatStrictOpenXMLDocument = 24
wdFormatTemplate = 1
wdFormatText = 2
wdFormatTextLineBreaks = 3
wdFormatUnicodeText = 7
wdFormatWebArchive = 9
wdFormatXML = 11
wdFormatDocument97 = 0
wdFormatDocumentDefault = 16
wdFormatPDF = 17
wdFormatTemplate97 = 1
wdFormatXMLDocument = 12
wdFormatXMLDocumentMacroEnabled = 13
wdFormatXMLTemplate = 14
wdFormatXMLTemplateMacroEnabled = 15
wdFormatXPS = 18

# WdSaveOptions enumeration
wdDoNotSaveChanges = 0
wdPromptToSaveChanges = -2
wdSaveChanges = -1

# WdScrollbarType enumeration
wdScrollbarTypeAuto = 0
wdScrollbarTypeNo = 2
wdScrollbarTypeYes = 1

# WdSectionDirection enumeration
wdSectionDirectionLtr = 1
wdSectionDirectionRtl = 0

# WdSectionStart enumeration
wdSectionContinuous = 0
wdSectionEvenPage = 3
wdSectionNewColumn = 1
wdSectionNewPage = 2
wdSectionOddPage = 4

# WdSeekView enumeration
wdSeekCurrentPageFooter = 10
wdSeekCurrentPageHeader = 9
wdSeekEndnotes = 8
wdSeekEvenPagesFooter = 6
wdSeekEvenPagesHeader = 3
wdSeekFirstPageFooter = 5
wdSeekFirstPageHeader = 2
wdSeekFootnotes = 7
wdSeekMainDocument = 0
wdSeekPrimaryFooter = 4
wdSeekPrimaryHeader = 1

# WdSelectionFlags enumeration
wdSelActive = 8
wdSelAtEOL = 2
wdSelOvertype = 4
wdSelReplace = 16
wdSelStartActive = 1

# WdSelectionType enumeration
wdNoSelection = 0
wdSelectionBlock = 6
wdSelectionColumn = 4
wdSelectionFrame = 3
wdSelectionInlineShape = 7
wdSelectionIP = 1
wdSelectionNormal = 2
wdSelectionRow = 5
wdSelectionShape = 8

# WdSeparatorType enumeration
wdSeparatorColon = 2
wdSeparatorEmDash = 3
wdSeparatorEnDash = 4
wdSeparatorHyphen = 0
wdSeparatorPeriod = 1

# WdShapePosition enumeration
wdShapeBottom = -999997
wdShapeCenter = -999995
wdShapeInside = -999994
wdShapeLeft = -999998
wdShapeOutside = -999993
wdShapeRight = -999996
wdShapeTop = -999999

# WdShapePositionRelative enumeration
wdShapePositionRelativeNone = -999999

# WdShapeSizeRelative enumeration
wdShapeSizeRelativeNone = -999999

# WdShowFilter enumeration
wdShowFilterFormattingAvailable = 4
wdShowFilterFormattingInUse = 3
wdShowFilterStylesAll = 2
wdShowFilterStylesAvailable = 0
wdShowFilterStylesInUse = 1
wdShowFilterFormattingRecommended = 5

# WdShowSourceDocuments enumeration
wdShowSourceDocumentsBoth = 3
wdShowSourceDocumentsNone = 0
wdShowSourceDocumentsOriginal = 1
wdShowSourceDocumentsRevised = 2

# WdSmartTagControlType enumeration
wdControlActiveX = 13
wdControlButton = 6
wdControlCheckbox = 9
wdControlCombo = 12
wdControlDocumentFragment = 14
wdControlDocumentFragmentURL = 15
wdControlHelp = 3
wdControlHelpURL = 4
wdControlImage = 8
wdControlLabel = 7
wdControlLink = 2
wdControlListbox = 11
wdControlRadioGroup = 16
wdControlSeparator = 5
wdControlSmartTag = 1
wdControlTextbox = 10

# WdSortFieldType enumeration
wdSortFieldAlphanumeric = 0
wdSortFieldDate = 2
wdSortFieldJapanJIS = 4
wdSortFieldKoreaKS = 6
wdSortFieldNumeric = 1
wdSortFieldStroke = 5
wdSortFieldSyllable = 3

# WdSortOrder enumeration
wdSortOrderAscending = 0
wdSortOrderDescending = 1

# WdSortSeparator enumeration
wdSortSeparateByCommas = 1
wdSortSeparateByDefaultTableSeparator = 2
wdSortSeparateByTabs = 0

# WdSpanishSpeller enumeration
wdSpanishTuteoAndVoseo = 1
wdSpanishTuteoOnly = 0
wdSpanishVoseoOnly = 2

# WdSpecialPane enumeration
wdPaneComments = 15
wdPaneCurrentPageFooter = 17
wdPaneCurrentPageHeader = 16
wdPaneEndnoteContinuationNotice = 12
wdPaneEndnoteContinuationSeparator = 13
wdPaneEndnotes = 8
wdPaneEndnoteSeparator = 14
wdPaneEvenPagesFooter = 6
wdPaneEvenPagesHeader = 3
wdPaneFirstPageFooter = 5
wdPaneFirstPageHeader = 2
wdPaneFootnoteContinuationNotice = 9
wdPaneFootnoteContinuationSeparator = 10
wdPaneFootnotes = 7
wdPaneFootnoteSeparator = 11
wdPaneNone = 0
wdPanePrimaryFooter = 4
wdPanePrimaryHeader = 1
wdPaneRevisions = 18
wdPaneRevisionsHoriz = 19
wdPaneRevisionsVert = 20

# WdSpellingErrorType enumeration
wdSpellingCapitalization = 2
wdSpellingCorrect = 0
wdSpellingNotInDictionary = 1

# WdSpellingWordType enumeration
wdAnagram = 2
wdSpellword = 0
wdWildcard = 1

# WdStatistic enumeration
wdStatisticCharacters = 3
wdStatisticCharactersWithSpaces = 5
wdStatisticFarEastCharacters = 6
wdStatisticLines = 1
wdStatisticPages = 2
wdStatisticParagraphs = 4
wdStatisticWords = 0

# WdStoryType enumeration
wdCommentsStory = 4
wdEndnoteContinuationNoticeStory = 17
wdEndnoteContinuationSeparatorStory = 16
wdEndnoteSeparatorStory = 15
wdEndnotesStory = 3
wdEvenPagesFooterStory = 8
wdEvenPagesHeaderStory = 6
wdFirstPageFooterStory = 11
wdFirstPageHeaderStory = 10
wdFootnoteContinuationNoticeStory = 14
wdFootnoteContinuationSeparatorStory = 13
wdFootnoteSeparatorStory = 12
wdFootnotesStory = 2
wdMainTextStory = 1
wdPrimaryFooterStory = 9
wdPrimaryHeaderStory = 7
wdTextFrameStory = 5

# WdStyleSheetLinkType enumeration
wdStyleSheetLinkTypeImported = 1
wdStyleSheetLinkTypeLinked = 0

# WdStyleSheetPrecedence enumeration
wdStyleSheetPrecedenceHigher = -1
wdStyleSheetPrecedenceHighest = 1
wdStyleSheetPrecedenceLower = -2
wdStyleSheetPrecedenceLowest = 0

# WdStyleSort enumeration
wdStyleSortByBasedOn = 3
wdStyleSortByFont = 2
wdStyleSortByName = 0
wdStyleSortByType = 4
wdStyleSortRecommended = 1

# WdStyleType enumeration
wdStyleTypeCharacter = 2
wdStyleTypeList = 4
wdStyleTypeParagraph = 1
wdStyleTypeTable = 3

# WdStylisticSet enumeration
wdStylisticSet01 = 1
wdStylisticSet02 = 2
wdStylisticSet03 = 4
wdStylisticSet04 = 8
wdStylisticSet05 = 16
wdStylisticSet06 = 32
wdStylisticSet07 = 64
wdStylisticSet08 = 128
wdStylisticSet09 = 256
wdStylisticSet10 = 512
wdStylisticSet11 = 1024
wdStylisticSet12 = 2048
wdStylisticSet13 = 4096
wdStylisticSet14 = 8192
wdStylisticSet15 = 16384
wdStylisticSet16 = 32768
wdStylisticSet17 = 65536
wdStylisticSet18 = 131072
wdStylisticSet19 = 262144
wdStylisticSet20 = 524288
wdStylisticSetDefault = 0

# WdSubscriberFormats enumeration
wdSubscriberBestFormat = 0
wdSubscriberPict = 4
wdSubscriberRTF = 1
wdSubscriberText = 2

# WdTabAlignment enumeration
wdAlignTabBar = 4
wdAlignTabCenter = 1
wdAlignTabDecimal = 3
wdAlignTabLeft = 0
wdAlignTabList = 6
wdAlignTabRight = 2

# WdTabLeader enumeration
wdTabLeaderDashes = 2
wdTabLeaderDots = 1
wdTabLeaderHeavy = 4
wdTabLeaderLines = 3
wdTabLeaderMiddleDot = 5
wdTabLeaderSpaces = 0

# WdTableDirection enumeration
wdTableDirectionLtr = 1
wdTableDirectionRtl = 0

# WdTableFieldSeparator enumeration
wdSeparateByCommas = 2
wdSeparateByDefaultListSeparator = 3
wdSeparateByParagraphs = 0
wdSeparateByTabs = 1

# WdTableFormat enumeration
wdTableFormat3DEffects1 = 32
wdTableFormat3DEffects2 = 33
wdTableFormat3DEffects3 = 34
wdTableFormatClassic1 = 4
wdTableFormatClassic2 = 5
wdTableFormatClassic3 = 6
wdTableFormatClassic4 = 7
wdTableFormatColorful1 = 8
wdTableFormatColorful2 = 9
wdTableFormatColorful3 = 10
wdTableFormatColumns1 = 11
wdTableFormatColumns2 = 12
wdTableFormatColumns3 = 13
wdTableFormatColumns4 = 14
wdTableFormatColumns5 = 15
wdTableFormatContemporary = 35
wdTableFormatElegant = 36
wdTableFormatGrid1 = 16
wdTableFormatGrid2 = 17
wdTableFormatGrid3 = 18
wdTableFormatGrid4 = 19
wdTableFormatGrid5 = 20
wdTableFormatGrid6 = 21
wdTableFormatGrid7 = 22
wdTableFormatGrid8 = 23
wdTableFormatList1 = 24
wdTableFormatList2 = 25
wdTableFormatList3 = 26
wdTableFormatList4 = 27
wdTableFormatList5 = 28
wdTableFormatList6 = 29
wdTableFormatList7 = 30
wdTableFormatList8 = 31
wdTableFormatNone = 0
wdTableFormatProfessional = 37
wdTableFormatSimple1 = 1
wdTableFormatSimple2 = 2
wdTableFormatSimple3 = 3
wdTableFormatSubtle1 = 38
wdTableFormatSubtle2 = 39
wdTableFormatWeb1 = 40
wdTableFormatWeb2 = 41
wdTableFormatWeb3 = 42

# WdTableFormatApply enumeration
wdTableFormatApplyAutoFit = 16
wdTableFormatApplyBorders = 1
wdTableFormatApplyColor = 8
wdTableFormatApplyFirstColumn = 128
wdTableFormatApplyFont = 4
wdTableFormatApplyHeadingRows = 32
wdTableFormatApplyLastColumn = 256
wdTableFormatApplyLastRow = 64
wdTableFormatApplyShading = 2

# WdTablePosition enumeration
wdTableBottom = -999997
wdTableCenter = -999995
wdTableInside = -999994
wdTableLeft = -999998
wdTableOutside = -999993
wdTableRight = -999996
wdTableTop = -999999

# WdTaskPanes enumeration
wdTaskPaneApplyStyles = 17
wdTaskPaneDocumentActions = 7
wdTaskPaneDocumentProtection = 6
wdTaskPaneFaxService = 11
wdTaskPaneFormatting = 0
wdTaskPaneHelp = 9
wdTaskPaneMailMerge = 2
wdTaskPaneProofing = 20
wdTaskPaneResearch = 10
wdTaskPaneRevealFormatting = 1
wdTaskPaneRevPaneFlex = 22
wdTaskPaneSearch = 4
wdTaskPaneSignature = 14
wdTaskPaneStyleInspector = 15
wdTaskPaneThesaurus = 23
wdTaskPaneTranslate = 3
wdTaskPaneXMLDocument = 12
wdTaskPaneXMLMapping = 21
wdTaskPaneXMLStructure = 5

# WdTCSCConverterDirection enumeration
wdTCSCConverterDirectionAuto = 2
wdTCSCConverterDirectionSCTC = 0
wdTCSCConverterDirectionTCSC = 1

# WdTemplateType enumeration
wdAttachedTemplate = 2
wdGlobalTemplate = 1
wdNormalTemplate = 0

# WdTextboxTightWrap enumeration
wdTightAll = 1
wdTightFirstAndLastLines = 2
wdTightFirstLineOnly = 3
wdTightLastLineOnly = 4
wdTightNone = 0

# WdTextFormFieldType enumeration
wdCalculationText = 5
wdCurrentDateText = 3
wdCurrentTimeText = 4
wdDateText = 2
wdNumberText = 1
wdRegularText = 0

# WdTextOrientation enumeration
wdTextOrientationDownward = 3
wdTextOrientationHorizontal = 0
wdTextOrientationHorizontalRotatedFarEast = 4
wdTextOrientationUpward = 2
wdTextOrientationVerticalFarEast = 1
wdTextOrientationVertical = 5

# WdTextureIndex enumeration
wdTexture10Percent = 100
wdTexture12Pt5Percent = 125
wdTexture15Percent = 150
wdTexture17Pt5Percent = 175
wdTexture20Percent = 200
wdTexture22Pt5Percent = 225
wdTexture25Percent = 250
wdTexture27Pt5Percent = 275
wdTexture2Pt5Percent = 25
wdTexture30Percent = 300
wdTexture32Pt5Percent = 325
wdTexture35Percent = 350
wdTexture37Pt5Percent = 375
wdTexture40Percent = 400
wdTexture42Pt5Percent = 425
wdTexture45Percent = 450
wdTexture47Pt5Percent = 475
wdTexture50Percent = 500
wdTexture52Pt5Percent = 525
wdTexture55Percent = 550
wdTexture57Pt5Percent = 575
wdTexture5Percent = 50
wdTexture60Percent = 600
wdTexture62Pt5Percent = 625
wdTexture65Percent = 650
wdTexture67Pt5Percent = 675
wdTexture70Percent = 700
wdTexture72Pt5Percent = 725
wdTexture75Percent = 750
wdTexture77Pt5Percent = 775
wdTexture7Pt5Percent = 75
wdTexture80Percent = 800
wdTexture82Pt5Percent = 825
wdTexture85Percent = 850
wdTexture87Pt5Percent = 875
wdTexture90Percent = 900
wdTexture92Pt5Percent = 925
wdTexture95Percent = 950
wdTexture97Pt5Percent = 975
wdTextureCross = -11
wdTextureDarkCross = -5
wdTextureDarkDiagonalCross = -6
wdTextureDarkDiagonalDown = -3
wdTextureDarkDiagonalUp = -4
wdTextureDarkHorizontal = -1
wdTextureDarkVertical = -2
wdTextureDiagonalCross = -12
wdTextureDiagonalDown = -9
wdTextureDiagonalUp = -10
wdTextureHorizontal = -7
wdTextureNone = 0
wdTextureSolid = 1000
wdTextureVertical = -8

# WdThemeColorIndex enumeration
wdNotThemeColor = -1
wdThemeColorAccent1 = 4
wdThemeColorAccent2 = 5
wdThemeColorAccent3 = 6
wdThemeColorAccent4 = 7
wdThemeColorAccent5 = 8
wdThemeColorAccent6 = 9
wdThemeColorBackground1 = 12
wdThemeColorBackground2 = 14
wdThemeColorHyperlink = 10
wdThemeColorHyperlinkFollowed = 11
wdThemeColorMainDark1 = 0
wdThemeColorMainDark2 = 2
wdThemeColorMainLight1 = 1
wdThemeColorMainLight2 = 3
wdThemeColorText1 = 13
wdThemeColorText2 = 15

# WdToaFormat enumeration
wdTOAClassic = 1
wdTOADistinctive = 2
wdTOAFormal = 3
wdTOASimple = 4
wdTOATemplate = 0

# WdTocFormat enumeration
wdTOCClassic = 1
wdTOCDistinctive = 2
wdTOCFancy = 3
wdTOCFormal = 5
wdTOCModern = 4
wdTOCSimple = 6
wdTOCTemplate = 0

# WdTofFormat enumeration
wdTOFCentered = 3
wdTOFClassic = 1
wdTOFDistinctive = 2
wdTOFFormal = 4
wdTOFSimple = 5
wdTOFTemplate = 0

# WdTrailingCharacter enumeration
wdTrailingNone = 2
wdTrailingSpace = 1
wdTrailingTab = 0

# WdTwoLinesInOneType enumeration
wdTwoLinesInOneAngleBrackets = 4
wdTwoLinesInOneCurlyBrackets = 5
wdTwoLinesInOneNoBrackets = 1
wdTwoLinesInOneNone = 0
wdTwoLinesInOneParentheses = 2
wdTwoLinesInOneSquareBrackets = 3

# WdUnderline enumeration
wdUnderlineDash = 7
wdUnderlineDashHeavy = 23
wdUnderlineDashLong = 39
wdUnderlineDashLongHeavy = 55
wdUnderlineDotDash = 9
wdUnderlineDotDashHeavy = 25
wdUnderlineDotDotDash = 10
wdUnderlineDotDotDashHeavy = 26
wdUnderlineDotted = 4
wdUnderlineDottedHeavy = 20
wdUnderlineDouble = 3
wdUnderlineNone = 0
wdUnderlineSingle = 1
wdUnderlineThick = 6
wdUnderlineWavy = 11
wdUnderlineWavyDouble = 43
wdUnderlineWavyHeavy = 27
wdUnderlineWords = 2

# WdUnits enumeration
wdCell = 12
wdCharacter = 1
wdCharacterFormatting = 13
wdColumn = 9
wdItem = 16
wdLine = 5
wdParagraph = 4
wdParagraphFormatting = 14
wdRow = 10
wdScreen = 7
wdSection = 8
wdSentence = 3
wdStory = 6
wdTable = 15
wdWindow = 11
wdWord = 2

# WdUpdateStyleListBehavior enumeration
wdListBehaviorAddBulletsNumbering = 1
wdListBehaviorKeepPreviousPattern = 0

# WdUseFormattingFrom enumeration
wdFormattingFromCurrent = 0
wdFormattingFromPrompt = 2
wdFormattingFromSelected = 1

# WdVerticalAlignment enumeration
wdAlignVerticalBottom = 3
wdAlignVerticalCenter = 1
wdAlignVerticalJustify = 2
wdAlignVerticalTop = 0

# WdViewType enumeration
wdMasterView = 5
wdNormalView = 1
wdOutlineView = 2
wdPrintPreview = 4
wdPrintView = 3
wdReadingView = 7
wdWebView = 6

# WdVisualSelection enumeration
wdVisualSelectionBlock = 0
wdVisualSelectionContinuous = 1

# WdWindowState enumeration
wdWindowStateMaximize = 1
wdWindowStateMinimize = 2
wdWindowStateNormal = 0

# WdWindowType enumeration
wdWindowDocument = 0
wdWindowTemplate = 1

# WdWordDialog enumeration
wdDialogBuildingBlockOrganizer = 2067
wdDialogConnect = 420
wdDialogConsistencyChecker = 1121
wdDialogContentControlProperties = 2394
wdDialogControlRun = 235
wdDialogConvertObject = 392
wdDialogCopyFile = 300
wdDialogCreateAutoText = 872
wdDialogCreateSource = 1922
wdDialogCSSLinks = 1261
wdDialogDocumentInspector = 1482
wdDialogDocumentStatistics = 78
wdDialogDrawAlign = 634
wdDialogDrawSnapToGrid = 633
wdDialogEditAutoText = 985
wdDialogEditCreatePublisher = 732
wdDialogEditFind = 112
wdDialogEditFrame = 458
wdDialogEditGoTo = 896
wdDialogEditGoToOld = 811
wdDialogEditLinks = 124
wdDialogEditObject = 125
wdDialogEditPasteSpecial = 111
wdDialogEditPublishOptions = 735
wdDialogEditReplace = 117
wdDialogEditStyle = 120
wdDialogEditSubscribeOptions = 736
wdDialogEditSubscribeTo = 733
wdDialogEditTOACategory = 625
wdDialogEmailOptions = 863
wdDialogFileDocumentLayout = 178
wdDialogFileFind = 99
wdDialogFileMacCustomPageSetupGX = 737
wdDialogFileMacPageSetup = 685
wdDialogFileMacPageSetupGX = 444
wdDialogFileNew = 79
wdDialogFileOpen = 80
wdDialogFilePageSetup = 178
wdDialogFilePrint = 88
wdDialogFilePrintOneCopy = 445
wdDialogFilePrintSetup = 97
wdDialogFileRoutingSlip = 624
wdDialogFileSaveAs = 84
wdDialogFileSaveVersion = 1007
wdDialogFileSummaryInfo = 86
wdDialogFileVersions = 945
wdDialogFitText = 983
wdDialogFontSubstitution = 581
wdDialogFormatAddrFonts = 103
wdDialogFormatBordersAndShading = 189
wdDialogFormatBulletsAndNumbering = 824
wdDialogFormatCallout = 610
wdDialogFormatChangeCase = 322
wdDialogFormatColumns = 177
wdDialogFormatDefineStyleBorders = 185
wdDialogFormatDefineStyleFont = 181
wdDialogFormatDefineStyleFrame = 184
wdDialogFormatDefineStyleLang = 186
wdDialogFormatDefineStylePara = 182
wdDialogFormatDefineStyleTabs = 183
wdDialogFormatDrawingObject = 960
wdDialogFormatDropCap = 488
wdDialogFormatEncloseCharacters = 1162
wdDialogFormatFont = 174
wdDialogFormatFrame = 190
wdDialogFormatPageNumber = 298
wdDialogFormatParagraph = 175
wdDialogFormatPicture = 187
wdDialogFormatRetAddrFonts = 221
wdDialogFormatSectionLayout = 176
wdDialogFormatStyle = 180
wdDialogFormatStyleGallery = 505
wdDialogFormatStylesCustom = 1248
wdDialogFormatTabs = 179
wdDialogFormatTheme = 855
wdDialogFormattingRestrictions = 1427
wdDialogFormFieldHelp = 361
wdDialogFormFieldOptions = 353
wdDialogFrameSetProperties = 1074
wdDialogHelpAbout = 9
wdDialogHelpWordPerfectHelp = 10
wdDialogHelpWordPerfectHelpOptions = 511
wdDialogHorizontalInVertical = 1160
wdDialogIMESetDefault = 1094
wdDialogInsertAddCaption = 402
wdDialogInsertAutoCaption = 359
wdDialogInsertBookmark = 168
wdDialogInsertBreak = 159
wdDialogInsertCaption = 357
wdDialogInsertCaptionNumbering = 358
wdDialogInsertCrossReference = 367
wdDialogInsertDatabase = 341
wdDialogInsertDateTime = 165
wdDialogInsertField = 166
wdDialogInsertFile = 164
wdDialogInsertFootnote = 370
wdDialogInsertFormField = 483
wdDialogInsertHyperlink = 925
wdDialogInsertIndex = 170
wdDialogInsertIndexAndTables = 473
wdDialogInsertMergeField = 167
wdDialogInsertNumber = 812
wdDialogInsertObject = 172
wdDialogInsertPageNumbers = 294
wdDialogInsertPicture = 163
wdDialogInsertPlaceholder = 2348
wdDialogInsertSource = 2120
wdDialogInsertSubdocument = 583
wdDialogInsertSymbol = 162
wdDialogInsertTableOfAuthorities = 471
wdDialogInsertTableOfContents = 171
wdDialogInsertTableOfFigures = 472
wdDialogInsertWebComponent = 1324
wdDialogLabelOptions = 1367
wdDialogLetterWizard = 821
wdDialogListCommands = 723
wdDialogMailMerge = 676
wdDialogMailMergeCheck = 677
wdDialogMailMergeCreateDataSource = 642
wdDialogMailMergeCreateHeaderSource = 643
wdDialogMailMergeFieldMapping = 1304
wdDialogMailMergeFindRecipient = 1326
wdDialogMailMergeFindRecord = 569
wdDialogMailMergeHelper = 680
wdDialogMailMergeInsertAddressBlock = 1305
wdDialogMailMergeInsertAsk = 4047
wdDialogMailMergeInsertFields = 1307
wdDialogMailMergeInsertFillIn = 4048
wdDialogMailMergeInsertGreetingLine = 1306
wdDialogMailMergeInsertIf = 4049
wdDialogMailMergeInsertNextIf = 4053
wdDialogMailMergeInsertSet = 4054
wdDialogMailMergeInsertSkipIf = 4055
wdDialogMailMergeOpenDataSource = 81
wdDialogMailMergeOpenHeaderSource = 82
wdDialogMailMergeQueryOptions = 681
wdDialogMailMergeRecipients = 1308
wdDialogMailMergeSetDocumentType = 1339
wdDialogMailMergeUseAddressBook = 779
wdDialogMarkCitation = 463
wdDialogMarkIndexEntry = 169
wdDialogMarkTableOfContentsEntry = 442
wdDialogMyPermission = 1437
wdDialogNewToolbar = 586
wdDialogNoteOptions = 373
wdDialogOMathRecognizedFunctions = 2165
wdDialogOrganizer = 222
wdDialogPermission = 1469
wdDialogPhoneticGuide = 986
wdDialogReviewAfmtRevisions = 570
wdDialogSchemaLibrary = 1417
wdDialogSearch = 1363
wdDialogShowRepairs = 1381
wdDialogSourceManager = 1920
wdDialogStyleManagement = 1948
wdDialogTableAutoFormat = 563
wdDialogTableCellOptions = 1081
wdDialogTableColumnWidth = 143
wdDialogTableDeleteCells = 133
wdDialogTableFormatCell = 612
wdDialogTableFormula = 348
wdDialogTableInsertCells = 130
wdDialogTableInsertRow = 131
wdDialogTableInsertTable = 129
wdDialogTableOfCaptionsOptions = 551
wdDialogTableOfContentsOptions = 470
wdDialogTableProperties = 861
wdDialogTableRowHeight = 142
wdDialogTableSort = 199
wdDialogTableSplitCells = 137
wdDialogTableTableOptions = 1080
wdDialogTableToText = 128
wdDialogTableWrapping = 854
wdDialogTCSCTranslator = 1156
wdDialogTextToTable = 127
wdDialogToolsAcceptRejectChanges = 506
wdDialogToolsAdvancedSettings = 206
wdDialogToolsAutoCorrect = 378
wdDialogToolsAutoCorrectExceptions = 762
wdDialogToolsAutoManager = 915
wdDialogToolsAutoSummarize = 874
wdDialogToolsBulletsNumbers = 196
wdDialogToolsCompareDocuments = 198
wdDialogToolsCreateDirectory = 833
wdDialogToolsCreateEnvelope = 173
wdDialogToolsCreateLabels = 489
wdDialogToolsCustomize = 152
wdDialogToolsCustomizeKeyboard = 432
wdDialogToolsCustomizeMenuBar = 615
wdDialogToolsCustomizeMenus = 433
wdDialogToolsDictionary = 989
wdDialogToolsEnvelopesAndLabels = 607
wdDialogToolsGrammarSettings = 885
wdDialogToolsHangulHanjaConversion = 784
wdDialogToolsHighlightChanges = 197
wdDialogToolsHyphenation = 195
wdDialogToolsLanguage = 188
wdDialogToolsMacro = 215
wdDialogToolsMacroRecord = 214
wdDialogToolsManageFields = 631
wdDialogToolsMergeDocuments = 435
wdDialogToolsOptions = 974
wdDialogToolsOptionsAutoFormat = 959
wdDialogToolsOptionsAutoFormatAsYouType = 778
wdDialogToolsOptionsBidi = 1029
wdDialogToolsOptionsCompatibility = 525
wdDialogToolsOptionsEdit = 224
wdDialogToolsOptionsEditCopyPaste = 1356
wdDialogToolsOptionsFileLocations = 225
wdDialogToolsOptionsFuzzy = 790
wdDialogToolsOptionsGeneral = 203
wdDialogToolsOptionsPrint = 208
wdDialogToolsOptionsSave = 209
wdDialogToolsOptionsSecurity = 1361
wdDialogToolsOptionsSmartTag = 1395
wdDialogToolsOptionsSpellingAndGrammar = 211
wdDialogToolsOptionsTrackChanges = 386
wdDialogToolsOptionsTypography = 739
wdDialogToolsOptionsUserInfo = 213
wdDialogToolsOptionsView = 204
wdDialogToolsProtectDocument = 503
wdDialogToolsProtectSection = 578
wdDialogToolsRevisions = 197
wdDialogToolsSpellingAndGrammar = 828
wdDialogToolsTemplates = 87
wdDialogToolsThesaurus = 194
wdDialogToolsUnprotectDocument = 521
wdDialogToolsWordCount = 228
wdDialogTwoLinesInOne = 1161
wdDialogUpdateTOC = 331
wdDialogViewZoom = 577
wdDialogWebOptions = 898
wdDialogWindowActivate = 220
wdDialogXMLElementAttributes = 1460
wdDialogXMLOptions = 1425

# WdWordDialogTab enumeration
wdDialogEmailOptionsTabQuoting = 1900002
wdDialogEmailOptionsTabSignature = 1900000
wdDialogEmailOptionsTabStationary = 1900001
wdDialogFilePageSetupTabCharsLines = 150004
wdDialogFilePageSetupTabLayout = 150003
wdDialogFilePageSetupTabMargins = 150000
wdDialogFilePageSetupTabPaper = 150001
wdDialogFormatBordersAndShadingTabBorders = 700000
wdDialogFormatBordersAndShadingTabPageBorder = 700001
wdDialogFormatBordersAndShadingTabShading = 700002
wdDialogFormatBulletsAndNumberingTabBulleted = 1500000
wdDialogFormatBulletsAndNumberingTabNumbered = 1500001
wdDialogFormatBulletsAndNumberingTabOutlineNumbered = 1500002
wdDialogFormatDrawingObjectTabColorsAndLines = 1200000
wdDialogFormatDrawingObjectTabHR = 1200007
wdDialogFormatDrawingObjectTabPicture = 1200004
wdDialogFormatDrawingObjectTabPosition = 1200002
wdDialogFormatDrawingObjectTabSize = 1200001
wdDialogFormatDrawingObjectTabTextbox = 1200005
wdDialogFormatDrawingObjectTabWeb = 1200006
wdDialogFormatDrawingObjectTabWrapping = 1200003
wdDialogFormatFontTabAnimation = 600002
wdDialogFormatFontTabCharacterSpacing = 600001
wdDialogFormatFontTabFont = 600000
wdDialogFormatParagraphTabIndentsAndSpacing = 1000000
wdDialogFormatParagraphTabTeisai = 1000002
wdDialogFormatParagraphTabTextFlow = 1000001
wdDialogInsertIndexAndTablesTabIndex = 400000
wdDialogInsertIndexAndTablesTabTableOfAuthorities = 400003
wdDialogInsertIndexAndTablesTabTableOfContents = 400001
wdDialogInsertIndexAndTablesTabTableOfFigures = 400002
wdDialogInsertSymbolTabSpecialCharacters = 200001
wdDialogInsertSymbolTabSymbols = 200000
wdDialogLetterWizardTabLetterFormat = 1600000
wdDialogLetterWizardTabOtherElements = 1600002
wdDialogLetterWizardTabRecipientInfo = 1600001
wdDialogLetterWizardTabSenderInfo = 1600003
wdDialogNoteOptionsTabAllEndnotes = 300001
wdDialogNoteOptionsTabAllFootnotes = 300000
wdDialogOrganizerTabAutoText = 500001
wdDialogOrganizerTabCommandBars = 500002
wdDialogOrganizerTabMacros = 500003
wdDialogOrganizerTabStyles = 500000
wdDialogTablePropertiesTabCell = 1800003
wdDialogTablePropertiesTabColumn = 1800002
wdDialogTablePropertiesTabRow = 1800001
wdDialogTablePropertiesTabTable = 1800000
wdDialogTemplates = 2100000
wdDialogTemplatesLinkedCSS = 2100003
wdDialogTemplatesXMLExpansionPacks = 2100002
wdDialogTemplatesXMLSchema = 2100001
wdDialogToolsAutoCorrectExceptionsTabFirstLetter = 1400000
wdDialogToolsAutoCorrectExceptionsTabHangulAndAlphabet = 1400002
wdDialogToolsAutoCorrectExceptionsTabIac = 1400003
wdDialogToolsAutoCorrectExceptionsTabInitialCaps = 1400001
wdDialogToolsAutoManagerTabAutoCorrect = 1700000
wdDialogToolsAutoManagerTabAutoFormat = 1700003
wdDialogToolsAutoManagerTabAutoFormatAsYouType = 1700001
wdDialogToolsAutoManagerTabAutoText = 1700002
wdDialogToolsAutoManagerTabSmartTags = 1700004
wdDialogToolsEnvelopesAndLabelsTabEnvelopes = 800000
wdDialogToolsEnvelopesAndLabelsTabLabels = 800001
wdDialogToolsOptionsTabAcetate = 1266
wdDialogToolsOptionsTabBidi = 1029
wdDialogToolsOptionsTabCompatibility = 525
wdDialogToolsOptionsTabEdit = 224
wdDialogToolsOptionsTabFileLocations = 225
wdDialogToolsOptionsTabFuzzy = 790
wdDialogToolsOptionsTabGeneral = 203
wdDialogToolsOptionsTabHangulHanjaConversion = 786
wdDialogToolsOptionsTabPrint = 208
wdDialogToolsOptionsTabProofread = 211
wdDialogToolsOptionsTabSave = 209
wdDialogToolsOptionsTabSecurity = 1361
wdDialogToolsOptionsTabTrackChanges = 386
wdDialogToolsOptionsTabTypography = 739
wdDialogToolsOptionsTabUserInfo = 213
wdDialogToolsOptionsTabView = 204
wdDialogWebOptionsBrowsers = 2000000
wdDialogWebOptionsEncoding = 2000003
wdDialogWebOptionsFiles = 2000001
wdDialogWebOptionsFonts = 2000004
wdDialogWebOptionsGeneral = 2000000
wdDialogWebOptionsPictures = 2000002
wdDialogStyleManagementTabEdit = 2200000
wdDialogStyleManagementTabRecommend = 2200001
wdDialogStyleManagementTabRestrict = 2200002

# WdWrapSideType enumeration
wdWrapBoth = 0
wdWrapLargest = 3
wdWrapLeft = 1
wdWrapRight = 2

# WdWrapType enumeration
wdWrapInline = 7
wdWrapNone = 3
wdWrapSquare = 0
wdWrapThrough = 2
wdWrapTight = 1
wdWrapTopBottom = 4
wdWrapBehind = 5
wdWrapFront = 3

# WdWrapTypeMerged enumeration
wdWrapMergeBehind = 3
wdWrapMergeFront = 4
wdWrapMergeInline = 0
wdWrapMergeSquare = 1
wdWrapMergeThrough = 5
wdWrapMergeTight = 2
wdWrapMergeTopBottom = 6

class WebOptions:

    def __init__(self, weboptions=None):
        self.weboptions = weboptions

    @property
    def AllowPNG(self):
        return self.weboptions.AllowPNG

    @AllowPNG.setter
    def AllowPNG(self, value):
        self.weboptions.AllowPNG = value

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

    @OptimizeForBrowser.setter
    def OptimizeForBrowser(self, value):
        self.weboptions.OptimizeForBrowser = value

    @property
    def OrganizeInFolder(self):
        return self.weboptions.OrganizeInFolder

    @OrganizeInFolder.setter
    def OrganizeInFolder(self, value):
        self.weboptions.OrganizeInFolder = value

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

    @RelyOnCSS.setter
    def RelyOnCSS(self, value):
        self.weboptions.RelyOnCSS = value

    @property
    def RelyOnVML(self):
        return self.weboptions.RelyOnVML

    @RelyOnVML.setter
    def RelyOnVML(self, value):
        self.weboptions.RelyOnVML = value

    @property
    def ScreenSize(self):
        return self.weboptions.ScreenSize

    @ScreenSize.setter
    def ScreenSize(self, value):
        self.weboptions.ScreenSize = value

    @property
    def TargetBrowser(self):
        return self.weboptions.TargetBrowser

    @TargetBrowser.setter
    def TargetBrowser(self, value):
        self.weboptions.TargetBrowser = value

    @property
    def UseLongFileNames(self):
        return self.weboptions.UseLongFileNames

    @UseLongFileNames.setter
    def UseLongFileNames(self, value):
        self.weboptions.UseLongFileNames = value

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

    @DisplayHorizontalScrollBar.setter
    def DisplayHorizontalScrollBar(self, value):
        self.window.DisplayHorizontalScrollBar = value

    @property
    def DisplayLeftScrollBar(self):
        return self.window.DisplayLeftScrollBar

    @DisplayLeftScrollBar.setter
    def DisplayLeftScrollBar(self, value):
        self.window.DisplayLeftScrollBar = value

    @property
    def DisplayRightRuler(self):
        return self.window.DisplayRightRuler

    @DisplayRightRuler.setter
    def DisplayRightRuler(self, value):
        self.window.DisplayRightRuler = value

    @property
    def DisplayRulers(self):
        return self.window.DisplayRulers

    @DisplayRulers.setter
    def DisplayRulers(self, value):
        self.window.DisplayRulers = value

    @property
    def DisplayScreenTips(self):
        return self.window.DisplayScreenTips

    @DisplayScreenTips.setter
    def DisplayScreenTips(self, value):
        self.window.DisplayScreenTips = value

    @property
    def DisplayVerticalRuler(self):
        return self.window.DisplayVerticalRuler

    @DisplayVerticalRuler.setter
    def DisplayVerticalRuler(self, value):
        self.window.DisplayVerticalRuler = value

    @property
    def DisplayVerticalScrollBar(self):
        return self.window.DisplayVerticalScrollBar

    @DisplayVerticalScrollBar.setter
    def DisplayVerticalScrollBar(self, value):
        self.window.DisplayVerticalScrollBar = value

    @property
    def Document(self):
        return Document(self.window.Document)

    @property
    def DocumentMap(self):
        return self.window.DocumentMap

    @DocumentMap.setter
    def DocumentMap(self, value):
        self.window.DocumentMap = value

    @property
    def EnvelopeVisible(self):
        return self.window.EnvelopeVisible

    @EnvelopeVisible.setter
    def EnvelopeVisible(self, value):
        self.window.EnvelopeVisible = value

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

    @Split.setter
    def Split(self, value):
        self.window.Split = value

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

    @Visible.setter
    def Visible(self, value):
        self.window.Visible = value

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

    def Close(self, SaveChanges=None, RouteDocument=None):
        arguments = com_arguments([SaveChanges, RouteDocument])
        self.window.Close(*arguments)

    def GetPoint(self, ScreenPixelsLeft=None, ScreenPixelsTop=None, ScreenPixelsWidth=None, ScreenPixelsHeight=None, obj=None):
        arguments = com_arguments([ScreenPixelsLeft, ScreenPixelsTop, ScreenPixelsWidth, ScreenPixelsHeight, obj])
        self.window.GetPoint(*arguments)

    def LargeScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = com_arguments([Down, Up, ToRight, ToLeft])
        self.window.LargeScroll(*arguments)

    def NewWindow(self):
        return self.window.NewWindow()

    def PageScroll(self, Down=None, Up=None):
        arguments = com_arguments([Down, Up])
        self.window.PageScroll(*arguments)

    def PrintOut(self, Background=None, Append=None, Range=None, OutputFileName=None, From=None, To=None, Item=None, Copies=None, Pages=None, PageType=None, PrintToFile=None, Collate=None, FileName=None, ActivePrinterMacGX=None, ManualDuplexPrint=None, PrintZoomColumn=None, PrintZoomRow=None, PrintZoomPaperWidth=None, PrintZoomPaperHeight=None):
        arguments = com_arguments([Background, Append, Range, OutputFileName, From, To, Item, Copies, Pages, PageType, PrintToFile, Collate, FileName, ActivePrinterMacGX, ManualDuplexPrint, PrintZoomColumn, PrintZoomRow, PrintZoomPaperWidth, PrintZoomPaperHeight])
        self.window.PrintOut(*arguments)

    def RangeFromPoint(self, x=None, y=None):
        arguments = com_arguments([x, y])
        return self.window.RangeFromPoint(*arguments)

    def ScrollIntoView(self, Obj=None, Start=None):
        arguments = com_arguments([Obj, Start])
        self.window.ScrollIntoView(*arguments)

    def SetFocus(self):
        self.window.SetFocus()

    def SmallScroll(self, Down=None, Up=None, ToRight=None, ToLeft=None):
        arguments = com_arguments([Down, Up, ToRight, ToLeft])
        self.window.SmallScroll(*arguments)

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

    @Type.setter
    def Type(self, value):
        self.wrapformat.Type = value


# XlConstants enumeration
xl3DBar = -4099
xl3DSurface = -4103
xlAbove = 0
xlAutomatic = -4105
xlBar = 2
xlBelow = 1
xlBoth = 1
xlBottom = -4107
xlCenter = -4108
xlChecker = 9
xlCircle = 8
xlColumn = 3
xlCombination = -4111
xlCorner = 2
xlCrissCross = 16
xlCross = 4
xlCustom = -4114
xlDefaultAutoFormat = -1
xlDiamond = 2
xlDistributed = -4117
xlFill = 5
xlFixedValue = 1
xlGeneral = 1
xlGray16 = 17
xlGray25 = -4124
xlGray50 = -4125
xlGray75 = -4126
xlGray8 = 18
xlGrid = 15
xlHigh = -4127
xlInside = 2
xlJustify = -4130
xlLeft = -4131
xlLightDown = 13
xlLightHorizontal = 11
xlLightUp = 14
xlLightVertical = 12
xlLow = -4134
xlMaximum = 2
xlMinimum = 4
xlMinusValues = 3
xlNextToAxis = 4
xlNone = -4142
xlOpaque = 3
xlOutside = 3
xlPercent = 2
xlPlus = 9
xlPlusValues = 2
xlRight = -4152
xlScale = 3
xlSemiGray75 = 10
xlShowLabel = 4
xlShowLabelAndPercent = 5
xlShowPercent = 3
xlShowValue = 2
xlSingle = 2
xlSolid = 1
xlSquare = 1
xlStar = 5
xlStError = 4
xlTop = -4160
xlTransparent = 2
xlTriangle = 3

# XlReadingOrder enumeration
xlContext = -5002
xlLTR = -5003
xlRTL = -5004

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

    def SetMapping(self, XPath=None, PrefixMapping=None, Source=None):
        arguments = com_arguments([XPath, PrefixMapping, Source])
        return self.xmlmapping.SetMapping(*arguments)

    def SetMappingByNode(self, Node=None):
        arguments = com_arguments([Node])
        return self.xmlmapping.SetMappingByNode(*arguments)


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

    def AttachToDocument(self, Document=None):
        arguments = com_arguments([Document])
        self.xmlnamespace.AttachToDocument(*arguments)

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

    def ValidationErrorText(self, Advanced=None):
        arguments = com_arguments([Advanced])
        if callable(self.xmlnode.ValidationErrorText):
            return self.xmlnode.ValidationErrorText(*arguments)
        else:
            return self.xmlnode.GetValidationErrorText(*arguments)

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

    def RemoveChild(self, ChildElement=None):
        arguments = com_arguments([ChildElement])
        return self.xmlnode.RemoveChild(*arguments)

    def SelectNodes(self, XPath=None, PrefixMapping=None, FastSearchSkippingTextNodes=None):
        arguments = com_arguments([XPath, PrefixMapping, FastSearchSkippingTextNodes])
        return self.xmlnode.SelectNodes(*arguments)

    def SelectSingleNode(self, XPath=None, PrefixMapping=None, FastSearchSkippingTextNodes=None):
        arguments = com_arguments([XPath, PrefixMapping, FastSearchSkippingTextNodes])
        return self.xmlnode.SelectSingleNode(*arguments)

    def SetValidationError(self, Status=None, ErrorText=None, ClearedAutomatically=None):
        arguments = com_arguments([Status, ErrorText, ClearedAutomatically])
        self.xmlnode.SetValidationError(*arguments)

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

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.xmlnodes.Item(*arguments)


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

    @HideValidationErrors.setter
    def HideValidationErrors(self, value):
        self.xmlschemareferences.HideValidationErrors = value

    @property
    def IgnoreMixedContent(self):
        return self.xmlschemareferences.IgnoreMixedContent

    @IgnoreMixedContent.setter
    def IgnoreMixedContent(self, value):
        self.xmlschemareferences.IgnoreMixedContent = value

    @property
    def Parent(self):
        return self.xmlschemareferences.Parent

    @property
    def ShowPlaceholderText(self):
        return self.xmlschemareferences.ShowPlaceholderText

    @ShowPlaceholderText.setter
    def ShowPlaceholderText(self, value):
        self.xmlschemareferences.ShowPlaceholderText = value

    def Add(self, NamespaceURI=None, Alias=None, FileName=None, InstallForAllUsers=None):
        arguments = com_arguments([NamespaceURI, Alias, FileName, InstallForAllUsers])
        return XMLSchemaReference(self.xmlschemareferences.Add(*arguments))

    def Item(self, Index=None):
        arguments = com_arguments([Index])
        return self.xmlschemareferences.Item(*arguments)

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

